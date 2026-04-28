#!/usr/bin/env python3
"""查询未发货清单 + 生成重试 plan.json.

用途:
  - 看 ledger/ 里跨多次 run 还没发的 (cluster, SKU, qty) 列表
  - 按 cluster 生成新的 plan.json, 方便重试当初失败的批次
  - 或"改派"某集群的未发到另一个集群

典型用法:
    python3 list_pending.py                     # 打印所有 pending
    python3 list_pending.py --by-cluster        # 按集群分组
    python3 list_pending.py --by-sku            # 按 SKU 分组
    python3 list_pending.py --reallocate 新西伯利亚=莫斯科,远东=圣彼得堡  # 改派
    python3 list_pending.py --emit-plan retry_plan.json  # 直接出新 plan.json
    python3 list_pending.py --filter-reason MATRIX_DROPPED_FROM_SUPPLY  # 只看矩阵过滤的
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
LEDGER_DIR = SCRIPT_DIR / "ledger"
RS = "http://127.0.0.1:4181"
FBO = "http://127.0.0.1:4182"


def load_pending_from_approved(account: str, since: str, to: str) -> list[dict]:
    """切源版: 缺货 = (4181 审批 cap) - (Ozon 已发 active).

    返回跟 load_pending 一样的 dict 列表 (cluster, offer_id, sku, quantity, reason_code).
    reason_code 用 "PENDING_VS_APPROVED" 表示这是审批 gap 不是 ledger skipped.
    """
    import urllib.parse, urllib.request
    enc = urllib.parse.quote(account, safe="")

    # 1) 拉 approved/dispatched plans
    with urllib.request.urlopen(f"{RS}/plans?account={enc}&limit=200", timeout=30) as f:
        lst = json.load(f)
    by_sku: dict[str, dict] = {}
    for meta in lst:
        if meta.get("status") not in ("approved", "dispatched"):
            continue
        prev = by_sku.get(meta["sku"])
        if prev is None or (meta.get("created_at", "") > prev.get("created_at", "")):
            by_sku[meta["sku"]] = meta
    caps: dict[str, dict[str, tuple[int, int]]] = {}  # offer_id → {cluster: (pieces, sku)}
    for offer_id, meta in by_sku.items():
        with urllib.request.urlopen(f"{RS}/plans/{meta['plan_id']}", timeout=30) as f:
            p = json.load(f)
        cluster_caps: dict[str, tuple[int, int]] = {}
        sku = 0
        for a in (p.get("allocations") or []):
            cl = a.get("cluster")
            if cl:
                cluster_caps[cl] = (int(a.get("pieces") or 0), sku)
        if cluster_caps:
            caps[offer_id] = cluster_caps

    # 2) 拉 Ozon supplies + bundles, 算 active 已发
    def _post(path: str, body: dict, retries: int = 3) -> dict:
        import time as _time
        last = None
        for i in range(retries):
            try:
                req = urllib.request.Request(f"{FBO}{path}",
                    data=json.dumps(body).encode(),
                    headers={"Content-Type": "application/json", "X-Ozon-Account": enc})
                return json.load(urllib.request.urlopen(req, timeout=60))
            except Exception as e:
                last = e
                _time.sleep(2 + i * 2)
        # 全 retry 失败, 抛出 (而不是 silent return {} 漏算 active)
        raise RuntimeError(f"POST {path} failed after {retries} retries: {last}")

    # wh→cluster (复用 build_shipment_overview 的逻辑, 简化版)
    cl_resp = _post("/cluster/list", {"cluster_ids": [], "cluster_type": "CLUSTER_TYPE_OZON"})
    wh_map: dict[str, str] = {}
    for c in cl_resp.get("clusters", []):
        cn = c.get("name", "")
        for lc in c.get("logistic_clusters", []):
            for wh in lc.get("warehouses", []):
                wn = wh.get("name", "")
                if wn and cn:
                    wh_map[wn] = cn
    stem_buckets: dict[str, set[str]] = {}
    for wh, cn in wh_map.items():
        stem_buckets.setdefault(wh.split("_")[0], set()).add(cn)
    stem_to_cluster = {s: list(cs)[0] for s, cs in stem_buckets.items() if len(cs) == 1}

    def _wh2cluster(wh: str) -> str:
        if wh in wh_map:
            return wh_map[wh]
        for known, cn in wh_map.items():
            if wh.startswith(known) or known.startswith(wh):
                return cn
        return stem_to_cluster.get(wh.split("_")[0], wh)

    # 列 supplies
    ids: list[int] = []
    last = ""
    for _ in range(20):
        body = {
            "filter": {
                "states": ["READY_TO_SUPPLY", "DATA_FILLING", "SUPPLIED", "SUPPLIED_PARTIALLY",
                           "SUPPLY_PROCESSING", "SUPPLY_REJECTED", "DRAFT"],
                "since": since, "to": to,
            },
            "limit": 50, "sort_by": 1,
        }
        if last: body["last_id"] = last
        r = _post("/supply-order/list", body)
        page = r.get("order_ids") or []
        ids.extend(page)
        last = r.get("last_id") or ""
        if not last or len(page) < 50:
            break
    active: dict[tuple[str, str], int] = defaultdict(int)
    if ids:
        for chunk in [ids[i:i+30] for i in range(0, len(ids), 30)]:
            r = _post("/supply-order/get", {"order_ids": chunk})
            for o in r.get("orders", []):
                for sup in (o.get("supplies") or []):
                    if sup.get("state") == "CANCELLED" or o.get("state") == "CANCELLED":
                        continue
                    wh = (sup.get("storage_warehouse") or {}).get("name", "")
                    cl = _wh2cluster(wh)
                    bid = sup.get("bundle_id")
                    if not bid:
                        continue
                    b = _post("/supply-order/bundle", {"bundle_ids": [bid], "limit": 100})
                    for it in (b.get("items") or []):
                        oid = it.get("offer_id")
                        if oid:
                            active[(cl, oid)] += int(it.get("quantity") or 0)

    # 3) 计算 cap - active = pending. 改派支持: 先按 SKU 总量, 总发齐就跳整 SKU
    #    (eg. 远东 Q_ChongQiZui 600 改派到 Москва, plan 仍 cap=600 但 active 在 Москва)
    BOX_SIZE = 300  # default; 注: 跟 plan box_size 一致 (4181 plan 里有 box_size 字段, 这里简化)
    out: list[dict] = []
    for offer_id, cluster_caps in caps.items():
        cap_total = sum(c[0] for c in cluster_caps.values())
        # active_total: 含 4181 plan 没规划的 cluster (如改派目标)
        active_total = sum(qty for (c, oid), qty in active.items() if oid == offer_id)
        if active_total >= cap_total:
            continue  # 整 offer 发齐 (含改派抵扣), 跳过所有 cluster

        remaining = cap_total - active_total
        for cluster, (cap_pieces, _sku) in cluster_caps.items():
            if remaining <= 0:
                break
            if cap_pieces <= 0:
                continue
            shipped = active.get((cluster, offer_id), 0)
            cluster_gap = cap_pieces - shipped
            if cluster_gap <= 0:
                continue
            effective_raw = min(cluster_gap, remaining)
            # 整箱规则: 残箱 < box_size/2 砍掉
            full_boxes = effective_raw // BOX_SIZE
            rem_pcs = effective_raw % BOX_SIZE
            if rem_pcs > 0 and rem_pcs < BOX_SIZE // 2:
                effective = full_boxes * BOX_SIZE
                note = f" (残 {rem_pcs} < {BOX_SIZE//2} floor 至 {effective})"
            else:
                effective = effective_raw
                note = ""
            if effective <= 0:
                continue
            out.append({
                "cluster": cluster,
                "offer_id": offer_id,
                "sku": "",
                "quantity": effective,
                "reason_code": "PENDING_VS_APPROVED",
                "reason": (f"审批总 {cap_total} - 全 SKU 已发 {active_total} = 缺 {remaining}; "
                           f"此集群 cap {cap_pieces} - 此集群已发 {shipped}{note}"),
                "_source_file": "4181 plans + Ozon API",
            })
            remaining -= effective
    return out


def load_pending(ledger_dir: Path) -> list[dict]:
    """扫 ledger/*.json, 汇总未发清单; 按 (cluster, offer_id, sku) 去重, 保留 timestamp 最新.

    同 `create_fbo_plan.merge_pending` 的去重规则.
    """
    shipped_keys: set = set()
    raw: list[dict] = []
    for f in sorted(ledger_dir.glob("shipment_ledger_*.json")):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        run_at = data.get("run_at","")
        for s in (data.get("shipped") or []):
            cluster_names = s.get("cluster_names") or ([s.get("cluster")] if s.get("cluster") else [])
            for it in (s.get("items") or []):
                for c in cluster_names:
                    shipped_keys.add((c, it.get("offer_id")))
        for s in (data.get("skipped") or []):
            s2 = dict(s)
            s2["_source_file"] = f.name
            s2["_source_run_at"] = run_at
            raw.append(s2)

    def _ts(e: dict) -> str:
        return e.get("timestamp") or e.get("_source_run_at") or ""

    by_key: dict[tuple, dict] = {}
    for p in raw:
        k = (p.get("cluster"), p.get("offer_id"), p.get("sku"))
        if k not in by_key or _ts(p) > _ts(by_key[k]):
            by_key[k] = p

    return [p for p in by_key.values()
            if (p.get("cluster"), p.get("offer_id")) not in shipped_keys]


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="未发货清单查询 + 重试 plan 生成",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--ledger-dir", default=str(LEDGER_DIR),
                   help="台账目录 (含 shipment_ledger_*.json) — 仅 --from-ledger 模式用")
    p.add_argument("--from-ledger", action="store_true",
                   help="老模式: 从 ledger/shipment_ledger_*.json 累加 skipped (数据可能过时); "
                        "默认走 --from-approved (4181 审批 plans + Ozon active)")
    p.add_argument("--account", default="丝绸生活",
                   help="拉 4181 plans 的账号")
    p.add_argument("--since", default="2026-04-22T00:00:00Z",
                   help="supply-order/list 起始")
    p.add_argument("--to", default="2026-04-30T00:00:00Z",
                   help="supply-order/list 终止")
    g = p.add_mutually_exclusive_group()
    g.add_argument("--by-cluster", action="store_true",
                   help="按集群分组打印")
    g.add_argument("--by-sku", action="store_true",
                   help="按 SKU/货号 分组打印")
    g.add_argument("--by-reason", action="store_true",
                   help="按 reason_code 分组打印")
    p.add_argument("--filter-reason", default="",
                   help="只保留指定 reason_code 的条目 (MATRIX_DROPPED_FROM_SUPPLY / NO_AVAILABLE_WAREHOUSE / ...)")
    p.add_argument("--filter-cluster", default="",
                   help="只保留指定集群关键字 (substring)")
    p.add_argument("--reallocate", default="",
                   help='"源集群=新集群,..." 把未发集群改派到其他集群')
    p.add_argument("--emit-plan", default="",
                   help="生成新 plan.json 路径 (用 --reallocate 后的集群映射)")
    p.add_argument("--drop-off-keyword", default="ЩЕРБИНКА",
                   help="emit-plan 的 drop_off_keyword")
    p.add_argument("--source-warehouse-keyword", default="ЖУКОВСКИЙ_РФЦ",
                   help="emit-plan 的 source_warehouse_keyword")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    ld = Path(args.ledger_dir)
    if not ld.exists():
        print(f"台账目录 {ld} 不存在"); return 2

    if args.from_ledger:
        print(f"[源] ledger/shipment_ledger_*.json (老模式; 数据可能过时)")
        pending = load_pending(ld)
    else:
        print(f"[源] 4181 审批 plans + Ozon API ({args.since} ~ {args.to})")
        pending = load_pending_from_approved(args.account, args.since, args.to)

    # 过滤
    if args.filter_reason:
        pending = [p for p in pending if p.get("reason_code") == args.filter_reason]
    if args.filter_cluster:
        pending = [p for p in pending if args.filter_cluster in (p.get("cluster") or "")]

    print(f"=== pending {len(pending)} 条 (过滤后) ===\n")

    # 改派
    realloc: dict[str, str] = {}
    for seg in (args.reallocate or "").split(","):
        seg = seg.strip()
        if "=" in seg:
            k, v = seg.split("=", 1)
            realloc[k.strip()] = v.strip()
    if realloc:
        for p in pending:
            orig = p.get("cluster")
            if orig in realloc:
                p["_original_cluster"] = orig
                p["cluster"] = realloc[orig]
        print(f"改派映射: {realloc}\n")

    # 分组打印
    if args.by_cluster:
        by = defaultdict(list)
        for p in pending: by[p.get("cluster","")].append(p)
        for cl in sorted(by):
            print(f"[{cl}] ({len(by[cl])} 条)")
            for p in by[cl]:
                print(f"  {p.get('offer_id')} ({p.get('sku')}) × {p.get('quantity')} — {p.get('reason_code','')}")
    elif args.by_sku:
        by = defaultdict(list)
        for p in pending: by[p.get("offer_id","")].append(p)
        for oid in sorted(by):
            total = sum(p.get("quantity",0) for p in by[oid])
            print(f"[{oid}] 合计 {total} 件, {len(by[oid])} 条")
            for p in by[oid]:
                print(f"  → {p.get('cluster')} × {p.get('quantity')} ({p.get('reason_code','')})")
    elif args.by_reason:
        by = defaultdict(list)
        for p in pending: by[p.get("reason_code","")].append(p)
        for rc in sorted(by):
            print(f"[{rc}] ({len(by[rc])} 条)")
            for p in by[rc]:
                print(f"  {p.get('cluster')} / {p.get('offer_id')} × {p.get('quantity')}")
    else:
        for p in pending:
            print(f"  {p.get('cluster','?'):<20} {p.get('offer_id','?'):<30} "
                  f"× {p.get('quantity',0):>5}  {p.get('reason_code','')}  "
                  f"({p.get('_source_file','')})")

    # 出新 plan.json
    if args.emit_plan:
        matrix: dict[str, dict[str, int]] = defaultdict(dict)
        for p in pending:
            cl = p.get("cluster")
            oid = p.get("offer_id")
            q = p.get("quantity", 0)
            if cl and oid and q > 0:
                matrix[cl][oid] = matrix[cl].get(oid, 0) + q
        plan = {
            "source_warehouse_keyword": args.source_warehouse_keyword,
            "drop_off_keyword": args.drop_off_keyword,
            "matrix": dict(matrix),
        }
        out = Path(args.emit_plan)
        out.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n[emit-plan] 已写: {out}")
        print(f"  用: python3 create_fbo_plan.py --config {out} --box-size <N> --account <账号>")

    return 0


if __name__ == "__main__":
    sys.exit(main())
