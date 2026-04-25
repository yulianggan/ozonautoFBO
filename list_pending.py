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
                   help="台账目录 (含 shipment_ledger_*.json)")
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

    pending = load_pending(ld)

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
