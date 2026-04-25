"""计划件数 vs 已发件数 对照表 (2026-04-25 ad-hoc).

数据源:
- 计划: 扫 SCRIPT_DIR/allocation_*.xlsx 取 (cluster, offer_id) → 合计件 (capped)
- 已发: Ozon /supply-order/list + /get + /bundle, 按 (cluster, offer_id) 聚合
       active = 非 CANCELLED 的 supply 累加;  cancelled 单独累加
- 输出: shipment_overview_<account>_<YYYYMMDD>.xlsx, 1 sheet 22 列

用法: python3 build_shipment_overview.py [--since 2026-04-22] [--to 2026-04-26]
"""
from __future__ import annotations

import argparse
import json
import time
import urllib.request
import urllib.parse
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

SCRIPT_DIR = Path(__file__).resolve().parent
SVC = "http://127.0.0.1:4182"
DEFAULT_ACCOUNT = "丝绸生活"


# ------- HTTP helpers -------

def _post(account: str, path: str, body: dict, retries: int = 3, timeout: int = 60) -> dict:
    acc = urllib.parse.quote(account)
    headers = {"Content-Type": "application/json", "X-Ozon-Account": acc}
    req = urllib.request.Request(f"{SVC}{path}", data=json.dumps(body).encode(), headers=headers)
    last_err = None
    for r in range(retries):
        try:
            return json.load(urllib.request.urlopen(req, timeout=timeout))
        except Exception as e:
            last_err = e
            time.sleep(2)
    raise RuntimeError(f"POST {path} failed after {retries}: {last_err}")


def list_supply_orders(account: str, since: str, to: str) -> list[int]:
    ids: list[int] = []
    last_id = ""
    for _ in range(20):
        body = {
            "filter": {
                "states": [
                    "READY_TO_SUPPLY", "DRAFT", "SUPPLIED", "SUPPLIED_PARTIALLY",
                    "SUPPLY_PROCESSING", "SUPPLY_REJECTED", "CANCELLED",
                ],
                "since": since,
                "to": to,
            },
            "limit": 50,
            "sort_by": 1,
        }
        if last_id:
            body["last_id"] = last_id
        r = _post(account, "/supply-order/list", body)
        page = r.get("order_ids") or []
        ids.extend(page)
        last_id = r.get("last_id") or ""
        if not last_id or len(page) < 50:
            break
    return ids


def get_orders(account: str, ids: list[int]) -> list[dict]:
    out = []
    for chunk_start in range(0, len(ids), 30):
        chunk = ids[chunk_start:chunk_start + 30]
        r = _post(account, "/supply-order/get", {"order_ids": chunk})
        out.extend(r.get("orders", []))
    return out


def get_bundle_items(account: str, bundle_id: str) -> list[dict]:
    r = _post(account, "/supply-order/bundle", {"bundle_ids": [bundle_id], "limit": 100})
    return r.get("items") or r.get("contents") or []


# ------- allocation_xlsx caps -------

def read_allocation_caps(script_dir: Path) -> dict[str, dict[str, int]]:
    caps: dict[str, dict[str, int]] = {}
    for p in sorted(script_dir.glob("allocation_*.xlsx")):
        try:
            wb = load_workbook(p, data_only=True)
            ws = wb.active
            offer_id = ws["B1"].value
            if not offer_id or not isinstance(offer_id, str):
                continue
            hdr = [c.value for c in ws[17]]
            if "集群" not in hdr:
                continue
            cluster_col = hdr.index("集群")
            total_col = next(
                (i for i, h in enumerate(hdr) if isinstance(h, str) and "合计件" in h),
                None,
            )
            if total_col is None:
                continue
            cluster_caps: dict[str, int] = {}
            for row in ws.iter_rows(min_row=18, values_only=True):
                cluster = row[cluster_col]
                pieces = row[total_col]
                if cluster and isinstance(pieces, (int, float)):
                    cluster_caps[str(cluster)] = int(pieces)
            if cluster_caps:
                caps[offer_id] = cluster_caps
        except Exception as e:
            print(f"  ! 读 {p.name}: {e}")
    return caps


# ------- storage warehouse → cluster mapping (从 /cluster/list 拉) -------

def build_wh_to_cluster(account: str) -> dict[str, str]:
    """从 Ozon /cluster/list 拉所有 cluster + 它们的 warehouses, 反建 wh→cluster 映射."""
    r = _post(account, "/cluster/list",
              {"cluster_ids": [], "cluster_type": "CLUSTER_TYPE_OZON"})
    out: dict[str, str] = {}
    for c in r.get("clusters", []):
        cluster_name = c.get("name", "")
        for lc in c.get("logistic_clusters", []):
            for wh in lc.get("warehouses", []):
                wh_name = wh.get("name", "")
                if wh_name and cluster_name:
                    out[wh_name] = cluster_name
    return out


# ------- main -------

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--account", default=DEFAULT_ACCOUNT)
    ap.add_argument("--since", default="2026-04-22T00:00:00Z")
    ap.add_argument("--to", default="2026-04-26T00:00:00Z")
    ap.add_argument("--out", default="")
    args = ap.parse_args()

    print(f"[1/4] 读 allocation_*.xlsx 取计划 caps")
    caps = read_allocation_caps(SCRIPT_DIR)
    print(f"  → {len(caps)} 个 offer_id 有计划")
    for o, cl in caps.items():
        total = sum(cl.values())
        print(f"    {o}: {len(cl)} 集群, 合计 {total} 件")

    print(f"\n[1.5/4] 拉 wh → cluster 映射")
    wh_map = build_wh_to_cluster(args.account)
    print(f"  → {len(wh_map)} 个仓库")

    # 反向 stem→cluster 用于 fuzzy fallback (老仓如 НОВОСИБИРСК_РФЦ_НОВЫЙ → 找含 НОВОСИБИРСК 的)
    stem_to_cluster: dict[str, str] = {}
    for wh, cl in wh_map.items():
        # 取 wh 第一个 _ 前的 stem (НОВОСИБИРСК_РФЦ → НОВОСИБИРСК)
        stem = wh.split("_")[0]
        stem_to_cluster.setdefault(stem, cl)

    def storage_to_cluster(wh_name: str) -> str:
        if not wh_name:
            return "?"
        if wh_name in wh_map:
            return wh_map[wh_name]
        # fuzzy: 查 stem (老仓 _НОВЫЙ 后缀等)
        stem = wh_name.split("_")[0]
        if stem in stem_to_cluster:
            return stem_to_cluster[stem]
        return f"?({wh_name})"

    print(f"\n[2/4] 拉 Ozon supply orders ({args.since} ~ {args.to})")
    ids = list_supply_orders(args.account, args.since, args.to)
    print(f"  → {len(ids)} orders: {ids}")

    print(f"\n[3/4] 拉每个 order 的详情 + bundle items")
    orders = get_orders(args.account, ids)
    rows = []  # 每行: (order_id, supply_id, state, cluster, offer_id, qty, is_cancelled)
    for o in orders:
        order_id = o.get("order_id")
        order_number = o.get("order_number", "")
        order_state = o.get("state", "")
        for sup in o.get("supplies") or []:
            sup_id = sup.get("supply_id")
            sup_state = sup.get("state", "")
            wh = (sup.get("storage_warehouse") or {}).get("name", "")
            cluster = storage_to_cluster(wh)
            bundle_id = sup.get("bundle_id")
            is_cancelled = (
                sup_state == "CANCELLED"
                or order_state == "CANCELLED"
            )
            if bundle_id:
                items = get_bundle_items(args.account, bundle_id)
            else:
                items = []
            for it in items:
                rows.append({
                    "order_id": order_id,
                    "order_number": order_number,
                    "supply_id": sup_id,
                    "supply_state": sup_state,
                    "order_state": order_state,
                    "storage_warehouse": wh,
                    "cluster": cluster,
                    "offer_id": it.get("offer_id", ""),
                    "sku": it.get("sku", ""),
                    "quantity": int(it.get("quantity") or 0),
                    "is_cancelled": is_cancelled,
                })
            print(
                f"  order={order_id} sup={sup_id} state={sup_state} "
                f"wh={wh} → cluster={cluster} items={len(items)}"
            )

    # 聚合: (cluster, offer_id) → {planned, shipped_active, shipped_cancelled, supply_ids_active, supply_ids_cancelled}
    agg: dict[tuple[str, str], dict] = {}
    for r in rows:
        key = (r["cluster"], r["offer_id"])
        a = agg.setdefault(key, {
            "planned": 0, "shipped_active": 0, "shipped_cancelled": 0,
            "supply_ids_active": [], "supply_ids_cancelled": [],
            "sku": r["sku"],
        })
        if r["is_cancelled"]:
            a["shipped_cancelled"] += r["quantity"]
            a["supply_ids_cancelled"].append(r["supply_id"])
        else:
            a["shipped_active"] += r["quantity"]
            a["supply_ids_active"].append(r["supply_id"])
    # 注入 planned 件数
    for offer_id, cluster_caps in caps.items():
        for cluster, planned in cluster_caps.items():
            key = (cluster, offer_id)
            agg.setdefault(key, {
                "planned": 0, "shipped_active": 0, "shipped_cancelled": 0,
                "supply_ids_active": [], "supply_ids_cancelled": [],
                "sku": "",
            })
            agg[key]["planned"] = planned

    print(f"\n[4/4] 输出 xlsx ({len(agg)} 行)")
    out = args.out or str(SCRIPT_DIR / f"shipment_overview_{args.account}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "计划 vs 已发"
    headers = [
        "集群", "offer_id (货号)", "SKU",
        "计划件数 (allocation 合计件)",
        "已发件数 (active)", "已发件数 (cancelled)",
        "差额 (计划-active)",
        "状态",
        "active supply_ids", "cancelled supply_ids",
    ]
    ws.append(headers)
    bold = Font(bold=True)
    for c in ws[1]:
        c.font = bold
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    fill_warn = PatternFill("solid", fgColor="FFE9C2")
    fill_done = PatternFill("solid", fgColor="D6F0C8")
    fill_cancel = PatternFill("solid", fgColor="F0D6D6")

    # 排序: 集群按字典序, offer_id 内按字典序
    for (cluster, offer_id), a in sorted(agg.items()):
        diff = a["planned"] - a["shipped_active"]
        if a["shipped_cancelled"] > 0 and a["shipped_active"] == 0:
            status = "全部取消"
            row_fill = fill_cancel
        elif diff <= 0 and a["planned"] > 0:
            status = "已完成"
            row_fill = fill_done
        elif a["planned"] == 0 and a["shipped_active"] > 0:
            status = "无计划-意外发货"
            row_fill = fill_warn
        elif diff > 0:
            status = "缺 {} 件".format(diff)
            row_fill = fill_warn
        else:
            status = "无计划"
            row_fill = None
        row = [
            cluster, offer_id, a["sku"] or "",
            a["planned"], a["shipped_active"], a["shipped_cancelled"],
            diff, status,
            ",".join(str(x) for x in a["supply_ids_active"]),
            ",".join(str(x) for x in a["supply_ids_cancelled"]),
        ]
        ws.append(row)
        if row_fill:
            for c in ws[ws.max_row]:
                c.fill = row_fill
    # 列宽
    widths = [16, 28, 12, 12, 12, 14, 12, 14, 30, 30]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(ord('A') + i - 1)].width = w
    # 总计行
    ws.append([])
    total_planned = sum(a["planned"] for a in agg.values())
    total_active = sum(a["shipped_active"] for a in agg.values())
    total_cancel = sum(a["shipped_cancelled"] for a in agg.values())
    ws.append([
        "合计", "", "",
        total_planned, total_active, total_cancel,
        total_planned - total_active,
        f"完成率 {total_active/max(total_planned,1)*100:.1f}%",
        "", "",
    ])
    for c in ws[ws.max_row]:
        c.font = bold

    wb.save(out)
    print(f"  → {out}")
    print(f"\n=== 摘要 ===")
    print(f"  计划合计:  {total_planned} 件")
    print(f"  已发 active: {total_active} 件 ({total_active/max(total_planned,1)*100:.1f}%)")
    print(f"  已发 cancelled: {total_cancel} 件")
    print(f"  仍缺:      {total_planned - total_active} 件")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
