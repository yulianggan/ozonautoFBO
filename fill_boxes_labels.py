#!/usr/bin/env python3
"""Ozon FBO 填箱 + 拉箱唛 PDF + 产出对照表.

承接 create_moscow_shipment.py / restock_planner.py 的产出:
  1. 用 order_id / supply_id 去 Ozon 后台查单 (/v3/supply-order/get)
  2. 按 qty + box_size 切箱: cargoes = [{key:"1", items:[{barcode, quantity}], type:"BOX"}, ...]
  3. 调 /flow/upload-cargoes 一把梭: cargoes/create → poll → label/create → poll → 下载 PDF
  4. 写 Excel 对照表 (交货ID / 供货ID / 货位ID / 集群 / SKU / 货号 / 箱唛 PDF 路径 ...)

典型用法:
  # 单个 order_id (已建好的 FBO 供货单)
  python3 fill_boxes_labels.py \
      --order-id 100452984 --offer-id Q_MeiGongDao-HeiSe-Free \
      --sku 3214793665 --qty 4200 --box-size 300 \
      --cluster "Москва, МО и Дальние регионы" \
      --account 丝绸生活

  # 从 fbo_plan_*.result.json 批量 (execute_fbo_plan 的结果)
  python3 fill_boxes_labels.py \
      --plan-result fbo_plan_Q_MeiGongDao-HeiSe-Free_20260422.result.json \
      --box-size 300 --account 丝绸生活

  # dry-run 只打箱+打表, 不调 Ozon
  python3 fill_boxes_labels.py --order-id 100452984 ... --dry-run
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from datetime import date
from pathlib import Path
from typing import Any
from urllib.parse import quote

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_FBO_SERVICE = "http://127.0.0.1:4182"


# ---------- HTTP ----------

def _post(service: str, account: str, path: str, body: dict[str, Any]) -> dict[str, Any]:
    r = requests.post(
        f"{service}{path}",
        json=body,
        headers={"X-Ozon-Account": quote(account, safe="")},
        timeout=300,
    )
    if r.status_code >= 400:
        raise RuntimeError(
            f"FBO {path} HTTP {r.status_code}: {r.text[:500]}"
        )
    return r.json()


# ---------- 查单 ----------

def fetch_order(service: str, account: str, order_id: int) -> dict[str, Any]:
    resp = _post(service, account, "/supply-order/get", {"order_ids": [order_id]})
    orders = resp.get("orders") or []
    if not orders:
        raise RuntimeError(f"order_id={order_id} 未查到")
    return orders[0]


# ---------- 切箱 ----------

def split_into_cargoes(
    barcode: str, qty: int, box_size: int
) -> list[dict[str, Any]]:
    """按 box_size 平均切 + 最后一箱装余数. 返回 /flow/upload-cargoes 的 cargoes 字段."""
    if qty <= 0:
        return []
    if box_size <= 0:
        raise ValueError("box_size 必须 >0")
    n = math.ceil(qty / box_size)
    cargoes: list[dict[str, Any]] = []
    remaining = qty
    for i in range(n):
        in_box = box_size if remaining >= box_size else remaining
        remaining -= in_box
        cargoes.append({
            "key": str(i + 1),
            "items": [{"barcode": barcode, "quantity": in_box}],
            "type": "BOX",
        })
    return cargoes


# ---------- 申报箱 + 拉标签 (封装 /flow/upload-cargoes) ----------

def upload_cargoes_and_label(
    service: str,
    account: str,
    supply_id: int,
    cargoes: list[dict[str, Any]],
    save_label_pdf: bool = True,
) -> dict[str, Any]:
    body = {
        "supply_id": supply_id,
        "cargoes": cargoes,
        "delete_current_version": True,
        "generate_label": True,
        "save_label_pdf": save_label_pdf,
        "return_pdf_base64": False,
        "poll_interval": 2.0,
        "poll_timeout": 180.0,
    }
    return _post(service, account, "/flow/upload-cargoes", body)


# ---------- 对照表输出 ----------

_HEADERS = [
    "交货 ID", "供货 ID", "状态",
    "集群", "macrolocal_cluster_id",
    "货位 ID", "货位名称",
    "drop-off ID", "drop-off 名称",
    "货号 (offer_id)", "SKU", "条码",
    "箱号", "cargo_id", "该箱件数",
    "总件数", "总箱数", "单箱装箱率",
    "时段 (MSK/UTC)",
    "箱唛 PDF 路径",
    "错误",
]


def write_mapping_xlsx(rows: list[dict[str, Any]], out_path: Path) -> Path:
    """rows 为 per-box 行 (每箱一条). 相邻同 order 的行合并视觉上的 order 信息."""
    wb = Workbook()
    ws = wb.active
    ws.title = "FBO 对照表"

    bold = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.append(_HEADERS)
    for c in range(1, len(_HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = center

    for r in rows:
        ws.append([r.get(h, "") for h in _HEADERS])

    widths = [12, 14, 14, 32, 12, 18, 22, 16, 22, 28, 14, 16, 6, 18, 10, 10, 8, 10, 30, 50, 30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


# ---------- 装配一行 ----------

def _timeslot_str(order: dict[str, Any]) -> str:
    ts = (order.get("timeslot") or {}).get("timeslot") or {}
    return f"{ts.get('from','')} ~ {ts.get('to','')}"


def build_rows(
    order: dict[str, Any],
    offer_id: str,
    sku: int,
    qty: int,
    box_size: int,
    cluster: str,
    cargo_result: dict[str, Any] | None,
    error: str = "",
) -> list[dict[str, Any]]:
    """一行一箱: 对每个 cargo_id 发一行, 含箱号+该箱件数. 无 cargo_ids 时按预期箱数发占位行."""
    drop_off = order.get("drop_off_warehouse") or {}
    supplies = order.get("supplies") or [{}]
    supply = supplies[0]
    storage = supply.get("storage_warehouse") or {}
    macro = supply.get("macrolocal_cluster_id")

    cargo_ids: list[int] = []
    pdf_path = ""
    if cargo_result:
        for c in cargo_result.get("cargoes") or []:
            value = c.get("value") or {}
            cid = value.get("cargo_id") or c.get("cargo_id")
            if cid:
                cargo_ids.append(int(cid))
        pdf_path = cargo_result.get("pdf_path") or ""

    n_boxes_expected = math.ceil(qty / box_size) if box_size > 0 else 0
    box_qtys: list[int] = []
    remaining = qty
    for _ in range(n_boxes_expected):
        in_box = box_size if remaining >= box_size else remaining
        box_qtys.append(in_box)
        remaining -= in_box

    n_rows = max(n_boxes_expected, len(cargo_ids), 1)

    base = {
        "交货 ID": order.get("order_id"),
        "供货 ID": supply.get("supply_id", ""),
        "状态": order.get("state", ""),
        "集群": cluster,
        "macrolocal_cluster_id": macro or "",
        "货位 ID": storage.get("warehouse_id", ""),
        "货位名称": storage.get("name", ""),
        "drop-off ID": drop_off.get("warehouse_id", ""),
        "drop-off 名称": drop_off.get("name", ""),
        "货号 (offer_id)": offer_id,
        "SKU": sku,
        "条码": f"OZN{sku}",
        "总件数": qty,
        "总箱数": n_boxes_expected,
        "单箱装箱率": box_size,
        "时段 (MSK/UTC)": _timeslot_str(order),
        "箱唛 PDF 路径": pdf_path,
    }

    rows: list[dict[str, Any]] = []
    for i in range(n_rows):
        row = dict(base)
        row["箱号"] = i + 1
        row["cargo_id"] = cargo_ids[i] if i < len(cargo_ids) else ""
        row["该箱件数"] = box_qtys[i] if i < len(box_qtys) else ""
        # 错误信息只写在第一行, 避免视觉污染
        row["错误"] = error if i == 0 else ""
        rows.append(row)
    return rows


# ---------- 单 order 处理 ----------

def process_one(
    service: str,
    account: str,
    order_id: int,
    offer_id: str,
    sku: int,
    qty: int,
    box_size: int,
    cluster: str,
    dry_run: bool,
) -> list[dict[str, Any]]:
    """返回 per-box 行列表 (一箱一行)."""
    print(f"[order {order_id}] 查单 ...", flush=True)
    order = fetch_order(service, account, order_id)
    supplies = order.get("supplies") or []
    if not supplies:
        raise RuntimeError(f"order {order_id} 无 supplies")
    supply_id = int(supplies[0]["supply_id"])
    state = order.get("state", "")
    print(f"[order {order_id}] supply_id={supply_id} state={state}", flush=True)

    cargoes = split_into_cargoes(f"OZN{sku}", qty, box_size)
    print(f"[order {order_id}] 切箱: {len(cargoes)} 箱, 末箱 {cargoes[-1]['items'][0]['quantity']} 件", flush=True)

    cargo_result: dict[str, Any] | None = None
    error = ""
    if dry_run:
        print(f"[order {order_id}] dry-run, 不申报", flush=True)
    else:
        try:
            cargo_result = upload_cargoes_and_label(
                service, account, supply_id, cargoes, save_label_pdf=True,
            )
            stage = cargo_result.get("stage")
            print(f"[order {order_id}] /flow/upload-cargoes stage={stage}", flush=True)
            if stage not in {"done", "cargoes_created"}:
                error = f"stage={stage}"
            if cargo_result.get("pdf_path"):
                print(f"[order {order_id}] 箱唛 PDF: {cargo_result['pdf_path']}", flush=True)
        except Exception as e:
            error = str(e)
            print(f"[order {order_id}] ERROR: {error}", flush=True)

    return build_rows(order, offer_id, sku, qty, box_size, cluster,
                      cargo_result, error)


# ---------- CLI ----------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Ozon FBO 填箱+箱唛 PDF+对照表",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--order-id", type=int, help="单个已建好的供货单 order_id")
    g.add_argument("--plan-result", type=str,
                   help="fbo_plan_*.result.json (execute_fbo_plan 输出)")

    p.add_argument("--offer-id", type=str, default="",
                   help="货号 (单 order 模式必填)")
    p.add_argument("--sku", type=int, default=0,
                   help="Ozon 数字 sku (单 order 模式必填)")
    p.add_argument("--qty", type=int, default=0,
                   help="总件数 (单 order 模式必填)")
    p.add_argument("--cluster", type=str, default="",
                   help="集群名, 给对照表用 (单 order 模式可传)")

    p.add_argument("--box-size", type=int, required=True,
                   help="单箱装箱率 (件/箱)")
    p.add_argument("--account", default="丝绸生活", help="X-Ozon-Account")
    p.add_argument("--fbo-service", default=DEFAULT_FBO_SERVICE)
    p.add_argument("--out-xlsx", default="",
                   help="对照表输出路径, 缺省: fbo_mapping_<date>.xlsx")
    p.add_argument("--dry-run", action="store_true",
                   help="只切箱+打表, 不调 Ozon (保护已建好的单)")
    return p.parse_args()


def main() -> int:
    args = parse_args()

    if args.order_id:
        if not (args.offer_id and args.sku and args.qty):
            print("[ERROR] 单 order 模式需 --offer-id / --sku / --qty", file=sys.stderr)
            return 2
        jobs: list[dict[str, Any]] = [{
            "order_id": args.order_id,
            "offer_id": args.offer_id,
            "sku": args.sku,
            "qty": args.qty,
            "cluster": args.cluster,
        }]
    else:
        plan_path = Path(args.plan_result)
        if not plan_path.exists():
            print(f"[ERROR] 找不到 {plan_path}", file=sys.stderr)
            return 2
        results = json.loads(plan_path.read_text(encoding="utf-8"))
        jobs = []
        for r in results:
            if r.get("error"):
                print(f"[skip] {r.get('cluster')} 建单失败: {r['error']}")
                continue
            for oid in (r.get("order_ids") or []):
                jobs.append({
                    "order_id": int(oid),
                    "offer_id": r.get("offer_id", ""),
                    "sku": int(r.get("sku", 0)),
                    "qty": int(r.get("qty_pieces", 0)),
                    "cluster": r.get("cluster", ""),
                })
        if not jobs:
            print(f"[ERROR] {plan_path.name} 里无 order_ids, 先跑 restock_planner --execute",
                  file=sys.stderr)
            return 2

    print(f"[total] {len(jobs)} 个 order 待处理 (dry_run={args.dry_run})")
    all_rows: list[dict[str, Any]] = []
    for i, job in enumerate(jobs, start=1):
        print(f"\n=== {i}/{len(jobs)} order={job['order_id']} ===")
        try:
            rows = process_one(
                args.fbo_service, args.account,
                job["order_id"], job["offer_id"], job["sku"],
                job["qty"], args.box_size, job["cluster"],
                args.dry_run,
            )
        except Exception as e:
            rows = [{
                "交货 ID": job["order_id"],
                "货号 (offer_id)": job["offer_id"],
                "SKU": job["sku"],
                "总件数": job["qty"],
                "集群": job["cluster"],
                "单箱装箱率": args.box_size,
                "箱号": 1,
                "错误": str(e),
            }]
        all_rows.extend(rows)

    out = args.out_xlsx
    if not out:
        tag = date.today().strftime("%Y%m%d")
        out = str(SCRIPT_DIR / f"fbo_mapping_{tag}.xlsx")
    out_path = write_mapping_xlsx(all_rows, Path(out))
    print(f"\n[done] 对照表: {out_path} ({len(all_rows)} 行)")

    print("\n=== 摘要 (按 order 聚合) ===")
    by_order: dict[Any, list[dict[str, Any]]] = {}
    for r in all_rows:
        by_order.setdefault(r.get("交货 ID"), []).append(r)
    for oid, rs in by_order.items():
        r0 = rs[0]
        print(f"  交货={oid} 供货={r0.get('供货 ID')} "
              f"货位={r0.get('货位 ID')} 集群={r0.get('集群')} "
              f"SKU={r0.get('SKU')} 箱数={len(rs)} "
              f"错误={r0.get('错误','')}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
