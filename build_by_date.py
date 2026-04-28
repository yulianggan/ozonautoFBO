"""按 Ozon timeslot 日期整理 FBO 对照表 + 合并箱唛 PDF.

数据源: Ozon /supply-order/list + /get + /bundle (含手工建单).
输出: by_date/YYYY-MM-DD/
        对照表.xlsx     # 22 列, 跟 create_fbo_plan.py 输出同
        箱唛合并.pdf    # 同日所有 supply PDF 用 pypdf 合并

用法:
  python3 build_by_date.py                              # 默认丝绸生活, 当月
  python3 build_by_date.py --account 个人之路 --since 2026-04-22 --to 2026-04-30
"""
from __future__ import annotations

import argparse
import json
import math
import time
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pypdf import PdfReader, PdfWriter

SCRIPT_DIR = Path(__file__).resolve().parent
SVC = "http://127.0.0.1:4182"
DEFAULT_ACCOUNT = "丝绸生活"
LABEL_DIRS = [
    Path("/Users/mac/Documents/ozns/github/ozon_fbo_shipment_service/labels"),
    SCRIPT_DIR / "labels",
]


_HEADERS = [
    "交货 ID", "供货 ID", "状态", "链路", "集群", "macrolocal_cluster_id",
    "货位 ID", "货位名称", "drop-off ID", "drop-off 名称",
    "货号 (offer_id)", "SKU", "条码",
    "箱号", "cargo_id", "该箱件数",
    "总件数", "总箱数", "单箱装箱率",
    "时段",
    "箱唛 PDF 路径",
    "备注/错误",
]


def _post(account: str, path: str, body: dict, retries: int = 3, timeout: int = 60) -> dict:
    acc = urllib.parse.quote(account)
    headers = {"Content-Type": "application/json", "X-Ozon-Account": acc}
    req = urllib.request.Request(f"{SVC}{path}", data=json.dumps(body).encode(), headers=headers)
    last_err: Exception | None = None
    for _ in range(retries):
        try:
            return json.load(urllib.request.urlopen(req, timeout=timeout))
        except Exception as e:
            last_err = e
            time.sleep(2)
    raise RuntimeError(f"POST {path} failed after {retries}: {last_err}")


def list_supply_orders(account: str, since: str, to: str, states: list[str]) -> list[int]:
    ids: list[int] = []
    last_id = ""
    for _ in range(40):
        body = {
            "filter": {
                "states": states,
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
    for s in range(0, len(ids), 30):
        chunk = ids[s:s + 30]
        r = _post(account, "/supply-order/get", {"order_ids": chunk})
        out.extend(r.get("orders", []))
    return out


def get_bundle_items(account: str, bundle_id: str) -> list[dict]:
    r = _post(account, "/supply-order/bundle", {"bundle_ids": [bundle_id], "limit": 100})
    return r.get("items") or r.get("contents") or []


def get_cargoes(account: str, supply_id: int) -> list[dict]:
    """返回 [{"cargo_id":..., "type":"BOX", ...}, ...] (cargo_id list, 没 items 嵌套).

    items 通过 /supply-order/bundle 单独拿; 每箱内容用 box_size 推算.
    """
    try:
        r = _post(account, "/cargoes/get", {"supply_ids": [supply_id]})
    except Exception:
        return []
    for s in r.get("supply") or []:
        if int(s.get("supply_id", 0)) == int(supply_id):
            return s.get("cargoes") or []
    return []


def find_pdf(supply_id: int) -> Path | None:
    for d in LABEL_DIRS:
        if not d.exists():
            continue
        # 多种命名: <supply_id>.pdf 或 supply_<supply_id>.pdf 或 包含 supply_id
        for name in (f"{supply_id}.pdf", f"supply_{supply_id}.pdf"):
            p = d / name
            if p.exists():
                return p
        # fallback: substring 匹配
        for f in d.iterdir():
            if f.is_file() and str(supply_id) in f.name and f.suffix.lower() == ".pdf":
                return f
    return None


def build_rows_for_supply(account: str, order: dict) -> list[dict]:
    """单 order 拆 supply 出 22 列 rows; mode 推断: state SUPPLIED/SUPPLIED_PARTIALLY → SUPPLIED, 其他原样."""
    order_id = order.get("supply_order_id") or order.get("order_id", "")
    state = order.get("state", "")
    drop_off_info = (order.get("drop_off_warehouse") or {})
    ts = (order.get("timeslot") or {}).get("timeslot") or {}
    ts_str = f"{ts.get('from','')} ~ {ts.get('to','')}"

    supplies = order.get("supplies") or []
    rows: list[dict] = []
    for supply in supplies:
        supply_id = supply.get("supply_id", "")
        bundle_id = supply.get("bundle_id", "")
        storage = supply.get("storage_warehouse") or {}
        cluster_name = (supply.get("cluster") or {}).get("name") or ""
        macrolocal = supply.get("macrolocal_cluster_id", "")
        # mode 推断: 1 个 supply 的 order 大概率单集群 CROSSDOCK; 多 supply order 是 multi-cluster
        mode = "MULTI_CLUSTER" if len(supplies) > 1 else "CROSSDOCK"

        items = get_bundle_items(account, bundle_id) if bundle_id else []
        cargoes = get_cargoes(account, supply_id) if supply_id else []

        # build per-cargo box rows
        # cargo[i] 含 items, 每 cargo 是 1 箱
        # offer_id / sku / barcode 从 items 拼回
        offer_by_sku = {it.get("sku"): it.get("offer_id", "") for it in items}
        bc_by_sku = {it.get("sku"): it.get("barcode", "") for it in items}
        # 每 SKU 总件 + 总箱数
        total_pieces_by_sku: dict = {}
        for it in items:
            sku = it.get("sku")
            total_pieces_by_sku[sku] = total_pieces_by_sku.get(sku, 0) + int(it.get("quantity", 0) or 0)

        # 从 cargoes 拆每箱: cargo_id list 来自 /cargoes/get (无 items 嵌套);
        # items 来自 /supply-order/bundle (sku 总件); 用 box_size 推算每箱内容
        pdf = find_pdf(int(supply_id)) if supply_id else None
        pdf_path = str(pdf) if pdf else "(未找到)"
        if cargoes and items:
            cargo_ids_list = [c.get("cargo_id", "") for c in cargoes]
            n_cargoes = len(cargoes)
            # 推算 box_size: 总件数 / 总箱数 (取整, 默认 300)
            total_all = sum(total_pieces_by_sku.values())
            inferred_box = total_all // n_cargoes if n_cargoes > 0 else 300
            box_size = inferred_box if inferred_box > 0 else 300
            # 按 SKU split 成 (sku, box_idx, qty_in_box)
            flat_boxes: list[tuple] = []
            for sku, qty in total_pieces_by_sku.items():
                n = math.ceil(qty / box_size) if box_size > 0 else 0
                remain = qty
                for i in range(n):
                    in_box = box_size if remain >= box_size else remain
                    remain -= in_box
                    flat_boxes.append((sku, i + 1, in_box))
            # 映射 cargo_id 到 flat_boxes (order)
            for global_idx, (sku, box_i, q) in enumerate(flat_boxes):
                cid = cargo_ids_list[global_idx] if global_idx < len(cargo_ids_list) else ""
                rows.append({
                    "交货 ID": order_id, "供货 ID": supply_id, "状态": state, "链路": mode,
                    "集群": cluster_name, "macrolocal_cluster_id": macrolocal,
                    "货位 ID": storage.get("warehouse_id", ""),
                    "货位名称": storage.get("name", ""),
                    "drop-off ID": drop_off_info.get("warehouse_id", ""),
                    "drop-off 名称": drop_off_info.get("name", ""),
                    "货号 (offer_id)": offer_by_sku.get(sku, ""),
                    "SKU": sku, "条码": bc_by_sku.get(sku, ""),
                    "箱号": box_i, "cargo_id": cid, "该箱件数": q,
                    "总件数": total_pieces_by_sku.get(sku, ""),
                    "总箱数": n_cargoes,
                    "单箱装箱率": box_size,
                    "时段": ts_str, "箱唛 PDF 路径": pdf_path,
                    "备注/错误": "",
                })
        else:
            # 没 cargoes (DATA_FILLING 状态, 还没切箱) — 按 items 一行
            for it in items:
                sku = it.get("sku")
                rows.append({
                    "交货 ID": order_id, "供货 ID": supply_id, "状态": state, "链路": mode,
                    "集群": cluster_name, "macrolocal_cluster_id": macrolocal,
                    "货位 ID": storage.get("warehouse_id", ""),
                    "货位名称": storage.get("name", ""),
                    "drop-off ID": drop_off_info.get("warehouse_id", ""),
                    "drop-off 名称": drop_off_info.get("name", ""),
                    "货号 (offer_id)": it.get("offer_id", ""),
                    "SKU": sku, "条码": it.get("barcode", ""),
                    "箱号": "", "cargo_id": "", "该箱件数": "",
                    "总件数": int(it.get("quantity", 0) or 0),
                    "总箱数": "", "单箱装箱率": "",
                    "时段": ts_str, "箱唛 PDF 路径": pdf_path,
                    "备注/错误": "(未切箱/无 cargoes)",
                })
            if not items:
                rows.append({
                    "交货 ID": order_id, "供货 ID": supply_id, "状态": state, "链路": mode,
                    "集群": cluster_name, "macrolocal_cluster_id": macrolocal,
                    "货位 ID": storage.get("warehouse_id", ""), "货位名称": storage.get("name", ""),
                    "drop-off ID": drop_off_info.get("warehouse_id", ""),
                    "drop-off 名称": drop_off_info.get("name", ""),
                    "货号 (offer_id)": "", "SKU": "", "条码": "",
                    "箱号": "", "cargo_id": "", "该箱件数": "",
                    "总件数": "", "总箱数": "", "单箱装箱率": "",
                    "时段": ts_str, "箱唛 PDF 路径": pdf_path,
                    "备注/错误": "(无 items, 可能 DRAFT/CANCELLED)",
                })
    return rows


_STRING_FIELDS = {"供货 ID", "货位 ID", "drop-off ID", "cargo_id"}


def write_xlsx(rows: list[dict], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "FBO 对照表"
    ws.append(_HEADERS)
    bold = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F4E78")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for c in range(1, len(_HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = bold
        cell.fill = fill
        cell.alignment = center
    for r in rows:
        out = []
        for h in _HEADERS:
            v = r.get(h, "")
            if h in _STRING_FIELDS and v not in ("", None):
                v = str(v)
            out.append(v)
        ws.append(out)
    # 把 4 列设成文本格式 (防 Excel 重新解析为科学计数法)
    string_col_idxs = [i + 1 for i, h in enumerate(_HEADERS) if h in _STRING_FIELDS]
    for col_idx in string_col_idxs:
        for row_idx in range(2, ws.max_row + 1):
            ws.cell(row=row_idx, column=col_idx).number_format = "@"
    widths = [12, 14, 18, 10, 32, 12, 18, 22, 16, 22, 28, 14, 16, 6, 18, 10, 10, 8, 10, 28, 50, 30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def merge_pdfs(pdf_paths: list[Path], out_path: Path) -> int:
    """合并 PDF 用 pypdf, 返回合并的页数."""
    writer = PdfWriter()
    n_pages = 0
    for p in pdf_paths:
        try:
            reader = PdfReader(str(p))
            for page in reader.pages:
                writer.add_page(page)
                n_pages += 1
        except Exception as e:
            print(f"  ! 跳过 PDF {p.name}: {e}")
    if n_pages > 0:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        with open(out_path, "wb") as f:
            writer.write(f)
    return n_pages


def main() -> int:
    today = date.today().strftime("%Y-%m-%d")
    p = argparse.ArgumentParser(description="按 timeslot 日期整理 FBO 对照表 + 合并箱唛")
    p.add_argument("--account", default=DEFAULT_ACCOUNT)
    p.add_argument("--since", default="2026-04-22T00:00:00Z", help="ISO 起始")
    p.add_argument("--to", default="2026-05-31T00:00:00Z", help="ISO 终止")
    p.add_argument("--out-root", default=str(SCRIPT_DIR / "by_date"))
    p.add_argument(
        "--states",
        default="READY_TO_SUPPLY",
        help="逗号分隔, 默认 READY_TO_SUPPLY (Ozon 后台'准备交货' Tab). "
             "其他: DATA_FILLING/DRAFT/SUPPLIED/SUPPLIED_PARTIALLY/SUPPLY_PROCESSING/"
             "SHIPPING_PREPARED/SUPPLY_REJECTED/CANCELLED. 'all' = 全部",
    )
    args = p.parse_args()

    if args.states.strip().lower() == "all":
        states = [
            "READY_TO_SUPPLY", "DATA_FILLING", "DRAFT", "SUPPLIED",
            "SUPPLIED_PARTIALLY", "SUPPLY_PROCESSING", "SUPPLY_REJECTED",
            "SHIPPING_PREPARED", "CANCELLED",
        ]
    else:
        states = [s.strip() for s in args.states.split(",") if s.strip()]

    print(f"[1/4] /supply-order/list states={states} since={args.since} to={args.to}")
    ids = list_supply_orders(args.account, args.since, args.to, states)
    print(f"  → {len(ids)} orders")

    print(f"[2/4] /supply-order/get + bundle + cargoes/get (per order)")
    orders = get_orders(args.account, ids)
    by_date: dict[str, list[dict]] = defaultdict(list)
    pdf_by_date: dict[str, set[Path]] = defaultdict(set)
    for o in orders:
        ts = (o.get("timeslot") or {}).get("timeslot") or {}
        ts_from = ts.get("from", "")
        if not ts_from:
            ds = "_no_timeslot"
        else:
            ds = ts_from[:10]  # YYYY-MM-DD
        rows = build_rows_for_supply(args.account, o)
        by_date[ds].extend(rows)
        for s in (o.get("supplies") or []):
            pdf = find_pdf(int(s.get("supply_id", 0)))
            if pdf:
                pdf_by_date[ds].add(pdf)

    out_root = Path(args.out_root)
    print(f"[3/4] 写出到 {out_root}/<日期>/")
    summary = []
    for ds, rows in sorted(by_date.items()):
        date_dir = out_root / ds
        xlsx_path = date_dir / "对照表.xlsx"
        write_xlsx(rows, xlsx_path)
        pdf_paths = sorted(pdf_by_date.get(ds, set()))
        merged_pdf = date_dir / "箱唛合并.pdf"
        n_pages = merge_pdfs(pdf_paths, merged_pdf) if pdf_paths else 0
        summary.append({
            "date": ds,
            "orders": len({r.get("交货 ID") for r in rows}),
            "supplies": len({r.get("供货 ID") for r in rows}),
            "rows": len(rows),
            "pdfs": len(pdf_paths),
            "pdf_pages": n_pages,
        })
        print(f"  {ds}: {len(rows)} 行 / {len({r.get('供货 ID') for r in rows})} supplies / "
              f"{len(pdf_paths)} PDF ({n_pages} 页) → {date_dir}")

    print(f"\n[4/4] 完成")
    print(f"{'日期':<12} {'orders':>6} {'supplies':>8} {'rows':>5} {'pdfs':>5} {'pages':>5}")
    for s in summary:
        print(f"{s['date']:<12} {s['orders']:>6} {s['supplies']:>8} {s['rows']:>5} "
              f"{s['pdfs']:>5} {s['pdf_pages']:>5}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
