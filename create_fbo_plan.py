#!/usr/bin/env python3
"""Ozon FBO 发货 SOP: 多 SKU × 多集群 配送计划 → 供货单 → 箱唛 → 对照表.

SOP (2026-04-24 定稿):
  1. 先尝试 multi-cluster 一票货件
  2. 轮询 /v2/draft/create/info 读 availability_status 矩阵
     - AVAILABLE / PARTIAL_AVAILABLE 的集群 → 进 multi-cluster 供货单
     - NOT_AVAILABLE 的集群 (常见 reason: NOT_AVAILABLE_MATRIX) → fallback 成单集群 CROSSDOCK
  3. 所有成功建单的 supply_order 都做 cargoes/labels
  4. 汇总对照表 Excel (per-box: 交货ID/供货ID/货位ID/集群/SKU/货号/cargo_id/箱唛 PDF)

典型用法:
    python3 create_fbo_plan.py --config plan.json --box-size 300 --account 丝绸生活

plan.json:
{
  "source_warehouse_keyword": "ЖУКОВСКИЙ_РФЦ",
  "drop_off_keyword":         "ЩЕРБИНКА",
  "matrix": {
    "Новосибирск":    {"Q_ChongQiZui-HuangSe-Free": 600,  "Q_MeiGongDao-HeiSe-Free": 1800},
    "Дальний Восток": {"Q_ChongQiZui-HuangSe-Free": 300,  "Q_MeiGongDao-HeiSe-Free": 900}
  }
}
"""

from __future__ import annotations

import argparse
import json
import math
import sys
import time
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Callable
from urllib.parse import quote

import requests


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_FBO_SERVICE = "http://127.0.0.1:4182"
LEDGER_DIR = SCRIPT_DIR / "ledger"


# ---------- allocation_xlsx sanity check (cap 校验, 防上游 raw 需求泄漏) ----------


def _read_allocation_caps(script_dir: Path) -> dict[str, dict[str, int]]:
    """扫 SCRIPT_DIR/allocation_*xlsx, 返回 {offer_id: {cluster: 合计件}}.

    格式假设 (allocation_2026-04-22_*.xlsx):
      B1 = offer_id, R17 = headers (含 "集群" 和某个含 "(合计件)" 的列), R18+ 为数据
    """
    try:
        from openpyxl import load_workbook
    except ImportError:
        return {}
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
        except Exception:
            continue
    return caps


def assert_cfg_within_alloc(cfg: dict, caps: dict[str, dict[str, int]]) -> list[str]:
    """检查 cfg.matrix 件数不超 allocation_xlsx 合计件; 返回 violations 列表."""
    violations: list[str] = []
    for cluster, sku_qtys in (cfg.get("matrix") or {}).items():
        for offer_id, qty in sku_qtys.items():
            cap = caps.get(offer_id, {}).get(cluster)
            if cap is None:
                continue
            if int(qty) > cap:
                violations.append(
                    f"{cluster} / {offer_id}: cfg={qty} > allocation合计件={cap} "
                    f"(超 {int(qty) - cap} 件)"
                )
    return violations


# ---------- 台账 (持久化 shipped/skipped, 跨运行追踪未发部分) ----------


def _now_iso() -> str:
    from datetime import datetime, timezone
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def write_ledger(ledger: dict, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(ledger, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def merge_pending(ledger_dir: Path, out_path: Path) -> dict:
    """扫 ledger/*.json, 汇总所有 skipped → pending_summary.

    去重规则: 按 (cluster, offer_id, sku) 维度, 保留 timestamp 最新的那条 (代表最新失败原因).
    已在后续 run 发成功的 (cluster, offer_id) 从 pending 移除进 resolved_later.
    """
    raw_pending: list[dict] = []
    shipped_records: list[tuple[str, tuple, str]] = []  # (file_name, (cluster, offer_id), run_at)
    runs: list[dict] = []
    # 读所有 ledger
    files = sorted(ledger_dir.glob("shipment_ledger_*.json"))
    for f in files:
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        run_at = data.get("run_at","")
        runs.append({
            "file": f.name, "run_at": run_at,
            "shipped_count": len(data.get("shipped") or []),
            "skipped_count": len(data.get("skipped") or []),
        })
        for s in (data.get("shipped") or []):
            # shipped 记录支持多集群 (cluster_names list), 也支持单 cluster 字段
            cluster_names = s.get("cluster_names") or ([s.get("cluster")] if s.get("cluster") else [])
            for it in (s.get("items") or []):
                for cn in cluster_names:
                    shipped_records.append((f.name, (cn, it.get("offer_id")), run_at))
        for s in (data.get("skipped") or []):
            s2 = dict(s)
            s2["_source_file"] = f.name
            s2["_source_run_at"] = run_at
            raw_pending.append(s2)

    # 去重: 同 (cluster, offer_id, sku) 保留 timestamp 最新的
    def _entry_key(e: dict) -> tuple:
        return (e.get("cluster"), e.get("offer_id"), e.get("sku"))

    def _ts(e: dict) -> str:
        return e.get("timestamp") or e.get("_source_run_at") or ""

    by_key: dict[tuple, dict] = {}
    for p in raw_pending:
        k = _entry_key(p)
        if k not in by_key or _ts(p) > _ts(by_key[k]):
            by_key[k] = p
    deduped = list(by_key.values())

    # resolved: 对应 (cluster, offer_id) 已在某 run shipped
    shipped_keys = {rec[1] for rec in shipped_records}
    resolved: list[dict] = []
    still_pending: list[dict] = []
    for p in deduped:
        key = (p.get("cluster"), p.get("offer_id"))
        if key in shipped_keys:
            # 找最新的 shipped_record
            resolved_at = ""
            resolved_file = ""
            for fn, rk, ra in shipped_records:
                if rk == key and ra > resolved_at:
                    resolved_at = ra; resolved_file = fn
            p2 = dict(p)
            p2["_resolved_in"] = resolved_file
            p2["_resolved_at"] = resolved_at
            resolved.append(p2)
        else:
            still_pending.append(p)

    summary = {
        "generated_at": _now_iso(),
        "runs": runs,
        "pending_count": len(still_pending),
        "resolved_count": len(resolved),
        "raw_skipped_count": len(raw_pending),
        "pending": still_pending,
        "resolved_later": resolved,
    }
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary

# Ozon warehouse_type 短字符串 (multi-cluster 官方 schema 接受)
_WH_TYPE_STR_FROM_INT = {
    1: "DELIVERY_POINT", 2: "ORDERS_RECEIVING_POINT",
    3: "SORTING_CENTER", 4: "FULL_FILLMENT", 5: "CROSS_DOCK",
}
_WH_TYPE_NUM_FROM_STR = {v: k for k, v in _WH_TYPE_STR_FROM_INT.items()}


# ---------- restock_decision_service (4181) 集成 ----------

DEFAULT_RESTOCK_SERVICE = "http://127.0.0.1:4181"


def _rs_get(service: str, path: str, timeout: int = 15) -> dict:
    r = requests.get(f"{service.rstrip('/')}{path}", timeout=timeout)
    if r.status_code >= 400:
        raise RuntimeError(f"restock_service GET {path} → HTTP {r.status_code} {r.text[:200]}")
    return r.json()


def _rs_post(service: str, path: str, body: dict, timeout: int = 15) -> dict:
    r = requests.post(
        f"{service.rstrip('/')}{path}", json=body, timeout=timeout,
        headers={"Content-Type": "application/json"},
    )
    if r.status_code >= 400:
        raise RuntimeError(f"restock_service POST {path} → HTTP {r.status_code} {r.text[:200]}")
    return r.json()


def fetch_plans_by_ids(service: str, plan_ids: list[str]) -> list[dict]:
    """按 plan_id 列表拉完整 plan 详情."""
    out = []
    for pid in plan_ids:
        p = _rs_get(service, f"/plans/{pid}")
        out.append(p)
    return out


def fetch_plans_by_batch(service: str, account: str, batch_id: str) -> list[dict]:
    """按 batch_id 取所有 plan (列表 view 不含 batch_id, 故需逐个拉详情筛选)."""
    from urllib.parse import quote as _q
    lst = _rs_get(service, f"/plans?account={_q(account, safe='')}&limit=500")
    out = []
    for meta in lst:
        # 只看非 cancelled 的候选, 减少 detail 请求
        if meta.get("status") in ("cancelled", "received"):
            continue
        p = _rs_get(service, f"/plans/{meta['plan_id']}")
        if (p.get("stats") or {}).get("batch_id") == batch_id:
            out.append(p)
    return out


def assert_plans_approved(plans: list[dict]) -> None:
    bad = [(p["plan_id"][:12], p["status"]) for p in plans if p.get("status") != "approved"]
    if bad:
        lines = "\n".join(f"  - {pid}… status={st}" for pid, st in bad)
        raise RuntimeError(
            f"以下 plan 非 approved 状态, 不能发货:\n{lines}\n"
            f"请先在 Bitable 审批或手工 transition 到 approved"
        )


def build_cfg_from_plans(
    plans: list[dict],
    source_warehouse_keyword: str,
    drop_off_keyword: str,
) -> tuple[dict, int]:
    """把 4181 plans 合并成 create_fbo_plan 的 cfg 格式 + 返回 box_size.

    合并规则: 同 cluster 下多 SKU 直接并入 matrix[cluster][offer_id] = pieces.
    box_size: 所有 plan 必须一致 (实际都是 sku_profile 里 SKU 装箱率, 不同 SKU 同值概率小).
              混用会报错, 提示 CLI 覆盖.
    """
    matrix: dict[str, dict[str, int]] = {}
    box_sizes = set()
    for p in plans:
        sku = p.get("sku")  # offer_id
        if not sku:
            continue
        bs = int(p.get("box_size") or 0)
        if bs > 0:
            box_sizes.add(bs)
        for a in (p.get("allocations") or []):
            if int(a.get("pieces") or 0) <= 0:
                continue
            cl = a.get("cluster")
            if not cl:
                continue
            matrix.setdefault(cl, {})
            matrix[cl][sku] = matrix[cl].get(sku, 0) + int(a["pieces"])
    if not matrix:
        raise RuntimeError("plans 合并后 matrix 为空 (所有 allocations pieces=0?)")
    if len(box_sizes) > 1:
        raise RuntimeError(
            f"多个 plan 装箱率不一致: {sorted(box_sizes)}. "
            f"用 --box-size N 手工覆盖, 或拆批次分别跑"
        )
    bs = box_sizes.pop() if box_sizes else 0
    cfg = {
        "source_warehouse_keyword": source_warehouse_keyword,
        "drop_off_keyword": drop_off_keyword,
        "matrix": matrix,
    }
    return cfg, bs


def transition_plan_dispatched(
    service: str, plan_id: str, note: str = "", actor: str = "fbo-plan-auto",
) -> dict:
    """approved → dispatched. 失败不抛, 返 {ok, error}."""
    try:
        d = _rs_post(service, f"/plans/{plan_id}/transition", {
            "to_status": "dispatched", "actor": actor, "note": note or "create_fbo_plan 建单成功",
        })
        return {"ok": d.get("status") == "dispatched", "status": d.get("status"), "plan_id": plan_id}
    except Exception as e:
        return {"ok": False, "error": str(e), "plan_id": plan_id}


# ---------- FBO 服务调用 ----------

class FBOClient:
    def __init__(self, service: str, account: str) -> None:
        self.service = service.rstrip("/")
        self.account = account
        self.hdr = {
            "X-Ozon-Account": quote(account, safe=""),
            "Content-Type": "application/json",
        }

    def post(self, path: str, body: dict[str, Any] | None = None) -> dict[str, Any]:
        r = requests.post(
            f"{self.service}{path}", json=body or {}, headers=self.hdr, timeout=120,
        )
        if r.status_code >= 400:
            raise RuntimeError(f"FBO {path} HTTP {r.status_code}: {r.text[:500]}")
        return r.json()

    def get(self, path: str) -> dict[str, Any]:
        r = requests.get(f"{self.service}{path}", headers=self.hdr, timeout=60)
        if r.status_code >= 400:
            raise RuntimeError(f"FBO GET {path} {r.status_code}: {r.text[:300]}")
        return r.json()


# ---------- 轮询 helper ----------

def poll(
    fetch: Callable[[], dict[str, Any]],
    done: Callable[[dict[str, Any]], bool],
    *,
    interval: float = 3.0,
    timeout: float = 180.0,
    label: str = "",
) -> dict[str, Any]:
    t0 = time.monotonic()
    body: dict[str, Any] = {}
    while time.monotonic() - t0 < timeout:
        body = fetch()
        if done(body):
            return body
        time.sleep(interval)
    raise TimeoutError(f"poll {label} timeout after {timeout}s; last={str(body)[:300]}")


def retry(fn: Callable[[], dict[str, Any]], *, tries: int = 6, wait: float = 10.0,
          retry_if: Callable[[Exception], bool] = lambda _: True,
          label: str = "") -> dict[str, Any]:
    last: Exception | None = None
    for i in range(1, tries + 1):
        try:
            return fn()
        except Exception as e:
            if not retry_if(e):
                raise
            last = e
            print(f"    retry {label} {i}/{tries}: {str(e)[:120]}", flush=True)
            time.sleep(wait)
    raise last or RuntimeError("retry exhausted")


# ---------- 解析: cluster / sku / warehouse ----------

def resolve_cluster_macros(cli: FBOClient, names: list[str]) -> dict[str, dict]:
    """集群名 (中/俄 substring) → {id, macrolocal_cluster_id, name}."""
    resp = cli.post("/cluster/list", {"cluster_type": "CLUSTER_TYPE_OZON"})
    out: dict[str, dict] = {}
    for kw in names:
        hits = [c for c in (resp.get("clusters") or []) if kw in (c.get("name") or "")]
        if not hits:
            raise RuntimeError(f"未找到匹配集群 '{kw}'")
        c = hits[0]
        macro = c.get("macrolocal_cluster_id")
        if not macro:
            raise RuntimeError(f"集群 {c.get('name')} 缺 macrolocal_cluster_id")
        out[kw] = {"id": int(c["id"]), "macro": str(macro), "name": c.get("name","")}
    return out


def resolve_skus(cli: FBOClient, offer_ids: list[str]) -> dict[str, int]:
    resp = cli.post("/product/info/list", {"offer_id": list(set(offer_ids))})
    out: dict[str, int] = {}
    for it in (resp.get("items") or []):
        oid = it.get("offer_id")
        sku = None
        for s in (it.get("sources") or []):
            if s.get("sku"):
                sku = int(s["sku"]); break
        if oid and sku:
            out[oid] = sku
    missing = set(offer_ids) - set(out.keys())
    if missing:
        raise RuntimeError(f"offer_id 解析失败: {missing}")
    return out


def resolve_dropoff(cli: FBOClient, kw: str) -> dict:
    """drop-off warehouse 解析 (SORTING_CENTER / CROSS_DOCK)."""
    if len(kw) < 4:
        raise RuntimeError("Ozon 要求 search >= 4 字符")
    resp = cli.post("/warehouse/fbo/list",
                    {"filter_by_supply_type":["CREATE_TYPE_CROSSDOCK"], "search": kw})
    hits = resp.get("search") or []
    if not hits:
        raise RuntimeError(f"没找到 drop-off '{kw}'")
    w = hits[0]
    wt_str = (w.get("warehouse_type","") or "").replace("WAREHOUSE_TYPE_","")
    return {
        "id": int(w["warehouse_id"]), "name": w.get("name",""),
        "type_str": wt_str, "type_num": _WH_TYPE_NUM_FROM_STR.get(wt_str, 3),
    }


def resolve_seller_warehouse(cli: FBOClient, kw: str) -> dict | None:
    """seller source warehouse (通常是 FULL_FILLMENT РФЦ)."""
    if not kw or len(kw) < 4:
        return None
    resp = cli.post("/warehouse/fbo/list",
                    {"filter_by_supply_type":["CREATE_TYPE_DIRECT"], "search": kw})
    hits = resp.get("search") or []
    if not hits:
        return None
    w = hits[0]
    return {"id": int(w["warehouse_id"]), "name": w.get("name","")}


# ---------- 多集群 draft + 可用性探测 ----------

def try_multi_cluster(
    cli: FBOClient, macros: dict[str, dict], skus: dict[str, int],
    matrix: dict[str, dict[str, int]], dropoff: dict,
    seller_warehouse: dict | None,
) -> dict:
    """尝试建多集群 draft 并返回 {draft_id, availability_by_cluster_kw}."""
    clusters_info = []
    for kw, sku_qtys in matrix.items():
        m = macros[kw]
        items = [{"sku": skus[oid], "quantity": q}
                 for oid, q in sku_qtys.items() if q > 0]
        if not items:
            continue
        clusters_info.append({
            "macrolocal_cluster_id": int(m["macro"]), "items": items,
        })
    delivery_info = {
        "type": "DROPOFF",
        "drop_off_warehouse": {
            "warehouse_id": dropoff["id"], "warehouse_type": dropoff["type_str"],
        },
    }
    if seller_warehouse:
        delivery_info["seller_warehouse_id"] = seller_warehouse["id"]
    body = {
        "clusters_info": clusters_info,
        "deletion_sku_mode": "PARTIAL",
        "delivery_info": delivery_info,
    }
    print(f"  → POST /draft/multi-cluster/create ({len(clusters_info)} clusters)")
    resp = cli.post("/draft/multi-cluster/create", body)
    did = resp.get("draft_id")
    if not did:
        raise RuntimeError(f"multi-cluster draft 创建失败: {resp}")
    print(f"    draft_id = {did}")

    print(f"  → 轮询 /draft/create/info (v2)")
    info = retry(
        lambda: poll(
            fetch=lambda: cli.post("/draft/create/info", {"draft_id": int(did)}),
            done=lambda b: str(b.get("status","")).upper() in {"SUCCESS","FAILED","EXPIRED"},
            interval=5.0, timeout=180.0, label=f"draft {did}",
        ),
        label=f"create/info {did}",
    )
    status = str(info.get("status",""))
    print(f"    status = {status}")

    # 映射 cluster_kw → {state, bundle_id, ...}
    by_macro: dict[int, dict] = {}
    for c in (info.get("clusters") or []):
        mid = int(c.get("macrolocal_cluster_id", 0))
        ws = (c.get("warehouses") or [{}])
        w = ws[0] if ws else {}
        by_macro[mid] = {
            "state": (w.get("availability_status") or {}).get("state",""),
            "invalid_reason": (w.get("availability_status") or {}).get("invalid_reason",""),
            "bundle_id": w.get("bundle_id") or w.get("restricted_bundle_id") or "",
            "storage_warehouse": w.get("storage_warehouse") or {},
            "supply_type": c.get("supply_type","MULTI_CLUSTER"),
            "cluster_name": c.get("cluster_name",""),
        }

    availability: dict[str, dict] = {}
    for kw, m in macros.items():
        mid = int(m["macro"])
        availability[kw] = by_macro.get(mid, {"state":"UNKNOWN"})

    return {
        "draft_id": int(did),
        "status": status,
        "availability_by_cluster_kw": availability,
        "raw": info,
    }


# ---------- 单集群 CROSSDOCK fallback ----------

def create_single_cluster_crossdock(
    cli: FBOClient, cluster_macro: str, skus_qtys: dict[int, int],
    dropoff: dict,
) -> int:
    """返回 draft_id (同步)."""
    body = {
        "cluster_info": {
            "macrolocal_cluster_id": str(cluster_macro),
            "items": [{"sku": sku, "quantity": q} for sku, q in skus_qtys.items()],
        },
        "deletion_sku_mode": "PARTIAL",
        "delivery_info": {
            "type": "DROPOFF",
            "drop_off_warehouse": {
                "warehouse_id": dropoff["id"],
                "warehouse_type": dropoff["type_num"],
            },
        },
    }
    resp = cli.post("/draft/crossdock/create", body)
    did = resp.get("draft_id")
    if not did:
        raise RuntimeError(f"CROSSDOCK draft 失败: {resp}")
    return int(did)


# ---------- timeslot + supply create 通用 ----------

def find_timeslot(
    cli: FBOClient, draft_id: int, warehouse_id: int, macro: str,
    supply_type: str, days_from: int = 3, days_to: int = 28,
) -> dict:
    today = date.today()
    body = {
        "draft_id": draft_id,
        "selected_cluster_warehouses": [
            {"warehouse_id": str(warehouse_id), "macrolocal_cluster_id": str(macro)},
        ],
        "supply_type": supply_type,
        "date_from": (today + timedelta(days=days_from)).strftime("%Y-%m-%d"),
        "date_to":   (today + timedelta(days=days_to)).strftime("%Y-%m-%d"),
    }
    resp = retry(
        lambda: cli.post("/draft/timeslot/info", body),
        tries=6, wait=8.0,
        retry_if=lambda e: "FailedPrecondition" in str(e) or "can't find" in str(e).lower(),
        label="timeslot/info",
    )
    result = resp.get("result") or resp
    outer = result.get("drop_off_warehouse_timeslots") or {}
    if isinstance(outer, list):
        outer = outer[0] if outer else {}
    for day in (outer.get("days") or []):
        slots = day.get("timeslots") or []
        if slots:
            return slots[0]
    raise RuntimeError(f"无可用时段. raw={json.dumps(resp, ensure_ascii=False)[:300]}")


def create_supply_and_wait(
    cli: FBOClient, draft_id: int,
    cluster_warehouses: list[dict],
    supply_type: str, slot: dict,
) -> int:
    """返 order_id (v2 supply/create 返 draft_id 无 operation, 轮询 status by draft_id)."""
    body = {
        "draft_id": draft_id,
        "selected_cluster_warehouses": cluster_warehouses,
        "supply_type": supply_type,
        "timeslot": {
            "from_in_timezone": slot["from_in_timezone"].rstrip("Z"),
            "to_in_timezone":   slot["to_in_timezone"].rstrip("Z"),
        },
    }
    resp = cli.post("/draft/supply/create", body)
    errs = resp.get("error_reasons") or []
    if errs:
        raise RuntimeError(f"supply/create 失败: {errs}; raw={resp}")

    final = poll(
        fetch=lambda: cli.post("/draft/supply/create/status", {"draft_id": draft_id}),
        done=lambda b: str(b.get("status","")).upper() in {"SUCCESS","FAILED","ERROR"},
        interval=4.0, timeout=180.0, label=f"supply status draft={draft_id}",
    )
    if str(final.get("status","")).upper() != "SUCCESS":
        raise RuntimeError(f"supply 建单失败: {json.dumps(final, ensure_ascii=False)[:300]}")
    oid = final.get("order_id")
    if not oid:
        raise RuntimeError(f"status=SUCCESS 但无 order_id: {final}")
    return int(oid)


# ---------- cargo + label ----------

def split_into_cargoes(barcode: str, qty: int, box_size: int) -> list[dict]:
    n = math.ceil(qty / box_size) if box_size > 0 else 0
    out: list[dict] = []
    remaining = qty
    for i in range(n):
        in_box = box_size if remaining >= box_size else remaining
        remaining -= in_box
        out.append({
            "key": str(i+1),
            "items":[{"barcode": barcode, "quantity": in_box}],
            "type":"BOX",
        })
    return out


def fetch_supply_actual_items(cli: FBOClient, order_detail: dict) -> dict[int, int]:
    """用 /supply-order/bundle 查 supply 实际含的 items (Ozon 矩阵可能悄悄过滤某些 SKU).
    返 {sku: quantity}.
    """
    supplies = order_detail.get("supplies") or []
    if not supplies:
        return {}
    bid = supplies[0].get("bundle_id")
    if not bid:
        return {}
    resp = cli.post("/supply-order/bundle", {"bundle_ids": [bid], "limit": 100})
    out: dict[int, int] = {}
    for it in (resp.get("items") or []):
        sku = it.get("sku")
        q = it.get("quantity", 0)
        if sku and q > 0:
            out[int(sku)] = out.get(int(sku), 0) + int(q)
    return out


def fill_boxes_and_label(
    cli: FBOClient, order_id: int, order_detail: dict,
    offer_id_by_sku: dict[int, str],  # sku→offer_id
    qty_by_sku: dict[int, int],       # sku→total qty (实际会按 /supply-order/bundle 覆盖)
    box_size: int,
) -> dict:
    """一个 order 可能含多 SKU; 按 SKU 切箱然后合并发 /flow/upload-cargoes.

    关键: Ozon 矩阵会在建 supply 时悄悄过滤掉某些 SKU (单集群/多集群都可能).
    所以必须先查 /supply-order/bundle 拿 supply 里**实际**的 items, 按实际切箱.
    不然 /v1/cargoes/create 会返 SUPPLY_ITEM_NOT_FOUND.
    """
    supply_id = int((order_detail.get("supplies") or [{}])[0]["supply_id"])

    # 按 Ozon 实际接受的 items 切箱 (覆盖原计划)
    actual = fetch_supply_actual_items(cli, order_detail)
    if actual:
        dropped = {sku: q for sku, q in qty_by_sku.items()
                   if sku not in actual or actual[sku] < q}
    else:
        actual = dict(qty_by_sku)
        dropped = {}

    cargoes: list[dict] = []
    running_key = 1
    for sku, qty in actual.items():
        barcode = f"OZN{sku}"
        chunks = split_into_cargoes(barcode, qty, box_size)
        for c in chunks:
            c["key"] = str(running_key); running_key += 1
            cargoes.append(c)

    if not cargoes:
        return {
            "supply_id": supply_id, "pdf_path": "", "file_url": "",
            "cargo_ids": [], "stage": "no_items",
            "actual_qty_by_sku": actual, "dropped_qty_by_sku": dropped,
        }

    resp = cli.post("/flow/upload-cargoes", {
        "supply_id": supply_id, "cargoes": cargoes,
        "delete_current_version": True, "generate_label": True,
        "save_label_pdf": True, "return_pdf_base64": False,
        "poll_interval": 3.0, "poll_timeout": 300.0,
    })
    # cargo_ids per key
    cargo_ids_by_key = {}
    for c in (resp.get("cargoes") or []):
        cargo_ids_by_key[c.get("key")] = (c.get("value") or {}).get("cargo_id")
    return {
        "supply_id": supply_id,
        "pdf_path": resp.get("pdf_path") or "",
        "file_url": resp.get("file_url") or "",
        "cargo_ids": [cargo_ids_by_key.get(c["key"]) for c in cargoes],
        "stage": resp.get("stage",""),
        "cargoes_sent": cargoes,
        "actual_qty_by_sku": actual,
        "dropped_qty_by_sku": dropped,
    }


# ---------- 对照表 ----------

_HEADERS = [
    "交货 ID","供货 ID","状态","链路","集群","macrolocal_cluster_id",
    "货位 ID","货位名称","drop-off ID","drop-off 名称",
    "货号 (offer_id)","SKU","条码",
    "箱号","cargo_id","该箱件数",
    "总件数","总箱数","单箱装箱率",
    "时段",
    "箱唛 PDF 路径",
    "备注/错误",
]


def write_excel(rows: list[dict], out_path: Path) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "FBO 对照表"

    ws.append(_HEADERS)
    bold = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F4E78")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for c in range(1, len(_HEADERS)+1):
        cell = ws.cell(row=1, column=c); cell.font=bold; cell.fill=fill; cell.alignment=center

    for r in rows:
        ws.append([r.get(h, "") for h in _HEADERS])

    widths = [12,14,18,10,32,12,18,22,16,22,28,14,16,6,18,10,10,8,10,28,50,30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def rows_per_box(
    *, order_id: int, order_detail: dict, mode: str,
    cluster_kw: str, cluster_name: str, macrolocal: str,
    dropoff: dict, sku_offer: dict[int, str],
    qty_by_sku: dict[int, int], box_size: int,
    fill_result: dict | None, note: str = "",
) -> list[dict]:
    drop_off_info = (order_detail.get("drop_off_warehouse") or {})
    supply = (order_detail.get("supplies") or [{}])[0]
    storage = supply.get("storage_warehouse") or {}
    ts = (order_detail.get("timeslot") or {}).get("timeslot") or {}
    ts_str = f"{ts.get('from','')} ~ {ts.get('to','')}"
    supply_id = supply.get("supply_id","")

    # flatten all (sku, box_idx) 组合
    flat_boxes: list[tuple[int, int, int]] = []  # (sku, box_1idx, qty_in_box)
    for sku, qty in qty_by_sku.items():
        n = math.ceil(qty / box_size) if box_size > 0 else 0
        remain = qty
        for i in range(n):
            in_box = box_size if remain >= box_size else remain
            remain -= in_box
            flat_boxes.append((sku, i+1, in_box))

    cargo_ids = (fill_result or {}).get("cargo_ids") or []
    pdf_path = (fill_result or {}).get("pdf_path","")

    rows: list[dict] = []
    global_idx = 0
    for sku, box_i, q in flat_boxes:
        cid = cargo_ids[global_idx] if global_idx < len(cargo_ids) else ""
        rows.append({
            "交货 ID": order_id,
            "供货 ID": supply_id,
            "状态": order_detail.get("state",""),
            "链路": mode,   # "MULTI_CLUSTER" / "CROSSDOCK"
            "集群": cluster_name or cluster_kw,
            "macrolocal_cluster_id": macrolocal,
            "货位 ID": storage.get("warehouse_id",""),
            "货位名称": storage.get("name",""),
            "drop-off ID": drop_off_info.get("warehouse_id","") or dropoff.get("id",""),
            "drop-off 名称": drop_off_info.get("name","") or dropoff.get("name",""),
            "货号 (offer_id)": sku_offer.get(sku, ""),
            "SKU": sku,
            "条码": f"OZN{sku}",
            "箱号": box_i + (global_idx - (box_i - 1)),   # 全局箱号, 同 order 顺序递增
            "cargo_id": cid,
            "该箱件数": q,
            "总件数": qty_by_sku.get(sku, 0),
            "总箱数": len(flat_boxes),
            "单箱装箱率": box_size,
            "时段": ts_str,
            "箱唛 PDF 路径": pdf_path,
            "备注/错误": note if global_idx == 0 else "",
        })
        global_idx += 1
    return rows


# ---------- 主流程 ----------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Ozon FBO SOP: multi-cluster 优先 + 不可发 fallback 单集群",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    src = p.add_mutually_exclusive_group(required=True)
    src.add_argument("--config", default="",
                     help="plan JSON: {source_warehouse_keyword, drop_off_keyword, matrix}")
    src.add_argument("--plan-id", default="",
                     help="4181 plan_id (逗号分隔多个) — 必须都是 approved, 发完自动 dispatched")
    src.add_argument("--batch-id", default="",
                     help="4181 stats.batch_id — 拉批次内全部 approved plans 合并发")

    p.add_argument("--box-size", type=int, default=0,
                   help="单箱装箱率 (plan-id/batch-id 模式自动用 plan.box_size; 非 0 覆盖)")
    p.add_argument("--account", default="丝绸生活")
    p.add_argument("--fbo-service", default=DEFAULT_FBO_SERVICE)
    p.add_argument("--restock-service", default=DEFAULT_RESTOCK_SERVICE,
                   help="4181 服务地址 (plan-id/batch-id 模式用)")
    p.add_argument("--source-warehouse-keyword", default="ЖУКОВСКИЙ_РФЦ",
                   help="发货仓关键字 (plan-id/batch-id 模式必传; config 里有则覆盖)")
    p.add_argument("--drop-off-keyword", default="ЩЕРБИНКА",
                   help="drop-off 关键字 (plan-id/batch-id 模式必传; config 里有则覆盖)")
    p.add_argument("--skip-dispatch-transition", action="store_true",
                   help="建单成功后不把 plan 推到 dispatched (用于调试)")
    p.add_argument("--out-xlsx", default="",
                   help="对照表输出路径, 默认 fbo_plan_<date>.xlsx")
    p.add_argument("--dry-run", action="store_true",
                   help="只走到 availability 探测 + 打印分拆方案, 不建单")
    p.add_argument("--skip-fill", action="store_true",
                   help="建完 supply_order 就停, 不申报装箱/不拉标签")
    p.add_argument("--skip-alloc-check", action="store_true",
                   help="跳过 allocation_xlsx vs cfg.matrix 件数 sanity check (谨慎用)")
    p.add_argument("--days-from", type=int, default=3)
    p.add_argument("--days-to", type=int, default=28)
    return p.parse_args()


def main() -> int:
    args = parse_args()

    # ---- 三种来源解析 ----
    rs_plans: list[dict] = []          # 4181 plan 详情 (plan-id/batch-id 模式)
    rs_plan_ids: list[str] = []
    rs_batch_id: str = ""
    config_source: str = ""

    if args.config:
        cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
        config_source = str(Path(args.config).resolve())
    elif args.plan_id or args.batch_id:
        if args.plan_id:
            rs_plan_ids = [x.strip() for x in args.plan_id.split(",") if x.strip()]
            print(f"[rs] 从 4181 拉 {len(rs_plan_ids)} 个 plan: {[p[:12]+'…' for p in rs_plan_ids]}")
            rs_plans = fetch_plans_by_ids(args.restock_service, rs_plan_ids)
        else:
            rs_batch_id = args.batch_id.strip()
            print(f"[rs] 从 4181 按 batch_id={rs_batch_id[:12]}… 筛选 approved plans")
            rs_plans = fetch_plans_by_batch(args.restock_service, args.account, rs_batch_id)
            rs_plan_ids = [p["plan_id"] for p in rs_plans]
            if not rs_plans:
                print(f"  ! batch_id 未匹配任何 plan"); return 2
            print(f"  → 命中 {len(rs_plans)} 个 plan")
        assert_plans_approved(rs_plans)
        cfg, bs_from_plan = build_cfg_from_plans(
            rs_plans, args.source_warehouse_keyword, args.drop_off_keyword,
        )
        if args.box_size == 0:
            if bs_from_plan <= 0:
                print(f"  ! plan 缺 box_size, 用 --box-size N 手工传入"); return 2
            args.box_size = bs_from_plan
            print(f"[rs] box_size 取自 plan: {args.box_size}")
        else:
            print(f"[rs] box_size CLI 覆盖: {args.box_size} (plan 里是 {bs_from_plan})")
        config_source = (
            f"4181://plans?ids={','.join(p[:12] for p in rs_plan_ids)}"
            + (f"&batch_id={rs_batch_id[:12]}" if rs_batch_id else "")
        )
    else:
        # 互斥组已保证不会走到这里
        print("[ERROR] 必须传 --config 或 --plan-id 或 --batch-id", file=sys.stderr); return 2

    if args.box_size <= 0:
        print(f"[ERROR] --box-size 必须 > 0"); return 2

    # ---- allocation cap 校验 (防上游 raw 需求绕过 allocation 限制泄漏到 supply) ----
    if not args.skip_alloc_check:
        caps = _read_allocation_caps(SCRIPT_DIR)
        if caps:
            offers_in_cfg = {o for r in cfg.get("matrix", {}).values() for o in r.keys()}
            covered = sum(1 for o in offers_in_cfg if o in caps)
            print(f"[alloc] 校验 cfg.matrix 不超 allocation 合计件 (已知 {covered}/{len(offers_in_cfg)} offer)")
            violations = assert_cfg_within_alloc(cfg, caps)
            if violations:
                print(f"[ERROR] allocation cap 违规 ({len(violations)} 条):", file=sys.stderr)
                for v in violations:
                    print(f"  - {v}", file=sys.stderr)
                print("[hint] 本次 cfg 件数超 allocation 决策的 cap, 真发可能造成超额建单 (见 supply 100580300 教训).", file=sys.stderr)
                print("       如果确认要发, 加 --skip-alloc-check 强制覆盖.", file=sys.stderr)
                return 2
        else:
            print(f"[alloc] 未找到 allocation_*.xlsx, 跳过 cap 校验")

    cli = FBOClient(args.fbo_service, args.account)

    # 台账: 记录本次所有 shipped / skipped
    ledger = {
        "run_at": _now_iso(),
        "account": args.account,
        "config_source": config_source,
        "restock_plan_ids": rs_plan_ids,   # 4181 plan_id 列表 (config 模式为空)
        "restock_batch_id": rs_batch_id,
        "original_matrix": cfg.get("matrix", {}),
        "source_warehouse_keyword": cfg.get("source_warehouse_keyword",""),
        "drop_off_keyword": cfg.get("drop_off_keyword",""),
        "box_size": args.box_size,
        "shipped": [],   # [{cluster, macrolocal, mode, order_id, supply_id, items:[{offer_id,sku,qty,boxes,cargo_ids,plan_id?}], pdf_path, ...}]
        "skipped": [],   # [{cluster, offer_id, sku, quantity, reason, reason_code, plan_id?, related_order_id, timestamp}]
    }

    print(f"[0/6] 健康检查 & 预解析")
    health = cli.get("/health")
    if args.account not in (health.get("accounts") or []):
        print(f"  ! 账号 '{args.account}' 不在 .env 里", file=sys.stderr); return 2

    matrix: dict[str, dict[str, int]] = cfg["matrix"]
    cluster_kws = list(matrix.keys())
    offer_ids = sorted({oid for r in matrix.values() for oid in r.keys()})

    print(f"[1/6] 解析 clusters / SKUs / drop-off / seller warehouse")
    macros = resolve_cluster_macros(cli, cluster_kws)
    for kw, m in macros.items():
        print(f"  cluster {kw} → macro={m['macro']} name={m['name']}")

    skus = resolve_skus(cli, offer_ids)
    for oid, sku in skus.items():
        print(f"  offer_id {oid} → sku={sku}")

    dropoff = resolve_dropoff(cli, cfg["drop_off_keyword"])
    print(f"  drop-off {dropoff['name']} id={dropoff['id']} type={dropoff['type_str']}")

    seller_wh = resolve_seller_warehouse(cli, cfg.get("source_warehouse_keyword",""))
    if seller_wh:
        print(f"  seller source: {seller_wh['name']} id={seller_wh['id']}")
    else:
        print(f"  seller source: 未指定/未找到 (可选)")

    # 反向映射 sku→offer_id
    offer_by_sku = {v: k for k, v in skus.items()}

    # ========== multi-cluster 尝试 ==========
    print(f"\n[2/6] 尝试 multi-cluster 一票货件")
    try:
        mc = try_multi_cluster(cli, macros, skus, matrix, dropoff, seller_wh)
    except Exception as e:
        print(f"  ! multi-cluster 建 draft 失败: {e}")
        mc = None

    multi_ok: list[str] = []       # cluster_kw list that went into multi-cluster
    single_fallback: list[str] = []
    if mc:
        for kw, avail in mc["availability_by_cluster_kw"].items():
            state = avail.get("state","")
            reason = avail.get("invalid_reason","")
            print(f"  {kw}: state={state} reason={reason}")
            # 规则 (2026-04-24 用户确认):
            # - AVAILABLE → 保留进 multi-cluster
            # - PARTIAL_AVAILABLE → fallback 到单集群 CROSSDOCK (单集群矩阵更宽松)
            #   因为我们无法精确识别哪些 SKU 属于 bundle_id 可发组, 强发会被 supply/create 拒
            # - NOT_AVAILABLE → fallback 单集群
            if state in ("AVAILABLE", "FULL_AVAILABLE"):
                multi_ok.append(kw)
            else:
                single_fallback.append(kw)
    else:
        # full fallback
        single_fallback = list(cluster_kws)

    # 规则 2: multi-cluster 至少 2 个集群才有意义; 剩 1 个也降级成 CROSSDOCK
    if len(multi_ok) < 2:
        print(f"  (multi-cluster 只剩 {len(multi_ok)} 个集群, 降级全部走 CROSSDOCK)")
        single_fallback.extend(multi_ok)
        multi_ok = []

    print(f"\n[3/6] 分拆方案:")
    print(f"  多集群一票: {multi_ok or '—'}")
    print(f"  单集群 fallback: {single_fallback or '—'}")

    if args.dry_run:
        print("\n[dry-run] 不建单. 去掉 --dry-run 再跑.")
        return 0

    # ========== 建单 ==========
    created: list[dict] = []

    # multi-cluster supply
    if multi_ok and mc:
        print(f"\n[4/6] multi-cluster supply 建单 ({len(multi_ok)} 集群)")
        scw = []
        for kw in multi_ok:
            scw.append({
                "warehouse_id": str(dropoff["id"]),
                "macrolocal_cluster_id": macros[kw]["macro"],
            })
        try:
            slot = find_timeslot(
                cli, mc["draft_id"], dropoff["id"], macros[multi_ok[0]]["macro"],
                "MULTI_CLUSTER", args.days_from, args.days_to,
            )
            print(f"  timeslot: {slot.get('from_in_timezone')} ~ {slot.get('to_in_timezone')}")
            order_id = create_supply_and_wait(cli, mc["draft_id"], scw, "MULTI_CLUSTER", slot)
            print(f"  order_id = {order_id}")
            created.append({
                "order_id": order_id, "mode": "MULTI_CLUSTER",
                "cluster_kws": multi_ok,
            })
        except Exception as e:
            print(f"  ! multi-cluster supply 失败: {e}")
            # 如果 supply 建失败, 全部 fallback
            print(f"  → 全部转单集群 fallback")
            single_fallback = list(cluster_kws)

    # 单集群 fallback
    for kw in single_fallback:
        print(f"\n[5/6] 单集群 CROSSDOCK fallback: {kw}")
        skus_qtys = {skus[oid]: q for oid, q in matrix[kw].items() if q > 0}
        if not skus_qtys:
            continue
        try:
            did = create_single_cluster_crossdock(
                cli, macros[kw]["macro"], skus_qtys, dropoff,
            )
            print(f"  draft_id = {did}")
            slot = find_timeslot(
                cli, did, dropoff["id"], macros[kw]["macro"],
                "CROSSDOCK", args.days_from, args.days_to,
            )
            print(f"  timeslot: {slot.get('from_in_timezone')} ~ {slot.get('to_in_timezone')}")
            order_id = create_supply_and_wait(
                cli, did,
                [{"warehouse_id": str(dropoff["id"]),
                  "macrolocal_cluster_id": macros[kw]["macro"]}],
                "CROSSDOCK", slot,
            )
            print(f"  order_id = {order_id}")
            created.append({
                "order_id": order_id, "mode": "CROSSDOCK",
                "cluster_kws": [kw],
            })
        except Exception as e:
            err_str = str(e)
            print(f"  ! {kw} CROSSDOCK 失败: {err_str}")
            # 记台账: 整个集群全部 SKU 都 skipped
            reason_code = "NO_AVAILABLE_WAREHOUSE" if "warehouse scoring" in err_str or "NotFound" in err_str \
                else "DRAFT_OR_SUPPLY_FAILED"
            for oid, q in matrix[kw].items():
                if q <= 0: continue
                ledger["skipped"].append({
                    "cluster": macros[kw]["name"], "cluster_kw": kw,
                    "macrolocal_cluster_id": macros[kw]["macro"],
                    "offer_id": oid, "sku": skus.get(oid, 0),
                    "barcode": f"OZN{skus.get(oid, '')}",
                    "quantity": q,
                    "reason": err_str[:300],
                    "reason_code": reason_code,
                    "timestamp": _now_iso(),
                })
            created.append({
                "order_id": 0, "mode": "CROSSDOCK", "cluster_kws": [kw],
                "error": err_str,
            })

    if not created:
        print("\n! 一个单也没建起来")
        # 还是把 ledger 写出来
        lpath = LEDGER_DIR / f"shipment_ledger_{_now_utc_stamp()}.json"
        write_ledger(ledger, lpath)
        print(f"  ledger: {lpath}")
        return 2

    # ========== 装箱 + 标签 + 对照表 ==========
    print(f"\n[6/6] 装箱 + 标签 + 对照表 ({len(created)} 个 orders)")
    all_rows: list[dict] = []
    for row in created:
        if row.get("error"):
            # 失败 row 仍写到表里
            for kw in row["cluster_kws"]:
                for oid, q in matrix[kw].items():
                    all_rows.append({
                        "交货 ID": 0, "链路": row["mode"], "集群": kw,
                        "货号 (offer_id)": oid, "SKU": skus.get(oid,""),
                        "条码": f"OZN{skus.get(oid,'')}",
                        "总件数": q, "备注/错误": row["error"],
                    })
            continue

        order_id = row["order_id"]
        detail_resp = cli.post("/supply-order/get", {"order_ids":[order_id]})
        od = (detail_resp.get("orders") or [{}])[0]

        # merge all SKUs of all clusters in this order
        qty_by_sku: dict[int, int] = {}
        for kw in row["cluster_kws"]:
            for oid, q in matrix[kw].items():
                qty_by_sku[skus[oid]] = qty_by_sku.get(skus[oid], 0) + q

        fill_r = None
        note = ""
        if not args.skip_fill:
            try:
                fill_r = fill_boxes_and_label(
                    cli, order_id, od, offer_by_sku, qty_by_sku, args.box_size,
                )
                print(f"  order {order_id}: PDF {fill_r.get('pdf_path')}")
            except Exception as e:
                note = f"fill/label 失败: {e}"
                print(f"  ! order {order_id}: {note}")

        # 台账: shipped 行 (每个 order 一条, items 列 Ozon 实际接受的)
        actual = (fill_r or {}).get("actual_qty_by_sku") or {}
        dropped = (fill_r or {}).get("dropped_qty_by_sku") or {}
        supply = (od.get("supplies") or [{}])[0]
        shipped_entry = {
            "order_id": order_id, "mode": row["mode"],
            "supply_id": supply.get("supply_id",""),
            "cluster_kws": row["cluster_kws"],
            "cluster_names": [macros[kw]["name"] for kw in row["cluster_kws"]],
            "macrolocal_cluster_ids": [macros[kw]["macro"] for kw in row["cluster_kws"]],
            "drop_off": {
                "id": (od.get("drop_off_warehouse") or {}).get("warehouse_id",""),
                "name": (od.get("drop_off_warehouse") or {}).get("name",""),
            },
            "storage_warehouse": {
                "id": supply.get("storage_warehouse",{}).get("warehouse_id",""),
                "name": supply.get("storage_warehouse",{}).get("name",""),
            },
            "timeslot": od.get("timeslot",{}).get("timeslot") or {},
            "state": od.get("state",""),
            "items": [
                {
                    "offer_id": offer_by_sku.get(sku, ""), "sku": int(sku),
                    "barcode": f"OZN{sku}", "quantity": int(q),
                    "boxes": math.ceil(q / args.box_size) if args.box_size else 0,
                    "cargo_ids": [cid for cid in ((fill_r or {}).get("cargo_ids") or []) if cid],
                }
                for sku, q in actual.items()
            ],
            "pdf_path": (fill_r or {}).get("pdf_path",""),
            "pdf_file_url": (fill_r or {}).get("file_url",""),
            "fill_note": note,
            "timestamp": _now_iso(),
        }
        ledger["shipped"].append(shipped_entry)
        # skipped: dropped by Ozon 矩阵
        for sku, q_missing in dropped.items():
            # 找 cluster: row 下 cluster_kws 里哪个含这个 sku (原计划)
            for kw in row["cluster_kws"]:
                for oid, qp in matrix[kw].items():
                    if skus.get(oid) == sku and qp > 0:
                        # 实际接受数 (如有)
                        accepted = actual.get(sku, 0)
                        if accepted < qp:
                            ledger["skipped"].append({
                                "cluster": macros[kw]["name"], "cluster_kw": kw,
                                "macrolocal_cluster_id": macros[kw]["macro"],
                                "offer_id": oid, "sku": sku,
                                "barcode": f"OZN{sku}",
                                "quantity": qp - accepted,
                                "reason": "Ozon 矩阵在建 supply 时静默过滤, 未进入实际 items",
                                "reason_code": "MATRIX_DROPPED_FROM_SUPPLY",
                                "related_order_id": order_id,
                                "related_supply_id": supply.get("supply_id",""),
                                "timestamp": _now_iso(),
                            })
                        break

        # 输出 rows 按 (集群×SKU) 独立行 — 链路=MULTI_CLUSTER 时 1 个 order 多集群, 每集群单独起
        if row["mode"] == "MULTI_CLUSTER":
            # 按 (cluster, sku) 切箱, 全 order 共享一个 PDF
            global_cargo_idx = 0
            cargo_ids = (fill_r or {}).get("cargo_ids") or []
            for kw in row["cluster_kws"]:
                for oid, q in matrix[kw].items():
                    if q == 0: continue
                    sku = skus[oid]
                    box_n = math.ceil(q / args.box_size)
                    remain = q
                    for bi in range(box_n):
                        in_box = args.box_size if remain >= args.box_size else remain
                        remain -= in_box
                        cid = cargo_ids[global_cargo_idx] if global_cargo_idx < len(cargo_ids) else ""
                        all_rows.append({
                            "交货 ID": order_id,
                            "供货 ID": (od.get("supplies") or [{}])[0].get("supply_id",""),
                            "状态": od.get("state",""),
                            "链路": "MULTI_CLUSTER",
                            "集群": macros[kw]["name"],
                            "macrolocal_cluster_id": macros[kw]["macro"],
                            "货位 ID": (od.get("supplies") or [{}])[0].get("storage_warehouse",{}).get("warehouse_id",""),
                            "货位名称": (od.get("supplies") or [{}])[0].get("storage_warehouse",{}).get("name",""),
                            "drop-off ID": (od.get("drop_off_warehouse") or {}).get("warehouse_id",""),
                            "drop-off 名称": (od.get("drop_off_warehouse") or {}).get("name",""),
                            "货号 (offer_id)": oid,
                            "SKU": sku, "条码": f"OZN{sku}",
                            "箱号": global_cargo_idx + 1,
                            "cargo_id": cid,
                            "该箱件数": in_box,
                            "总件数": q,
                            "总箱数": box_n,
                            "单箱装箱率": args.box_size,
                            "时段": f"{(od.get('timeslot',{}).get('timeslot') or {}).get('from','')} ~ {(od.get('timeslot',{}).get('timeslot') or {}).get('to','')}",
                            "箱唛 PDF 路径": (fill_r or {}).get("pdf_path",""),
                            "备注/错误": note if global_cargo_idx == 0 else "",
                        })
                        global_cargo_idx += 1
        else:
            # single-cluster: cluster_kws 只有 1 个
            kw = row["cluster_kws"][0]
            global_cargo_idx = 0
            cargo_ids = (fill_r or {}).get("cargo_ids") or []
            for oid, q in matrix[kw].items():
                if q == 0: continue
                sku = skus[oid]
                box_n = math.ceil(q / args.box_size)
                remain = q
                for bi in range(box_n):
                    in_box = args.box_size if remain >= args.box_size else remain
                    remain -= in_box
                    cid = cargo_ids[global_cargo_idx] if global_cargo_idx < len(cargo_ids) else ""
                    all_rows.append({
                        "交货 ID": order_id,
                        "供货 ID": (od.get("supplies") or [{}])[0].get("supply_id",""),
                        "状态": od.get("state",""),
                        "链路": "CROSSDOCK",
                        "集群": macros[kw]["name"],
                        "macrolocal_cluster_id": macros[kw]["macro"],
                        "货位 ID": (od.get("supplies") or [{}])[0].get("storage_warehouse",{}).get("warehouse_id",""),
                        "货位名称": (od.get("supplies") or [{}])[0].get("storage_warehouse",{}).get("name",""),
                        "drop-off ID": (od.get("drop_off_warehouse") or {}).get("warehouse_id",""),
                        "drop-off 名称": (od.get("drop_off_warehouse") or {}).get("name",""),
                        "货号 (offer_id)": oid,
                        "SKU": sku, "条码": f"OZN{sku}",
                        "箱号": global_cargo_idx + 1,
                        "cargo_id": cid,
                        "该箱件数": in_box,
                        "总件数": q,
                        "总箱数": box_n,
                        "单箱装箱率": args.box_size,
                        "时段": f"{(od.get('timeslot',{}).get('timeslot') or {}).get('from','')} ~ {(od.get('timeslot',{}).get('timeslot') or {}).get('to','')}",
                        "箱唛 PDF 路径": (fill_r or {}).get("pdf_path",""),
                        "备注/错误": note if global_cargo_idx == 0 else "",
                    })
                    global_cargo_idx += 1

    from datetime import datetime as _dt
    _now_tag = _dt.now().strftime("%Y%m%d_%H%M%S")  # 带时分秒, 避免同日多次 run 互相覆盖

    out = args.out_xlsx
    if not out:
        out = str(SCRIPT_DIR / f"fbo_plan_{_now_tag}.xlsx")
    out_path = write_excel(all_rows, Path(out))

    # 写台账 (shipped + skipped) + 更新 pending_summary
    tag = _now_tag
    ledger_path = LEDGER_DIR / f"shipment_ledger_{tag}.json"
    write_ledger(ledger, ledger_path)
    pending_path = LEDGER_DIR / "pending_summary.json"
    ps = merge_pending(LEDGER_DIR, pending_path)
    print(f"\n[ledger] 本次: {ledger_path}")
    print(f"         shipped={len(ledger['shipped'])} skipped={len(ledger['skipped'])}")
    print(f"[pending] 累计未发: {pending_path} ({ps['pending_count']} 条)")
    print(f"\n[done] 对照表: {out_path} ({len(all_rows)} 行)")

    # ---- 4181 plan 状态推进: approved → dispatched (仅 plan-id/batch-id 模式) ----
    if rs_plans and not args.skip_dispatch_transition and not args.dry_run:
        # 按 plan.sku 找到它在本次 run 里有没有 shipped
        shipped_offers: set[str] = set()
        for s in ledger["shipped"]:
            for it in (s.get("items") or []):
                if it.get("quantity", 0) > 0 and it.get("offer_id"):
                    shipped_offers.add(it["offer_id"])
        print(f"\n[rs-transition] 本次 run 发出 SKU: {sorted(shipped_offers) or '(无)'}")
        for p in rs_plans:
            pid = p.get("plan_id")
            sku = p.get("sku")
            if sku in shipped_offers:
                note = f"fbo_plan run {tag}: ledger={ledger_path.name}"
                r = transition_plan_dispatched(args.restock_service, pid, note=note)
                status = "✓" if r.get("ok") else "✗"
                print(f"  {status} {pid[:12]}… ({sku:30}) → dispatched  {r.get('error','')}")
            else:
                print(f"  ~ {pid[:12]}… ({sku:30}) 未发任何件, 保留 approved (可 /fbo-retry 后续)")
    elif rs_plans and args.skip_dispatch_transition:
        print(f"\n[rs-transition] 跳过 (--skip-dispatch-transition): 不推进 plan 状态")
    elif rs_plans and args.dry_run:
        print(f"\n[rs-transition] dry-run 模式, 不推进 plan 状态")

    # 摘要
    print("\n=== 摘要 (按 order 聚合) ===")
    by_order: dict[Any, list[dict]] = {}
    for r in all_rows:
        by_order.setdefault(r.get("交货 ID"), []).append(r)
    for oid, rs in by_order.items():
        r0 = rs[0]
        print(f"  交货={oid} 链路={r0.get('链路')} "
              f"集群={r0.get('集群')} 货位={r0.get('货位 ID','')} "
              f"箱={len(rs)} 备注={r0.get('备注/错误','')}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
