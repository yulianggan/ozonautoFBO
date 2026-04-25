#!/usr/bin/env python3
"""Ozon FBO 越库发货一键脚本 (借道 localhost:4182 FBO shipment service).

典型用法:
    # dry-run 看选仓和时段, 不实际建单
    python3 create_moscow_shipment.py --cluster Москва --dropoff ЩЕРБИНКА \
        --sku 3214793665 --qty 4200 --dry-run

    # 真实创建 (会在 Ozon 后台产生一条 FBO 供货请求)
    python3 create_moscow_shipment.py --cluster Москва --dropoff ЩЕРБИНКА \
        --sku 3214793665 --qty 4200

选仓策略: 在指定集群内, 筛 is_available=true, 按 travel_time_days 升序取第一.
选时段策略: 对选中的仓查 date_from~date_to 区间, 取最早有 slot 的日的第 N 个时段.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import date, timedelta
from typing import Any, Callable
from urllib.parse import quote

import requests


SUPPLY_TYPE_NUM: dict[str, int] = {
    "CREATE_TYPE_CROSSDOCK": 1,
    "CREATE_TYPE_DIRECT": 2,
    "CREATE_TYPE_MULTI_CLUSTER": 3,
}


# ---------- HTTP + polling helpers ----------

_ARGS: argparse.Namespace  # 全局, _post 用


def _post(path: str, body: dict[str, Any]) -> dict[str, Any]:
    try:
        r = requests.post(
            f"{_ARGS.api}{path}",
            json=body,
            headers={"X-Ozon-Account": quote(_ARGS.account, safe="")},
            timeout=60,
        )
    except requests.RequestException as e:
        print(f"[FATAL] 连接 {_ARGS.api}{path} 失败: {e}", file=sys.stderr)
        print("  -> 是否忘了起服务? cd /Users/mac/Documents/ozns/github && "
              "python3 -m uvicorn ozon_fbo_shipment_service.app:app --port 4182",
              file=sys.stderr)
        sys.exit(2)
    if r.status_code >= 400:
        print(f"[API ERROR] {path} -> HTTP {r.status_code}", file=sys.stderr)
        print(f"  body: {r.text[:500]}", file=sys.stderr)
        sys.exit(2)
    return r.json()


def _get(path: str) -> dict[str, Any]:
    r = requests.get(f"{_ARGS.api}{path}", timeout=30)
    if r.status_code >= 400:
        print(f"[API ERROR] GET {path} -> {r.status_code}: {r.text[:300]}",
              file=sys.stderr)
        sys.exit(2)
    return r.json()


def _poll(
    fetch: Callable[[], dict[str, Any]],
    is_done: Callable[[dict[str, Any]], bool],
    *,
    interval: float = 2.0,
    timeout: float = 180.0,
    step_name: str = "",
) -> dict[str, Any]:
    t0 = time.monotonic()
    body: dict[str, Any] = {}
    while time.monotonic() - t0 < timeout:
        body = fetch()
        if is_done(body):
            return body
        time.sleep(interval)
    print(f"[TIMEOUT] {step_name} poll 超时 {timeout}s, last={body}", file=sys.stderr)
    sys.exit(3)


def _step(n: int, total: int, msg: str) -> None:
    print(f"[{n}/{total}] {msg}", flush=True)


# ---------- Task 2: cluster 解析 ----------

def resolve_cluster_id(keyword: str) -> tuple[int, int, str]:
    """返回 (cluster_id, macrolocal_cluster_id, name). 新 draft 接口要 macrolocal."""
    resp = _post("/cluster/list", {"cluster_type": "CLUSTER_TYPE_OZON"})
    clusters = resp.get("clusters") or []
    hits = [c for c in clusters if keyword in (c.get("name") or "")]
    if not hits:
        names = [c.get("name", "") for c in clusters]
        print(f"[ERROR] 未找到匹配 '{keyword}' 的集群. 可选:", file=sys.stderr)
        for n in names:
            print(f"  - {n}", file=sys.stderr)
        sys.exit(1)
    if len(hits) > 1:
        names = [c.get("name") for c in hits]
        print(f"[ERROR] '{keyword}' 命中多个集群, 请给更精确关键字: {names}",
              file=sys.stderr)
        sys.exit(1)
    c = hits[0]
    macro = c.get("macrolocal_cluster_id")
    if macro in (None, 0, ""):
        print(f"[ERROR] 集群 {c.get('name')} 缺 macrolocal_cluster_id, 新 draft 接口无法用",
              file=sys.stderr)
        sys.exit(2)
    return int(c["id"]), int(macro), c.get("name", "")


# ---------- Task 3: drop-off 解析 ----------

_WH_TYPE_NUM = {
    "WAREHOUSE_TYPE_DELIVERY_POINT": 1,
    "WAREHOUSE_TYPE_ORDERS_RECEIVING_POINT": 2,
    "WAREHOUSE_TYPE_SORTING_CENTER": 3,
    "WAREHOUSE_TYPE_FULL_FILLMENT": 4,
    "WAREHOUSE_TYPE_CROSS_DOCK": 5,
}


def resolve_dropoff_id(keyword: str, create_type: str) -> tuple[int, str, int]:
    """返回 (warehouse_id, name, warehouse_type_num). 新 draft 接口要 type 数字."""
    if len(keyword) < 4:
        print("[ERROR] Ozon 要求 search 关键字 >=4 字符", file=sys.stderr)
        sys.exit(1)
    resp = _post(
        "/warehouse/fbo/list",
        {"filter_by_supply_type": [create_type], "search": keyword},
    )
    hits = resp.get("search") or []
    if not hits:
        print(f"[ERROR] 未找到 drop-off 匹配 '{keyword}' (type={create_type})",
              file=sys.stderr)
        sys.exit(1)
    if len(hits) > 1:
        print(f"[WARN] '{keyword}' 命中 {len(hits)} 个, 取第一个:", file=sys.stderr)
        for h in hits[:5]:
            print(f"  - {h.get('name')} (id={h.get('warehouse_id')})", file=sys.stderr)
    w = hits[0]
    wt_str = w.get("warehouse_type", "")
    wt_num = _WH_TYPE_NUM.get(wt_str, 3)  # 默认 3=SORTING_CENTER
    return int(w["warehouse_id"]), w.get("name", ""), wt_num


# ---------- Task 4: 创建 draft ----------

def create_draft(
    macrolocal_cluster_id: int,
    dropoff_id: int,
    dropoff_type_num: int,
    sku: int,
    qty: int,
    create_type: str,
) -> int:
    """新版 draft 接口同步返回 draft_id (不再是 operation_id 两段式).

    CROSSDOCK → /v1/draft/crossdock/create
    DIRECT    → /v1/draft/direct/create
    (/v1/draft/create 已废弃, 2026-03-16 禁用)
    """
    items = [{"sku": sku, "quantity": qty}]
    if create_type == "CREATE_TYPE_CROSSDOCK":
        path = "/draft/crossdock/create"
        body: dict[str, Any] = {
            "cluster_info": {
                "items": items,
                "macrolocal_cluster_id": str(macrolocal_cluster_id),
            },
            "deletion_sku_mode": "PARTIAL",
            "delivery_info": {
                "type": "DROPOFF",
                "drop_off_warehouse": {
                    "warehouse_id": dropoff_id,
                    "warehouse_type": dropoff_type_num,
                },
            },
        }
    elif create_type == "CREATE_TYPE_DIRECT":
        path = "/draft/direct/create"
        body = {
            "cluster_info": {
                "items": items,
                "macrolocal_cluster_id": str(macrolocal_cluster_id),
            },
            "deletion_sku_mode": "PARTIAL",
        }
    else:
        print(f"[ERROR] 不支持的 type: {create_type}", file=sys.stderr)
        sys.exit(1)
    resp = _post(path, body)
    draft_id = resp.get("draft_id")
    errs = resp.get("errors") or []
    if errs or not draft_id:
        print(f"[ERROR] {path} 失败: {json.dumps(resp, ensure_ascii=False, indent=2)}",
              file=sys.stderr)
        sys.exit(2)
    return int(draft_id)


# ---------- Task 5: 列目标集群仓库候选 (新 API 下不再 poll draft/create/info) ----------

def list_cluster_fulfillment_warehouses(cluster_name_keyword: str) -> list[dict[str, Any]]:
    """从 /cluster/list 拿指定集群内 type=FULL_FILLMENT 的仓列表 (越库目的地)."""
    resp = _post("/cluster/list", {"cluster_type": "CLUSTER_TYPE_OZON"})
    out: list[dict[str, Any]] = []
    for c in resp.get("clusters") or []:
        if cluster_name_keyword not in (c.get("name") or ""):
            continue
        for lc in c.get("logistic_clusters") or []:
            for w in lc.get("warehouses") or []:
                if w.get("type") == "FULL_FILLMENT":
                    out.append(
                        {
                            "warehouse_id": int(w["warehouse_id"]),
                            "name": w.get("name", ""),
                        }
                    )
    return out


# ---------- Task 6: 查时段 ----------

def _pick_earliest_timeslot(
    ts_resp: dict[str, Any], timeslot_index: int
) -> dict[str, Any]:
    result = ts_resp.get("result") or ts_resp
    outer = result.get("drop_off_warehouse_timeslots") or {}
    if isinstance(outer, list):
        outer = outer[0] if outer else {}
    days = outer.get("days") or []
    for day in days:
        slots = day.get("timeslots") or []
        if slots:
            idx = timeslot_index if timeslot_index < len(slots) else 0
            return slots[idx]
    print(f"[ERROR] 查询区间内无可用时段. 原始响应: "
          f"{json.dumps(ts_resp, ensure_ascii=False)[:400]}", file=sys.stderr)
    sys.exit(2)


def find_timeslot(
    draft_id: int,
    warehouse_id: int,
    macrolocal_cluster_id: int | str,
    create_type: str,
    days_from: int,
    days_to: int,
    timeslot_index: int,
) -> tuple[dict[str, Any], str]:
    """v2 timeslot/info: selected_cluster_warehouses + YYYY-MM-DD + supply_type int.

    split-draft 是异步计算; timeslot 接口有时会返回 FailedPrecondition
    "Can't find success calculation result for draft" — 需隔几秒重试.
    """
    if days_to - days_from > 28:
        print("[ERROR] Ozon 限制 date_from/date_to 跨度 <=28 天", file=sys.stderr)
        sys.exit(1)
    supply_type_num = SUPPLY_TYPE_NUM.get(create_type)
    if supply_type_num is None:
        print(f"[ERROR] 未知 create_type: {create_type}", file=sys.stderr)
        sys.exit(1)
    today = date.today()
    date_from = (today + timedelta(days=days_from)).strftime("%Y-%m-%d")
    date_to = (today + timedelta(days=days_to)).strftime("%Y-%m-%d")
    body = {
        "draft_id": draft_id,
        "selected_cluster_warehouses": [
            {
                "warehouse_id": str(warehouse_id),
                "macrolocal_cluster_id": str(macrolocal_cluster_id),
            }
        ],
        "supply_type": supply_type_num,
        "date_from": date_from,
        "date_to": date_to,
    }
    last_err: dict[str, Any] = {}
    for attempt in range(1, 7):  # 最多 6 次, 间隔 10s → 累计 50s
        r = requests.post(
            f"{_ARGS.api}/draft/timeslot/info",
            json=body,
            headers={"X-Ozon-Account": quote(_ARGS.account, safe="")},
            timeout=60,
        )
        if r.status_code < 400:
            resp = r.json()
            slot = _pick_earliest_timeslot(resp, timeslot_index)
            day_str = str(slot.get("from_in_timezone", ""))[:10]
            return slot, day_str
        try:
            last_err = r.json()
        except Exception:
            last_err = {"text": r.text[:400]}
        raw_msg = ((last_err.get("error") or {}).get("ozon_raw") or {}).get("body", {}).get("message", "")
        if "Can't find success calculation result" in raw_msg:
            print(f"    (draft 计算未就绪, 第 {attempt}/6 次重试, 等 10s)", flush=True)
            time.sleep(10)
            continue
        print(f"[API ERROR] /draft/timeslot/info -> HTTP {r.status_code}",
              file=sys.stderr)
        print(f"  body: {r.text[:500]}", file=sys.stderr)
        sys.exit(2)
    print(f"[ERROR] timeslot 重试 6 次仍失败: {json.dumps(last_err, ensure_ascii=False)[:400]}",
          file=sys.stderr)
    sys.exit(2)


# ---------- Task 7: 创建 supply + 轮询 ----------

def create_supply_and_wait(
    draft_id: int,
    warehouse_id: int,
    macrolocal_cluster_id: int | str,
    create_type: str,
    slot: dict[str, Any],
) -> list[int]:
    """v2 /draft/supply/create:
    - body 新 schema: selected_cluster_warehouses[] + supply_type + timeslot
    - 返回 {draft_id, error_reasons}, 无 operation_id
    - 轮询 v2 /draft/supply/create/status by draft_id → {order_id, status, error_reasons}
    - timeslot 字段必须无 Z 后缀 (v1 要 Z, v2 反过来要 '2026-04-26T21:00:00')
    """
    supply_type_num = SUPPLY_TYPE_NUM.get(create_type)
    if supply_type_num is None:
        print(f"[ERROR] 未知 create_type: {create_type}", file=sys.stderr)
        sys.exit(1)
    r = _post(
        "/draft/supply/create",
        {
            "draft_id": draft_id,
            "selected_cluster_warehouses": [
                {
                    "warehouse_id": str(warehouse_id),
                    "macrolocal_cluster_id": str(macrolocal_cluster_id),
                }
            ],
            "supply_type": supply_type_num,
            "timeslot": {
                "from_in_timezone": slot["from_in_timezone"].rstrip("Z"),
                "to_in_timezone": slot["to_in_timezone"].rstrip("Z"),
            },
        },
    )
    errs = r.get("error_reasons") or []
    if errs:
        print(f"[ERROR] /draft/supply/create 返回 error_reasons: {errs}",
              file=sys.stderr)
        sys.exit(2)

    final = _poll(
        fetch=lambda: _post(
            "/draft/supply/create/status", {"draft_id": draft_id}
        ),
        is_done=lambda b: str(b.get("status", "")).upper()
        in {"SUCCESS", "FAILED", "ERROR"},
        step_name="supply/create/status",
    )
    if str(final.get("status", "")).upper() != "SUCCESS":
        print(
            f"[ERROR] supply 创建失败: "
            f"{json.dumps(final, ensure_ascii=False, indent=2)}",
            file=sys.stderr,
        )
        sys.exit(2)
    order_id = final.get("order_id")
    if not order_id:
        print(f"[ERROR] status=SUCCESS 但无 order_id: {final}", file=sys.stderr)
        sys.exit(2)
    return [int(order_id)]


# ---------- Task 8: main ----------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Ozon FBO 越库一键发货脚本 (调 localhost:4182 服务)",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--cluster", required=True,
                   help="目标集群名关键字 (如 Москва)")
    p.add_argument("--dropoff", required=True,
                   help="drop-off 仓名关键字, >=4 字符 (如 ЩЕРБИНКА)")
    p.add_argument("--sku", type=int, required=True, help="SKU 数字 id")
    p.add_argument("--qty", type=int, required=True, help="件数")
    p.add_argument("--type", default="CREATE_TYPE_CROSSDOCK",
                   choices=["CREATE_TYPE_CROSSDOCK", "CREATE_TYPE_DIRECT"],
                   help="发货类型")
    p.add_argument("--api", default="http://127.0.0.1:4182",
                   help="FBO shipment service base URL")
    p.add_argument("--account", default="丝绸生活",
                   help="X-Ozon-Account header")
    p.add_argument("--days-from", type=int, default=3,
                   help="从今天起至少 N 天后才可预约")
    p.add_argument("--days-to", type=int, default=28,
                   help="查询时段区间 (从 days-from 起算跨度 <=28 天)")
    p.add_argument("--timeslot-index", type=int, default=0,
                   help="在首个有 slot 的日挑第几个时段 (0 起)")
    p.add_argument("--dry-run", action="store_true",
                   help="只走到选仓+选时段就停, 不实际建 supply")
    p.add_argument("--step-delay", type=float, default=0.0,
                   help="每步之间主动 sleep N 秒 (防御性; 默认 0 交由服务端限流器管)")
    return p.parse_args()


def _sleep_between(n: int, total: int) -> None:
    if _ARGS.step_delay > 0 and n < total:
        print(f"    (step-delay {_ARGS.step_delay}s)", flush=True)
        time.sleep(_ARGS.step_delay)


def main() -> None:
    global _ARGS
    _ARGS = parse_args()

    health = _get("/health")
    if not health.get("ok"):
        print(f"[ERROR] /health 异常: {health}", file=sys.stderr)
        sys.exit(2)
    if _ARGS.account not in (health.get("accounts") or []):
        print(
            f"[ERROR] 账号 '{_ARGS.account}' 不在 .env 配置里 "
            f"(已知: {health.get('accounts')})",
            file=sys.stderr,
        )
        sys.exit(1)

    total = 7

    _step(1, total, f"查集群 '{_ARGS.cluster}'")
    cluster_id, macrolocal_cluster_id, cluster_name = resolve_cluster_id(_ARGS.cluster)
    print(
        f"    → cluster_id={cluster_id}, macrolocal={macrolocal_cluster_id}, "
        f"name={cluster_name}"
    )
    _sleep_between(1, total)

    _step(2, total, f"查 drop-off '{_ARGS.dropoff}' (type={_ARGS.type})")
    dropoff_id, dropoff_name, dropoff_type_num = resolve_dropoff_id(
        _ARGS.dropoff, _ARGS.type
    )
    print(
        f"    → warehouse_id={dropoff_id}, name={dropoff_name}, "
        f"type_num={dropoff_type_num}"
    )
    _sleep_between(2, total)

    _step(3, total,
          f"创建 draft (SKU={_ARGS.sku}, qty={_ARGS.qty}, type={_ARGS.type})")
    draft_id = create_draft(
        macrolocal_cluster_id, dropoff_id, dropoff_type_num,
        _ARGS.sku, _ARGS.qty, _ARGS.type,
    )
    print(f"    → draft_id={draft_id}")
    _sleep_between(3, total)

    _step(4, total, "选 supply 目的仓")
    if _ARGS.type == "CREATE_TYPE_CROSSDOCK":
        supply_warehouse_id = dropoff_id
        supply_warehouse_name = dropoff_name
        print(f"    → CROSSDOCK 直接用 drop-off: {supply_warehouse_name} "
              f"(id={supply_warehouse_id})")
    else:
        candidates = list_cluster_fulfillment_warehouses(_ARGS.cluster)
        if not candidates:
            print(f"[ERROR] 集群 '{_ARGS.cluster}' 下没有 FULL_FILLMENT 仓",
                  file=sys.stderr)
            sys.exit(2)
        w = candidates[0]
        supply_warehouse_id = int(w["warehouse_id"])
        supply_warehouse_name = w.get("name", "")
        print(f"    → DIRECT 选 FULL_FILLMENT: {supply_warehouse_name} "
              f"(id={supply_warehouse_id})")
    _sleep_between(4, total)

    _step(5, total,
          f"查 {_ARGS.days_from}~{_ARGS.days_to}d 内可用时段")
    slot, day = find_timeslot(
        draft_id, supply_warehouse_id, macrolocal_cluster_id, _ARGS.type,
        _ARGS.days_from, _ARGS.days_to, _ARGS.timeslot_index,
    )
    print(f"    → {day} {slot.get('from_in_timezone')} ~ "
          f"{slot.get('to_in_timezone')}")
    _sleep_between(5, total)

    if _ARGS.dry_run:
        print()
        print("[dry-run] 未实际建单. 去掉 --dry-run 再跑可真实创建.")
        summary = {
            "cluster": cluster_name,
            "cluster_id": cluster_id,
            "macrolocal_cluster_id": macrolocal_cluster_id,
            "dropoff": dropoff_name,
            "dropoff_id": dropoff_id,
            "draft_id": draft_id,
            "warehouse": supply_warehouse_name,
            "warehouse_id": supply_warehouse_id,
            "timeslot": slot,
            "sku": _ARGS.sku,
            "qty": _ARGS.qty,
            "type": _ARGS.type,
        }
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return

    _step(6, total, "创建 supply + 轮询 status")
    order_ids = create_supply_and_wait(
        draft_id, supply_warehouse_id, macrolocal_cluster_id, _ARGS.type, slot,
    )
    print(f"    → order_ids={order_ids}")

    _step(7, total, "完成 ✓")
    summary = {
        "account": _ARGS.account,
        "cluster": cluster_name,
        "macrolocal_cluster_id": macrolocal_cluster_id,
        "dropoff": dropoff_name,
        "warehouse": supply_warehouse_name,
        "warehouse_id": supply_warehouse_id,
        "draft_id": draft_id,
        "timeslot": slot,
        "order_ids": order_ids,
        "sku": _ARGS.sku,
        "qty": _ARGS.qty,
        "type": _ARGS.type,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
