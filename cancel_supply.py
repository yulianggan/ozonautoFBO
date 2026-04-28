"""安全撤 supply order: 先把 timeslot 推 +82h (>80h 防罚), 然后 cancel.

用户规则 (2026-04-25): 取消货件需要把预约的时间改到离当前俄罗斯时间 80 小时之后再取消, 否则平台会处罚.

用法:
    python3 cancel_supply.py 100696121
    python3 cancel_supply.py 100696121 100680603 100696122   # 多个
    python3 cancel_supply.py --account 丝绸生活 100696121
"""
from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime, timezone, timedelta

SVC = "http://127.0.0.1:4182"


def post(account: str, path: str, body: dict, timeout: int = 30) -> dict:
    enc = urllib.parse.quote(account, safe="")
    req = urllib.request.Request(
        f"{SVC}{path}",
        data=json.dumps(body).encode(),
        headers={"Content-Type": "application/json", "X-Ozon-Account": enc},
    )
    return json.load(urllib.request.urlopen(req, timeout=timeout))


def poll(account: str, path: str, op_id: str, timeout: int = 60) -> dict:
    """简易轮询 — 兼容 SUCCESS / STATUS_SUCCESS (Ozon protobuf enum)."""
    body = {"operation_id": op_id}
    deadline = time.time() + timeout
    while time.time() < deadline:
        r = post(account, path, body)
        st = str(r.get("status", "")).upper()
        if any(t in st for t in ("SUCCESS", "FAILED", "ERROR", "READY", "COMPLETED")):
            return r
        time.sleep(2)
    raise TimeoutError(f"poll {path} {op_id} timeout")


def safe_cancel(account: str, order_id: int) -> bool:
    moscow = timezone(timedelta(hours=3))
    target_utc = datetime.now(timezone.utc) + timedelta(hours=82)
    target_utc = target_utc.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
    new_from = target_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
    new_to = (target_utc + timedelta(hours=1)).strftime("%Y-%m-%dT%H:%M:%SZ")

    # 1) 当前 timeslot
    needs_shift = True
    try:
        cur = post(account, "/supply-order/timeslot/get", {"supply_order_id": order_id})
        cur_from = (cur.get("timeslot") or {}).get("from") or cur.get("from", "")
        if cur_from:
            cur_dt = datetime.fromisoformat(cur_from.replace("Z", "+00:00"))
            if cur_dt.tzinfo is None:
                cur_dt = cur_dt.replace(tzinfo=moscow)
            hrs = (cur_dt - datetime.now(moscow)).total_seconds() / 3600
            if hrs >= 80:
                needs_shift = False
                print(f"  当前 timeslot {cur_from} ({hrs:.1f}h, ≥80h 直接 cancel)")
    except Exception as e:
        print(f"  ! timeslot/get 失败: {e}; 仍试 shift")

    # 2) shift if needed
    if needs_shift:
        print(f"  shift timeslot → {new_from} ~ {new_to} (+82h)")
        try:
            up = post(account, "/supply-order/timeslot/update", {
                "supply_order_id": order_id,
                "timeslot": {"from": new_from, "to": new_to},
            })
            op = up.get("operation_id")
            if op:
                r = poll(account, "/supply-order/timeslot/status", op)
                print(f"  shift status: {r.get('status')}")
        except Exception as e:
            print(f"  ! shift 失败: {e}; 仍试 cancel (可能扣分)")

    # 3) cancel
    try:
        c = post(account, "/supply-order/cancel", {"order_id": order_id})
        op = c.get("operation_id")
        if op:
            r = poll(account, "/supply-order/cancel/status", op)
            cancelled = (r.get("result") or {}).get("is_order_cancelled")
            print(f"  cancel: {r.get('status')} cancelled={cancelled}")
            return bool(cancelled)
    except Exception as e:
        print(f"  ! cancel 失败: {e}")
    return False


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("order_ids", nargs="+", type=int, help="要撤的 order_id (可多个)")
    ap.add_argument("--account", default="丝绸生活")
    args = ap.parse_args()

    ok, fail = 0, 0
    for oid in args.order_ids:
        print(f"\n=== order {oid} ===")
        if safe_cancel(args.account, oid):
            ok += 1
        else:
            fail += 1
    print(f"\n总: ok={ok} fail={fail}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
