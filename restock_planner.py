"""Ozon 海外仓补货规划器.

输入 SKU，拉取 `/api/overseas_inventory` 与 `/api/rows`，按策略 (d) 分配发往各集群的箱数:
    Phase1  先把每个集群补到 FBO 安全天数 (= `集群需补货数量` 件).
            若总箱数不够, 两种兜底策略:
                greedy  严格按 (可售天数↑, 日销↓) 贪心, 一个吃满再下一个.
                prorate 同紧急度 (可售天数相同) 的集群按需求件数按比例分配, 剩余 1 箱按零头大的优先.
    Phase2  Phase1 覆盖后的剩余箱数, 按每日销量占比再分.

装箱率 (单箱装多少件) 优先级:
    --box-size CLI > 本地缓存 > 飞书海外仓库存表 (按 --account 查 / 未指定则轮询) > --web 表单 > 终端交互.

用法:
    python restock_planner.py --sku Q_ChongQiZui-HuangSe-Free
    python restock_planner.py --sku X --box-size 300 --date 2026/04/22
    python restock_planner.py --sku X --account 丝绸生活
    python restock_planner.py --sku X --fallback prorate
    python restock_planner.py --sku X --web            # 打开浏览器填装箱率
    python restock_planner.py --sku X --blacklist Красноярск,Новосибирск
"""

from __future__ import annotations

import argparse
import json
import math
import sys
import webbrowser
from dataclasses import dataclass, field
from datetime import date, timedelta
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Iterable
from urllib.parse import parse_qs, quote, urlparse

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


SCRIPT_DIR = Path(__file__).resolve().parent
ENV_FILE = SCRIPT_DIR / ".env"
CACHE_FILE = SCRIPT_DIR / ".box_size_cache.json"
LARK_RESOLVE_CACHE = SCRIPT_DIR / ".lark_resolve_cache.json"
DEFAULT_API_BASE = "http://localhost:4180"
LARK_HOST = "https://open.feishu.cn"
SKU_FIELD_NAME = "sku"
BOX_SIZE_FIELD_NAME = "单箱数量"


def load_env() -> dict[str, str]:
    env: dict[str, str] = {}
    if not ENV_FILE.exists():
        return env
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        env[k.strip()] = v.strip()
    return env


def parse_feishu_url(url: str) -> dict:
    p = urlparse(url)
    q = parse_qs(p.query)
    seg = p.path.rstrip("/").split("/")[-1]
    return {
        "raw_token": seg,
        "is_wiki": "/wiki/" in p.path,
        "is_base": "/base/" in p.path,
        "table_id": q.get("table", [""])[0],
        "view_id": q.get("view", [""])[0],
    }


def lark_accounts(env: dict[str, str]) -> list[dict]:
    accounts = []
    idx = 1
    while True:
        name = env.get(f"LARK_ACCOUNT_{idx}_NAME")
        url = env.get(f"LARK_ACCOUNT_{idx}_URL")
        if not name or not url:
            break
        accounts.append({"name": name, "url": url})
        idx += 1
    return accounts


def lark_tenant_token(env: dict[str, str]) -> str:
    r = requests.post(
        f"{LARK_HOST}/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": env["LARK_APP_ID"], "app_secret": env["LARK_APP_SECRET"]},
        timeout=15,
    )
    r.raise_for_status()
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"飞书认证失败: {d}")
    return d["tenant_access_token"]


def _load_resolve_cache() -> dict:
    if LARK_RESOLVE_CACHE.exists():
        try:
            return json.loads(LARK_RESOLVE_CACHE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}
    return {}


def _save_resolve_cache(c: dict) -> None:
    LARK_RESOLVE_CACHE.write_text(json.dumps(c, ensure_ascii=False, indent=2), encoding="utf-8")


def resolve_app_token(url: str, token: str) -> tuple[str, str]:
    """返回 (app_token, table_id). wiki 链接先解析到 base."""
    info = parse_feishu_url(url)
    cache = _load_resolve_cache()
    if url in cache:
        return cache[url]["app_token"], info["table_id"]
    if info["is_wiki"]:
        r = requests.get(
            f"{LARK_HOST}/open-apis/wiki/v2/spaces/get_node",
            params={"token": info["raw_token"]},
            headers={"Authorization": f"Bearer {token}"},
            timeout=15,
        )
        d = r.json()
        if d.get("code") != 0:
            raise RuntimeError(f"wiki 节点解析失败: {d}")
        app_token = d["data"]["node"]["obj_token"]
    else:
        app_token = info["raw_token"]
    cache[url] = {"app_token": app_token, "table_id": info["table_id"]}
    _save_resolve_cache(cache)
    return app_token, info["table_id"]


def lark_find_box_size(
    token: str, app_token: str, table_id: str, sku: str
) -> int | None:
    r = requests.post(
        f"{LARK_HOST}/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/search",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json; charset=utf-8",
        },
        params={"page_size": 5},
        json={
            "filter": {
                "conjunction": "and",
                "conditions": [
                    {"field_name": SKU_FIELD_NAME, "operator": "is", "value": [sku]}
                ],
            },
            "field_names": [SKU_FIELD_NAME, BOX_SIZE_FIELD_NAME],
        },
        timeout=15,
    )
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"搜索失败 app={app_token}: {d}")
    for it in d.get("data", {}).get("items", []):
        v = it.get("fields", {}).get(BOX_SIZE_FIELD_NAME)
        if isinstance(v, (int, float)) and v > 0:
            return int(v)
    return None


@dataclass
class ClusterRow:
    cluster: str
    cluster_cn_ru: str
    fbo_on_shelf: float
    fbo_in_transit: float
    daily_sales: float
    safety_days: int
    days_of_supply: int
    need_pieces: float
    sales_7d: float
    sales_28d: float

    @classmethod
    def from_api(cls, d: dict) -> "ClusterRow":
        return cls(
            cluster=d.get("集群", ""),
            cluster_cn_ru=d.get("集群中俄", d.get("集群", "")),
            fbo_on_shelf=float(d.get("FBO上架数量(万)", 0) or 0),
            fbo_in_transit=float(d.get("FBO越库在途数量(万)", 0) or 0),
            daily_sales=float(d.get("每日销量", 0) or 0),
            safety_days=int(d.get("FBO安全天数", 0) or 0),
            days_of_supply=int(d.get("FBO上架可售天数", 0) or 0),
            need_pieces=float(d.get("集群需补货数量", 0) or 0),
            sales_7d=float(d.get("7日销量", 0) or 0),
            sales_28d=float(d.get("28日销量", 0) or 0),
        )


@dataclass
class Allocation:
    row: ClusterRow
    boxes: int = 0
    phase1_boxes: int = 0
    phase2_boxes: int = 0

    @property
    def pieces(self) -> int:
        return self.boxes * self._box_size

    _box_size: int = field(default=0, repr=False)


def fetch_overseas_inventory(api_base: str, date_str: str, sku: str) -> float:
    url = f"{api_base}/api/overseas_inventory"
    r = requests.get(url, params={"date": date_str, "sku": sku}, timeout=15)
    r.raise_for_status()
    data = r.json()
    return float(data.get("value", 0) or 0)


def fetch_rows(api_base: str, date_str: str, sku: str, blacklist: str | None) -> list[ClusterRow]:
    url = f"{api_base}/api/rows"
    params = {"date": date_str, "sku": sku}
    if blacklist:
        params["blacklist"] = blacklist
    r = requests.get(url, params=params, timeout=30)
    r.raise_for_status()
    data = r.json()
    return [ClusterRow.from_api(x) for x in data.get("rows", [])]


def load_cache() -> dict:
    if CACHE_FILE.exists():
        try:
            return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}
    return {}


def save_cache(cache: dict) -> None:
    CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def resolve_box_size_from_lark(
    sku: str, account: str | None, env: dict[str, str]
) -> tuple[int | None, str | None]:
    """按 account 查 1 个表；未指定 account 时轮询. 返回 (box_size, 命中的账号名)."""
    if not env.get("LARK_APP_ID") or not env.get("LARK_APP_SECRET"):
        return None, None
    accounts = lark_accounts(env)
    if not accounts:
        return None, None
    try:
        token = lark_tenant_token(env)
    except Exception as e:
        print(f"  !! 飞书认证失败: {e}")
        return None, None

    if account:
        accounts = [a for a in accounts if a["name"] == account]
        if not accounts:
            print(f"  !! 未找到账号 {account!r}, 跳过飞书查询")
            return None, None

    for acc in accounts:
        try:
            app_token, table_id = resolve_app_token(acc["url"], token)
            box_size = lark_find_box_size(token, app_token, table_id, sku)
        except Exception as e:
            print(f"  !! 查询 {acc['name']} 失败: {e}")
            continue
        if box_size:
            print(f"  ✓ 飞书命中: {acc['name']} / sku={sku} / 单箱数量={box_size}")
            return box_size, acc["name"]
    return None, None


def prompt_box_size(sku: str) -> int:
    while True:
        raw = input(f"[装箱率未知] 请输入 SKU '{sku}' 的单箱装箱数量 (件/箱): ").strip()
        if raw.isdigit() and int(raw) > 0:
            return int(raw)
        print("  需要正整数, 重试.")


class _BoxSizeHandler(BaseHTTPRequestHandler):
    sku: str = ""
    collected: dict = {}

    def log_message(self, format: str, *args) -> None:  # silence stderr
        return

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/submit":
            qs = parse_qs(parsed.query)
            raw = (qs.get("box_size") or [""])[0]
            if raw.isdigit() and int(raw) > 0:
                self.__class__.collected["box_size"] = int(raw)
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(
                    f"<h2>已收到: {raw} 件/箱, 可关闭页面</h2>".encode("utf-8")
                )
                return
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b"invalid")
            return

        html = f"""<!doctype html>
<html lang='zh'>
<meta charset='utf-8'>
<title>装箱率填写 - {self.__class__.sku}</title>
<style>
body{{font-family:-apple-system,Helvetica,sans-serif;max-width:480px;margin:60px auto;padding:24px;}}
h1{{font-size:20px}} .sku{{color:#666;font-family:Menlo,monospace}}
input[type=number]{{font-size:18px;padding:10px;width:100%;box-sizing:border-box}}
button{{margin-top:16px;padding:10px 20px;font-size:16px;cursor:pointer}}
</style>
<h1>海外仓补货 · 装箱率</h1>
<p>SKU: <span class='sku'>{self.__class__.sku}</span></p>
<p>请填写该 SKU 的<strong>单箱装箱数量 (件/箱)</strong>:</p>
<form action='/submit' method='get'>
  <input type='number' name='box_size' min='1' required autofocus>
  <button type='submit'>提交</button>
</form>
"""
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(html.encode("utf-8"))


def prompt_box_size_web(sku: str, port: int = 4188) -> int:
    _BoxSizeHandler.sku = sku
    _BoxSizeHandler.collected = {}
    server = HTTPServer(("127.0.0.1", port), _BoxSizeHandler)
    url = f"http://127.0.0.1:{port}/"
    print(f"[web] 已启动装箱率填写页: {url}")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    while "box_size" not in _BoxSizeHandler.collected:
        server.handle_request()
    server.server_close()
    return _BoxSizeHandler.collected["box_size"]


def resolve_box_size(
    sku: str,
    arg_box_size: int | None,
    account: str | None,
    use_web: bool,
    env: dict[str, str],
) -> tuple[int, str]:
    """返回 (装箱率, 来源说明)."""
    if arg_box_size and arg_box_size > 0:
        return arg_box_size, "CLI --box-size"

    cache = load_cache()
    if sku in cache and isinstance(cache[sku], int) and cache[sku] > 0:
        print(f"  ✓ 本地缓存命中: 装箱率 = {cache[sku]} 件/箱")
        return cache[sku], "本地缓存"

    got, hit_account = resolve_box_size_from_lark(sku, account, env)
    if got:
        cache[sku] = got
        save_cache(cache)
        return got, f"飞书[{hit_account}]"

    box_size = prompt_box_size_web(sku) if use_web else prompt_box_size(sku)
    cache[sku] = box_size
    save_cache(cache)
    return box_size, "手动输入"


def allocate(
    rows: list[ClusterRow],
    total_pieces_available: float,
    box_size: int,
    fallback: str = "greedy",
) -> tuple[list[Allocation], dict]:
    """返回 (每集群分配, 统计信息).

    fallback: 'greedy' | 'prorate'
        greedy  按 (可售天数↑, 日销↓) 一个集群吃满再下一个.
        prorate 按可售天数分桶, 同桶内按 need_boxes 占比分配, 余数按零头大的优先.
    """
    available_boxes = int(total_pieces_available // box_size)

    allocs = [Allocation(row=r, _box_size=box_size) for r in rows]

    phase1_need_pieces = {a.row.cluster: a.row.need_pieces for a in allocs if a.row.need_pieces > 0}
    phase1_need_boxes = {
        k: math.ceil(v / box_size) for k, v in phase1_need_pieces.items() if v > 0
    }
    total_phase1_boxes = sum(phase1_need_boxes.values())

    remaining = available_boxes

    if total_phase1_boxes <= remaining:
        for a in allocs:
            need = phase1_need_boxes.get(a.row.cluster, 0)
            a.phase1_boxes = need
            a.boxes = need
        remaining -= total_phase1_boxes
    elif fallback == "prorate":
        by_dos: dict[int, list[Allocation]] = {}
        for a in allocs:
            if phase1_need_boxes.get(a.row.cluster, 0) > 0:
                by_dos.setdefault(a.row.days_of_supply, []).append(a)
        for dos in sorted(by_dos.keys()):
            if remaining <= 0:
                break
            bucket = by_dos[dos]
            bucket_need = sum(phase1_need_boxes[a.row.cluster] for a in bucket)
            if bucket_need <= remaining:
                for a in bucket:
                    need = phase1_need_boxes[a.row.cluster]
                    a.phase1_boxes = need
                    a.boxes = need
                    remaining -= need
            else:
                raw = {
                    a.row.cluster: remaining * phase1_need_boxes[a.row.cluster] / bucket_need
                    for a in bucket
                }
                floored = {k: int(v) for k, v in raw.items()}
                leftover = remaining - sum(floored.values())
                frac_order = sorted(
                    raw.items(), key=lambda kv: (kv[1] - int(kv[1]), -phase1_need_boxes[kv[0]]),
                    reverse=True,
                )
                for cluster, _ in frac_order:
                    if leftover <= 0:
                        break
                    if floored[cluster] >= phase1_need_boxes[cluster]:
                        continue
                    floored[cluster] += 1
                    leftover -= 1
                for a in bucket:
                    give = min(floored.get(a.row.cluster, 0), phase1_need_boxes[a.row.cluster])
                    a.phase1_boxes = give
                    a.boxes = give
                remaining = leftover
    else:
        urgent = [a for a in allocs if phase1_need_boxes.get(a.row.cluster, 0) > 0]
        urgent.sort(key=lambda a: (a.row.days_of_supply, -a.row.daily_sales))
        for a in urgent:
            need = phase1_need_boxes[a.row.cluster]
            give = min(need, remaining)
            a.phase1_boxes = give
            a.boxes = give
            remaining -= give
            if remaining <= 0:
                break

    warnings: list[str] = []
    if available_boxes == 0 and total_pieces_available > 0:
        warnings.append(
            f"海外仓 {int(total_pieces_available)} 件 < 装箱率 {box_size}, 凑不满一整箱"
        )

    if remaining > 0:
        sales_total = sum(a.row.daily_sales for a in allocs if a.row.daily_sales > 0)
        if sales_total > 0:
            raw = {
                a.row.cluster: remaining * (a.row.daily_sales / sales_total)
                for a in allocs
            }
            floored = {k: int(v) for k, v in raw.items()}
            leftover = remaining - sum(floored.values())
            frac_order = sorted(
                raw.items(),
                key=lambda kv: (kv[1] - int(kv[1])),
                reverse=True,
            )
            for cluster, _ in frac_order:
                if leftover <= 0:
                    break
                floored[cluster] += 1
                leftover -= 1
            for a in allocs:
                add = floored.get(a.row.cluster, 0)
                a.phase2_boxes = add
                a.boxes += add
            remaining = 0
        else:
            warnings.append(
                f"所有集群日销量均为 0, Phase2 无法按销量分配, 保留 {remaining} 箱未分配 (建议人工复核)"
            )

    waste_pieces = sum(
        a.phase1_boxes * box_size - int(a.row.need_pieces)
        for a in allocs
        if a.phase1_boxes > 0 and a.phase1_boxes * box_size > int(a.row.need_pieces)
    )

    stats = {
        "total_pieces_available": total_pieces_available,
        "box_size": box_size,
        "available_boxes": available_boxes,
        "phase1_need_pieces": sum(phase1_need_pieces.values()),
        "phase1_need_boxes": total_phase1_boxes,
        "phase1_satisfied": total_phase1_boxes <= available_boxes,
        "phase2_distributed_boxes": sum(a.phase2_boxes for a in allocs),
        "remaining_boxes_after_alloc": remaining,
        "shortfall_boxes": max(0, total_phase1_boxes - available_boxes),
        "waste_pieces": waste_pieces,
        "inventory_too_low_for_one_box": available_boxes == 0 and total_pieces_available > 0,
        "warnings": warnings,
    }
    return allocs, stats


def print_plan(sku: str, date_str: str, allocs: list[Allocation], stats: dict) -> None:
    box_size = stats["box_size"]
    print()
    print(f"== 补货方案 · SKU {sku} · 日期 {date_str} ==")
    print(
        f"海外仓: {int(stats['total_pieces_available']):>8} 件 "
        f"= {stats['available_boxes']:>5} 箱 (装箱率 {box_size} 件/箱)"
    )
    print(
        f"Phase1 (补到安全天数) 需求: {int(stats['phase1_need_pieces']):>8} 件 "
        f"= {stats['phase1_need_boxes']:>5} 箱"
    )
    if stats["phase1_satisfied"]:
        print(
            f"Phase1 全部满足, Phase2 再分: {stats['phase2_distributed_boxes']} 箱 "
            f"(按每日销量占比)"
        )
    else:
        mode = stats.get("fallback", "greedy")
        desc = "一个集群吃满再下一个" if mode == "greedy" else "同紧急度按需求比例分配"
        print(
            f"!! 箱数不足, 缺口 {stats['shortfall_boxes']} 箱. 兜底策略: {mode} ({desc})."
        )
    if stats.get("waste_pieces", 0) > 0:
        print(f"箱级取整浪费: {stats['waste_pieces']} 件 (装满整箱超出需求的部分)")
    for w in stats.get("warnings", []):
        print(f"⚠️  {w}")
    print()

    header = f"{'集群中俄':<32} {'上架':>6} {'日销':>6} {'可售天':>7} {'需补件':>8} {'Ph1箱':>6} {'Ph2箱':>6} {'合计箱':>7} {'合计件':>8}"
    print(header)
    print("-" * len(header))
    allocs_sorted = sorted(allocs, key=lambda a: -a.boxes)
    total_boxes = 0
    total_pieces = 0
    for a in allocs_sorted:
        r = a.row
        pieces = a.boxes * box_size
        total_boxes += a.boxes
        total_pieces += pieces
        print(
            f"{r.cluster_cn_ru[:32]:<32} "
            f"{int(r.fbo_on_shelf):>6} "
            f"{r.daily_sales:>6.1f} "
            f"{r.days_of_supply:>7} "
            f"{int(r.need_pieces):>8} "
            f"{a.phase1_boxes:>6} "
            f"{a.phase2_boxes:>6} "
            f"{a.boxes:>7} "
            f"{pieces:>8}"
        )
    print("-" * len(header))
    print(
        f"{'合计':<32} {'':>6} {'':>6} {'':>7} {'':>8} "
        f"{sum(a.phase1_boxes for a in allocs):>6} "
        f"{sum(a.phase2_boxes for a in allocs):>6} "
        f"{total_boxes:>7} "
        f"{total_pieces:>8}"
    )
    print()


def write_excel(
    sku: str,
    date_str: str,
    allocs: list[Allocation],
    stats: dict,
    out_dir: Path,
) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    safe_date = date_str.replace("/", "-")
    out_path = out_dir / f"allocation_{safe_date}_{sku}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "补货分配"

    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    head_fill = PatternFill("solid", fgColor="FFE8A1")
    total_fill = PatternFill("solid", fgColor="D5E8D4")

    meta = [
        ("SKU", sku),
        ("日期", date_str),
        ("装箱率 (件/箱)", stats["box_size"]),
        ("海外仓总件数", int(stats["total_pieces_available"])),
        ("海外仓可用箱数", stats["available_boxes"]),
        ("Phase1 需求 (件)", int(stats["phase1_need_pieces"])),
        ("Phase1 需求 (箱)", stats["phase1_need_boxes"]),
        ("Phase1 是否满足", "是" if stats["phase1_satisfied"] else f"否 (缺 {stats['shortfall_boxes']} 箱)"),
        ("Phase2 分配 (箱)", stats["phase2_distributed_boxes"]),
        ("兜底策略", stats.get("fallback", "greedy")),
        ("装箱率来源", stats.get("box_size_source", "")),
        ("箱级取整浪费 (件)", stats.get("waste_pieces", 0)),
        ("剩余未分配箱数", stats.get("remaining_boxes_after_alloc", 0)),
        ("警告", "; ".join(stats.get("warnings", [])) or "无"),
    ]
    for i, (k, v) in enumerate(meta, start=1):
        ws.cell(row=i, column=1, value=k).font = bold
        ws.cell(row=i, column=2, value=v)
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 40

    start_row = len(meta) + 3
    headers = [
        "集群中俄",
        "集群",
        "FBO上架",
        "FBO在途",
        "每日销量",
        "FBO安全天数",
        "FBO可售天数",
        "需补货件数",
        sku + " (Phase1箱)",
        sku + " (Phase2箱)",
        sku + " (合计箱)",
        sku + " (合计件)",
    ]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c, value=h)
        cell.font = bold
        cell.fill = head_fill
        cell.alignment = center

    allocs_sorted = sorted(allocs, key=lambda a: -a.boxes)
    for i, a in enumerate(allocs_sorted, start=1):
        r = a.row
        pieces = a.boxes * stats["box_size"]
        row = [
            r.cluster_cn_ru,
            r.cluster,
            int(r.fbo_on_shelf),
            int(r.fbo_in_transit),
            r.daily_sales,
            r.safety_days,
            r.days_of_supply,
            int(r.need_pieces),
            a.phase1_boxes,
            a.phase2_boxes,
            a.boxes,
            pieces,
        ]
        for c, v in enumerate(row, start=1):
            ws.cell(row=start_row + i, column=c, value=v)

    total_row_idx = start_row + len(allocs_sorted) + 1
    totals = [
        "合计", "", "", "", "", "", "", int(sum(a.row.need_pieces for a in allocs)),
        sum(a.phase1_boxes for a in allocs),
        sum(a.phase2_boxes for a in allocs),
        sum(a.boxes for a in allocs),
        sum(a.boxes for a in allocs) * stats["box_size"],
    ]
    for c, v in enumerate(totals, start=1):
        cell = ws.cell(row=total_row_idx, column=c, value=v)
        cell.font = bold
        cell.fill = total_fill

    for c in range(1, len(headers) + 1):
        letter = get_column_letter(c)
        ws.column_dimensions[letter].width = max(12, len(headers[c - 1]) + 4)
    ws.column_dimensions["A"].width = 36

    wb.save(out_path)
    return out_path


def default_date() -> str:
    return (date.today() - timedelta(days=1)).strftime("%Y/%m/%d")


# ---------- FBO 发货集成 (调 localhost:4182 ozon_fbo_shipment_service) ----------


def _fbo_post(
    fbo_service: str, account: str, path: str, body: dict
) -> dict:
    r = requests.post(
        f"{fbo_service}{path}",
        json=body,
        headers={"X-Ozon-Account": quote(account, safe="")},
        timeout=60,
    )
    if r.status_code >= 400:
        raise RuntimeError(
            f"FBO {path} HTTP {r.status_code}: {r.text[:400]}"
        )
    return r.json()


def parse_dropoff_map(spec: str) -> dict[str, str]:
    """"Москва=ЩЕРБИНКА,Новосибирск=НСК_ХАБ" → {"Москва":"ЩЕРБИНКА",...}"""
    out: dict[str, str] = {}
    for seg in (spec or "").split(","):
        seg = seg.strip()
        if not seg or "=" not in seg:
            continue
        k, v = seg.split("=", 1)
        out[k.strip()] = v.strip()
    return out


def resolve_ozon_sku(fbo_service: str, account: str, offer_id: str) -> int:
    """POST /product/info/list {"offer_id":[...]} → 数字 sku.

    Ozon 新体系下 sku 跨 SDS/FBO 统一. 先认 FBO source, 没有再退回任何 source 的 sku.
    """
    resp = _fbo_post(
        fbo_service, account, "/product/info/list",
        {"offer_id": [offer_id]},
    )
    items = (resp.get("items") or resp.get("result", {}).get("items") or [])
    for it in items:
        sources = it.get("sources") or []
        for s in sources:
            if str(s.get("source", "")).upper() == "FBO" and s.get("sku"):
                return int(s["sku"])
        for s in sources:
            if s.get("sku"):
                return int(s["sku"])
        for k in ("fbo_sku", "sku"):
            if it.get(k):
                return int(it[k])
    raise RuntimeError(f"未从 /product/info/list 拿到 offer_id={offer_id} 的数字 sku; resp={resp}")


def build_fbo_plan(
    offer_id: str,
    ozon_sku: int,
    allocs: list[Allocation],
    box_size: int,
    create_type: str,
    dropoff_map: dict[str, str],
) -> list[dict]:
    """每个 boxes>0 的集群出一条 plan 行."""
    out: list[dict] = []
    for a in allocs:
        if a.boxes <= 0:
            continue
        pieces = a.boxes * box_size
        cluster_name = a.row.cluster
        dropoff = ""
        if create_type == "CREATE_TYPE_CROSSDOCK":
            dropoff = dropoff_map.get(cluster_name, "")
            if not dropoff:
                for k, v in dropoff_map.items():
                    if k in cluster_name or cluster_name in k:
                        dropoff = v
                        break
        out.append({
            "cluster": cluster_name,
            "cluster_cn_ru": a.row.cluster_cn_ru,
            "offer_id": offer_id,
            "sku": ozon_sku,
            "boxes": a.boxes,
            "qty_pieces": pieces,
            "create_type": create_type,
            "dropoff_keyword": dropoff,
        })
    return out


def execute_fbo_plan(
    plan: list[dict],
    fbo_service: str,
    account: str,
    days_from: int,
    days_to: int,
) -> list[dict]:
    """对 plan 每行调 /flow/create-supply. 先 /cluster/list 补 macrolocal,
    CROSSDOCK 再 /warehouse/fbo/list 补 dropoff id."""
    cluster_resp = _fbo_post(
        fbo_service, account, "/cluster/list",
        {"cluster_type": "CLUSTER_TYPE_OZON"},
    )
    clusters = cluster_resp.get("clusters") or []

    def _find_cluster(keyword: str) -> dict:
        hits = [c for c in clusters if keyword in (c.get("name") or "")]
        if not hits:
            raise RuntimeError(f"未找到匹配 '{keyword}' 的集群")
        return hits[0]

    today = date.today()
    date_from = (today + timedelta(days=days_from)).strftime("%Y-%m-%d")
    date_to = (today + timedelta(days=days_to)).strftime("%Y-%m-%d")

    results: list[dict] = []
    for row in plan:
        out: dict = {
            "cluster": row["cluster"],
            "offer_id": row["offer_id"],
            "sku": row["sku"],
            "qty_pieces": row["qty_pieces"],
            "create_type": row["create_type"],
        }
        try:
            cluster = _find_cluster(row["cluster"])
            macro = cluster.get("macrolocal_cluster_id")
            if macro in (None, 0, ""):
                raise RuntimeError(f"集群 {cluster.get('name')} 缺 macrolocal_cluster_id")
            out["macrolocal_cluster_id"] = str(macro)

            drop_id: int | None = None
            if row["create_type"] == "CREATE_TYPE_CROSSDOCK":
                kw = row["dropoff_keyword"]
                if len(kw) < 4:
                    raise RuntimeError("CROSSDOCK 需 dropoff_keyword ≥4 字符")
                wh_resp = _fbo_post(
                    fbo_service, account, "/warehouse/fbo/list",
                    {"filter_by_supply_type": ["CREATE_TYPE_CROSSDOCK"], "search": kw},
                )
                hits = wh_resp.get("search") or []
                if not hits:
                    raise RuntimeError(f"未找到 drop-off '{kw}'")
                drop_id = int(hits[0]["warehouse_id"])
                out["drop_off_warehouse_id"] = drop_id
                out["drop_off_name"] = hits[0].get("name", "")

            flow_body: dict = {
                "type": row["create_type"],
                "items": [{"sku": int(row["sku"]), "quantity": int(row["qty_pieces"])}],
                "macrolocal_cluster_id": str(macro),
                "date_from": date_from,
                "date_to": date_to,
            }
            if drop_id is not None:
                flow_body["drop_off_point_warehouse_id"] = drop_id

            flow_resp = _fbo_post(
                fbo_service, account, "/flow/create-supply", flow_body,
            )
            out["stage"] = flow_resp.get("stage")
            out["draft_id"] = flow_resp.get("draft_id")
            out["warehouse_id"] = flow_resp.get("warehouse_id")
            out["timeslot"] = flow_resp.get("timeslot")
            out["order_ids"] = flow_resp.get("order_ids") or []
            if flow_resp.get("error"):
                out["error"] = flow_resp["error"]
        except Exception as e:
            out["error"] = str(e)
        results.append(out)
    return results


def main(argv: Iterable[str] | None = None) -> int:
    env = load_env()
    default_api = env.get("API_BASE", DEFAULT_API_BASE)

    p = argparse.ArgumentParser(description="Ozon 海外仓补货规划器")
    p.add_argument("--sku", required=True, help="SKU, 例: Q_ChongQiZui-HuangSe-Free")
    p.add_argument("--date", default=default_date(), help="日期 YYYY/MM/DD (默认昨天)")
    p.add_argument("--box-size", type=int, default=0, help="单箱装箱数量 (件/箱)")
    p.add_argument("--account", default=None, help="账号名 (丝绸生活/个人之路); 不传则轮询")
    p.add_argument("--blacklist", default="", help="排除的集群, 逗号分隔")
    p.add_argument("--api-base", default=default_api)
    p.add_argument(
        "--fallback",
        choices=["greedy", "prorate"],
        default="greedy",
        help="箱不足兜底: greedy (吃满再下一个) / prorate (同紧急度按需求比例)",
    )
    p.add_argument("--web", action="store_true", help="缺装箱率时启动本地网页表单")
    p.add_argument("--out-dir", default=str(SCRIPT_DIR), help="Excel 输出目录")
    p.add_argument("--fbo-service", default="http://127.0.0.1:4182",
                   help="ozon_fbo_shipment_service base URL")
    p.add_argument("--fbo-account", default=None,
                   help="覆盖 --account 作为 X-Ozon-Account; 缺省同 --account 或 丝绸生活")
    p.add_argument("--create-type", default="CREATE_TYPE_CROSSDOCK",
                   choices=["CREATE_TYPE_CROSSDOCK", "CREATE_TYPE_DIRECT"],
                   help="Ozon 发货类型")
    p.add_argument("--dropoff-map", default="",
                   help='CROSSDOCK 每集群 dropoff: "Москва=ЩЕРБИНКА,Новосибирск=НСК_ХАБ"')
    p.add_argument("--execute", action="store_true",
                   help="实际按集群建 FBO draft+supply; 默认只出 plan.json 清单")
    p.add_argument("--days-from", type=int, default=3,
                   help="FBO 时段查询从今日起 N 天后开始")
    p.add_argument("--days-to", type=int, default=28,
                   help="FBO 时段查询截止到今日起 N 天内 (跨度 <=28)")
    args = p.parse_args(list(argv) if argv is not None else None)

    blacklist = args.blacklist.strip()

    print(f"[1/4] 拉取 /api/overseas_inventory ...")
    total_pieces = fetch_overseas_inventory(args.api_base, args.date, args.sku)
    print(f"      → 海外仓: {int(total_pieces)} 件")

    print(f"[2/4] 拉取 /api/rows ...")
    rows = fetch_rows(args.api_base, args.date, args.sku, blacklist or None)
    if not rows:
        print(f"  !! 未查到任何集群数据, 请确认 SKU / 日期")
        return 1
    print(f"      → {len(rows)} 个集群")

    print(f"[3/4] 解析装箱率 ...")
    box_size, src = resolve_box_size(args.sku, args.box_size, args.account, args.web, env)
    print(f"      → 装箱率: {box_size} 件/箱 (来源: {src})")

    print(f"[4/4] 计算分配 (兜底={args.fallback}) ...")
    allocs, stats = allocate(rows, total_pieces, box_size, fallback=args.fallback)
    stats["fallback"] = args.fallback
    stats["box_size_source"] = src

    print_plan(args.sku, args.date, allocs, stats)

    xlsx = write_excel(args.sku, args.date, allocs, stats, Path(args.out_dir))
    print(f"[done] Excel 已写入: {xlsx}")

    fbo_account = args.fbo_account or args.account or "丝绸生活"
    print(f"[fbo] 解析 offer_id={args.sku} 的数字 sku (account={fbo_account}) ...")
    try:
        ozon_sku = resolve_ozon_sku(args.fbo_service, fbo_account, args.sku)
    except Exception as e:
        print(f"      ! 解析失败: {e}. 跳过 FBO 集成, 只出 Excel.")
        return 0
    print(f"      → ozon_sku={ozon_sku}")

    dropoff_map = parse_dropoff_map(args.dropoff_map)
    fbo_plan = build_fbo_plan(
        args.sku, ozon_sku, allocs, box_size,
        args.create_type, dropoff_map,
    )
    date_tag = args.date.replace("/", "")
    plan_path = Path(args.out_dir) / f"fbo_plan_{args.sku}_{date_tag}.json"
    plan_path.write_text(
        json.dumps(fbo_plan, ensure_ascii=False, indent=2), encoding="utf-8",
    )
    print(f"[fbo-plan] 已写: {plan_path} ({len(fbo_plan)} 条)")

    if args.execute:
        if not fbo_plan:
            print("[fbo-execute] 无需建单 (plan 为空)")
            return 0
        missing_dropoff = [
            row["cluster"] for row in fbo_plan
            if row["create_type"] == "CREATE_TYPE_CROSSDOCK" and not row["dropoff_keyword"]
        ]
        if missing_dropoff:
            print(
                f"[ERROR] CROSSDOCK 模式下集群 {missing_dropoff} 缺 dropoff_keyword; "
                f"请用 --dropoff-map 补齐", file=sys.stderr,
            )
            return 2
        print(f"[fbo-execute] 开始按集群建单 ({len(fbo_plan)} 条) ...")
        results = execute_fbo_plan(
            fbo_plan, args.fbo_service, fbo_account,
            args.days_from, args.days_to,
        )
        result_path = plan_path.with_suffix(".result.json")
        result_path.write_text(
            json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8",
        )
        ok = sum(1 for r in results if r.get("order_ids"))
        fail = len(results) - ok
        print(f"[fbo-execute] 完成: 成功 {ok}, 失败 {fail}. 详情: {result_path}")
    else:
        print(f"[dry-run] 未建单. 复核 {plan_path.name} 后加 --execute 重跑,"
              f" 或挨个手动:")
        for row in fbo_plan[:3]:
            dropoff_arg = f" --dropoff {row['dropoff_keyword']}" if row["dropoff_keyword"] else ""
            print(
                f"  python3 create_moscow_shipment.py --cluster {row['cluster']}"
                f"{dropoff_arg} --sku {ozon_sku} --qty {row['qty_pieces']}"
                f" --type {row['create_type']} --account {fbo_account}"
            )

    return 0


if __name__ == "__main__":
    sys.exit(main())
