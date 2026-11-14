#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数据源滞后哨兵 —— 把「官网停更」从静默变成可见信号。

为什么需要它
------------
2026-11-12、11-13 是正常交易日（腾讯行情有这两天 K 线），但中证官网的
perf 接口与估值指标 .xls 双双停在 11-11。脚本的预检按设计判定
「官网最新交易日未超过已发布 → 本轮跳过」并 sys.exit(0)，
于是两天里 69 次 Actions 运行全部 success、全部绿色、零提交，
GitHub Pages 静默停留在 11-11。

轮询早退本身是对的（省掉空跑重活），错的是它和「一切正常」在日志里
长得一模一样。本脚本独立比对「官网最新交易日」与「真实交易日」，
把滞后量算出来并据此产出 annotation / 退出码。

判定规则
--------
  lag = 参考交易日历中，晚于官网最新交易日的交易日个数
  lag == 0        正常
  lag >= 1        告警（::warning::），退出 0      —— 官网当晚发布偏晚属常态
  lag >= 2 且处于当日收尾窗口  报错（::error::），退出 1 —— 让 Actions 变红，触发失败通知

自身任何异常（网络、解析、字段变更）都退回退出 0：哨兵不该因为自己坏了
把生产流水线染红。唯一的非零退出是上面那条明确算出来的滞后结论。

用法
----
  python scripts/check_staleness.py            # CI / 日常
  python scripts/check_staleness.py --force    # 忽略收尾窗口，立即按 lag 判定
  python scripts/check_staleness.py --quiet    # 只输出结论行
"""

import argparse
import json
import os
import sys
import urllib.request
from datetime import date, datetime, timedelta, timezone

BJ = timezone(timedelta(hours=8))

CSINDEX_PERF = "https://www.csindex.com.cn/csindex-home/perf/index-perf"
# 注意 param 的第二段字面量 "day" 不能省：少了它腾讯会返回空 data。
TX_KLINE = ("https://web.ifzq.gtimg.cn/appstock/app/fqkline/get"
            "?param={sym},day,{start},{end},320,qfq")

INDEX_CODE = "H00922"      # 主标的：中证红利全收益
TX_SYMBOL = "sh000001"     # 交易日历参考：上证指数（与生产脚本同源）
LOOKBACK_DAYS = 30

# 收尾窗口：北京时间 23:45 之后那一轮才允许判红。
# 中证官网当日数据实测首现 17:01~23:21、中位 19:30，所以白天与傍晚的
# lag=1 属正常等待，不能报红；只有到当日收尾仍滞后才值得打扰。
FAIL_WINDOW = (23, 45)

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"


def bjnow():
    return datetime.now(BJ)


def _get_json(url, referer=None, timeout=40):
    h = {"User-Agent": UA}
    if referer:
        h["Referer"] = referer
    req = urllib.request.Request(url, headers=h)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def _ymd(d):
    return d.strftime("%Y%m%d")


def _parse_ymd(s):
    s = str(s).strip()
    if len(s) == 8 and s.isdigit():
        return date(int(s[:4]), int(s[4:6]), int(s[6:8]))
    try:
        return date.fromisoformat(s[:10])
    except Exception:
        return None


def official_latest(today):
    """官网 perf 接口返回的最新交易日。失败返回 (None, 原因)。"""
    url = (f"{CSINDEX_PERF}?indexCode={INDEX_CODE}"
           f"&startDate={_ymd(today - timedelta(days=LOOKBACK_DAYS))}"
           f"&endDate={_ymd(today)}")
    d = _get_json(url, referer="https://www.csindex.com.cn/")
    if str(d.get("code")) != "200":
        return None, f"官网返回 code={d.get('code')} msg={d.get('message')}"
    ds = [_parse_ymd(x.get("tradeDate")) for x in (d.get("data") or [])]
    ds = [x for x in ds if x]
    if not ds:
        return None, "官网近 30 天无任何数据行"
    return max(ds), ""


def reference_calendar(today):
    """真实交易日历（腾讯上证指数日线）。返回升序 date 列表。"""
    url = TX_KLINE.format(sym=TX_SYMBOL,
                          start=(today - timedelta(days=LOOKBACK_DAYS)).isoformat(),
                          end=today.isoformat())
    d = _get_json(url, timeout=30)
    node = ((d.get("data") or {}).get(TX_SYMBOL) or {})
    bars = node.get("qfqday") or node.get("day") or []
    out = []
    for b in bars:
        dt = _parse_ymd(b[0]) if b else None
        if dt:
            out.append(dt)
    return sorted(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true", help="忽略收尾窗口，立即按 lag 判定")
    ap.add_argument("--quiet", action="store_true", help="只输出结论行")
    args = ap.parse_args()

    now = bjnow()
    today = now.date()

    def log(msg):
        if not args.quiet:
            print(msg, flush=True)

    log(f"[哨兵] 当前北京 {now:%Y-%m-%d %H:%M:%S}")

    # ---------- 取两侧日期 ----------
    try:
        off, why = official_latest(today)
    except Exception as e:
        log(f"[哨兵] 官网探测失败（不判定）：{type(e).__name__}: {e}")
        return 0
    try:
        cal = reference_calendar(today)
    except Exception as e:
        log(f"[哨兵] 交易日历探测失败（不判定）：{type(e).__name__}: {e}")
        return 0

    if not cal:
        log("[哨兵] 交易日历为空（不判定）")
        return 0
    ref = cal[-1]

    if off is None:
        log(f"[哨兵] 官网最新交易日读不到（{why}）→ 参考交易日 {ref}，不判定")
        print(f"::warning::官网 perf 接口无数据，无法判定滞后（参考交易日 {ref}）")
        return 0

    lag_days = [d for d in cal if d > off]
    lag = len(lag_days)

    log(f"[哨兵] 官网最新交易日 {off} / 参考交易日 {ref} → 滞后 {lag} 个交易日")
    if lag:
        log(f"[哨兵] 缺失交易日：{', '.join(d.isoformat() for d in lag_days)}")

    # ---------- 判定 ----------
    if lag == 0:
        print(f"::notice::数据源正常：官网 {off} 已覆盖最新交易日 {ref}")
        return 0

    span = f"官网停在 {off}，参考交易日已到 {ref}，缺 {lag} 个交易日（{', '.join(d.isoformat() for d in lag_days)}）"

    in_fail_window = (now.hour, now.minute) >= FAIL_WINDOW
    should_fail = lag >= 2 and (in_fail_window or args.force) \
        and os.environ.get("STALE_FAIL", "1") != "0"

    if should_fail:
        msg = f"数据源滞后 {lag} 个交易日：{span}"
        log(f"[哨兵] 判红：{msg}")
        print(f"::error::{msg}")
        print(f"::error::页面会静默停留在旧日期；请确认中证官网是否停更（不是本仓库故障）。")
        return 1

    msg = f"数据源滞后 {lag} 个交易日：{span}"
    log(f"[哨兵] 告警（不改退出码）：{msg}")
    if lag >= 2 and not in_fail_window:
        log(f"[哨兵] 已满足 lag>=2，但未到收尾窗口 {FAIL_WINDOW[0]:02d}:{FAIL_WINDOW[1]:02d}，暂不判红")
    print(f"::warning::{msg}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception as e:            # 兜底：哨兵自己出问题绝不染红流水线
        print(f"[哨兵] 未预期异常（不判定）：{type(e).__name__}: {e}")
        sys.exit(0)
