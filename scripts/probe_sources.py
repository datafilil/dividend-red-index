#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""红利日报数据源可达性探测。

用途：在把日报迁到任何新环境（Antigravity Remote Agent、新的 CI runner、海外 VPS）
之前，先回答一个问题——这个环境能不能访问中证官网和东方财富。

只依赖标准库，不需要 pip install。退出码 0 = 全部关键源可用。

用法：
    python3 probe_sources.py
    python3 probe_sources.py --verbose     # 打印每个源的前 200 字节
"""

import json
import ssl
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone

BJ = timezone(timedelta(hours=8))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
TIMEOUT = 25

# 中证官网对 Referer 有校验，缺了可能返回空
CSINDEX_HEADERS = {"User-Agent": UA, "Referer": "https://www.csindex.com.cn/"}
EM_HEADERS = {"User-Agent": UA, "Referer": "https://quote.eastmoney.com/"}
TX_HEADERS = {"User-Agent": UA, "Referer": "https://gu.qq.com/"}

INDICATOR_URL = ("https://oss-ch.csindex.com.cn/static/html/csindex/public/uploads"
                 "/file/autofile/indicator/000922indicator.xls")
CONS_URL = ("https://oss-ch.csindex.com.cn/static/html/csindex/public/uploads"
            "/file/autofile/closeweight/000922closeweight.xls")

OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _get(url, headers, timeout=TIMEOUT):
    """返回 (status, body_bytes, elapsed_ms)。异常不抛出，转成 status=None。"""
    req = urllib.request.Request(url, headers=headers)
    ctx = ssl.create_default_context()
    t0 = time.time()
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
            body = r.read()
            return r.status, body, int((time.time() - t0) * 1000)
    except urllib.error.HTTPError as e:
        return e.code, b"", int((time.time() - t0) * 1000)
    except Exception as e:  # 超时 / DNS / TLS / 连接重置
        return None, repr(e).encode()[:160], int((time.time() - t0) * 1000)


def check_csindex_perf():
    """主标的与基准的日线。15 年全史都走这一个接口。"""
    end = datetime.now(BJ)
    start = end - timedelta(days=30)
    ok_all, detail = True, []
    for code in ("H00922", "H00985"):
        url = (f"https://www.csindex.com.cn/csindex-home/perf/index-perf"
               f"?indexCode={code}&startDate={start:%Y%m%d}&endDate={end:%Y%m%d}")
        st, body, ms = _get(url, CSINDEX_HEADERS)
        n, latest = 0, "-"
        if st == 200:
            try:
                d = json.loads(body.decode("utf-8"))
                rows = d.get("data") or []
                n = len(rows)
                ds = [str(x.get("tradeDate")) for x in rows if x.get("tradeDate")]
                if ds:
                    t = max(ds)
                    latest = f"{t[:4]}-{t[4:6]}-{t[6:8]}"
            except Exception:
                pass
        good = st == 200 and n > 0
        ok_all &= good
        detail.append(f"{code}:{n}根/{latest}" + ("" if good else " 失败"))
    return ok_all, f"HTTP {st} · {ms}ms · " + " / ".join(detail)


def check_xls(url, label, min_bytes=4096):
    """官网估值指标 / 成分股权重文件。固定 .xls，校验 OLE2 头。"""
    st, body, ms = _get(url, CSINDEX_HEADERS)
    size = len(body)
    ole = body[:8] == OLE2_MAGIC
    good = st == 200 and size >= min_bytes and ole
    return good, f"HTTP {st} · {ms}ms · {size} 字节 · OLE2={'是' if ole else '否'}"


def check_eastmoney_kline():
    """交易日历的备用源（生产脚本里腾讯是第一顺位，这里只是保底）。

    参数形态与 dividend_index_reminder.py 一致：secid=1.000001（上证指数）、beg/end，
    不要改成 lmt —— 商用 kline 接口不带 beg/end 时经常直接断连。
    """
    year = datetime.now(BJ).year
    url = ("https://push2his.eastmoney.com/api/qt/stock/kline/get"
           f"?secid=1.000001&klt=101&fqt=0&beg={year}0101&end={year}1231"
           "&fields1=f1,f2,f3&fields2=f51")
    st, body, ms = _get(url, EM_HEADERS)
    n = 0
    err = ""
    if st == 200:
        try:
            n = len((json.loads(body.decode("utf-8")).get("data") or {}).get("klines") or [])
        except Exception:
            pass
    elif st is None:
        err = " (" + body.decode("utf-8", "replace")[:60] + ")"
    return st == 200 and n > 0, f"HTTP {st} · {ms}ms · {n} 根{err}"


def check_eastmoney_clist():
    """全市场扫描的第一页。周更才用，但决定扫描能不能跑。"""
    url = ("https://push2.eastmoney.com/api/qt/clist/get"
           "?pn=1&pz=5&po=1&np=1&fltt=2&invt=2&fid=f3&fs=m:0+t:6,m:0+t:80&fields=f12,f14,f2")
    st, body, ms = _get(url, EM_HEADERS)
    n = 0
    if st == 200:
        try:
            n = len(((json.loads(body.decode("utf-8")).get("data") or {}).get("diff")) or [])
        except Exception:
            pass
    return st == 200 and n > 0, f"HTTP {st} · {ms}ms · {n} 行"


def check_eastmoney_datacenter():
    """分红送配数据，股息率的一路来源。"""
    url = ("https://datacenter-web.eastmoney.com/api/data/v1/get"
           "?reportName=RPT_SHAREBONUS_DET&columns=ALL&pageNumber=1&pageSize=5")
    st, body, ms = _get(url, EM_HEADERS)
    n = 0
    if st == 200:
        try:
            n = len((((json.loads(body.decode("utf-8")).get("result") or {})
                      .get("data")) or []))
        except Exception:
            pass
    return st == 200 and n > 0, f"HTTP {st} · {ms}ms · {n} 行"


def check_tencent_kline():
    """腾讯行情，作为东财的备用源。"""
    url = "https://web.ifzq.gtimg.cn/appstock/app/fqkline/get?param=sh000922,day,,,5,qfq"
    st, body, ms = _get(url, TX_HEADERS)
    n = 0
    if st == 200:
        try:
            d = json.loads(body.decode("utf-8")).get("data") or {}
            for v in d.values():
                if isinstance(v, dict) and v.get("day"):
                    n = len(v["day"])
                    break
        except Exception:
            pass
    return st == 200 and n > 0, f"HTTP {st} · {ms}ms · {n} 根"


CHECKS = [
    # (显示名, 探测函数, 是否关键, 不通时的实际后果)
    ("中证 perf（H00922/H00985 日线）", check_csindex_perf, True,
     "彻底跑不了：日线是全部指标的唯一来源"),
    ("中证 估值指标 .xls", lambda: check_xls(INDICATOR_URL, "indicator"), True,
     "股息率与官网 PE1 全部缺失"),
    ("中证 成分股权重 .xls", lambda: check_xls(CONS_URL, "closeweight"), True,
     "成分股表缺失（可用 cache/000922cons.json 降级）"),
    ("东财 K线（交易日历备用）", check_eastmoney_kline, False,
     "无影响：腾讯是主源，这条只是保底"),
    ("东财 全市场列表", check_eastmoney_clist, False,
     "全市场扫描退化，日报主体不受影响"),
    ("东财 分红送配", check_eastmoney_datacenter, False,
     "成分股股息率的参照口径退化"),
    ("腾讯 K线（备用）", check_tencent_kline, False,
     "交易日历失去主源，改由东财保底"),
]


def main():
    print(f"探测环境：{sys.platform} · Python {sys.version.split()[0]}")
    print(f"当前北京：{datetime.now(BJ):%Y-%m-%d %H:%M:%S}")
    print("=" * 74)

    failed = []
    for name, fn, critical, impact in CHECKS:
        try:
            ok, detail = fn()
        except Exception as e:
            ok, detail = False, f"探测本身出错：{type(e).__name__}: {e}"
        flag = "OK  " if ok else "FAIL"
        print(f"[{flag}] {name:<28} {detail}")
        if not ok:
            failed.append((name, critical, impact))

    print("=" * 74)
    crit = [x for x in failed if x[1]]
    if crit:
        print("结论：不可用 —— 关键源不通，不要在此环境上线。")
        for name, _c, impact in crit:
            print(f"  · {name}：{impact}")
        return 1

    print("结论：可以跑日报（关键源全部通过）。")
    if failed:
        print("以下可选源不通，按点降级，不影响主流程：")
        for name, _c, impact in failed:
            print(f"  · {name}：{impact}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
