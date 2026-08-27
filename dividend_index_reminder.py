#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中证红利(全收益)指数低频提醒 · 取数+计算+报告生成
数据源：中证指数有限公司官网  https://www.csindex.com.cn
       接口：/csindex-home/perf/index-perf  (免费、无需 Key、返回 16 年日线)
指标：
  - 主标的(全收益)收盘价 MA250 / MA350 / MA500
  - 主标的 40 日收益 − 基准(中证全指 000985) 40 日收益 = 40日收益差值
  - 主标的 PE 历史分位(由官网 peg 字段自算)
  - 股息率(官网每日更新的指数估值指标文件, 价格指数 000922 口径)
  - 近五年"收盘价低于 MA500"区间汇总
说明：
  - 主标的 H00922 中证红利全收益指数：官网权威直取(全收益=含分红再投资)。
  - 全部数据源均为中证指数官网(csindex.com.cn)及其官方指标文件，免费、无需授权。
  - 全部为离线自包含 HTML(SVG 图表,无外部依赖)。输出注明来源与"非投资建议"。
依赖：xlrd(解析官网指标 .xls, 已装于托管 Python；缺失时自动跳过股息率, 不中断主流程)
"""
import os, sys, json, math, datetime, urllib.request, urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))

# ============ 可配置区（标的切换只改这里）============
PRIMARY_CODE = "H00922"            # 中证红利全收益指数(全收益=含分红)
PRIMARY_NAME = "中证红利全收益指数"
BENCH_CODE   = "000985"            # 中证全指(基准)
BENCH_NAME   = "中证全指"
HIST_START   = "20000101"          # 尽量早取，保证 MA500 前置充足
INDICATOR_CODE = "000922"          # 股息率指标文件挂在价格指数上(成分与 H00922 一致)
INDICATOR_URL = ("https://oss-ch.csindex.com.cn/static/html/csindex/public/uploads"
                 f"/file/autofile/indicator/{INDICATOR_CODE}indicator.xls")
# ====================================================

MAS = [250, 350, 500]
RET_WINDOW = 40
HISTORY_FILE = os.path.join(HERE, "dividend_index_history.json")
REPORT_FILE  = os.path.join(HERE, "dividend_index_report.html")
SITE_DIR     = os.path.join(HERE, "site")   # 发布目录(复制为 index.html 供部署)
CACHE_DIR    = os.path.join(HERE, "cache")  # 全历史日线本地缓存(增量取数用)
CACHE_P      = os.path.join(CACHE_DIR, f"{PRIMARY_CODE.lower()}.json")
CACHE_B      = os.path.join(CACHE_DIR, f"{BENCH_CODE}.json")
CSINDEX_PERF = "https://www.csindex.com.cn/csindex-home/perf/index-perf"
DATA_NOTE = ("数据来源：中证指数有限公司官网 csindex.com.cn（指数表现接口 + 每日更新的指数估值指标文件），免费、无需授权。"
             "主标的 H00922 为全收益口径(含分红再投资)；基准 000985 中证全指。"
             "PE 分位由官网 peg 字段在自身历史中计算；股息率取自官网指标文件(000922 价格指数口径)。")

def fetch_index(code, start=HIST_START, end=None, allow_empty=False):
    """返回 [{date:datetime, close, peg}]，按日期升序。失败即报错(不伪造数据)。
    allow_empty=True 时允许区间内无数据(返回 [])——用于增量拉取(官网当日数据未更新时)。"""
    end = end or datetime.date.today().strftime("%Y%m%d")
    url = f"{CSINDEX_PERF}?indexCode={code}&startDate={start}&endDate={end}"
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Referer": "https://www.csindex.com.cn/",
    })
    d = None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=40) as r:
                d = json.loads(r.read().decode("utf-8"))
            if d.get("code") == 200 and d.get("data"):
                break
        except Exception as e:
            if attempt == 2:
                raise RuntimeError(f"中证官网取数失败 {code}: {e}")
        import time as _t; _t.sleep(1.5)
    if not d or str(d.get("code")) != "200":
        raise RuntimeError(f"中证官网返回异常 {code}: code={d.get('code') if d else None} msg={d.get('message') if d else None}")
    if not d.get("data"):
        if allow_empty:
            return []
        raise RuntimeError(f"中证官网 {code} 无有效数据")
    out = []
    for row in d["data"]:
        td = row.get("tradeDate")
        cl = row.get("close")
        if not td or cl is None:
            continue
        dt = datetime.datetime.strptime(str(td), "%Y%m%d").date()
        peg = row.get("peg")
        out.append({"date": dt, "close": float(cl), "peg": (float(peg) if peg not in (None, "", "null") else None)})
    out.sort(key=lambda x: x["date"])
    if not out:
        raise RuntimeError(f"中证官网 {code} 无有效数据")
    return out

# ---------- 全历史本地缓存 + 增量取数 ----------
def load_cache(path):
    """读缓存 [{"date":"YYYY-MM-DD","close":..,"peg":..}]，失败返回 []。"""
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def save_cache(path, bars):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(bars, f, ensure_ascii=False)

def get_index_cached(code, path):
    """全历史缓存 + 增量拉取缺口(缓存次日~今天)，合并去重后写回。
    返回 [{date: date对象, close, peg}] 升序。首跑无缓存时全量拉取。"""
    cached = load_cache(path)
    # 防御：缓存日期必须严格升序且无重复，否则视为损坏、全量重拉(避免静默错位)
    if cached:
        ds = [c["date"] for c in cached]
        if any(ds[i] >= ds[i + 1] for i in range(len(ds) - 1)):
            print(f"      {code} 缓存日期异常(重复/乱序), 全量重拉修复")
            cached = []
    start, mode = HIST_START, "全量(首跑)"
    if cached:
        last_dt = datetime.date.fromisoformat(cached[-1]["date"])
        if last_dt >= datetime.date.today():
            print(f"      {code} 缓存已含最新日({cached[-1]['date']}), 直接复用 bars={len(cached)}")
            return [{"date": datetime.date.fromisoformat(c["date"]),
                     "close": c["close"], "peg": c.get("peg")} for c in cached]
        start = (last_dt + datetime.timedelta(days=1)).strftime("%Y%m%d")
        mode = f"增量({cached[-1]['date']}次日→今天)"
    new_bars = fetch_index(code, start=start, allow_empty=True)
    if not new_bars:
        print(f"      {code} 官网在 {start} 后无新数据(当日未更新), 沿用缓存 bars={len(cached)}")
        return [{"date": datetime.date.fromisoformat(c["date"]),
                 "close": c["close"], "peg": c.get("peg")} for c in cached]
    merged = {c["date"]: c for c in cached}
    for nb in new_bars:
        key = nb["date"].isoformat()
        merged[key] = {"date": key, "close": nb["close"], "peg": nb["peg"]}
    bars = sorted(merged.values(), key=lambda x: x["date"])
    save_cache(path, bars)
    print(f"      {code} {mode} 拉 {len(new_bars)} 日 → 缓存 {len(bars)} 日 ({bars[0]['date']}~{bars[-1]['date']})")
    return [{"date": datetime.date.fromisoformat(x["date"]),
             "close": x["close"], "peg": x.get("peg")} for x in bars]

def fetch_indicator():
    """官网每日更新的指数估值指标文件(.xls)：返回按日期升序的
    [{date:'YYYYMMDD', pe1, pe2, dp1, dp2}]（股息率1/2, 单位:%）。
    失败返回 []——股息率是增强项，不中断主流程(不伪造数据)。"""
    try:
        import xlrd
    except ImportError:
        print("      [提示] 未安装 xlrd，跳过股息率(其余指标不受影响)")
        return []
    tmp = os.path.join(HERE, "indicator_tmp.xls")
    for attempt in range(3):
        try:
            req = urllib.request.Request(INDICATOR_URL, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
            with urllib.request.urlopen(req, timeout=40) as r:
                raw = r.read()
            if len(raw) < 512:
                raise RuntimeError(f"指标文件过小({len(raw)}B)")
            open(tmp, "wb").write(raw)
            sh = xlrd.open_workbook(tmp).sheet_by_index(0)
            def fnum(x):
                try: return float(x)
                except (TypeError, ValueError): return None
            rows = []
            for r_ in range(1, sh.nrows):
                v = [sh.cell_value(r_, c) for c in range(sh.ncols)]
                d = str(v[0]).split(".")[0]
                if len(d) != 8 or not d.isdigit():
                    continue
                rows.append({"date": d, "pe1": fnum(v[6]), "pe2": fnum(v[7]),
                             "dp1": fnum(v[8]), "dp2": fnum(v[9])})
            if rows:
                rows.sort(key=lambda x: x["date"])
                return rows
        except Exception as e:
            if attempt == 2:
                print(f"      [提示] 股息率指标文件获取失败: {e}")
                return []
            import time as _t; _t.sleep(1.5)
    return []

def sma_series(closes, w):
    """返回与 closes 等长的列表，前 w-1 个为 None。"""
    n = len(closes)
    res = [None] * n
    cum = 0.0
    for i, v in enumerate(closes):
        cum += v
        if i >= w:
            cum -= closes[i - w]
        if i >= w - 1:
            res[i] = cum / w
    return res

def pct_rank(series, value):
    vals = [x for x in series if x is not None]
    if not vals:
        return None
    below = sum(1 for x in vals if x <= value)
    return below / len(vals) * 100.0

# ---------- SVG ----------
def svg_line_chart(title, series, width=880, height=360, zero_line=False, y_precision=2, val_formatter=None, x_ticks=None):
    """x_ticks: [(x值, 标签)] —— 绘制 X 轴时间刻度(竖网格线+日期标签)。"""
    if not series:
        return ""
    allpts = [p for s in series for p in s["points"]]
    if not allpts:
        return ""
    xs = [p[0] for p in allpts]; ys = [p[1] for p in allpts]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    if zero_line:
        min_y = min(min_y, 0); max_y = max(max_y, 0)
    if max_y == min_y:
        max_y += 1; min_y -= 1
    pad = (max_y - min_y) * 0.08
    min_y -= pad; max_y += pad
    L, R, T, B = 60, 18, 34, 34
    pw, ph = width - L - R, height - T - B
    def mx(x): return L + (0 if max_x == min_x else (x - min_x) / (max_x - min_x)) * pw
    def my(y): return T + (1 - (y - min_y) / (max_y - min_y)) * ph
    def fmt(v): return val_formatter(v) if val_formatter else f"{v:.{y_precision}f}"
    svg = [f'<svg viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" font-family="system-ui,Segoe UI,Arial,sans-serif">']
    svg.append(f'<rect x="0" y="0" width="{width}" height="{height}" fill="#ffffff"/>')
    svg.append(f'<text x="{L}" y="20" font-size="14" font-weight="600" fill="#1f2937">{title}</text>')
    for i in range(5):
        gy = min_y + (max_y - min_y) * i / 4
        yy = my(gy)
        svg.append(f'<line x1="{L}" y1="{yy:.1f}" x2="{width-R}" y2="{yy:.1f}" stroke="#eef0f3" stroke-width="1"/>')
        svg.append(f'<text x="{L-6}" y="{yy+4:.1f}" font-size="11" fill="#9ca3af" text-anchor="end">{fmt(gy)}</text>')
    if zero_line:
        zy = my(0)
        svg.append(f'<line x1="{L}" y1="{zy:.1f}" x2="{width-R}" y2="{zy:.1f}" stroke="#cbd5e1" stroke-width="1" stroke-dasharray="4 3"/>')
    # X 轴时间刻度: 竖网格线
    if x_ticks:
        for tx, tlab in x_ticks:
            xx = mx(tx)
            if xx < L or xx > width - R:
                continue
            svg.append(f'<line x1="{xx:.1f}" y1="{T}" x2="{xx:.1f}" y2="{T+ph:.1f}" stroke="#eef0f3" stroke-width="1"/>')
    for s in series:
        if not s["points"]:
            continue
        d = " ".join(f"{'M' if i==0 else 'L'}{mx(x):.1f},{my(y):.1f}" for i,(x,y) in enumerate(s["points"]))
        dash = ' stroke-dasharray="5 3"' if s.get("dashed") else ""
        svg.append(f'<path d="{d}" fill="none" stroke="{s["color"]}" stroke-width="{s.get("width",1.8)}"{dash}/>')
    # X 轴时间刻度: 日期标签
    if x_ticks:
        for tx, tlab in x_ticks:
            xx = mx(tx)
            if xx < L or xx > width - R:
                continue
            anchor = "middle"
            if xx < L + 16: anchor = "start"
            elif xx > width - R - 16: anchor = "end"
            svg.append(f'<text x="{xx:.1f}" y="{height-12}" font-size="11" fill="#9ca3af" text-anchor="{anchor}">{tlab}</text>')
    lx = L + 8
    for s in series:
        svg.append(f'<rect x="{lx}" y="{T+2}" width="12" height="12" rx="2" fill="{s["color"]}"/>')
        svg.append(f'<text x="{lx+16}" y="{T+12}" font-size="11.5" fill="#374151">{s["name"]}</text>')
        lx += 22 + len(s["name"]) * 13 + 14
    svg.append('</svg>')
    return "\n".join(svg)

def year_ticks(dates, i_min, i_max, max_ticks=10):
    """取区间内每年首个交易日的 (索引, '年份') 作为刻度；过多时均匀抽稀。"""
    ticks = []
    last_y = None
    for i in range(max(0, i_min), min(len(dates), i_max + 1)):
        y = dates[i].year
        if y != last_y:
            ticks.append((i, str(y))); last_y = y
    if len(ticks) > max_ticks:
        step = math.ceil(len(ticks) / max_ticks)
        ticks = ticks[::step]
    return ticks

# ================= 主流程 =================
print(f"[1/4] 取数 {PRIMARY_CODE} {PRIMARY_NAME} (缓存增量) ...")
p = get_index_cached(PRIMARY_CODE, CACHE_P)
print(f"[2/4] 取数 {BENCH_CODE} {BENCH_NAME} (缓存增量) ...")
b = get_index_cached(BENCH_CODE, CACHE_B)
print(f"[2b] 股息率指标文件(官网每日更新, {INDICATOR_CODE}) ...")
ind = fetch_indicator()
ind_last = ind[-1] if ind else None
if ind_last:
    print(f"      D/P2={ind_last['dp2']}% PE1={ind_last['pe1']} (最新 {ind_last['date']})")
else:
    print("      股息率不可得(报告中该栏显示—)")

p_close = [x["close"] for x in p]
p_dates = [x["date"] for x in p]
b_close = {x["date"]: x["close"] for x in b}
b_peg   = {x["date"]: x["peg"] for x in b}

ma = {w: sma_series(p_close, w) for w in MAS}
last_close = p_close[-1]
last_date = p_dates[-1]

# 40日收益差值（按日期对齐）
diff_series = []; ret40_p = []; ret40_b = []
for i in range(RET_WINDOW, len(p)):
    d = p_dates[i]
    bc = b_close.get(d)
    if bc is None or b_close.get(p_dates[i - RET_WINDOW]) is None:
        continue
    rp = p_close[i] / p_close[i - RET_WINDOW] - 1
    rb = bc / b_close[p_dates[i - RET_WINDOW]] - 1
    diff = rp - rb
    diff_series.append((i, diff)); ret40_p.append((i, rp)); ret40_b.append((i, rb))
cur_rp = ret40_p[-1][1]; cur_rb = ret40_b[-1][1]; cur_diff = diff_series[-1][1]
diff_pct = pct_rank([x[1] for x in diff_series], cur_diff)

# PE 历史分位
peg_series = [x["peg"] for x in p if x["peg"] is not None]
cur_peg = p[-1]["peg"]
peg_pct = pct_rank(peg_series, cur_peg) if cur_peg is not None else None

print(f"[3/4] 计算完成。收盘={last_close:.2f} MA="
      f"{ma[250][-1]:.2f}/{ma[350][-1]:.2f}/{ma[500][-1]:.2f} 40日差值={cur_diff*100:.2f}% PE分位={peg_pct:.0f}%")

# 近五年"低于 MA500"汇总
cut = last_date - datetime.timedelta(days=int(365.25*5))
win = [(i, p_dates[i], p_close[i]) for i in range(len(p)) if p_dates[i] >= cut]
below = [(i, dt, cl) for (i, dt, cl) in win if ma[500][i] is not None and cl < ma[500][i]]
below_rate = len(below) / len(win) * 100 if win else 0
# 连续破位段
episodes = []
if below:
    start_i, prev_i = below[0][0], below[0][0]
    run = [below[0]]
    for it in below[1:]:
        if it[0] == prev_i + 1:
            run.append(it); prev_i = it[0]
        else:
            episodes.append(run); run = [it]; start_i = it[0]; prev_i = it[0]
    episodes.append(run)
def max_dev(run):
    return min(cl / ma[500][i] - 1 for (i, dt, cl) in run) * 100
ep_stats = []
for ep in episodes:
    ep_stats.append({
        "start": ep[0][1], "end": ep[-1][1], "n": len(ep),
        "max_dev": max_dev(ep),
        "end_close": ep[-1][2], "end_ma": ma[500][ep[-1][0]],
    })
ep_stats.sort(key=lambda e: e["n"], reverse=True)

# 各均线近五年破位率(快速对照)
below_rates = {}
for w in MAS:
    bl = sum(1 for (i, dt, cl) in win if ma[w][i] is not None and cl < ma[w][i])
    below_rates[w] = bl / len(win) * 100 if win else 0

# 历史快照累计
history = []
if os.path.exists(HISTORY_FILE):
    try: history = json.load(open(HISTORY_FILE, encoding="utf-8"))
    except Exception: history = []
snap = {"date": last_date.isoformat(),
        "close": round(last_close, 3),
        "ma250": round(ma[250][-1], 3), "ma350": round(ma[350][-1], 3), "ma500": round(ma[500][-1], 3),
        "ret40_p": round(cur_rp*100, 3), "ret40_b": round(cur_rb*100, 3),
        "diff": round(cur_diff*100, 3), "pe": round(cur_peg, 2) if cur_peg else None,
        "pe_pct": round(peg_pct, 1) if peg_pct else None,
        "dy1": (ind_last["dp1"] if ind_last else None),
        "dy2": (ind_last["dp2"] if ind_last else None)}
history = [h for h in history if h.get("date") != snap["date"]]
history.append(snap); history.sort(key=lambda x: x["date"])
json.dump(history, open(HISTORY_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
print(f"[4/4] 近五年破MA500占比={below_rate:.1f}% ({len(episodes)}段)")

# ================= 信号 =================
def pos(v, ma_v):
    if ma_v is None or v is None: return "—"
    return f"{(v-ma_v)/ma_v*100:+.2f}%"
diff_state = "跑赢基准" if cur_diff > 0 else "跑输基准"

# ================= 图表 =================
# 用近五年做横轴
x0 = len(p) - len(win)
c1 = [{"name": f"{PRIMARY_NAME}收盘", "color": "#2563eb",
       "points": [(i, p_close[i]) for i in range(x0, len(p))], "width": 1.8}]
for w in MAS:
    c1.append({"name": f"MA{w}", "color": ["#f59e0b","#10b981","#ef4444"][MAS.index(w)],
               "points": [(i, ma[w][i]) for i in range(x0, len(p)) if ma[w][i] is not None], "width": 1.5})
svg1 = svg_line_chart(f"{PRIMARY_NAME} 收盘与长期均线 (近5年, MA250/350/500)", c1, y_precision=0,
                      x_ticks=year_ticks(p_dates, x0, len(p) - 1))

c2 = [{"name": "40日收益差值(%)", "color": "#7c3aed",
       "points": [(i, v*100) for i, v in diff_series], "width": 1.6}]
svg2 = svg_line_chart(f"{PRIMARY_NAME} − {BENCH_NAME} 40日收益差值(%)", c2, zero_line=True, y_precision=1,
                      x_ticks=year_ticks(p_dates, diff_series[0][0], diff_series[-1][0], max_ticks=9))

print(f"[5] 生成报告 ...")

def fmt_pct(x): return ("—" if x is None else f"{x*100:+.2f}%")
def fmt_num(x): return ("—" if x is None else f"{x:.2f}")

hist_rows = ""
for h in history[-12:]:
    hist_rows += (f"<tr><td>{h['date']}</td><td>{h['close']}</td>"
                  f"<td>{h.get('ma250')}</td><td>{h.get('ma350')}</td><td>{h.get('ma500')}</td>"
                  f"<td class='{'pos' if (h.get('diff') or 0)>=0 else 'neg'}'>{h.get('diff')}</td>"
                  f"<td>{h.get('pe_pct')}</td><td>{h.get('dy2') if h.get('dy2') is not None else '—'}</td></tr>")

ep_rows = ""
for e in ep_stats[:15]:
    ep_rows += (f"<tr><td>{e['start']}</td><td>{e['end']}</td><td>{e['n']}</td>"
                f"<td class='neg'>{e['max_dev']:.2f}%</td>"
                f"<td>{e['end_close']:.2f}</td><td>{e['end_ma']:.2f}</td></tr>")
if not ep_rows:
    ep_rows = "<tr><td colspan='6' style='text-align:center;color:#16a34a'>近五年无破位</td></tr>"

html = f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>红利指数低频提醒 · {last_date}</title>
<style>
*{{box-sizing:border-box}}
body{{font-family:system-ui,'Segoe UI','PingFang SC','Microsoft YaHei',Arial,sans-serif;
  margin:0;background:#f5f7fa;color:#1f2937;padding:24px}}
.wrap{{max-width:980px;margin:0 auto;background:#fff;border-radius:12px;
  box-shadow:0 2px 12px rgba(0,0,0,.06);overflow:hidden}}
.head{{padding:20px 24px;border-bottom:1px solid #eef0f3}}
.head h1{{margin:0 0 4px;font-size:20px}}
.head .sub{{color:#6b7280;font-size:13px}}
.ok{{margin:16px 24px;padding:10px 14px;background:#ecfdf5;border:1px solid #a7f3d0;
  border-radius:8px;font-size:13px;color:#065f46}}
.refbox{{margin:16px 24px;padding:14px 16px;background:#f8fafc;border:1px solid #e2e8f0;
  border-radius:8px;font-size:13px;color:#334155;line-height:1.7}}
.reftitle{{font-weight:600;color:#0f172a;margin-bottom:8px}}
.refnote{{color:#94a3b8;font-size:12px;margin-top:8px}}
.grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;padding:0 24px 8px}}
.card{{background:#f9fafb;border:1px solid #eef0f3;border-radius:10px;padding:14px}}
.card .k{{font-size:12px;color:#6b7280}}
.card .v{{font-size:20px;font-weight:700;margin-top:4px}}
.card .d{{font-size:12px;margin-top:2px}}
.card.mas{{grid-column:span 1}}
.ma-flex{{display:flex;gap:10px;margin-top:6px}}
.ma500{{flex:1 1 50%;border-right:1px solid #e5e7eb;padding-right:6px}}
.ma500 .v{{font-size:19px;line-height:1.25;margin-top:2px}}
.maside{{flex:1 1 50%;display:flex;flex-direction:column;gap:8px}}
.maitem{{font-size:12.5px;color:#374151;line-height:1.4}}
.maitem b{{font-size:13.5px;font-weight:700;margin-left:4px}}
.mk{{color:#6b7280;font-size:11.5px}}
.maitem .d{{font-size:11.5px}}
.pos{{color:#dc2626}} .neg{{color:#16a34a}}
section{{padding:8px 24px 4px}}
section h2{{font-size:15px;margin:18px 0 8px;color:#374151}}
table{{width:100%;border-collapse:collapse;font-size:13px;margin-top:6px}}
th,td{{padding:8px 10px;text-align:right;border-bottom:1px solid #f0f2f5}}
th:first-child,td:first-child{{text-align:left}}
th{{color:#6b7280;font-weight:600;background:#fafbfc}}
.note{{padding:16px 24px 24px;color:#9ca3af;font-size:12px;line-height:1.6}}
.chart{{padding:4px 24px}}
</style></head>
<body><div class="wrap">
<div class="head">
  <h1>红利指数低频提醒</h1>
  <div class="sub">主标的：{PRIMARY_CODE} {PRIMARY_NAME} ｜ 基准：{BENCH_CODE} {BENCH_NAME} ｜ 生成日期：{last_date}</div>
</div>
<div class="ok"><b>说明：</b>主标的 <b>H00922 中证红利全收益指数</b>（含分红再投资），基准 <b>000985 中证全指</b>。全部数据来自中证指数官网（免费、无需授权，含 2000 年至今完整历史）。</div>

<div class="grid">
  <div class="card"><div class="k">最新收盘 ({last_date})</div><div class="v">{last_close:.2f}</div><div class="d">全收益口径</div></div>
  <div class="card mas"><div class="k">长期均线 MA250/350/500</div>
    <div class="ma-flex">
      <div class="ma500">
        <div class="mk">MA500</div>
        <div class="v">{fmt_num(ma[500][-1])}</div>
        <div class="d {'pos' if last_close>ma[500][-1] else 'neg'}">{pos(last_close,ma[500][-1])}</div>
      </div>
      <div class="maside">
        <div class="maitem"><span class="mk">MA250</span><b>{fmt_num(ma[250][-1])}</b><div class="d {'pos' if last_close>ma[250][-1] else 'neg'}">{pos(last_close,ma[250][-1])}</div></div>
        <div class="maitem"><span class="mk">MA350</span><b>{fmt_num(ma[350][-1])}</b><div class="d {'pos' if last_close>ma[350][-1] else 'neg'}">{pos(last_close,ma[350][-1])}</div></div>
      </div>
    </div></div>
  <div class="card"><div class="k">40日收益差值</div><div class="v {'pos' if cur_diff>=0 else 'neg'}">{fmt_pct(cur_diff)}</div><div class="d">{diff_state} · 历史分位 {'—' if diff_pct is None else f'{diff_pct:.0f}%'}（自2000年）</div></div>
</div>
<div class="grid">
  <div class="card"><div class="k">PE(TTM) 历史分位</div><div class="v">{('—' if peg_pct is None else f'{peg_pct:.0f}%')}</div><div class="d">PE={('—' if cur_peg is None else f'{cur_peg:.2f}')}</div></div>
  <div class="card"><div class="k">股息率(计算用股本 D/P2)</div><div class="v">{('—' if not ind_last or ind_last['dp2'] is None else f"{ind_last['dp2']:.2f}%")}</div><div class="d">官网指标文件(000922)</div></div>
  <div class="card"><div class="k">近5年破MA500占比</div><div class="v">{below_rate:.1f}%</div><div class="d">共 {len(episodes)} 段</div></div>
</div>

<div class="chart">{svg1}</div>
<div class="chart">{svg2}</div>

<section><h2>近五年"收盘价低于 MA500"区间汇总</h2>
<div class="refbox">
  <div class="reftitle">统计口径</div>
  <p>区间：{cut} ~ {last_date}（近五年，共 {len(win)} 个交易日）。以<b>全历史</b>计算 MA500（前置充足，无截断误差）。
  破位定义为当日收盘价 &lt; 当日 MA500；连续交易日计为一段。</p>
  <table><thead><tr><th>均线</th><th>近五年破位占比</th><th>说明</th></tr></thead><tbody>
    <tr><td>MA250</td><td>{below_rates[250]:.1f}%</td><td>覆盖近 ~4.7 年(前置250)</td></tr>
    <tr><td>MA350</td><td>{below_rates[350]:.1f}%</td><td>覆盖近 ~4.5 年(前置350)</td></tr>
    <tr><td>MA500</td><td class="neg">{below_rate:.1f}%</td><td>覆盖完整近 5 年(前置500)</td></tr>
  </tbody></table>
  <p style="margin-top:10px"><b>破 MA500 明细（按持续天数降序，前 15 段）：</b></p>
  <table><thead><tr><th>起始</th><th>结束</th><th>天数</th><th>最深偏离</th><th>末收盘</th><th>末 MA500</th></tr></thead>
  <tbody>{ep_rows}</tbody></table>
  <p class="refnote">提示：单日假破位噪音较大；若作提醒条件，建议"连续 ≥3 日破位 且 偏离 &gt;2%"再触发，可过滤短假破位。</p>
</div></section>

<section><h2>历史快照（最近 {min(12,len(history))} 次）</h2>
<table><thead><tr><th>日期</th><th>收盘</th><th>MA250</th><th>MA350</th><th>MA500</th><th>40日差值%</th><th>PE分位%</th><th>股息率%</th></tr></thead>
<tbody>{hist_rows}</tbody></table></section>

<div class="note">{DATA_NOTE}<br>本报告由自动化脚本生成，仅供研究与跟踪参考，<b>不构成任何投资建议</b>。市场有风险，投资需谨慎。</div>
</div></body></html>"""

open(REPORT_FILE, "w", encoding="utf-8").write(html)

# ---------- 同步到发布目录(供云端部署) ----------
try:
    import shutil
    os.makedirs(SITE_DIR, exist_ok=True)
    shutil.copyfile(REPORT_FILE, os.path.join(SITE_DIR, "index.html"))
    print("site  ->", os.path.join(SITE_DIR, "index.html"))
except Exception as e:
    print(f"[提示] 发布目录同步失败(不影响本地报告): {e}")

print("DONE. report ->", REPORT_FILE)
