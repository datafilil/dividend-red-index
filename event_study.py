#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""⑤ 指数纳入效应事件研究（中证红利 000922 实证）

框架：以 ADJ_HISTORY 的历年「纳入(ins)」名单为事件，每只新纳入个股以
「公告日 ann(t=0) → 生效日 eff」为事件窗口，计算相对红利指数(000922)的
累计异常收益 CAR，验证「纳入效应」是否可交易。

- 异常收益 AR_t = r_个股,t − r_红利指数,t（基准=所加入指数，刻画指数基金调仓效应）
- 事件窗口：[-5,0] 事前泄露；[0,Te] 公告→生效；[0,+5/+10/+20] 持有
- 数据：东财 push2his 日线(前复权) + 中证全指(000985)名称→代码映射解析历史名
- 缓存：namemap / kline 落盘，断网可续跑；结果写 event_study.html + 控制台摘要

用法：python event_study.py
"""
import os, re, ast, json, ssl, time, urllib.request, datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
REMINDER = os.path.join(ROOT, "dividend_index_reminder.py")
if not os.path.exists(REMINDER):
    REMINDER = os.path.join(ROOT, "..", "dividend_index_reminder.py")
CACHE = os.path.join(ROOT, "event_study_cache")
os.makedirs(CACHE, exist_ok=True)

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

BENCH_SECID = "1.000922"   # 红利指数(基准=所加入指数)

def http_get(url, timeout=25, retries=6):
    last = None
    for attempt in range(1, retries + 1):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0", "Referer": "https://quote.eastmoney.com/"})
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception as ex:
            last = ex
            if attempt < retries:
                time.sleep(3 * attempt)
    raise last

# ---------- 1) 解析 ADJ_HISTORY（不执行主脚本，仅 ast 解析字面量） ----------
def load_adj_history():
    src = open(REMINDER, encoding="utf-8").read()
    tree = ast.parse(src)
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name) and t.id == "ADJ_HISTORY":
                    return ast.literal_eval(node.value)
    raise RuntimeError("ADJ_HISTORY not found in " + REMINDER)

# ---------- 2) 全指名称→代码映射 ----------
def build_name_map():
    p = os.path.join(CACHE, "namemap.json")
    if os.path.exists(p):
        return json.load(open(p, encoding="utf-8"))
    url = ("https://datacenter-web.eastmoney.com/api/data/v1/get?reportName=RPT_INDEX_CONSTITUENT"
           "&columns=SECURITY_CODE,SECURITY_NAME_ABBR&filter=(INDEX_CODE%3D%22000985%22)&pageSize=6000")
    j = http_get(url)
    data = (j.get("result") or {}).get("data") or []
    m = {x["SECURITY_NAME_ABBR"]: x["SECURITY_CODE"] for x in data}
    json.dump(m, open(p, "w", encoding="utf-8"), ensure_ascii=False)
    print(f"  全指名称映射: {len(m)} 只")
    return m

def extract_code(name):
    """'天健集团(000090)' -> ('000090','天健集团')；裸名 -> (None, name)"""
    mm = re.search(r"\((\d{6})\)", name)
    if mm:
        return mm.group(1), name.split("(")[0]
    return None, name

def resolve_events(namemap):
    evs = []
    for h in load_adj_history():
        ann, eff, year = h["ann"], h["eff"], h["year"]
        for raw in h.get("ins", []):
            code, nm = extract_code(raw)
            if not code:
                code = namemap.get(nm)
            rec = {"year": year, "name": nm, "code": code, "ann": ann, "eff": eff}
            if not code:
                rec["skip"] = "未解析代码"
            evs.append(rec)
    return evs

# ---------- 3) K线（前复权日线，缓存） ----------
def secid_of(code):
    return ("1." if code.startswith("6") else "0.") + code

def _tx_code(code):
    return ("sh" if code.startswith("6") else "sz") + code

def fetch_kline(code, beg, end):
    """beg/end: 'YYYYMMDD'。优先腾讯历史K线(稳定)，东财兜底；返回 [(date, close)]。
    缓存仅落非空集，避免把失败误当空结果缓存。"""
    fn = os.path.join(CACHE, f"kl_{code}_{beg}_{end}.json")
    if os.path.exists(fn):
        return json.load(open(fn, encoding="utf-8"))
    b = f"{beg[:4]}-{beg[4:6]}-{beg[6:]}"
    e = f"{end[:4]}-{end[4:6]}-{end[6:]}"
    tx = _tx_code(code)
    recs = None
    try:
        url = f"https://web.ifzq.gtimg.cn/appstock/app/fqkline/get?param={tx},day,{b},{e},320,qfq"
        j = http_get(url)
        node = (j.get("data") or {}).get(tx) or {}
        arr = node.get("qfqday") or node.get("day") or []
        if arr:
            recs = [(row[0], float(row[2])) for row in arr]   # date, close
    except Exception:
        recs = None
    if not recs:  # 东财兜底
        secid = secid_of(code)
        try:
            url2 = (f"https://push2his.eastmoney.com/api/qt/stock/kline/get?secid={secid}"
                    f"&fields1=f1,f2&fields2=f51,f52,f53,f54,f55,f56&klt=101&fqt=1&beg={beg}&end={end}")
            j = http_get(url2)
            k = (j.get("data") or {}).get("klines") or []
            recs = [(row.split(",")[0], float(row.split(",")[2])) for row in k]
        except Exception:
            recs = None
    if recs:
        json.dump(recs, open(fn, "w", encoding="utf-8"))
    return recs or []

def returns(series):
    """date -> 日收益（基于前收）"""
    d = {}
    for i in range(1, len(series)):
        d[series[i][0]] = series[i][1] / series[i - 1][1] - 1.0
    return d

def d2i(s):
    return datetime.date.fromisoformat(s).toordinal()

# ---------- 4) 单事件研究 ----------
def study_one(ev, idx_series):
    code = ev["code"]
    a, e = d2i(ev["ann"]), d2i(ev["eff"])
    beg = datetime.date.fromordinal(a - 45).strftime("%Y%m%d")
    end = datetime.date.fromordinal(e + 45).strftime("%Y%m%d")
    stk = fetch_kline(code, beg, end)
    if len(stk) < 2:
        return None
    r_stk, r_idx = returns(stk), returns(idx_series)
    idx_dates = [d for d, _ in idx_series]
    pos = {d: i for i, d in enumerate(idx_dates)}
    start = min((pos[d] for d in idx_dates if d2i(d) >= a), default=None)
    e_pos = max((pos[d] for d in idx_dates if d2i(d) <= e), default=None)
    if start is None or e_pos is None:
        return None

    def ar_between(i0, i1):
        if i0 < 0 or i1 >= len(idx_dates):
            return None
        tot, n = 0.0, 0
        for i in range(i0, i1 + 1):
            d = idx_dates[i]
            if d in r_stk:
                tot += r_stk[d] - r_idx[d]
                n += 1
        return (tot, n)

    def ar_off(n1, n2):
        return ar_between(start + n1, start + n2)

    car_ae = ar_between(start, e_pos)                 # [公告, 生效]
    car05 = ar_off(0, 5)
    car10 = ar_off(0, 10)
    car20 = ar_off(0, 20)
    car_pre = ar_off(-5, 0)                            # 事前泄露
    out = {}
    for k, v in (("car_ae", car_ae), ("car05", car05), ("car10", car10),
                 ("car20", car20), ("car_pre", car_pre)):
        out[k] = (round(v[0] * 100, 2), v[1]) if v else None
    return out

# ---------- 5) 聚合 + 报告 ----------
def aggregate(rows):
    stats = {}
    for key in ("car_ae", "car05", "car10", "car20", "car_pre"):
        vals = [r["res"][key][0] for r in rows if r.get("res") and r["res"].get(key)]
        if vals:
            vals_s = sorted(vals)
            med = vals_s[len(vals_s) // 2]
            stats[key] = {
                "n": len(vals), "mean": round(sum(vals) / len(vals), 2),
                "median": med, "hit": round(sum(1 for v in vals if v > 0) / len(vals) * 100, 1),
                "min": round(min(vals), 2), "max": round(max(vals), 2)}
        else:
            stats[key] = None
    return stats

def render_html(rows, stats):
    def fmt(v):
        if not v:
            return "—"
        cls = "pos" if v[0] > 0 else ("neg" if v[0] < 0 else "")
        return f"<span class='{cls}'>{v[0]:+.2f}%</span> <span class='dim'>({v[1]}d)</span>"
    body = ""
    for r in rows:
        res = r.get("res")
        if not res:
            body += (f"<tr><td>{r['year']}</td><td>{r['code'] or '—'}</td><td>{r['name']}</td>"
                     f"<td>{r['ann']}</td><td>{r['eff']}</td>"
                     f"<td colspan='5' class='dim'>{r.get('skip','无数据')}</td></tr>")
            continue
        body += (f"<tr><td>{r['year']}</td><td>{r['code']}</td><td>{r['name']}</td>"
                 f"<td>{r['ann']}</td><td>{r['eff']}</td>"
                 f"<td>{fmt(res['car_pre'])}</td><td>{fmt(res['car_ae'])}</td>"
                 f"<td>{fmt(res['car05'])}</td><td>{fmt(res['car10'])}</td><td>{fmt(res['car20'])}</td></tr>")
    def srow(k, label):
        s = stats.get(k)
        if not s:
            return f"<tr><td>{label}</td><td colspan='5' class='dim'>样本不足</td></tr>"
        return (f"<tr><td><b>{label}</b></td>"
                f"<td>{s['mean']:+.2f}%</td><td>{s['median']:+.2f}%</td><td>{s['hit']}%</td>"
                f"<td>{s['min']:+.2f}%</td><td>{s['max']:+.2f}%</td></tr>")
    html = f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>中证红利 纳入效应事件研究</title>
<style>
*{{box-sizing:border-box}} body{{font-family:system-ui,'PingFang SC','Microsoft YaHei',Arial,sans-serif;
margin:0;background:#f5f7fa;color:#1f2937;padding:24px}}
.wrap{{max-width:1000px;margin:0 auto;background:#fff;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,.06);padding:22px}}
h1{{font-size:19px;margin:0 0 4px}} .sub{{color:#6b7280;font-size:13px;margin-bottom:14px}}
.pos{{color:#dc2626}} .neg{{color:#16a34a}} .dim{{color:#94a3b8;font-size:12px}}
table{{border-collapse:collapse;width:100%;font-size:13px;margin:6px 0}}
th,td{{border:1px solid #eef0f3;padding:5px 7px;text-align:center}}
th{{background:#f8fafc;color:#475569}} .ref{{background:#f8fafc;color:#475569;font-size:12px;padding:10px;border-radius:8px;margin-top:12px}}
b{{color:#0f172a}}
</style></head><body><div class="wrap">
<h1>中证红利(000922) 指数纳入效应事件研究</h1>
<p class="sub">事件 = 历年定期调整「纳入(ins)」名单；基准 = 红利指数自身(000922)；异常收益 AR=个股−指数；CAR 为区间累加。</p>
<h2>整体统计（CAR，单位为 %）</h2>
<table><thead><tr><th>窗口</th><th>均值</th><th>中位数</th><th>胜率(>0)</th><th>最小</th><th>最大</th></tr></thead>
<tbody>
{srow('car_pre','[-5,0] 事前泄露')}
{srow('car_ae','[公告→生效] 主调仓窗口')}
{srow('car05','[0,+5] 持有')}
{srow('car10','[0,+10] 持有')}
{srow('car20','[0,+20] 持有')}
</tbody></table>
<h2>逐事件明细</h2>
<table><thead><tr><th>年份</th><th>代码</th><th>名称</th><th>公告日</th><th>生效日</th>
<th>事前[-5,0]</th><th>公告→生效</th><th>[0,+5]</th><th>[0,+10]</th><th>[0,+20]</th></tr></thead>
<tbody>{body}</tbody></table>
<div class="ref"><b>方法论与口径：</b>① 仅研究「纳入」事件（剔除事件为对称负向，未纳入本表）；
② 收益用前复权收盘价日收益，AR=个股日收益−红利指数日收益，CAR 为窗口内 AR 累加；
③ 个股停牌日不计入 CAR（跳过该日，避免用指数收益冒充）；
④ 基准选红利指数本身——这正是经典「指数效应」定义（被纳入个股相对所加入指数的超额，由指数基金被动调仓驱动）；
⑤ 数据来源东方财富行情，事件日期取自中证指数官网公告（ADJ_HISTORY）；2022/2023 年官方未长期存档完整名单，部分仅部分恢复。
<b>可交易性结论：</b>若 [公告→生效] 窗口均值 CAR 显著为正且胜率较高，则「公告日买入、生效日前卖出」是可执行策略；
若效应主要集中在事前（[-5,0]），则信息已提前反映，公开策略空间收窄。</div>
</div></body></html>"""
    return html

def main():
    print("=== ⑤ 指数纳入效应事件研究 ===")
    namemap = build_name_map()
    evs = resolve_events(namemap)
    print(f"  事件总数: {len(evs)}；可解析代码: {sum(1 for e in evs if e['code'])}")
    # 基准指数按年取（每年约 15 个月窗口，分拆避免大请求被代理丢弃；落盘缓存）
    idx_by_year = {}
    years = sorted({ev["year"] for ev in evs if ev["code"]})
    for y in years:
        yb = datetime.date(y - 1, 11, 1).strftime("%Y%m%d")
        ye = datetime.date(y + 1, 1, 31).strftime("%Y%m%d")
        try:
            idx_by_year[y] = fetch_kline("000922", yb, ye)
            print(f"  基准红利指数K线({y}区间): {len(idx_by_year[y])} 根")
        except Exception as ex:
            print(f"  基准指数 {y} 获取失败 {ex}（该年事件跳过）")
    rows = []
    for i, ev in enumerate(evs):
        if not ev["code"]:
            rows.append(ev)
            continue
        idx = idx_by_year.get(ev["year"])
        if not idx:
            ev["skip"] = "基准指数缺失"
            rows.append(ev)
            continue
        try:
            res = study_one(ev, idx)
            ev["res"] = res
            rows.append(ev)
            tag = f"CAR[公告→生效]={res['car_ae'][0]:+.2f}%" if res and res["car_ae"] else "无窗口"
            print(f"  [{i+1}/{len(evs)}] {ev['year']} {ev['name']}({ev['code']}) {tag}")
        except Exception as ex:
            ev["skip"] = f"ERR:{type(ex).__name__}"
            rows.append(ev)
            print(f"  [{i+1}/{len(evs)}] {ev['name']} 失败 {ex}")
    stats = aggregate(rows)
    html = render_html(rows, stats)
    out = os.path.join(ROOT, "event_study.html")
    open(out, "w", encoding="utf-8").write(html)
    print(f"\n报告已写: {out}")
    print("=== 整体统计 ===")
    for k, label in (("car_pre","事前[-5,0]"),("car_ae","公告→生效"),("car05","[0,+5]"),
                     ("car10","[0,+10]"),("car20","[0,+20]")):
        s = stats.get(k)
        if s:
            print(f"  {label:12s} 均值={s['mean']:+.2f}% 中位={s['median']:+.2f}% 胜率={s['hit']}% n={s['n']}")
    # 可交易性结论
    ae = stats.get("car_ae")
    if ae:
        verdict = "可交易(公告买入→生效卖出)" if ae["mean"] > 1.0 and ae["hit"] >= 55 else \
                  ("效应偏弱/已被提前反映" if ae["mean"] <= 1.0 else "存在正向效应但胜率一般")
        print(f"  结论: {verdict}")

if __name__ == "__main__":
    main()
