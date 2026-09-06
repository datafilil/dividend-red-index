#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中证红利(全收益)指数低频提醒 · 取数+计算+报告生成
数据源：中证指数有限公司官网  https://www.csindex.com.cn
       接口：/csindex-home/perf/index-perf  (免费、无需 Key、返回 16 年日线)
指标：
  - 主标的(全收益)收盘价 MA250 / MA350 / MA500
  - 主标的 40 日收益 − 基准(中证全指全收益 H00985) 40 日收益 = 40日收益差值
  - 主标的 PE 历史分位(由官网 peg 字段自算)
  - 股息率(官网每日更新的指数估值指标文件, 价格指数 000922 口径)
  - 近五年"收盘价低于 MA500"区间汇总
  - 成分股(官网权重文件)：前十大权重 + 股息率(最近会计年度/现价, 东财参照口径) + 定期调整预告 + 近五年调整史
说明：
  - 主标的 H00922 中证红利全收益指数：官网权威直取(全收益=含分红再投资)。
  - 基准 H00985 中证全指全收益：口径与主标的保持一致(全收益 − 全收益)，避免"全收益减价格"
    造成红利超额收益被系统性高估。如需改用价格口径，务必同时把 PRIMARY_CODE 换成 000922。
  - 滚动窗口一律在"主标的与基准的交易日交集"上按下标滚动，任一源缺失/多余交易日都不会错位。
  - 写入缓存前会清洗非交易日与幽灵行(详见 sanitize_bars)，并在尾部做数据质量校验。
  - 全部数据源均为中证指数官网(csindex.com.cn)及其官方指标文件，免费、无需授权。
  - 全部为离线自包含 HTML(SVG 图表,无外部依赖)。输出注明来源与"非投资建议"。
依赖：xlrd(解析官网指标 .xls, 已装于托管 Python；缺失时自动跳过股息率, 不中断主流程)
"""
import os, sys, json, math, time, datetime, urllib.request, urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))

# ============ 可配置区（标的切换只改这里）============
PRIMARY_CODE = "H00922"            # 中证红利全收益指数(全收益=含分红)
PRIMARY_NAME = "中证红利全收益指数"
BENCH_CODE   = "H00985"            # 中证全指全收益(基准；口径须与 PRIMARY_CODE 同为全收益)
BENCH_NAME   = "中证全指全收益"
HIST_START   = "20000101"          # 尽量早取，保证 MA500 前置充足
INDICATOR_CODE = "000922"          # 股息率指标文件挂在价格指数上(成分与 H00922 一致)
INDICATOR_URL = ("https://oss-ch.csindex.com.cn/static/html/csindex/public/uploads"
                 f"/file/autofile/indicator/{INDICATOR_CODE}indicator.xls")
CONS_URL = ("https://oss-ch.csindex.com.cn/static/html/csindex/public/uploads"
            f"/file/autofile/closeweight/{INDICATOR_CODE}closeweight.xls")  # 成分股权重文件(每月末更新)
# ====================================================

MAS = [250, 350, 500]
RET_WINDOW = 40
HISTORY_FILE = os.path.join(HERE, "dividend_index_history.json")
REPORT_FILE  = os.path.join(HERE, "dividend_index_report.html")
SITE_DIR     = os.path.join(HERE, "site")   # 发布目录(复制为 index.html 供部署)
CACHE_DIR    = os.path.join(HERE, "cache")  # 全历史日线本地缓存(增量取数用)
CACHE_P      = os.path.join(CACHE_DIR, f"{PRIMARY_CODE.lower()}.json")
CACHE_B      = os.path.join(CACHE_DIR, f"{BENCH_CODE}.json")
CONS_CACHE   = os.path.join(CACHE_DIR, f"{INDICATOR_CODE}cons.json")  # 成分股缓存(取数失败降级用)
CSINDEX_PERF = "https://www.csindex.com.cn/csindex-home/perf/index-perf"
DATA_NOTE = ("数据来源：中证指数有限公司官网 csindex.com.cn（指数表现接口 + 每日更新的指数估值指标文件），免费、无需授权。"
             "主标的 H00922 与基准 H00985 均为全收益口径(含分红再投资)，两侧可比。"
             "40 日收益差在主标的与基准的交易日交集上按 40 个交易日滚动计算。"
             "PE 分位由官网 peg 字段在自身历史中计算；股息率取自官网指标文件(000922 价格指数口径)。")

# ============ 数据质量：交易日清洗 & 告警 ============
# 背景：中证官网曾返回一行 2026-08-29(周六) 的伪数据，收盘价与次日 8/31 完全相同。
# 该行写入缓存后，使按数组下标滚动的窗口整体错位 1 个交易日，污染此后所有 40 日收益差。
# 另发现 2018-06-18(端午节休市) 同样是幽灵行：000922/000985/H00922 三条指数在该日的
# 收盘价与 6/15 分毫不差，全收益口径 H00985 仅因股息累积微增 0.012%——真实交易日
# 不可能完全持平，据此可判定该日休市。
WEEKDAY_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
# 权威交易日历(首选)：分年从东方财富拉取并本地缓存，可同时剔除周末与法定节假日。
# 拉取失败(网络/限流)时自动降级为「仅周末过滤」并打印告警，绝不阻断主流程。
CALENDAR_FILE = os.path.join(CACHE_DIR, "trading_calendar.json")
CALENDAR_SYMBOL = "1.000001"      # 东财口径：上证指数，仅用于取 A 股交易日历
CALENDAR_SYMBOL_TX = "sh000001"   # 腾讯口径：上证指数(实测稳定，作为主源)
CALENDAR_START_YEAR = 2005
# 手工补充的节假日(工作日但休市)，按 "YYYY-MM-DD" 追加；交易日历可用时本表为冗余保险
KNOWN_HOLIDAYS = set()
# 相邻两日收盘价完全相同时是否告警(只告警、不删除，原因见 sanitize_bars 规则 C)
WARN_DUP_CLOSE = True
# 尾部校验：最新交易日距今超过该天数视为数据陈旧(覆盖春节/国庆等最长休市)
MAX_STALE_DAYS = 10
_TRADE_CAL = None                 # 惰性加载的交易日历缓存


def load_trading_calendar():
    """加载 A 股权威交易日历(日期字符串集合)。优先读本地缓存，缺失则分年从东财拉取。
    失败返回 None(调用方降级为仅周末过滤)，绝不因日历不可用而中断主流程。"""
    global _TRADE_CAL
    if _TRADE_CAL is not None:
        return _TRADE_CAL
    if os.path.exists(CALENDAR_FILE):
        try:
            with open(CALENDAR_FILE, encoding="utf-8") as f:
                _TRADE_CAL = set(json.load(f))
            return _TRADE_CAL
        except Exception as e:
            print(f"      [数据质量] 交易日历缓存读取失败({e})，将重新拉取")
    this_year = datetime.date.today().year
    # 先拉当年：网络故障/接口限流通常是全局性的，当年失败即快速降级，
    # 避免 20 余个年份逐个重试把日常流程拖到几分钟。
    days = _fetch_calendar_year(this_year)
    if days is None:
        print(f"      [数据质量·告警] 交易日历拉取失败(疑似网络或限流)，降级为仅剔除周末；"
              f"节假日幽灵行(如 2018-06-18)本轮无法识别")
        return None
    for y in range(CALENDAR_START_YEAR, this_year):
        chunk = _fetch_calendar_year(y)
        if chunk is None:
            print(f"      [数据质量·告警] 交易日历 {y} 年拉取失败，降级为仅剔除周末")
            return None
        days |= chunk
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        with open(CALENDAR_FILE, "w", encoding="utf-8") as f:
            json.dump(sorted(days), f)
        print(f"      [数据质量] 交易日历已缓存 {len(days)} 天 -> {CALENDAR_FILE}")
    except Exception as e:
        print(f"      [数据质量] 交易日历写缓存失败({e})，不影响本次运行")
    _TRADE_CAL = days
    return _TRADE_CAL


def _parse_calendar_tx(d):
    """解析腾讯日线：data.<code>.qfqday / day -> [日期, 开, 收, 高, 低, 量, ...]"""
    data = d.get("data") or {}
    node = data.get(CALENDAR_SYMBOL_TX) if isinstance(data, dict) else None
    if not node:
        return set()
    rows = node.get("qfqday") or node.get("day") or []
    return {str(r[0]) for r in rows if r and str(r[0])}


def _parse_calendar_em(d):
    """解析东财日线：data.klines -> 'YYYY-MM-DD,...'"""
    kl = (d.get("data") or {}).get("klines") or []
    return {str(x.split(",")[0]) for x in kl if x}


def _fetch_calendar_year(year, retry=2):
    """拉取某一年的 A 股交易日集合，失败返回 None。
    双源容灾：优先腾讯(实测稳定、不限流)，失败回退东方财富(易受限流)。"""
    sources = (
        ("腾讯", f"https://web.ifzq.gtimg.cn/appstock/app/fqkline/get"
                 f"?param={CALENDAR_SYMBOL_TX},{year}-01-01,{year}-12-31,320,qfq",
         _parse_calendar_tx),
        ("东财", f"https://push2his.eastmoney.com/api/qt/stock/kline/get"
                 f"?secid={CALENDAR_SYMBOL}&klt=101&fqt=0"
                 f"&beg={year}0101&end={year}1231&fields1=f1,f2,f3&fields2=f51",
         _parse_calendar_em),
    )
    for name, url, parser in sources:
        for k in range(retry):
            try:
                req = urllib.request.Request(url, headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                    "Referer": "https://quote.eastmoney.com/"})
                with urllib.request.urlopen(req, timeout=25) as r:
                    d = json.loads(r.read().decode("utf-8"))
                got = parser(d)
                if got:
                    return got
            except Exception:
                pass
            if k < retry - 1:
                time.sleep(2 * (k + 1))
    return None


def sanitize_bars(code, bars):
    """清洗非交易日与幽灵行，返回 (清洗后 bars, 问题列表)。
    规则 A  剔除周六/周日：指数不可能在非工作日产生收盘价。(硬删除，零误杀)
    规则 B  剔除不在权威交易日历中的日期(含法定节假日)。日历不可用则自动跳过本规则。
    规则 C  仅告警、绝不删除：相邻两日收盘价完全相同。
            原因——幽灵行可能落在真实日的任一侧(2026-08-29 在前、2018-06-18 在后)，
            仅凭收盘价无法判定该删哪一日；早期版本按「删前一日」处理已误删真实交易日
            2018-06-15，故本规则只提示人工核查，不改动数据。
    """
    issues = []
    if not bars:
        return bars, issues
    cal = load_trading_calendar()
    # 日历覆盖不到的早期日期(如指数基日 2004-12-31)不参与日历判定，避免误删基准点
    cal_min = min(cal) if cal else None
    kept = []
    for x in bars:
        d = x["date"]
        wd, key = d.weekday(), d.isoformat()
        if wd >= 5:
            issues.append(f"{code} {key} 是{WEEKDAY_CN[wd]}，非交易日，已剔除")
            continue
        if cal is not None and cal_min and key >= cal_min and key not in cal:
            issues.append(f"{code} {key} 不在 A 股交易日历中(疑似节假日幽灵行)，已剔除")
            continue
        if key in KNOWN_HOLIDAYS:
            issues.append(f"{code} {key} 是法定节假日，已剔除")
            continue
        kept.append(x)
    bars = kept

    if WARN_DUP_CLOSE:
        for i in range(1, len(bars)):
            if bars[i]["close"] == bars[i - 1]["close"]:
                issues.append(f"[告警·未删除] {code} {bars[i]['date']} 收盘价与前一交易日"
                              f"({bars[i - 1]['date']})完全相同({bars[i]['close']})，"
                              f"疑似休市幽灵行，但无法判定应删除哪一日，请人工核查")
    return bars, issues


def report_quality(code, issues):
    """打印数据质量问题；发现异常时集中提示，便于 CI 日志检索。"""
    for m in issues:
        print(f"      [数据质量] {m}")
    if issues:
        print(f"      [数据质量] {code} 共 {len(issues)} 处异常，已按规则处理")


def check_tail(bars, label):
    """尾部质量校验：最新交易日必须是工作日且不得陈旧；重复收盘价仅告警。异常即中断。"""
    if not bars:
        raise RuntimeError(f"{label} 序列为空")
    last = bars[-1]
    wd = last["date"].weekday()
    if wd >= 5:
        raise RuntimeError(f"{label} 最新交易日 {last['date']} 是{WEEKDAY_CN[wd]}，数据异常，已中断")
    stale = (datetime.date.today() - last["date"]).days
    if stale > MAX_STALE_DAYS:
        raise RuntimeError(f"{label} 最新交易日 {last['date']} 距今 {stale} 天"
                           f"(超过 {MAX_STALE_DAYS} 天)，数据可能陈旧，已中断")
    tail = bars[-5:]
    for i in range(1, len(tail)):
        if tail[i]["close"] == tail[i - 1]["close"]:
            print(f"      [数据质量·告警] {label} 最近连续两日({tail[i - 1]['date']} 与 {tail[i]['date']})"
                  f"收盘价完全相同({tail[i]['close']})，请人工确认是否为真实收平")

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
    out, issues = sanitize_bars(code, out)
    report_quality(code, issues)
    if not out:
        raise RuntimeError(f"中证官网 {code} 数据经交易日清洗后为空，疑似接口异常")
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
    """全历史缓存 + 增量拉取缺口(缓存次日~今天)，合并去重、清洗交易日/幽灵行后写回。
    返回 [{date: date对象, close, peg}] 升序。首跑无缓存时全量拉取。

    注意：清洗同时作用于「存量缓存」与「新拉数据」。历史缓存中若已混入幽灵交易日，
    也会在此被剔除并回写磁盘，无需手工清理缓存文件。
    """
    def _clean(rows):
        obj = [{"date": datetime.date.fromisoformat(r["date"]) if isinstance(r["date"], str) else r["date"],
                "close": r["close"], "peg": r.get("peg")} for r in rows]
        obj, issues = sanitize_bars(code, obj)
        report_quality(code, issues)
        if not obj:
            raise RuntimeError(f"{code} 数据经交易日清洗后为空，疑似接口异常")
        save_cache(path, [{"date": o["date"].isoformat(), "close": o["close"],
                           "peg": o.get("peg")} for o in obj])
        return obj

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
            return _clean(cached)
        start = (last_dt + datetime.timedelta(days=1)).strftime("%Y%m%d")
        mode = f"增量({cached[-1]['date']}次日→今天)"
    new_bars = fetch_index(code, start=start, allow_empty=True)
    if not new_bars:
        print(f"      {code} 官网在 {start} 后无新数据(当日未更新), 沿用缓存 bars={len(cached)}")
        return _clean(cached)
    merged = {c["date"]: c for c in cached}
    for nb in new_bars:
        key = nb["date"].isoformat()
        merged[key] = {"date": key, "close": nb["close"], "peg": nb["peg"]}
    bars = sorted(merged.values(), key=lambda x: x["date"])
    obj = _clean(bars)
    print(f"      {code} {mode} 拉 {len(new_bars)} 日 → 清洗后缓存 {len(obj)} 日 "
          f"({obj[0]['date']}~{obj[-1]['date']})")
    return obj

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

# ---------- 成分股模块（官网权重文件 + 近五年调整史 + 定期调整预告） ----------
ADJ_HISTORY = [
    # 近五年样本定期调整史(每年12月生效, 20进20出)。名单以中证指数官网当期公告为准;
    # 官网自2024年起不长期存档历史进出名单, 以下名单经官方公告/权威媒体转载互证整理。
    {"year": 2021, "ann": "2021-11-26", "eff": "2021-12-13", "full": True,
     "ins": ["天健集团(000090)", "江铃汽车(000550)", "冀中能源(000937)", "中南建设(000961)",
             "伟星股份(002003)", "森马服饰(002563)", "汇洁股份(002763)", "佳士科技(300193)",
             "华电国际(600027)", "阳光照明(600261)", "江山股份(600389)", "千金药业(600479)",
             "益佰制药(600594)", "重庆百货(600729)", "唐山港(601000)", "四方股份(601126)",
             "君正集团(601216)", "陕西煤业(601225)", "平煤股份(601666)", "元祖股份(603886)"],
     "outs": ["潍柴动力(000338)", "晨鸣纸业(000488)", "佛山照明(000541)", "苏宁环球(000718)",
              "鲁泰A(000726)", "承德露露(000848)", "华孚时尚(002042)", "九阳股份(002242)",
              "科华数据(002335)", "永兴材料(002756)", "阳谷华泰(300121)", "永利股份(300230)",
              "汉宇集团(300403)", "招商银行(600036)", "东睦股份(600114)", "香江控股(600162)",
              "迪马股份(600565)", "广汇物流(600603)", "福耀玻璃(600660)", "依顿电子(603328)"]},
    {"year": 2022, "ann": "2022-11-25", "eff": "2022-12-12", "full": False,
     "ins": ["鲁西化工", "新钢股份", "格力电器"],
     "outs": ["冀中能源", "宇通客车", "佳士科技", "重庆百货", "大东方", "建新股份"]},
    {"year": 2023, "ann": "2023-11-24", "eff": "2023-12-11", "full": False,
     "ins": ["兰花科创", "山西焦煤", "开滦股份", "山煤国际", "潞安环能", "中国石油",
             "四川路桥", "洪城环境", "山东出版", "宁波华翔", "贵阳银行"],
     "outs": ["东莞控股", "国投电力", "双汇发展", "养元饮品", "步长制药", "华电国际",
              "上海石化", "凌霄泵业", "阳光照明", "森马服饰", "深高速", "首开股份",
              "金地集团", "华联控股", "金融街"]},
    {"year": 2024, "ann": "2024-11-29", "eff": "2024-12-16", "full": True,
     "ins": ["粤高速A", "双汇发展", "冀中能源", "陕天然气", "森马服饰", "周大生",
             "广汇能源", "深高速", "大商股份", "重庆百货", "昊华能源", "西部矿业",
             "邮储银行", "沪农商行", "成都银行", "中远海控", "中国平安", "重庆银行",
             "东方环宇", "同力股份(北交所首个样本)"],
     "outs": ["万科A", "万年青", "伟星股份", "达安基因", "三钢闽光", "电投能源",
              "康力电梯", "金洲管道", "保利发展", "江山股份", "千金药业", "申能股份",
              "川投能源", "物产中大", "华新水泥", "马钢股份", "梅花生物", "长江电力",
              "中国太保", "元祖股份"]},
    {"year": 2025, "ann": "2025-11-28", "eff": "2025-12-15", "full": True,
     "ins": ["藏格矿业(000408)", "上峰水泥(000672)", "神火股份(000933)", "兔宝宝(002043)",
             "报喜鸟(002154)", "亚太科技(002540)", "索菲亚(002572)", "永兴材料(002756)",
             "军信股份(301109)", "招商银行(600036)", "云天化(600096)", "安徽建工(600502)",
             "中粮糖业(600737)", "中国海油(600938)", "晋控煤业(601001)", "厦门银行(601187)",
             "中国外运(601598)", "中创智领(601717)", "浙商银行(601916)", "中创物流(603967)"],
     "outs": ["威孚高科(000581)", "鲁西化工(000830)", "华菱钢铁(000932)", "冀中能源(000937)",
              "宁波华翔(002048)", "鲁阳节能(002088)", "富安娜(002327)", "明德生物(002932)",
              "浦发银行(600000)", "宝钢股份(600019)", "华发股份(600325)", "宁沪高速(600377)",
              "盘江股份(600395)", "深高速(600548)", "大商股份(600694)", "新钢股份(600782)",
              "旗滨集团(601636)", "武进不锈(603878)", "蓝天燃气(605368)", "奥泰生物(688606)"]},
]

def fetch_constituents():
    """官网成分股权重文件(.xls, 每月末更新)：{date:'YYYYMMDD', items:[{code,name,weight}]}。
    失败时读取上次缓存；缓存也没有则返回空(成分股为增强模块, 不中断主流程)。"""
    try:
        import xlrd
    except ImportError:
        print("      [提示] 未安装 xlrd，成分股模块降级")
        xlrd = None
    items, d = [], None
    if xlrd is not None:
        tmp = os.path.join(HERE, "cons_tmp.xls")
        for attempt in range(3):
            try:
                req = urllib.request.Request(CONS_URL, headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
                with urllib.request.urlopen(req, timeout=40) as r:
                    raw = r.read()
                if len(raw) < 4096:
                    raise RuntimeError(f"权重文件过小({len(raw)}B)")
                open(tmp, "wb").write(raw)
                sh = xlrd.open_workbook(tmp).sheet_by_index(0)
                for r_ in range(1, sh.nrows):
                    v = [sh.cell_value(r_, c) for c in range(sh.ncols)]
                    dd = str(v[0]).split(".")[0]
                    if len(dd) == 8 and dd.isdigit():
                        d = dd
                    code = str(v[4]).split(".")[0].strip()
                    if len(code) != 6 or not code.isdigit():
                        continue
                    try: w = float(v[9])
                    except (TypeError, ValueError): w = None
                    items.append({"code": code, "name": str(v[5]).strip(), "weight": w})
                if len(items) < 90:
                    raise RuntimeError(f"成分股数量异常({len(items)})")
                obj = {"date": d, "items": items}
                try:
                    json.dump(obj, open(CONS_CACHE, "w", encoding="utf-8"), ensure_ascii=False)
                except Exception:
                    pass
                print(f"      成分股: 官网权重文件({d}) 取得 {len(items)} 只")
                return obj
            except Exception as e:
                if attempt == 2:
                    print(f"      [提示] 成分股权重文件获取失败: {e}")
                else:
                    time.sleep(1.5)
    if os.path.exists(CONS_CACHE):
        try:
            obj = json.load(open(CONS_CACHE, encoding="utf-8"))
            print(f"      成分股: 使用缓存({obj.get('date')}) 共 {len(obj.get('items') or [])} 只")
            return obj
        except Exception:
            pass
    return {"date": None, "items": []}

def fetch_div_yields(codes):
    """自算个股价息率 = 最近一个完整会计年度每股税前分红合计 ÷ 最新收盘价。
    分红明细来自东财数据中心(RPT_SHAREBONUS_DET, PRETAX_BONUS_RMB 为每10股税前红利，
    REPORT_DATE 归属会计年度)；收盘价来自东财批量行情(push2delay)。返回 {code: 股息率%}。
    口径：按"最近一个有分红记录的会计年度(中报+年报合计)"归集，与官方缓冲区条款
    "过去一年现金股息率"对齐；比滚动365天窗口更稳(后者会因除息日跨年漂移而错误归零)。
    早期版本直接取行情接口股息率字段(f115)，实测对小盘/特殊分红个股失真(如-43%、20%)，故弃用。
    注意：仍非中证官方选样口径(官方选样按"过去三年平均现金股息率")，仅作剔除候选参照。"""
    out = {}
    if not codes:
        return out
    # 1) 批量最新价
    px = {}
    for i in range(0, len(codes), 50):
        chunk = codes[i:i + 50]
        secids = ",".join(("1." if c.startswith("6") else "0.") + c for c in chunk)
        for host in ("https://push2delay.eastmoney.com", "https://push2.eastmoney.com"):
            try:
                url = (f"{host}/api/qt/ulist.np/get?secids={secids}"
                       "&fields=f12,f2&fltt=2&invt=2")
                req = urllib.request.Request(url, headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                    "Referer": "https://quote.eastmoney.com/"})
                with urllib.request.urlopen(req, timeout=25) as r:
                    j = json.loads(r.read().decode("utf-8"))
                for d_ in (j.get("data") or {}).get("diff") or []:
                    v = d_.get("f2")
                    if isinstance(v, (int, float)):
                        px[str(d_.get("f12"))] = float(v)
                break
            except Exception:
                continue
    # 2) 逐股分红明细 → 最近一个会计年度每股税前分红合计
    ok = 0
    for c in codes:
        price = px.get(c)
        if not price:
            continue
        try:
            url = (f"https://datacenter-web.eastmoney.com/api/data/v1/get"
                   f"?reportName=RPT_SHAREBONUS_DET&columns=SECURITY_CODE,REPORT_DATE,EX_DIVIDEND_DATE,PRETAX_BONUS_RMB"
                   f"&filter=(SECURITY_CODE%3D%22{c}%22)&pageSize=30&pageNumber=1&sortColumns=REPORT_DATE&sortTypes=-1")
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0", "Referer": "https://data.eastmoney.com/"})
            with urllib.request.urlopen(req, timeout=20) as r:
                j = json.loads(r.read().decode("utf-8"))
            rows = (j.get("result") or {}).get("data") or []
            per_year = {}     # 报告期年份 → 每股税前分红合计(含未除息记录)
            ex_years = set()  # 已有分红实施(已除息)的报告期年份
            for row in rows:
                bonus = row.get("PRETAX_BONUS_RMB")
                y = (row.get("REPORT_DATE") or "")[:4]
                if not y.isdigit() or bonus is None:
                    continue
                per_year[y] = per_year.get(y, 0.0) + float(bonus) / 10.0
                if (row.get("EX_DIVIDEND_DATE") or ""):
                    ex_years.add(y)
            # 只从"已实施分红"的年份中取最近一年——避免选中仅有未除息中期分红的当年，
            # 否则"年报分红已除息+下一中期未除息"的个股会被错误压缩到近乎为零(实测陕西煤业0.22%)。
            if per_year and ex_years:
                out[c] = per_year[max(ex_years)] / price * 100.0
                ok += 1
        except Exception:
            continue
    print(f"      股息率(最近会计年度): 自算完成 {ok}/{len(codes)} 只(东财分红明细/现价)")
    return out

def next_adjustment(today=None):
    """定期调整预告：生效日 = 当年12月第二个星期五的下一交易日(跨周末顺延)。
    已过生效日则推算下一年。返回 (年份, 生效日date, 公告惯例说明)。"""
    today = today or datetime.date.today()
    def sec_friday(y):
        d = datetime.date(y, 12, 1)
        d += datetime.timedelta(days=(4 - d.weekday()) % 7)  # 12月第一个周五
        return d + datetime.timedelta(days=7)                # 第二个周五
    def eff_day(y):
        e = sec_friday(y) + datetime.timedelta(days=1)
        while e.weekday() >= 5:
            e += datetime.timedelta(days=1)
        return e
    y = today.year
    eff = eff_day(y)
    if today >= eff:
        y += 1
        eff = eff_day(y)
    return y, eff, f"公告惯例于生效前约两周({y}年11月下旬)发布，以中证指数官网公告为准"

ADD_CACHE = os.path.join(CACHE_DIR, "add_candidates.json")

def fetch_industry_map(codes):
    """东财行业分类(f100字段)：{code: 行业名}。批量50只/次仅2个请求，失败返回 {}(行业分布模块降级)。
    注：编制方案无行业权重条款，行业分布为股息率选样的自然结果，此处仅作展示。"""
    out = {}
    if not codes:
        return out
    for i in range(0, len(codes), 50):
        chunk = codes[i:i + 50]
        secids = ",".join(("1." if c.startswith("6") else "0.") + c for c in chunk)
        for host in ("https://push2delay.eastmoney.com", "https://push2.eastmoney.com"):
            try:
                url = (f"{host}/api/qt/ulist.np/get?secids={secids}"
                       "&fields=f12,f100&fltt=2&invt=2")
                req = urllib.request.Request(url, headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                    "Referer": "https://quote.eastmoney.com/"})
                with urllib.request.urlopen(req, timeout=25) as r:
                    j = json.loads(r.read().decode("utf-8"))
                for d_ in (j.get("data") or {}).get("diff") or []:
                    if d_.get("f100"):
                        out[str(d_.get("f12"))] = str(d_["f100"])
                break
            except Exception:
                continue
    if codes:
        print(f"      行业分类: 取得 {len(out)}/{len(codes)} 只(东财口径)")
    return out

def fetch_add_candidates(max_age_days=7):
    """纳入候选参照：按官方选样规则(近似口径)对全市场做规则化筛选，与剔除候选对称。
    口径：连续3个会计年度现金分红 + 各年股利支付率(0,1)(每股分红/EPS) + 总市值前80%近似样本空间，
    按 三年平均每股分红/现价 排名；非成分股进入Top100者，按历年20进20出惯例取前若干只为候选。
    与官方差异：官方用历年年末市值算股息率、流动性口径为成交额，此处以现价/总市值近似。
    全市场扫描较重(分红明细约35页+行情快照约56页)，结果缓存 max_age_days 天；
    扫描失败自动降级读旧缓存(不限龄)；完全无数据返回 None(不中断主流程)。"""
    def _load_cache():
        try:
            return json.load(open(ADD_CACHE, encoding="utf-8"))
        except Exception:
            return None
    cached = _load_cache()
    if cached and cached.get("generated"):
        try:
            age = (datetime.date.today() - datetime.date.fromisoformat(cached["generated"])).days
            if age <= max_age_days:
                print(f"      纳入候选: 使用{age}天前缓存({cached['generated']})")
                return cached
        except ValueError:
            pass
    today = datetime.date.today()
    years = [str(today.year - 3), str(today.year - 2), str(today.year - 1)]
    y_last = years[-1]
    def get_json(url, timeout=25, retries=3):
        for a in range(retries):
            try:
                req = urllib.request.Request(url, headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                    "Referer": "https://data.eastmoney.com/"})
                with urllib.request.urlopen(req, timeout=timeout) as r:
                    return json.loads(r.read().decode("utf-8"))
            except Exception:
                if a == retries - 1:
                    raise
                time.sleep(1.5)
    try:
        # 1) 全市场分红明细(REPORT_DATE 归年)
        divs = {}
        page = 1
        while True:
            j = get_json("https://datacenter-web.eastmoney.com/api/data/v1/get"
                         "?reportName=RPT_SHAREBONUS_DET"
                         "&columns=SECURITY_CODE,SECURITY_NAME_ABBR,REPORT_DATE,EX_DIVIDEND_DATE,PRETAX_BONUS_RMB,BASIC_EPS"
                         f"&filter=(EX_DIVIDEND_DATE%3E%3D'{years[0]}-01-01')"
                         f"&pageNumber={page}&pageSize=500&sortColumns=EX_DIVIDEND_DATE,SECURITY_CODE&sortTypes=-1,-1")
            res = j.get("result") or {}
            for row in res.get("data") or []:
                code = row.get("SECURITY_CODE")
                bonus = row.get("PRETAX_BONUS_RMB")
                rd = (row.get("REPORT_DATE") or "")[:10]
                if not code or bonus is None or not rd:
                    continue
                yy = rd[:4]
                if yy not in years:
                    continue
                d = divs.setdefault(code, {"name": row.get("SECURITY_NAME_ABBR") or "", "years": {}, "eps": {}})
                d["years"][yy] = d["years"].get(yy, 0.0) + float(bonus) / 10.0
                if rd.endswith("12-31") and row.get("BASIC_EPS"):
                    try:
                        d["eps"][yy] = float(row["BASIC_EPS"])
                    except (TypeError, ValueError):
                        pass
            if page >= (res.get("pages") or 1):
                break
            page += 1
        # 2) 全A行情快照(现价/总市值)
        quotes = {}
        pn = 1
        while True:
            j = get_json("https://push2delay.eastmoney.com/api/qt/clist/get"
                         f"?pn={pn}&pz=100&po=1&np=1&fltt=2&invt=2&fid=f20"
                         "&fs=m:0+t:6,m:0+t:80,m:1+t:2,m:1+t:23"
                         "&fields=f12,f14,f2,f6,f20", retries=4)
            data = j.get("data") or {}
            diff = data.get("diff") or []
            if not diff:
                break
            for d_ in diff:
                c = d_.get("f12")
                if not c:
                    continue
                def fnum(x):
                    return float(x) if isinstance(x, (int, float)) else None
                quotes[c] = {"name": d_.get("f14") or "", "price": fnum(d_.get("f2")), "mcap": fnum(d_.get("f20"))}
            if pn * 100 >= (data.get("total") or 0):
                break
            pn += 1
        # 3) 资格筛选与排名
        members = set()
        try:
            members = {c["code"] for c in json.load(open(CONS_CACHE, encoding="utf-8")).get("items") or []}
        except Exception:
            pass
        mcaps = sorted(q["mcap"] for q in quotes.values() if q["mcap"])
        mcap_cut = mcaps[int(len(mcaps) * 0.8)] if mcaps else None
        cands = []
        for code, d in divs.items():
            ys = d["years"]
            if not all(ys.get(yy, 0) > 0 for yy in years):
                continue
            q = quotes.get(code)
            if not q or not q["price"]:
                continue
            pays = {}
            ok_payout = True
            for yy in years:
                eps = d["eps"].get(yy)
                if not eps or eps <= 0:
                    ok_payout = False
                    break
                p = ys[yy] / eps
                if not (0 < p < 1):
                    ok_payout = False
                    break
                pays[yy] = p
            if not ok_payout or not (0 < sum(pays.values()) / 3 < 1):
                continue
            name = q["name"] or d["name"]
            if "ST" in name.upper() or "退" in name:
                continue
            if code.startswith(("4", "8", "92")):   # 北交所不在中证全指样本空间
                continue
            if mcap_cut and (q["mcap"] or 0) < mcap_cut:
                continue
            y3 = sum(ys[yy] for yy in years) / 3 / q["price"] * 100
            cands.append({"code": code, "name": name, "y3": y3,
                          "y_last": ys[y_last] / q["price"] * 100,
                          "mcap": q["mcap"], "member": code in members})
        cands.sort(key=lambda x: -x["y3"])
        newcomers = [c for c in cands[:100] if not c["member"]]
        out = {"generated": today.isoformat(), "years": years,
               "n_elig": len(cands), "n_member_top100": sum(1 for c in cands[:100] if c["member"]),
               "n_newcomers": len(newcomers), "newcomers": newcomers[:25]}
        try:
            json.dump(out, open(ADD_CACHE, "w", encoding="utf-8"), ensure_ascii=False)
        except Exception:
            pass
        print(f"      纳入候选: 全市场扫描完成(合格池{len(cands)}, 非成分Top100 {len(newcomers)})")
        return out
    except Exception as e:
        print(f"      [提示] 纳入候选扫描失败({e})，降级用旧缓存" if cached else f"      [提示] 纳入候选扫描失败({e})，无缓存可用")
        return cached

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

def svg_loglog_trend(dates, closes, i0, i1,
                     title="中证红利全收益指数 长期趋势（对数-对数 · 二阶拟合）",
                     width=880, height=380):
    """对数-对数趋势：ln(P)=c0+c1·ln(年)+c2·ln²(年)，价格取对数轴，叠加 ±1.5σ 置信带。
    σ 用稳健估计(MAD×1.4826)而非经典标准差——A股崩盘日的极端残差会把经典σ撑大(实测约+16%)，
    导致平时的带偏宽；MAD 只反映常态散布，包裹区间更贴近大多数交易日。
    置信带 = 拟合曲线 ± 1.5×σ，正态假设下约覆盖 86.6% 的单日观测，视觉上包住绝大部分价格走势。
    另标注编制方案修订日竖线(2022-12-12 / 2025-10-13)作结构性断点参考。"""
    idx = list(range(i0, i1 + 1))
    if len(idx) < 4:
        return ""
    d0 = dates[i0]
    # 自变量：自 i0 起算的年数(+1，避免 ln(0))
    t = [(dates[i] - d0).days / 365.25 + 1.0 for i in idx]
    lt = [math.log(v) for v in t]
    lnP = [math.log(closes[i]) for i in idx]
    n = len(lt)
    # 最小二乘二次拟合 lnP = c0 + c1*lt + c2*lt^2 (正规方程 3x3)
    s = [sum(lt), sum(x * x for x in lt), sum(x ** 3 for x in lt), sum(x ** 4 for x in lt)]
    r_ = [sum(lnP), sum(lt[i] * lnP[i] for i in range(n)), sum(lt[i] ** 2 * lnP[i] for i in range(n))]
    m = [[n, s[0], s[1]], [s[0], s[1], s[2]], [s[1], s[2], s[3]]]
    # 高斯消元
    A = [row[:] + [r_[k]] for k, row in enumerate(m)]
    for col in range(3):
        piv = max(range(col, 3), key=lambda rr: abs(A[rr][col]))
        A[col], A[piv] = A[piv], A[col]
        for rr in range(col + 1, 3):
            f = A[rr][col] / A[col][col]
            for cc in range(col, 4):
                A[rr][cc] -= f * A[col][cc]
    c = [0.0, 0.0, 0.0]
    for rr in (2, 1, 0):
        c[rr] = (A[rr][3] - sum(A[rr][cc] * c[cc] for cc in range(rr + 1, 3))) / A[rr][rr]
    c0, c1, c2 = c
    fit_ln = [c0 + c1 * lt[i] + c2 * lt[i] ** 2 for i in range(n)]
    resid = [lnP[i] - fit_ln[i] for i in range(n)]
    # 稳健σ(MAD法)：中位数绝对偏差×1.4826。经典σ会被崩盘日极端残差撑大(实测约+16%)，
    # MAD 只反映常态散布，是本场景更合适的尺度估计。
    srt = sorted(resid)
    med = srt[n // 2] if n % 2 else (srt[n // 2 - 1] + srt[n // 2]) / 2
    sigma = 1.4826 * sorted(abs(r - med) for r in resid)[n // 2]
    mean_y = sum(lnP) / n
    ss_tot = sum((v - mean_y) ** 2 for v in lnP)
    R2 = 1 - sum(x * x for x in resid) / ss_tot if ss_tot else 0.0

    # 置信带 = 拟合曲线 ± 1.5×σ（残差标准差），包裹价格曲线：
    # 正态假设下约覆盖 86.6% 的单日观测，带宽恒定(对数空间)，包裹绝大部分走势。
    fit = [math.exp(v) for v in fit_ln]
    up = [math.exp(v + 1.5 * sigma) for v in fit_ln]
    lo = [math.exp(v - 1.5 * sigma) for v in fit_ln]
    cur = closes[idx[-1]]
    dev = (cur / fit[-1] - 1) * 100
    in_band = lo[-1] <= cur <= up[-1]
    band_half = (math.exp(1.5 * sigma) - 1) * 100

    L, R, T, B = 60, 18, 40, 34
    pw, ph = width - L - R, height - T - B
    ys_all = [closes[i] for i in idx] + up + lo
    lmin = math.log10(min(ys_all) * 0.97); lmax = math.log10(max(ys_all) * 1.03)

    def mx(i): return L + (i - i0) / (i1 - i0) * pw if i1 > i0 else L
    def my(v): return T + (1 - (math.log10(v) - lmin) / (lmax - lmin)) * ph

    yt = []
    for e_ in range(int(math.floor(lmin)), int(math.ceil(lmax)) + 1):
        for m_ in (1, 2, 5):
            val = m_ * 10 ** e_
            if lmin <= math.log10(val) <= lmax:
                yt.append(val)

    sign = "+" if c2 >= 0 else "−"
    sub = (f"ln(P)={c0:.3f}{c1:+.3f}·ln(年){sign}{abs(c2):.3f}·ln²(年) ｜ R²={R2:.3f} ｜ 稳健σ(MAD)={sigma:.4f} ｜ "
           f"当前 {cur:,.0f} 偏离拟合 {dev:+.1f}%（{'带内' if in_band else '带外'}）｜ ±1.5σ 置信带半宽≈{band_half:.1f}%")
    svg = [f'<svg viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" font-family="system-ui,Segoe UI,Arial,sans-serif">']
    svg.append(f'<rect x="0" y="0" width="{width}" height="{height}" fill="#ffffff"/>')
    svg.append(f'<text x="{L}" y="18" font-size="14" font-weight="600" fill="#1f2937">{title}</text>')
    svg.append(f'<text x="{L}" y="33" font-size="11.5" fill="#6b7280">{sub}</text>')
    for v in yt:
        yy = my(v)
        svg.append(f'<line x1="{L}" y1="{yy:.1f}" x2="{width-R}" y2="{yy:.1f}" stroke="#eef0f3" stroke-width="1"/>')
        svg.append(f'<text x="{L-6}" y="{yy+4:.1f}" font-size="11" fill="#9ca3af" text-anchor="end">{v:,.0f}</text>')
    for tx, tlab in year_ticks(dates, i0, i1, max_ticks=11):
        xx = mx(tx)
        svg.append(f'<line x1="{xx:.1f}" y1="{T}" x2="{xx:.1f}" y2="{T+ph:.1f}" stroke="#eef0f3" stroke-width="1"/>')
        anchor = "middle"
        if xx < L + 16: anchor = "start"
        elif xx > width - R - 16: anchor = "end"
        svg.append(f'<text x="{xx:.1f}" y="{height-12}" font-size="11" fill="#9ca3af" text-anchor="{anchor}">{tlab}</text>')
    # 编制方案修订日竖线（结构性断点参考）：2013-07-02 改股息率加权(2016起点下在样本外)、
    # 2022-12-12 提高分红连续性要求+权重上限、2025-10-13 延续2022框架的进一步修订。
    for dstr, lab in (("2022-12-12", "2022修订"), ("2025-10-13", "2025修订")):
        dd = datetime.date.fromisoformat(dstr)
        if dates[i0] <= dd <= dates[i1]:
            j = min(idx, key=lambda i_: abs((dates[i_] - dd).days))
            xv = mx(j)
            svg.append(f'<line x1="{xv:.1f}" y1="{T}" x2="{xv:.1f}" y2="{T+ph:.1f}" stroke="#94a3b8" stroke-width="1" stroke-dasharray="3 3"/>')
            svg.append(f'<text x="{xv:.1f}" y="{T+27}" font-size="10" fill="#94a3b8" text-anchor="middle">{lab}</text>')
    # 下沿链必须 x/y 同序反向：k 与 idx[k] 一一对应(此前用 enumerate(reversed(idx)) 导致
    # x 反向、y 正向，下沿被画成对角线，多边形扭曲成喇叭口)。
    band_pts = " ".join(f"{mx(i):.1f},{my(up[k]):.1f}" for k, i in enumerate(idx)) + " " + \
               " ".join(f"{mx(idx[k]):.1f},{my(lo[k]):.1f}" for k in range(n - 1, -1, -1))
    svg.append(f'<polygon points="{band_pts}" fill="#d97706" fill-opacity="0.10"/>')
    fit_d = " ".join(f"{'M' if k == 0 else 'L'}{mx(i):.1f},{my(fit[k]):.1f}" for k, i in enumerate(idx))
    svg.append(f'<path d="{fit_d}" fill="none" stroke="#d97706" stroke-width="1.6" stroke-dasharray="5 3"/>')
    act_d = " ".join(f"{'M' if k == 0 else 'L'}{mx(i):.1f},{my(closes[i]):.1f}" for k, i in enumerate(idx))
    svg.append(f'<path d="{act_d}" fill="none" stroke="#2563eb" stroke-width="1.8"/>')
    svg.append(f'<rect x="{L+8}" y="{T+2}" width="12" height="12" rx="2" fill="#2563eb"/><text x="{L+26}" y="{T+12}" font-size="11.5" fill="#374151">实际收盘</text>')
    svg.append(f'<rect x="{L+110}" y="{T+2}" width="12" height="12" rx="2" fill="#d97706"/><text x="{L+128}" y="{T+12}" font-size="11.5" fill="#374151">对数-对数二阶拟合</text>')
    svg.append(f'<rect x="{L+238}" y="{T+2}" width="12" height="12" rx="2" fill="#d97706" fill-opacity="0.25"/><text x="{L+256}" y="{T+12}" font-size="11.5" fill="#374151">±1.5σ 置信带（包裹曲线）</text>')
    svg.append('</svg>')
    return "\n".join(svg)



# ================= 主流程 =================
print(f"[1/4] 取数 {PRIMARY_CODE} {PRIMARY_NAME} (缓存增量) ...")
p = get_index_cached(PRIMARY_CODE, CACHE_P)
check_tail(p, f"{PRIMARY_CODE} {PRIMARY_NAME}")
print(f"[2/4] 取数 {BENCH_CODE} {BENCH_NAME} (缓存增量) ...")
b = get_index_cached(BENCH_CODE, CACHE_B)
check_tail(b, f"{BENCH_CODE} {BENCH_NAME}")
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

# 40日收益差值（在主标的与基准的「交易日交集」上按 40 个交易日滚动）
# 关键点：先取交集、再按下标滚动。若沿用旧的 p_close[i - RET_WINDOW]，
# 任一源缺失或多余一个交易日(例如缓存混入幽灵行 2026-08-29)都会让窗口整体错位，
# 且不抛错、不告警，属静默污染——本次故障即由此产生。
common_dates = sorted(set(p_dates) & set(b_close.keys()))
_bad = sorted(set(p_dates) ^ set(b_close.keys()))
if _bad:
    print(f"      [数据质量] 主标的与基准交易日不一致共 {len(_bad)} 天，"
          f"已按交集({len(common_dates)} 天)计算，不影响窗口对齐。"
          f"不一致日期：{_bad[0]} ~ {_bad[-1]}")
pos_in_p = {d: i for i, d in enumerate(p_dates)}
p_close_by_date = {d: p_close[i] for i, d in enumerate(p_dates)}
diff_series = []; ret40_p = []; ret40_b = []
for k in range(RET_WINDOW, len(common_dates)):
    d0, d1 = common_dates[k - RET_WINDOW], common_dates[k]
    rp = p_close_by_date[d1] / p_close_by_date[d0] - 1
    rb = b_close[d1] / b_close[d0] - 1
    diff = rp - rb
    i = pos_in_p[d1]
    diff_series.append((i, diff)); ret40_p.append((i, rp)); ret40_b.append((i, rb))
if not diff_series:
    raise RuntimeError("40日收益差序列为空：主标的与基准的交易日交集不足 "
                       f"{RET_WINDOW + 1} 天，请检查数据源")
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
# 破位段：间隔 ≤ GAP_MAX 个非破位交易日的破位视为同一波行情, 合并为一段
# (避免相邻的二次探底被切成独立段, 导致展示的"最深偏离"远低于真实行情极值)
GAP_MAX = 5  # ≈ 1 周
episodes = []
if below:
    run = [below[0]]
    for it in below[1:]:
        gap = it[0] - run[-1][0] - 1  # 间隔的非破位交易日数
        if gap <= GAP_MAX:
            run.append(it)
        else:
            episodes.append(run); run = [it]
    episodes.append(run)
def deep_info(run):
    """返回 (最深偏离%, 最深日, 最深日收盘, 最深日 MA500) ——
    所有列对齐到最深日, 消除"最深偏离"与"末收盘/末MA500"不同日的口径矛盾。"""
    items = [(cl / ma[500][i] - 1, dt, cl, ma[500][i]) for (i, dt, cl) in run]
    dev, dt, cl, mav = min(items, key=lambda x: x[0])
    return dev * 100, dt, cl, mav
ep_stats = []
for ep in episodes:
    dev, dt, cl, mav = deep_info(ep)
    ep_stats.append({
        "start": ep[0][1], "end": ep[-1][1], "n": len(ep),
        "max_dev": dev,
        "deep_date": dt, "deep_close": cl, "deep_ma": mav,
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
# 图1 用近10年做横轴(覆盖更长周期, 看清长期均线趋势)
cut10 = last_date - datetime.timedelta(days=int(365.25*10))
x0 = next((i for i, d in enumerate(p_dates) if d >= cut10), 0)
c1 = [{"name": f"{PRIMARY_NAME}收盘", "color": "#2563eb",
       "points": [(i, p_close[i]) for i in range(x0, len(p))], "width": 1.8}]
for w in MAS:
    c1.append({"name": f"MA{w}", "color": ["#f59e0b","#10b981","#ef4444"][MAS.index(w)],
               "points": [(i, ma[w][i]) for i in range(x0, len(p)) if ma[w][i] is not None], "width": 1.5})
svg1 = svg_line_chart(f"{PRIMARY_NAME} 收盘与长期均线 (近10年, MA250/350/500)", c1, y_precision=0,
                      x_ticks=year_ticks(p_dates, x0, len(p) - 1))

cut16 = datetime.date(2016, 1, 1)
diff_16 = [(i, v) for i, v in diff_series if p_dates[i] >= cut16]
c2 = [{"name": "40日收益差值(%)", "color": "#7c3aed",
       "points": [(i, v*100) for i, v in diff_16], "width": 1.6}]
svg2 = svg_line_chart(f"{PRIMARY_NAME} − {BENCH_NAME} 40日收益差值(%) (2016年起)", c2, zero_line=True, y_precision=1,
                      x_ticks=year_ticks(p_dates, diff_16[0][0], diff_16[-1][0], max_ticks=9))

# 图3 长期趋势（对数-对数二阶拟合，自2016-01起，价格对数轴 + ±1.5σ(MAD) 置信带）
# 起点依据：2016-01-01 起（与收益差图表周期对齐）；2013-07-02 加权方式已改为股息率加权，
# 故 2016 起的样本同处股息率加权框架内。σ 用 MAD 稳健估计，修订日竖线作断点参考。
i0_trend = next((i for i, d in enumerate(p_dates) if d >= datetime.date(2016, 1, 1)), 0)
svg3 = svg_loglog_trend(p_dates, p_close, i0_trend, len(p) - 1)

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
                f"<td>{e['deep_date']}</td>"
                f"<td>{e['deep_close']:.2f}</td><td>{e['deep_ma']:.2f}</td></tr>")
if not ep_rows:
    ep_rows = "<tr><td colspan='7' style='text-align:center;color:#16a34a'>近五年无破位</td></tr>"

# ---------- 成分股模块：当前成分(官网) + 剔除候选(规则参照) + 调整预告 + 近五年调整史 ----------
cons = fetch_constituents()
cons_items = cons.get("items") or []
cons_dy = fetch_div_yields([c["code"] for c in cons_items]) if cons_items else {}
for c in cons_items:
    c["dy"] = cons_dy.get(c["code"])
cons_date = cons.get("date")
cons_date_cn = (f"{cons_date[:4]}-{cons_date[4:6]}-{cons_date[6:]}" if cons_date else "—")
top10 = sorted([c for c in cons_items if c["weight"] is not None],
               key=lambda x: -x["weight"])[:10]
top10_sum = sum(c["weight"] for c in top10) if top10 else None
# 行业权重分布(东财口径)：编制方案无行业权重条款, 行业集中是股息率选样的自然涌现结果
ind_map = fetch_industry_map([c["code"] for c in cons_items]) if cons_items else {}
ind_agg, ind_cnt = {}, {}
for c in cons_items:
    nm = ind_map.get(c["code"])
    if nm and c["weight"] is not None:
        ind_agg[nm] = ind_agg.get(nm, 0.0) + c["weight"]
        ind_cnt[nm] = ind_cnt.get(nm, 0) + 1
ind_sorted = sorted(ind_agg.items(), key=lambda x: -x[1])
_dy_have = [c for c in cons_items if c["dy"] is not None and c["weight"]]
wavg_dy = None
if _dy_have:
    _wsum = sum(c["weight"] for c in _dy_have)
    wavg_dy = sum(c["weight"] * c["dy"] for c in _dy_have) / _wsum if _wsum else None
# 剔除候选：按股息率(最近会计年度)升序取最低20只(公告惯例20进20出；缓冲区硬条件为"过去一年现金股息率>0.5%"，此处为参照口径)
cand = sorted(_dy_have, key=lambda x: x["dy"])[:20]
# 纳入候选：全市场按官方规则近似筛选(缓存7天,失败降级旧缓存),与剔除候选对称的参照
addc = fetch_add_candidates()
adj_y, adj_eff, adj_note = next_adjustment()
adj_eff_cn = f"{adj_y}年12月第二个星期五的下一交易日 {adj_eff.isoformat()}（{WEEKDAY_CN[adj_eff.weekday()]}）"

top10_rows = ""
for i, c in enumerate(top10):
    dy_s = "—" if c["dy"] is None else f"{c['dy']:.2f}%"
    top10_rows += (f"<tr><td>{i+1}</td><td>{c['code']}</td><td>{c['name']}</td>"
                   f"<td>{c['weight']:.3f}%</td><td>{dy_s}</td></tr>")
cand_rows = "".join(
    f"<tr><td>{c['code']}</td><td>{c['name']}</td>"
    f"<td class='{'neg' if c['dy'] <= 0.5 else ''}'>{c['dy']:.2f}%</td></tr>"
    for c in cand)
# 行业权重分布折叠表(默认收起,点击展开)
if ind_sorted:
    ind_rows = "".join(
        f"<tr><td>{i+1}</td><td>{nm}</td><td>{w:.2f}%</td><td>{ind_cnt[nm]}</td></tr>"
        for i, (nm, w) in enumerate(ind_sorted))
    _top2 = " / ".join(f"{nm} {w:.2f}%" for nm, w in ind_sorted[:2])
    ind_block = f"""
  <details><summary style="cursor:pointer;margin-top:8px"><b>行业权重分布（东财行业口径 · {len(ind_sorted)} 个行业，点击展开/收起）</b></summary>
  <p style="margin:6px 0 2px">编制方案无行业权重上限条款(仅个股上限:单一样本≤10%、总市值&lt;100亿样本≤0.5%)，
  行业集中为股息率选样的自然结果；当前前两大行业 <b>{_top2}</b>，合计 {sum(w for _, w in ind_sorted[:2]):.1f}%。</p>
  <table><thead><tr><th>#</th><th>行业</th><th>权重合计</th><th>只数</th></tr></thead>
  <tbody>{ind_rows}</tbody></table>
  </details>"""
else:
    ind_block = ""
# 纳入候选折叠表(默认收起,点击展开)
if addc and addc.get("newcomers"):
    _nc = addc["newcomers"][:20]
    _ylab = (addc.get("years") or ["", "", ""])[-1]
    add_rows = "".join(
        f"<tr><td>{i+1}</td><td>{c['code']}</td><td>{c['name']}</td>"
        f"<td>{c['y3']:.2f}%</td><td>{c['y_last']:.2f}%</td><td>{c['mcap']/1e8:.0f}</td></tr>"
        for i, c in enumerate(_nc))
    add_head = (f"（全市场按官方选样规则近似筛选：连续{len(addc.get('years') or [0,0,0])}个会计年度分红+支付率0~1+"
                f"总市值前80%，按三年平均股息率排名，非成分股Top100中取前20；数据生成于 {addc.get('generated')}，每7天刷新）")
    add_block = f"""
  <p style="margin-top:10px"><b>纳入候选参照</b>{add_head}：当前头名 <b>{_nc[0]['name']} {_nc[0]['y3']:.2f}%</b>，
  共 {addc.get('n_elig', '—')} 只过资格线、非成分股进入Top100 {addc.get('n_newcomers', len(addc.get('newcomers', [])))} 只。</p>
  <details><summary style="cursor:pointer"><b>纳入候选 Top20（点击展开/收起）</b></summary>
  <table><thead><tr><th>#</th><th>代码</th><th>名称</th><th>三年平均股息率</th><th>{_ylab}年股息率</th><th>总市值(亿)</th></tr></thead>
  <tbody>{add_rows}</tbody></table>
  </details>
  <p class="refnote">与剔除候选同为规则参照非预测名单；官方按历年年末市值计股息率，此处以现价近似，年内大涨个股排名会偏高。
  官方纳入名单以 {adj_y} 年 11 月下旬公告为准。</p>"""
else:
    add_block = """
  <p style="margin-top:10px"><b>纳入候选参照</b>：本轮全市场扫描不可用（接口失败且无历史缓存），下次运行自动重试。</p>"""
adj_hist_html = ""
for h in ADJ_HISTORY:
    ins_txt = "、".join(h["ins"])
    outs_txt = "、".join(h["outs"])
    mark = "" if h["full"] else " <span style='color:#b45309'>（公开渠道仅部分恢复，完整名单以官网当期公告为准）</span>"
    adj_hist_html += (
        f"<details><summary><b>{h['year']}年调整</b> · 公告 {h['ann']} · 新样本自 {h['eff']} 启用"
        f"（公告20进20出{mark}）</summary>"
        f"<p style='margin:6px 0 2px'><b>纳入：</b>{ins_txt}</p>"
        f"<p style='margin:2px 0 8px'><b>剔除：</b>{outs_txt}</p></details>")

cons_html = f"""
<section><h2>成分股（官网权重文件 · 数据日期 {cons_date_cn}）</h2>
<details><summary style="cursor:pointer;font-weight:600;color:#0f172a;margin:4px 0">当前样本概览（点击展开/收起）</summary>
<div class="refbox" style="margin-top:4px">
  <p>成分股 <b>{len(cons_items)}</b> 只；前十大权重合计 <b>{'—' if top10_sum is None else f'{top10_sum:.1f}%'}</b>；
  成分股股息率加权均值(最近会计年度) <b>{'—' if wavg_dy is None else f'{wavg_dy:.2f}%'}</b>
  （个股价息率取自东方财富行情，为参照口径，非中证官方选样口径）。</p>
  <table><thead><tr><th>#</th><th>代码</th><th>名称</th><th>权重</th><th>股息率(年)</th></tr></thead>
  <tbody>{top10_rows}</tbody></table>
  <p class="refnote">仅列前十大权重，全部100只见官网权重文件（autofile/closeweight）。</p>{ind_block}
</div>
</details>
<details><summary style="cursor:pointer;font-weight:600;color:#0f172a;margin:8px 0 4px">下次定期调整预告 · 规则推算，非官方名单（点击展开/收起）</summary>
<div class="refbox" style="margin-top:4px">
  <p><b>生效日：{adj_eff_cn}</b>。{adj_note}。按编制方案，每次调整的样本比例一般不超过20%（即最多更换约20只），
  除非因不满足「过去一年现金股息率大于 0.5%」而剔除的原样本超过20%。</p>
  <p><b>缓冲区条款</b>（2022-12修订版，原样本不满足以下任一条件即失去样本资格）：
  ① 过去一年现金股息率 &gt; 0.5%；② 过去一年日均总市值位于中证全指样本空间前90%；
  ③ 过去一年日均成交金额位于中证全指样本空间前90%；④ 过去三年股利支付率均值在 0～1 之间。</p>
  <p><b>剔除候选参照</b>（当前样本中股息率最低的20只，按公告惯例20进20出取满额；口径为最近会计年度分红/现价，≤ 0.5% 将触发缓冲区硬条件，红色标注）：
  由于完整选样需全市场「过去三年平均股息率」排名（官方未公开逐股数据），下表仅为规则参照，<b>不构成调整名单预测</b>。</p>
  <table><thead><tr><th>代码</th><th>名称</th><th>股息率(年)</th></tr></thead>
  <tbody>{cand_rows}</tbody></table>{add_block}
  <p class="refnote">临时调整：样本退市即剔除；收购、合并、分拆等按指数计算与维护细则处理。</p>
</div>
</details>
<div class="refbox">
  <div class="reftitle">近五年样本调整史（每年12月生效，惯例20进20出）</div>
  {adj_hist_html}
  <p class="refnote">来源：中证指数官网当期公告及附件、权威媒体转载互证（2022/2023年官方未长期存档完整名单，仅部分恢复）。
  注：招商银行 2021 年被剔除、2025 年重新纳入；深高速/森马服饰/重庆百货/冀中能源等呈现"调出-回调入"的样本轮动特征。</p>
</div>
</section>
"""

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
<div class="ok"><b>说明：</b>主标的 <b>H00922 中证红利全收益指数</b>（含分红再投资），基准 <b>H00985 中证全指全收益</b>（同为全收益口径，两侧可比）。全部数据来自中证指数官网（免费、无需授权，含 2000 年至今完整历史）。</div>

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
<section><h2>长期趋势（对数-对数 · 二阶拟合）</h2>
<div class="refbox">
  <div class="reftitle">模型说明</div>
  <p><b>起点取 2016-01-01</b>（与 40日收益差图表周期对齐；样本同处 2013-07-02 股息率加权改革后的编制框架内，此前的市值加权序列不参与拟合）。图中灰色竖线标注 <b>2022-12-12</b>（分红连续性要求 2年→3年 + 股利支付率约束 + 个股权重上限）与 <b>2025-10-13</b> 两次编制方案修订日，属同一框架下的调整，不作断点剔除。</p>
  <p>ln(P) 对 ln(年) 做<b>二阶拟合</b>——指数为复利增长，log-log 空间下真实趋势是曲线，一阶幂律直线会穿中段、漏首尾（初始值塌陷），二阶项让曲线同时贴住起点与当前。</p>
  <p>图中置信带为拟合曲线 <b>±1.5×稳健标准差</b> 的包裹区间——σ 用 MAD 法估计（中位数绝对偏差×1.4826），避免崩盘日极端残差把带撑宽（经典 σ 实测被高估约 16%）；正态假设下约覆盖 86.6% 的单日观测。用于直观判断当前点位相对长期趋势的位置；注意该读数对起点选择敏感，宜作位置参照而非买卖信号。</p>
</div>
<div class="chart">{svg3}</div>
</section>

<section><h2>近五年"收盘价低于 MA500"区间汇总</h2>
<div class="refbox">
  <div class="reftitle">统计口径</div>
  <p>区间：{cut} ~ {last_date}（近五年，共 {len(win)} 个交易日）。以<b>全历史</b>计算 MA500（前置充足，无截断误差）。
  破位定义为当日收盘价 &lt; 当日 MA500；间隔 &le;5 个交易日的破位段视为同一波行情合并显示(避免二次探底被切独立段而极值低估)。</p>
  <table><thead><tr><th>均线</th><th>近五年破位占比</th><th>说明</th></tr></thead><tbody>
    <tr><td>MA250</td><td>{below_rates[250]:.1f}%</td><td>覆盖近 ~4.7 年(前置250)</td></tr>
    <tr><td>MA350</td><td>{below_rates[350]:.1f}%</td><td>覆盖近 ~4.5 年(前置350)</td></tr>
    <tr><td>MA500</td><td class="neg">{below_rate:.1f}%</td><td>覆盖完整近 5 年(前置500)</td></tr>
  </tbody></table>
  <p style="margin-top:10px"><b>破 MA500 明细（按持续天数降序，前 15 段）：</b></p>
  <table><thead><tr><th>起始</th><th>结束</th><th>天数</th><th>最深偏离</th><th>最深日</th><th>最深日收盘</th><th>最深日 MA500</th></tr></thead>
  <tbody>{ep_rows}</tbody></table>
  <p class="refnote">提示：单日假破位噪音较大；若作提醒条件，建议"连续 ≥3 日破位 且 偏离 &gt;2%"再触发，可过滤短假破位。</p>
</div></section>

{cons_html}

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
