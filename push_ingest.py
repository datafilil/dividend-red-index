# -*- coding: utf-8 -*-
"""把「报告」或「全市场扫描缓存」推给线上 Worker（CI 用，零 Cloudflare 凭据）。

为什么要走 Worker 而不是直写 KV
    写 Workers KV 需要一份 Cloudflare 凭据（API Token / OAuth）。把 CF 凭据放进
    GitHub Actions 并不理想（secret 分散、权限过大、要人工轮换）。Worker 侧因此
    开了两个端点，CI 只需要一个共享口令，由 Worker 用**它自己的 KV 绑定**写库：
        POST /ingest        body = 报告 HTML
        POST /ingest-scan   body = add_candidates JSON（全市场扫描结果）
    于是 CI 里 **零 Cloudflare 凭据**。

为什么扫描缓存也要推
    全市场扫描（约 35 页分红明细 + 56 页行情快照）是整条链路唯一的重活，线上
    Worker 跑它会撞 128MB 资源上限。让外部算好后推进 KV，Worker 就能一直命中
    缓存、永不自己扫 —— 这是让 Worker 具备日更能力的关键一步。

幂等性
    两个端点都带单调守卫：更旧/相同的推送一律 stored=false，不会把新数据覆盖回旧版本。

用法
    python push_ingest.py                             # 报告 -> /ingest
    python push_ingest.py path/to/report.html
    python push_ingest.py --scan                      # cache/add_candidates.json -> /ingest-scan
    python push_ingest.py --scan path/to/add_candidates.json
    python push_ingest.py --dry-run [--scan]          # 只校验不发送

环境变量
    INGEST_URL     默认 https://dividend-index-reminder.datat.workers.dev/ingest
                   （--scan 时自动改用同域的 /ingest-scan）
    INGEST_TOKEN   必填；与 Worker 的 INGEST_TOKEN secret 一致。
                   本地调试可放在同目录 .ingest_token 文件里（不要提交）。
"""
import json
import os
import re
import sys
import urllib.error
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
BASE_URL = "https://dividend-index-reminder.datat.workers.dev"
DEFAULT_REPORT = "dividend_index_report.html"
DEFAULT_SCAN = os.path.join("cache", "add_candidates.json")
UA = "dividend-index-uploader/1.0"


def load_token():
    t = os.environ.get("INGEST_TOKEN", "").strip()
    if t:
        return t
    for p in (os.path.join(HERE, ".ingest_token"), os.path.join(os.getcwd(), ".ingest_token")):
        if os.path.exists(p):
            return open(p, encoding="utf-8").read().strip()
    return ""


def endpoint(scan_mode):
    base = os.environ.get("INGEST_URL", BASE_URL + "/ingest")
    if not scan_mode:
        return base
    # 把结尾的 /ingest 换成 /ingest-scan（保留自定义 host）
    return re.sub(r"/ingest/?$", "/ingest-scan", base) if re.search(r"/ingest/?$", base) \
        else base.rstrip("/") + "/ingest-scan"


def resolve(path):
    return path if os.path.isabs(path) else os.path.join(HERE, path)


def main():
    argv = sys.argv[1:]
    dry = "--dry-run" in argv
    scan_mode = "--scan" in argv
    args = [a for a in argv if not a.startswith("--")]

    if scan_mode:
        path = resolve(args[0] if args else DEFAULT_SCAN)
    else:
        path = resolve(args[0] if args else DEFAULT_REPORT)

    if not os.path.exists(path):
        print("找不到文件：%s" % path, file=sys.stderr)
        return 2

    raw = open(path, "rb").read()
    label = "扫描缓存" if scan_mode else "报告"

    if scan_mode:
        try:
            payload = json.loads(raw)
        except Exception as e:
            print("不是合法 JSON：%s" % e, file=sys.stderr)
            return 3
        gen = payload.get("generated")
        if not gen or not payload.get("divs_slim"):
            print("拒绝发送：缺少 generated / divs_slim 字段", file=sys.stderr)
            return 3
        print("%s: %s (%d bytes)" % (label, path, len(raw)))
        print("  generated=%s  divs=%d  quotes=%d"
              % (gen, len(payload.get("divs_slim") or {}),
                 len(payload.get("quotes_slim") or {})))
    else:
        text = raw[:400000].decode("utf-8", "replace")
        m = re.search(r"生成日期：(\d{4}-\d{2}-\d{2})", text)
        if not m:
            print("拒绝发送：HTML 里没有「生成日期：YYYY-MM-DD」标记", file=sys.stderr)
            return 3
        print("%s: %s (%d bytes)" % (label, path, len(raw)))
        print("  生成日期=%s" % m.group(1))

    if dry:
        print("--dry-run：校验通过，未发送")
        return 0

    token = load_token()
    if not token:
        print("缺少 INGEST_TOKEN（环境变量或 .ingest_token 文件）", file=sys.stderr)
        return 4

    url = endpoint(scan_mode)
    ctype = "application/json; charset=utf-8" if scan_mode else "text/html; charset=utf-8"
    req = urllib.request.Request(
        url, data=raw, method="POST",
        headers={"X-Ingest-Token": token, "Content-Type": ctype, "User-Agent": UA})
    try:
        with urllib.request.urlopen(req, timeout=150) as r:
            out = r.read().decode("utf-8", "replace")
            print("HTTP %s -> %s" % (r.status, out))
            return 0 if json.loads(out).get("ok") is True else 5
    except urllib.error.HTTPError as e:
        print("HTTP %s -> %s" % (e.code, e.read().decode("utf-8", "replace")[:400]),
              file=sys.stderr)
        return 1
    except Exception as e:
        print("发送失败：%s: %s" % (type(e).__name__, e), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
