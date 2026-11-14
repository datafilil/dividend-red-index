# AGENTS.md

本仓库是中证红利指数（H00922 全收益）的低频提醒系统。核心只有一件事：**每天取到官网当日数据，生成一份 HTML 日报，发布到 GitHub Pages。**

你的角色是**执行器**，不是重构者。除非用户明确要求改代码，否则不要修改任何 `.py` 文件。

---

## 一、日报的唯一正确跑法

```bash
pip install -r requirements.txt          # 只有 xlrd 一个依赖
python3 dividend_index_reminder.py
```

产物：

| 文件 | 说明 |
|---|---|
| `dividend_index_report.html` | 本轮生成的报告，约 320 KB |
| `cache/done_today.json` | 今日锁，出报后写入 |
| `cache/anchor.json` | 收敛锚点（首次探测到官网更新的时点） |
| `cache/add_candidates.json` | 全市场扫描结果（带 7 天保鲜守卫） |

脚本会往 `$GITHUB_OUTPUT` 写两个值（不在 CI 里则静默跳过）：`report_updated`、`today_done`。

**不要**给它传参数，除非用户明确要求强制重跑。`--force` 或 `FORCE_RUN=1` 会绕过时间门控、预检和今日锁，等于无视所有保护直接全量重跑。

---

## 二、四条硬约束（违反会静默出错，不会报异常）

### 1. 时间门控：北京 19:26 之前不取数

官网数据傍晚才放出（实测首现 17:01–23:21，中位约 19:30）。脚本第 183 行起有硬门控，19:26 前直接 `sys.exit(0)` 并打印「时间门控」——**这是正常退出，不是失败**，不要重试、不要加 `--force` 绕过。

时区固定按 UTC+8 判定（`_BJ` 常量），与机器本地时区无关。所以云端 VM 是 UTC 也不影响。

### 2. 收敛锚点：探测到更新后要等 30 分钟

官网的 perf / indicator / 成分三个文件是分批发布的。脚本探测到数据前进后不会立刻跑全量，而是先写 `cache/anchor.json`，等到「首次探测时点 + 30 分钟」再跑，避免拿到发布到一半的数据。

**因此：单次运行中途退出、提示「收敛等待中」是设计行为。** 不要在同一会话里反复催促重跑。

### 3. 幂等守卫：数据日期不严格前进就不发布

这是这个仓库最重要的一条规则，历史上因为缺它出现过「一天 16 次重复提交」。

发布前必须比较两个日期，**两者都从文件内容里 grep，不要用运行日期代替**：

```bash
NEW_DD=$(grep -oP "数据日期：\K[0-9-]+" dividend_index_report.html | head -1)
OLD_DD=$(git show HEAD:index.html | grep -oP "数据日期：\K[0-9-]+" | head -1)
```

判定表：

| 情况 | 动作 |
|---|---|
| `OLD_DD` 为空 | **不发布**（保守） |
| `NEW_DD` 为空 | **不发布** |
| `NEW_DD == OLD_DD` | **不发布**，连 `cp` 都不要执行 |
| `NEW_DD < OLD_DD` | **拒绝回退**，不发布 |
| `NEW_DD > OLD_DD` | 发布：`cp dividend_index_report.html index.html` |

### 4. `cache/` 在 .gitignore 里，但文件是被追踪的

`git add cache/xxx.json` 会被静默拒绝。必须 `git add -f`。

`cache/h00922.json`（约 290 KB）是 16 年全史缓存，增量取数依赖它。**永远不要删除、不要重排、不要「清理」它**——删了只能从官网重新拉全史。

---

## 三、发布流程

只有在幂等守卫判定为「发布」时才执行：

```bash
git config user.name "antigravity-agent"
git config user.email "agent@localhost"

cp dividend_index_report.html index.html
git add index.html
git add -f cache/add_candidates.json
[ -f cache/anchor.json ] && git add -f cache/anchor.json
[ -f cache/done_today.json ] && git add -f cache/done_today.json

# 提交信息用「生成日期」，不是运行日期
GEN=$(grep -oP "生成日期：\K[0-9-]+" dividend_index_report.html | head -1)
git commit -m "chore: 自动更新日报 ${GEN}"
git push origin HEAD:main
```

若 `git diff --cached --quiet` 为真（无变化），**跳过提交**，不要造空提交。

### 关于 `push_ingest.py`

它把报告推到线上 Cloudflare Worker 的 KV，需要环境变量 `INGEST_TOKEN`。

**在 Remote Agent 里通常没有这个 secret。没有就跳过，不要把 `INGEST_TOKEN` 写进代码、不要猜、不要换成别的地址。** GitHub Pages 是主发布通道，KV 只是备份，跳过后日报照常可访问。

---

## 四、禁止事项

- ❌ 手工编辑 `index.html` 或 `dividend_index_report.html`
- ❌ 删除或改写 `cache/*.json`
- ❌ 修改 `_BJ`、`CONVERGE_MINUTES`、19:26 门控这些常量
- ❌ 为了「让它跑起来」而加 `--force`
- ❌ 把 `dividend_index_report.html` 提交进仓库（历史上就是因为这个文件陈旧导致幂等失效）
- ❌ 改 `requirements.txt` 里的依赖版本，除非用户要求

---

## 五、验收标准

一次成功的运行应当能回答：

1. 报告里的「数据日期」是多少？比 `git show HEAD:index.html` 里的新吗？
2. 报告里的「生成日期」是多少？
3. `cache/done_today.json` 的 `data_date` 是否等于报告数据日期？
4. 有没有产生提交？提交信息里的日期是否等于生成日期？

四项对不上任何一项，说明流程有问题，**如实报告，不要自行修补后声称成功**。

---

## 六、环境自检

在任何新环境（新的 runner、云端 VM、海外主机）首次运行前，先跑：

```bash
python3 scripts/probe_sources.py
```

退出码 0 才继续。它会分别探测中证官网三个源和东财、腾讯的接口，并给出「关键源不通时具体缺什么」。
