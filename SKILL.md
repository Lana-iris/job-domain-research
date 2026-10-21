---
name: job-domain-research
description: Use when 需要对某岗位域/职能做持续研究（岗位能力要求、招聘市场信号、AI 对某职能的影响、职业定位校准、岗位域趋势报告），或维护可追溯的招聘/岗位数据与周期性研究输出时。
---

# job-domain-research

对指定岗位域持续采集公开信号与真实招聘信息，按证据强度分级，产出带反证与验证实验的周期性研究输出，并维护可追溯的岗位数据。

## 何时使用
- 持续跟踪某岗位域的市场信号与真实招聘
- 研究岗位能力要求 / 招聘市场信号
- 跟踪 AI 对某职能的影响
- 职业定位校准
- 输出岗位域趋势报告

## 不适用 / 边界
- 不做调度：本技能只定义"一次运行"。
- 不替使用者做最终职业决策。
- 不写死任何领域词表、来源或输出路径。

## 最终目标
用真实市场证据持续推进定位与产出，不追求"证明某个选择正确"。
闭环：市场证据 → 职业假设 → 反证 → 能力模型 → 作品集 → 实验 → 新证据。

## 运行前
读取当前岗位域项目下的知识层：
- `profile.md`（个人）
- `research-state.md`（研究，唯一真源；含 检索位矩阵 与 信号优先类别）
缺失任一 → 执行 `references/onboarding-flow.md`。
若 `profile.md` 主方向暂缺 → 先协助拆分候选方向（见 `references/onboarding-flow.md`），不强行先选定。

## 核心方法
1. 研究闭环：市场信号 → 证据 → 判断 → 假设 → 反证 → 实验 → 新证据（能力模型 / 产出物由 O3 承接）。
2. 证据分级：见 `references/evidence-grading.md`；结论强度不得超过证据强度。
3. 反证内联：每条假设下同时写"支持"与"反证"。
4. 增量输出：以 `research-state.md` 为基线，只写相较上期的变化；并参考最近 2–3 天信号库以保持连续性。
5. 强度守恒：证据不足时写"当前证据不足"，不补全。
6. 采集与筛选口径（默认）：关注过去 24 小时，必要时回溯 3 天；每日 3–5 个重要信号，同源合并；筛选判据见 `references/signal-selection.md`；岗位域的「信号优先类别」见 `research-state.md`。
7. 数据层：信号库目录与文件命名由 `research-state.md` 声明（默认 `YYYY-MM-DD-日报.md`）。
8. 待验证问题回写：O1 完成后逐条更新 research-state.md §四的「本期进展」「下一步验证」；有新证据写结论+核心证据+强度，无变化写“YYYY-MM-DD：无变化（本轮未获新证据）”；O1 不改状态，状态由 O3 改。

## 默认风格（可被 `research-state.md` 输出偏好覆盖）
- 中文；结论先行；简明、清晰、具体；战略分析而非新闻摘要。
- 区分事实 / 推断 / 建议。
- 当日无重大变化时，直接写"今日无重大变化"。
- 英文来源提供中文摘要，并说明其意义。

## 产出
| 产出 | 模板 | 说明 |
|---|---|---|
| O1 信号日报 | `templates/signal-daily.template.md` | 每日 |
| O2a 岗位扫描采集 | `references/jd-scan-rules.md` + `templates/platform-coverage.xlsx` + `resources/coverage_tool.py describe` | 每日（主动检索） |
| O2b 岗位数据维护 | `references/jd-maintenance-rules.md` + `templates/jd-library.template.xlsx` + `resources/jd_tool.py describe` | 每日（刷新/失效/去重/排序） |
| O3 定位校准 | `templates/positioning-review.template.md` | 周，条件触发 |
| O4 趋势报告 | `templates/domain-trend-report.template.md` | 月/季，多方向横比 |
| O5 单次分析 | 复用 O3/O4 | 触发式 |

## 工作流
- 日：O1 + O2a + O2b
- 周：O3（仅当 profile 目标含"探索/求职"）；O3 默认独立成文，如需并轨到某文件末尾，在 `research-state.md` 声明。
- 月/季：O4
- 触发式：O5

节奏的绝对时间（几点/周几/几号）由 `research-state.md` 声明。

## O2 四步与交接边界

1. 扫描并记覆盖 → `jd-scan-rules.md` + `coverage_tool.py`
   产出：`platform-coverage.xlsx`
2. 核验候选 → `jd-scan-rules.md`「候选核验」
   只把详情页可读、仍在招的岗位送入下一步
3. 更新岗位库 → `jd-maintenance-rules.md` + `jd_tool.py`
   产出：`JD扫描-活跃库.xlsx` + 迭代日志
4. 产出日报 → `signal-daily.template.md`
   读取覆盖与库变化，写 O1

边界：扫描不写库；核验不改分；维护不重新检索；refresh 只核验既有岗位详情页；日报不改库。

## O2a 运行检查单（按序执行）

1. 先读字段规格：`python3 coverage_tool.py describe`。
2. 按 `research-state.md` §二矩阵执行 query；逐条 `append` 到覆盖台账。
3. 收尾运行 `coverage_tool.py audit --date YYYY-MM-DD --plan plan.json`；`plan.json` 由 §二矩阵生成（`platforms` + `qps[].min_queries`）。
4. O2b 完成后用 `coverage_tool.py update` 回填 `入库数`。

## O2b 运行检查单（按序执行）

1. 先读字段规格：`python3 jd_tool.py describe`（字段名、取值、写入规则以它为准）。
2. 按规格入库 / 刷新 / 失效 / 去重排序；重复项默认合并来源，重新发布必须带证据。
3. 写一行运行级日志：`python3 jd_tool.py log --change "..." --reason "..."`（4 列：日期 / 变更项 / 统计 / 原因）。
4. 收尾运行 `audit --strict`，确认日志统计与当前库一致、无新增问题。

## 交付清单（缺任一项 = 本轮未完成，不得宣称完成）
| 产出 | 完成判据 |
|---|---|
| O1 信号日报 | 3–5 条 ×（证据等级 E1–E5 + Evidence / Why it matters / Career Implication / Counterevidence / Experiment）；今日学习建议 3 问；今日 JD 扫描报告；发布日期精确到日（月/年须标"日未披露"）；待验证问题表已逐条判断并回写（无变化已注明） |
| O2a 岗位扫描采集 | 检索位矩阵全部 query 执行；平台覆盖台账逐平台有状态；覆盖审计已运行，缺失/受限/未执行已报告 |
| O2b 岗位数据维护 | 新增 / 刷新 / 失效；`全部来源` 已合并；迭代日志 4 列已写入且统计与当前库一致；`log` 后 `audit --strict` 通过 |
| O3 定位校准 | ①–④；支持↔反证同框；只写相较上期变化；回写 `research-state.md` |
| O4 趋势报告 | ①–④；多方向横比；证据等级 |

## 运行清单（每次运行按阶段按需读取）
| 阶段 | 读取 | 用到的字段 |
|---|---|---|
| 通用 | `profile.md`；`research-state.md` | profile：目标方向 / 目标 / 限制；research-state：研究问题 / 信号域 / 输出偏好 / 节奏表 |
| O1 信号日报 | `references/evidence-grading.md`；`references/signal-selection.md`；`templates/signal-daily.template.md` | research-state：§二 检索位矩阵 + 信号优先类别、§三 输出偏好、§五 已知反证、§四 待验证表、§六 产出物状态；profile：主方向、能力现状 |
| O2a 岗位扫描采集 | `references/jd-scan-rules.md`；`templates/platform-coverage.xlsx`；`resources/coverage_tool.py describe` | research-state：§二 检索位矩阵；profile：目标方向、能力现状、限制（地点） |
| O2b 岗位数据维护 | `references/jd-maintenance-rules.md`；`templates/jd-library.template.xlsx`；`resources/jd_tool.py describe` | research-state：§二 检索位矩阵；profile：目标方向、能力现状、限制（地点）；工具可用 --config 覆盖默认权重/城市/阈值 |
| O3 定位校准 | `templates/positioning-review.template.md` | research-state：§四 待验证表、§五 已知反证、§六 产出物状态、§一 研究问题；profile：目标方向、背景、限制、目标 |
| O4 趋势报告 | `templates/domain-trend-report.template.md` | research-state：§二 信号域、§四 / §五 / §六；profile：主方向 + 候选方向 |
| 档案缺失 | `references/onboarding-flow.md` | — |

## 参考文件
- templates/：profile · research-state · signal-daily · positioning-review · domain-trend-report · jd-library.template.xlsx · platform-coverage.xlsx
- references/：evidence-grading · signal-selection · jd-scan-rules · jd-maintenance-rules · onboarding-flow
- resources/：jd_tool.py（岗位库维护执行器）· coverage_tool.py（检索覆盖台账执行器）
