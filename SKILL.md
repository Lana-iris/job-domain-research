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
- `research-state.md`（研究，唯一真源；含 §二 信号域 与 §三 岗位域）
缺失任一 → 执行 `references/onboarding-flow.md`。
若 `profile.md` 主方向暂缺 → 先协助拆分候选方向（见 `references/onboarding-flow.md`），不强行先选定。

## 核心方法
1. 研究闭环：市场信号 → 证据 → 判断 → 假设 → 反证 → 实验 → 新证据（能力模型 / 产出物由 O3 承接）。
2. 证据分级：见 `references/evidence-grading.md`；结论强度不得超过证据强度。
3. 反证内联：每条假设下同时写"支持"与"反证"。
4. 增量输出：以 `research-state.md` 为基线，只写相较上期的变化；并参考最近 2–3 天日报（信号库即日报的存放目录）以保持连续性。
5. 强度守恒：证据不足时写"当前证据不足"，不补全。
6. 采集与筛选口径（默认）：关注过去 24 小时，必要时回溯 3 天；每日 3–5 个重要信号，同源合并；来源范围与筛选判据见 `references/signal-selection.md`；信号主题优先级见 `research-state.md` §二。
7. 待验证问题回写：O1 完成后按 `research-state.md` §五的回写要求更新「本期进展」「下一步验证」；O1 不改状态，状态由 O3 改。

## 默认风格（可被 `research-state.md` 输出偏好覆盖）
- 中文；结论先行；简明、清晰、具体；战略分析而非新闻摘要。
- 区分事实 / 推断 / 建议。
- 当日无重大变化时，直接写"今日无重大变化"。
- 英文来源提供中文摘要，并说明其意义。

## 产出
| 产出 | 模板 | 说明 |
|---|---|---|
| 日报（O1 市场信号 + O2 JD 扫描报告） | `templates/signal-daily.template.md` | 每日 |
| O2a 岗位扫描采集 | `references/jd-scan-rules.md` + `templates/platform-coverage.xlsx` + `resources/coverage_tool.py describe` | 每日（主动检索） |
| O2b 岗位数据维护 | `references/jd-maintenance-rules.md` + `templates/jd-library.template.xlsx` + `resources/jd_tool.py describe` | 每日（刷新/失效/去重/排序） |
| O3 定位校准 | `templates/positioning-review.template.md` | 周，条件触发 |
| O4 趋势报告 | `templates/domain-trend-report.template.md` | 月/季，多方向横比 |
| O5 单次分析 | 复用 O3/O4 | 触发式 |

## 工作流
- 日：独立执行 O1 信号采集和 O2 JD 流程。两者可按实际情况安排先后，不互相替代。
- 日报：将 O1 信号和 O2 JD 扫描报告分成两个部分；O2 数据不计入 O1 信号数量；O1 只使用自行检索到的公开来源，不读取 O2 岗位库。
- 周：O3（仅当 profile 目标含“探索/求职”）；输出 positioning-review 时同步更新 research-state.md §五 的状态（开放/已闭环/已并入产出/已降级）；默认独立成文，如需并入其他文件，在 research-state.md 声明。
- 月/季：O4
- 触发式：O5

各项产出的具体时间由 `research-state.md` 声明。

### O2 JD 流程与交接边界

1. 扫描：按岗位域检索矩阵执行查询，并逐条记录到覆盖台账。
2. 核验：检查候选岗位详情页、岗位信息和在招状态；合格候选交给维护阶段。
3. 维护：对合格候选入库，并按规则刷新、去重合并、处理失效和排序。
4. 交接：回填入库数，进入日终收尾。

产出：覆盖台账、JD 活跃库和变更计数（迭代日志由日终收尾写入）。
边界：扫描不写 JD 库；核验不修改评分；维护不重新搜索；refresh 只核验库内岗位。

### O2a 运行检查单

1. 先读覆盖台账字段口径（`coverage_tool.py describe`）。
2. 按 `research-state.md` §三 岗位域检索矩阵执行查询。
3. 每条查询逐条写入覆盖台账；未执行、受限、零结果分别如实标注。
4. 将详情页可读且仍在招的候选交给 O2b；其他候选记录原因。

### O2b 运行检查单

1. 先读岗位库字段与写入口径（`jd_tool.py describe`）。
2. 处理 O2a 核验合格的候选，以及需要刷新的库内岗位；执行入库、来源合并、去重、刷新和失效。
3. 变更落盘后，将新增/合并/去重/核验淘汰计数交接给日报 O2 部分；运行日志由日终收尾统一写入。

岗位库严格审计属于日终收尾，不在 O2b 阶段执行。

### 日终收尾与日报汇总

日终收尾按下方「日终收尾顺序」执行。日报分为两部分：
- O1 市场信号：独立按 §二 信号域采集和筛选。
- O2 今日 JD 扫描报告：总结入库、刷新、失效及岗位能力信号。

## 运行清单（每次运行按阶段读取）

| 阶段 | 读取 | 用途 |
|---|---|---|
| 通用 | `profile.md`；`research-state.md` | 目标、限制、研究问题、信号域、岗位域、输出偏好与节奏、待验证问题、反证、产出物、数据路径。 |
| O1 信号采集 | `references/evidence-grading.md`；`references/signal-selection.md`；`templates/signal-daily.template.md` | 按 §二 信号域与 signal-selection 采集、筛选、写日报。 |
| O2a 岗位扫描 | `references/jd-scan-rules.md`；`templates/platform-coverage.xlsx`；`resources/coverage_tool.py describe` | 按 §三 岗位域检索矩阵扫描、记录覆盖、核验候选。 |
| O2b 岗位维护 | `references/jd-maintenance-rules.md`；`templates/jd-library.template.xlsx`；`resources/jd_tool.py describe` | 处理 O2a 候选与库内刷新，维护 JD 库。 |
| 日终收尾与日报汇总 | `templates/signal-daily.template.md`；覆盖台账；岗位库及迭代日志 | 按日终收尾顺序执行；日报分 O1/O2 两部分。 |
| O3 定位校准 | `templates/positioning-review.template.md` | 定位校准（条件触发）。 |
| O4 趋势报告 | `templates/domain-trend-report.template.md` | 趋势报告。 |
| 档案缺失 | `references/onboarding-flow.md` | 补齐档案。 |

### 日终收尾顺序

覆盖审计 = 报告与缺口清单；岗位库严格审计 = 门禁。

1. 回填入库数（`coverage_tool.py update`）。
2. 运行覆盖审计（`coverage_tool.py audit`）。
3. 写运行日志（`jd_tool.py log`）。
4. 运行岗位库严格审计（`jd_tool.py audit --strict`）。若失败，按 audit 提示修复（`sort --apply` / `repair` / 补 `log`）后重跑；仍失败则在日报 O2 部分列为未完成。

## 交付清单（缺项时说明未完成或受限）

| 产出 | 完成判据 |
|---|---|
| O1 信号部分 | 按 §二 信号域与 signal-selection 完成采集和筛选；输出 3–5 条重要信号，不足时说明原因；每条信号含证据等级 E1–E5、Evidence、Why it matters、Career Implication、Counterevidence、Experiment；来源日期精确到日，未披露时注明；包含今日学习建议 3 问；待验证问题按 §五 回写。 |
| O2a 岗位扫描采集 | §三 检索矩阵中所有 query 均有记录（已检索/受限/未执行）；未执行与受限已写原因；合格候选已交接。 |
| O2b 岗位数据维护 | 变更已落盘；新增/合并/去重/核验淘汰计数已交接。 |
| 日终收尾 | 回填入库数；覆盖审计已运行；日志已写入；`audit --strict` 通过；日报分节呈现。 |
| 日报整体 | O1 市场信号与 O2 今日 JD 扫描报告分节呈现；O2 报告总结入库、刷新、失效及岗位能力信号。 |
| O3 定位校准 | 完成模板要求；支持与反证同框；只写相较上期的变化；回写 `research-state.md`（含 §五 状态：开放/已闭环/已并入产出/已降级）。 |
| O4 趋势报告 | 完成模板要求；包含多方向横向比较和证据等级。 |

## 参考文件
- templates/：profile · research-state · signal-daily · positioning-review · domain-trend-report · jd-library.template.xlsx · platform-coverage.xlsx
- references/：evidence-grading · signal-selection · jd-scan-rules · jd-maintenance-rules · onboarding-flow
- resources/：jd_tool.py（岗位库维护执行器）· coverage_tool.py（检索覆盖台账执行器）
