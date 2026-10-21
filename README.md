# job-domain-research

对指定岗位域持续采集公开信号与真实招聘信息，按证据强度分级，产出带反证与验证实验的周期性研究输出，并维护可追溯的岗位数据。

当前版本：**0.4.0**

## 安装

```bash
git clone https://github.com/Lana-iris/job-domain-research.git ~/.config/opencode/skills/personal/job-domain-research
```

重启 OpenCode 后生效。

## 首次使用

在岗位域项目目录准备 `profile.md` 与 `research-state.md`；缺失时技能会自动进入 onboarding 流程。主方向可暂缺，技能会先协助拆分候选方向。

## 产出

- O1 信号日报
- O2a 岗位扫描采集
- O2b 岗位数据维护
- O3 定位校准
- O4 趋势报告
- O5 单次分析

## 岗位库工具（O2b）

字段规格以工具内置 `FIELD_SPEC` 为唯一真源：

```bash
python3 resources/jd_tool.py describe                 # Markdown 字段表
python3 resources/jd_tool.py describe --format json   # 机器可读规格
```

常用命令：

```bash
python3 resources/jd_tool.py --workbook <路径> audit
python3 resources/jd_tool.py --workbook <路径> add new.json
python3 resources/jd_tool.py --workbook <路径> refresh --evidence refresh.json
python3 resources/jd_tool.py --workbook <路径> purge --apply
python3 resources/jd_tool.py --workbook <路径> log --change "新增 JD270–JD272" --reason "日常扫描"
```

### 0.4.0 行为变化

- **14 个必填字段**：`add` / `audit` 按同一份 `FIELD_SPEC` 校验 14 项必填输入；`--strict` 额外要求 3 个分析字段（Market Signal / Portfolio Relevance / Portfolio Benchmark）。
- **去重与合并来源**：活跃表命中严格键或归一化键即视为重复；默认把新来源合并进 `全部来源`（`--no-merge-sources` 可关闭），不再新增行。
- **重新发布需证据**：`--allow-repost` 必须搭配 `--repost-evidence`，证据含原 JD、核实日期与状态。
- **刷新需证据**：首选 `refresh --evidence`（仅在招/可投更新最后发现日期）；手工路径须 `--confirm-visible --reason`。
- **运行级迭代日志**：日志固定 4 列（日期 / 变更项 / 统计 / 原因），一次 O2b 运行由 `log` 写一行，其它写操作不再自动追加。
- **列按表头定位**：失效表 `失效日期` / `失效原因` / `全部来源` 均按表头名补齐。

## 覆盖台账工具（O2a）

字段规格以工具内置 `COVERAGE_SPEC` 为唯一真源：

```bash
python3 resources/coverage_tool.py describe
python3 resources/coverage_tool.py --workbook <路径> append scan.json
python3 resources/coverage_tool.py --workbook <路径> update --date 2026-10-21 --platform 猎聘 --qp QP1 --city 深圳 --inbound 2
python3 resources/coverage_tool.py --workbook <路径> audit --date 2026-10-21 --plan plan.json
```

`plan.json` 由 `research-state.md` §三 岗位域 / 检索位矩阵生成：

```json
{"platforms": ["猎聘", "智联招聘"], "qps": [{"id": "QP1", "min_queries": 2}]}
```

`audit` 报告缺失平台、QP 查询数不足、受限、零结果与待回填入库数；无 `--plan` 时只能报告观测值，不能判断漏检。

## 依赖

- Python 3
- `openpyxl`（O2b 岗位库维护需要）

## License

MIT
