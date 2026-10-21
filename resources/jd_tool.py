#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
jd_tool.py — JD 活跃库维护工具（通用）

把原先「每天手写一个新脚本」的做法收敛成一个可复用工具：同一套权重公式、
同一套去重 / 失效 / 备份 / 重排逻辑，只写一次，不再每天重抄。

用法：
    python3 jd_tool.py describe [--format markdown|json]      # 字段规格（唯一真源）
    python3 jd_tool.py --workbook <路径> audit                # 只读体检，不改数据
    python3 jd_tool.py --workbook <路径> add new.json         # 追加当日新 JD（JSON 数组）
    python3 jd_tool.py --workbook <路径> add new.json --strict # strict 三列缺任一即拒收
    python3 jd_tool.py --workbook <路径> add new.json --no-merge-sources  # 重复时不合并来源
    python3 jd_tool.py --workbook <路径> add new.json --allow-repost \
        --repost-evidence repost.json                         # 重新发布（必须带证据）
    python3 jd_tool.py --workbook <路径> refresh --evidence refresh.json  # 首选：带证据刷新
    python3 jd_tool.py --workbook <路径> refresh --ids JD251,JD256 \
        --confirm-visible --reason "详情页可见"               # 手工刷新（必须确认可见）
    python3 jd_tool.py --workbook <路径> purge                 # 超期移入失效（默认只预览）
    python3 jd_tool.py --workbook <路径> purge --apply         # 真正执行
    python3 jd_tool.py --workbook <路径> sort                  # 按综合优先级分降序（默认只预览）
    python3 jd_tool.py --workbook <路径> sort --apply
    python3 jd_tool.py --workbook <路径> repair                # 修脏数据（默认只预览）
    python3 jd_tool.py --workbook <路径> repair --apply
    python3 jd_tool.py --workbook <路径> set --file fix.json   # 补空值单元格（默认只预览）
    python3 jd_tool.py --workbook <路径> recompute             # 按存量输入列重算分值（默认只预览）
    python3 jd_tool.py --workbook <路径> recompute --apply
    python3 jd_tool.py --workbook <路径> backup                # 手动备份
    python3 jd_tool.py --workbook <路径> log --change "新增 JD270–JD272" \
        --reason "日常扫描" [--before-active N --before-dead M] [--dry-run]

    --workbook 默认当前目录下的「JD扫描-活跃库.xlsx」；--config 可覆盖默认规格（见下）。

规则真源：见本技能 references/jd-maintenance-rules.md（权重/失效/文件策略等默认）。
字段规格真源：本文件 FIELD_SPEC，describe 直接由它生成；add / audit 用同一份规格校验。

R 时效「未知」口径（2026-10-16 用户裁定）：
    未知取 0.4（与 >90 天同级，不把「发布日期不明」当成新鲜）；
    存量数据按需执行 recompute 重算。--r-unknown 可临时覆盖该默认值。

设计要点
- 列一律按「表头名」定位，绝不按列号硬编码（2026-09-22 曾因列号错位漏了 10 条）。
- 任何写操作前自动备份到工作簿同目录下的 `备份/`。
- purge / sort / repair / set / recompute 默认 dry-run，必须显式 --apply 才落盘；
  add / refresh 即时写入；audit 只读；log 写一行运行级日志（默认即写入，--dry-run 预览）。
- add 默认校验 14 个必填输入列；加 --strict 时还要求分析三列（Market Signal /
  Portfolio Relevance / Portfolio Benchmark）非空。校验口径见 FIELD_SPEC。
- 活跃表命中严格键或归一化键 → 视为重复；默认把新来源合并进 `全部来源`，不新增行。
- 失效表命中同样视为重复；重新发布必须 `--allow-repost` + `--repost-evidence` 同时具备。
- 迭代日志为「一次运行一行」的 4 列（日期 / 变更项 / 统计 / 原因），由 `log` 显式写入；
  其它子命令不再自动追加日志。
- 查重同时用「严格键」与「归一化键」（公司名剥离尾部标注 + 城市取主体）；
  归一化**只用于查重，不改写表内原始值**（2026-10-14 新增，修 JD262/JD182 类漏判）。
"""

from __future__ import annotations

import argparse
import copy as _copy
import datetime as dt
import json
import re
import shutil
import sys
from pathlib import Path
from typing import Any

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("需要 openpyxl：pip install openpyxl")

# ---------------------------------------------------------------- 默认规格（可被 --config 覆盖）

DEFAULT_CONFIG = {
    "sheets": {"active": "JD扫描(活跃)", "dead": "JD历史·已失效"},
    "weights": {"R": 0.30, "C": 0.28, "F": 0.22, "Ci": 0.10, "A": 0.10},
    "city": {"深圳": 1.0, "广州": 1.0, "杭州": 1.0,
             "北京": 0.9, "上海": 0.9, "香港": 0.9, "远程": 0.9},
    "city_default": 0.5,
    "stale_days": 14,  # 最后发现日期距今 > 该天数未再发现 -> 移入失效
}


def load_config(path):
    """载入 --config JSON，逐顶层键覆盖默认规格；缺省返回默认副本。"""
    cfg = _copy.deepcopy(DEFAULT_CONFIG)
    if path is None:
        return cfg
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    for k, v in data.items():
        if isinstance(v, dict) and isinstance(cfg.get(k), dict):
            cfg[k].update(v)
        else:
            cfg[k] = v
    return cfg


# 运行时生效的规格（main 中按 --config 重置）
CONFIG = _copy.deepcopy(DEFAULT_CONFIG)


# ---------------------------------------------------------------- 字段规格（唯一真源）
#
# 分组：
#   必填输入      —— 入库时必须非空（共 14 项）
#   严格分析输入  —— 仅 --strict / audit --strict 要求非空（共 3 项）
#   可选          —— 有则写，无则留空
#   生成          —— 由 add / recompute 计算，禁止从 JSON 输入
#   失效表        —— 只在失效表出现

GROUP_REQUIRED = "必填输入"
GROUP_STRICT = "严格分析输入"
GROUP_OPTIONAL = "可选"
GROUP_GENERATED = "生成"
GROUP_DEAD = "失效表"

FIELD_SPEC: list[dict[str, Any]] = [
    {
        "name": "去重ID",
        "meaning": "库内唯一编号",
        "group": GROUP_GENERATED,
        "required": "系统生成",
        "type": "JD+数字",
        "write_rule": "由 add 按现有最大编号 +1 生成，禁止手填或从 JSON 传入",
        "example": "JD270",
    },
    {
        "name": "抓取日期",
        "meaning": "本条首次入库的业务日期",
        "group": GROUP_GENERATED,
        "required": "系统生成",
        "type": "YYYY-MM-DD",
        "write_rule": "由 add 写入 --today（业务日），不随刷新变化",
        "example": "2026-10-21",
    },
    {
        "name": "来源平台",
        "meaning": "最优单一来源平台",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "非空文本（平台名）",
        "write_rule": "写平台本名，如猎聘、BOSS直聘；多平台合并只保留最优来源",
        "example": "猎聘",
    },
    {
        "name": "公司名称",
        "meaning": "招聘方 / 雇主",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "非空文本",
        "write_rule": "写招聘页上的公司名；括号标注可保留，归一化只用于查重",
        "example": "某科技公司",
    },
    {
        "name": "公司行业分",
        "meaning": "行业与目标领域的匹配度",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "0.5 / 0.8 / 1.0 / 未知",
        "write_rule": "只能取 0.5、0.8、1.0；不明写「未知」（计分按 0.5）",
        "example": "1.0",
    },
    {
        "name": "公司梯队分",
        "meaning": "公司梯队 / 规模评分",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "0.6–1.0 数字 / 未知",
        "write_rule": "0.6–1.0 的数字（大厂 1.0、外企 0.9、中厂 0.8、创业 0.6）；不明写「未知」（计分按 0.7）",
        "example": "0.8",
    },
    {
        "name": "职位名称",
        "meaning": "岗位名称（含批次/层级差异）",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "非空文本",
        "write_rule": "原样抄录职位名；校招/社招、层级差异写入职位名称，不新增字段",
        "example": "AI 产品经理",
    },
    {
        "name": "工作城市",
        "meaning": "岗位所在城市",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "非空文本",
        "write_rule": "写城市主体（如「北京」）；区/商圈可附「-」，归一化只用于查重",
        "example": "北京",
    },
    {
        "name": "经验要求",
        "meaning": "经验年限要求",
        "group": GROUP_OPTIONAL,
        "required": "否",
        "type": "文本",
        "write_rule": "照原文摘要，如「3–5 年」；未写明可留空",
        "example": "3–5 年",
    },
    {
        "name": "学历要求",
        "meaning": "学历门槛",
        "group": GROUP_OPTIONAL,
        "required": "否",
        "type": "文本",
        "write_rule": "照原文写，如「本科」「硕士优先」",
        "example": "本科",
    },
    {
        "name": "薪资范围(原文)",
        "meaning": "招聘页原始薪资",
        "group": GROUP_OPTIONAL,
        "required": "否",
        "type": "文本",
        "write_rule": "照抄原文，不改写单位",
        "example": "30–50k·15薪",
    },
    {
        "name": "薪资标准化(万/年)",
        "meaning": "换算后的年薪（万元）",
        "group": GROUP_OPTIONAL,
        "required": "否",
        "type": "数字（万/年）",
        "write_rule": "按原文中值折算为万元/年；无法折算留空",
        "example": "45",
    },
    {
        "name": "发布日期",
        "meaning": "岗位发布日",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "YYYY-MM-DD / YYYY-MM（日未披露） / 未知",
        "write_rule": "优先精确到日；仅月/年写 YYYY-MM（日未披露）；拿不到写「未知」（R=0.4）",
        "example": "2026-10-16",
    },
    {
        "name": "截止日期",
        "meaning": "投递截止日",
        "group": GROUP_OPTIONAL,
        "required": "否",
        "type": "YYYY-MM-DD / YYYY-MM（日未披露）",
        "write_rule": "与发布日期同格式；多为估算，不据此自动清理",
        "example": "2026-11-01",
    },
    {
        "name": "首次发现日期",
        "meaning": "首次被本库记录的日期",
        "group": GROUP_GENERATED,
        "required": "系统生成",
        "type": "YYYY-MM-DD",
        "write_rule": "由 add 写入 --today，之后不再改动",
        "example": "2026-10-21",
    },
    {
        "name": "最后发现日期",
        "meaning": "最近一次核实仍在招/可投的日期",
        "group": GROUP_GENERATED,
        "required": "系统生成",
        "type": "YYYY-MM-DD",
        "write_rule": "仅 refresh 在「在招/可投」证据下更新；>14 天未更新会被 purge",
        "example": "2026-10-21",
    },
    {
        "name": "核心职责",
        "meaning": "岗位职责摘要",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "非空文本",
        "write_rule": "摘要 2–4 条职责，不整段照抄",
        "example": "负责 AI 产品规划与落地",
    },
    {
        "name": "AI·数字化技能要求",
        "meaning": "岗位要求的 AI / 数字化能力",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "非空文本",
        "write_rule": "摘要原文中的 AI / 数据 / 数字化技能要求",
        "example": "熟悉大模型应用与数据分析",
    },
    {
        "name": "岗位/职能专有能力要求",
        "meaning": "该岗位域的专业能力要求",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "非空文本",
        "write_rule": "写领域专有能力，不写通用软技能",
        "example": "具备 B 端产品方法论",
    },
    {
        "name": "人才标准信号",
        "meaning": "招聘方衡量人才的显性/隐性标准",
        "group": GROUP_OPTIONAL,
        "required": "否",
        "type": "文本",
        "write_rule": "摘录能反映人才标准的原句或推断，标注推断",
        "example": "偏好有 0→1 经验者",
    },
    {
        "name": "契合度(1-5)",
        "meaning": "与个人主方向的契合程度",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "整数 1–5",
        "write_rule": "按主方向判断，写 1–5 的整数",
        "example": "4",
    },
    {
        "name": "当前可达性(1-5)",
        "meaning": "以当前资历投进的可行性",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "整数 1–5",
        "write_rule": "不等于契合度；随产出物推进应逐月上升，写 1–5 的整数",
        "example": "3",
    },
    {
        "name": "时效权重R",
        "meaning": "按发布日计算的时效分",
        "group": GROUP_GENERATED,
        "required": "系统生成",
        "type": "0.3–1.0（未知=0.4）",
        "write_rule": "≤7 天 1.0；8–30 天 0.75；31–90 天 0.5；>90 天 0.3；未知 0.4",
        "example": "1.0",
    },
    {
        "name": "公司权重C",
        "meaning": "公司维度得分",
        "group": GROUP_GENERATED,
        "required": "系统生成",
        "type": "0–1",
        "write_rule": "C = 0.6×行业分 + 0.4×梯队分（未知分别按 0.5 / 0.7）",
        "example": "0.92",
    },
    {
        "name": "契合权重F",
        "meaning": "契合度归一化分",
        "group": GROUP_GENERATED,
        "required": "系统生成",
        "type": "0–1",
        "write_rule": "F = 契合度(1–5) / 5",
        "example": "0.8",
    },
    {
        "name": "城市权重Ci",
        "meaning": "城市维度得分",
        "group": GROUP_GENERATED,
        "required": "系统生成",
        "type": "0–1",
        "write_rule": "深/广/杭 1.0；北/上/港/远程 0.9；其他 0.5（可由 --config 覆盖）",
        "example": "0.9",
    },
    {
        "name": "综合优先级分",
        "meaning": "库内排序分",
        "group": GROUP_GENERATED,
        "required": "系统生成",
        "type": "0–100",
        "write_rule": "100×(0.30R + 0.28C + 0.22F + 0.10Ci + 0.10A)，仅用于内部排序",
        "example": "88.4",
    },
    {
        "name": "Market Signal",
        "meaning": "市场为何种能力付费",
        "group": GROUP_STRICT,
        "required": "是（--strict）",
        "type": "非空文本",
        "write_rule": "写市场信号判断，不写空泛描述；--strict 时缺则拒收",
        "example": "企业愿为 AI 落地能力付溢价",
    },
    {
        "name": "Portfolio Relevance",
        "meaning": "对个人产出物的参考价值",
        "group": GROUP_STRICT,
        "required": "是（--strict）",
        "type": "非空文本",
        "write_rule": "写该岗位对作品集的参考点；--strict 时缺则拒收",
        "example": "可复用其 B 端案例结构",
    },
    {
        "name": "Portfolio Benchmark",
        "meaning": "需要用作品证明什么",
        "group": GROUP_STRICT,
        "required": "是（--strict）",
        "type": "非空文本",
        "write_rule": "写作品需覆盖的能力证据；--strict 时缺则拒收",
        "example": "需证明 0→1 落地能力",
    },
    {
        "name": "标签",
        "meaning": "状态 / 主题标签",
        "group": GROUP_OPTIONAL,
        "required": "否",
        "type": "文本",
        "write_rule": "自由标签，多个用「；」分隔",
        "example": "高信号；AI",
    },
    {
        "name": "定位启发",
        "meaning": "对职业定位的启发",
        "group": GROUP_OPTIONAL,
        "required": "否",
        "type": "文本",
        "write_rule": "写一句话启发，区分事实 / 推断",
        "example": "可转向 AI 解决方案岗",
    },
    {
        "name": "关键词命中",
        "meaning": "命中的检索关键词",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "非空文本",
        "write_rule": "写实际命中的检索词，多个用「；」分隔",
        "example": "AI 产品经理",
    },
    {
        "name": "原文链接",
        "meaning": "岗位详情页证据链接",
        "group": GROUP_REQUIRED,
        "required": "是",
        "type": "http(s) URL",
        "write_rule": "写详情页 URL，不写搜索页或聚合页",
        "example": "https://example.com/job/123",
    },
    {
        "name": "全部来源",
        "meaning": "同岗多平台来源汇总",
        "group": GROUP_OPTIONAL,
        "required": "否",
        "type": "平台|URL；平台|URL",
        "write_rule": "最优来源在前；为空时用「来源平台|原文链接」初始化；同 URL 不重复追加",
        "example": "猎聘|https://a.example/job/1；BOSS直聘|https://b.example/job/2",
    },
    {
        "name": "备注/质量说明",
        "meaning": "自由备注与质量标记",
        "group": GROUP_OPTIONAL,
        "required": "否",
        "type": "文本",
        "write_rule": "写数据质量说明、日报来源、重新发布等标记",
        "example": "【来自日报】",
    },
    {
        "name": "失效日期",
        "meaning": "移入失效表的日期",
        "group": GROUP_DEAD,
        "required": "失效时必填",
        "type": "YYYY-MM-DD",
        "write_rule": "由 purge 写入业务日；按表头名定位，不按列号",
        "example": "2026-10-21",
    },
    {
        "name": "失效原因",
        "meaning": "判定失效的依据",
        "group": GROUP_DEAD,
        "required": "失效时必填",
        "type": "非空文本",
        "write_rule": "由 purge 写入，如「最后发现 … 距今 >14 天未再发现」",
        "example": "最后发现 2026-09-30，距今 >14 天未再发现",
    },
]

# add / audit 的必填与 strict 清单都从 FIELD_SPEC 派生，避免维护第二份名单
REQUIRED_INPUTS = [f["name"] for f in FIELD_SPEC if f["group"] == GROUP_REQUIRED]
STRICT_INPUTS = [f["name"] for f in FIELD_SPEC if f["group"] == GROUP_STRICT]
GENERATED_FIELDS = [f["name"] for f in FIELD_SPEC if f["group"] == GROUP_GENERATED]
OPTIONAL_FIELDS = [f["name"] for f in FIELD_SPEC if f["group"] == GROUP_OPTIONAL]
DEAD_FIELDS = [f["name"] for f in FIELD_SPEC if f["group"] == GROUP_DEAD]

# 「未知」只允许出现在这三个必填字段
UNKNOWN_OK = {"发布日期", "公司行业分", "公司梯队分"}
INDUSTRY_ALLOWED = {0.5, 0.8, 1.0}
DEFAULT_INDUSTRY = 0.5
DEFAULT_TIER = 0.7

URL_RE = re.compile(r"https?://\S+", re.I)
DATE_FULL_RE = re.compile(r"\d{4}-\d{2}-\d{2}")
DATE_MONTH_RE = re.compile(r"\d{4}-\d{2}（日未披露）")
SOURCES_SPLIT_RE = re.compile(r"[；;]")


def _rel(p: Path) -> str:
    try:
        return str(Path(p).relative_to(Path.cwd()))
    except ValueError:
        return str(p)


# ---------------------------------------------------------------- 日期

def parse_date(v):
    """容错解析：datetime / date / '2026-10-07' / '2026-07-31(更新)' / '2026-09' / '2026'。

    解析失败（如 2026-13-40）返回 None，不抛异常。
    """
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if v is None:
        return None
    s = str(v).strip()
    try:
        m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)          # 2026-10-07（可带括号后缀）
        if m:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        m = re.match(r"(\d{4})-(\d{1,2})(?!\d)", s)              # 2026-09
        if m:
            return dt.date(int(m.group(1)), int(m.group(2)), 1)
        m = re.match(r"(\d{4})(?!\d)", s)                        # 2026
        if m:
            return dt.date(int(m.group(1)), 1, 1)
    except ValueError:
        return None
    return None


def date_value_problem(v, allow_unknown=False, allow_month=False):
    """检查单元格日期写法；正常返回 None，否则返回问题描述。

    allow_month：是否允许「YYYY-MM（日未披露）」。日级列（首次/最后发现、失效、核实日期）必须为 False。
    """
    if isinstance(v, (dt.datetime, dt.date)):
        return None
    s = str(v or "").strip()
    if not s:
        return None
    if s == "未知":
        return None if allow_unknown else "不允许「未知」"
    if DATE_FULL_RE.fullmatch(s):
        try:
            dt.date.fromisoformat(s)
            return None
        except ValueError:
            return "日期不存在"
    if DATE_MONTH_RE.fullmatch(s):
        return None if allow_month else "不允许月精度（该列必须精确到日 YYYY-MM-DD）"
    if re.fullmatch(r"\d{4}", s):
        return ("裸年份（应写 YYYY-MM（日未披露）或 YYYY-MM-DD）" if allow_month
                else "裸年份（该列必须精确到日 YYYY-MM-DD）")
    if re.fullmatch(r"\d{4}-\d{1,2}", s):
        return ("仅有月份（应写 YYYY-MM（日未披露））" if allow_month
                else "仅有月份（该列必须精确到日 YYYY-MM-DD）")
    return ("格式不符（应为 YYYY-MM-DD / YYYY-MM（日未披露））" if allow_month
            else "格式不符（该列必须精确到日 YYYY-MM-DD）")


def r_weight(pub, scan, unknown=0.4):
    """R 时效（默认规格见 references/jd-maintenance-rules.md §三）。发布日期未知 → unknown（**默认 0.4**）。

    口径已裁定（2026-10-16，用户确认）：未知取 0.4，与「>90 天」同级，
    不把「发布日期不明」当成新鲜；存量数据按需 recompute。
    如需临时采用其他口径，传 --r-unknown。
    """
    d = parse_date(pub)
    if d is None:
        return unknown
    days = (scan - d).days
    if days <= 7:
        return 1.0
    if days <= 30:
        return 0.75
    if days <= 90:
        return 0.5
    return 0.3


def c_weight(industry, tier):
    """C = 0.6×行业分 + 0.4×梯队分（默认规格见 references/jd-maintenance-rules.md §三）。"""
    return 0.6 * float(industry) + 0.4 * float(tier)


def ci_weight(city, given=None):
    if isinstance(given, (int, float)):
        return float(given)
    s = str(city or "")
    for name, w in CONFIG["city"].items():
        if name in s:
            return w
    return CONFIG["city_default"]


def total_score(R, C, F, Ci, A):
    """综合优先级分 = 100×(wR·R + wC·C + wF·F + wCi·Ci + wA·A)。

    权重取自 CONFIG["weights"]（默认 0.30/0.28/0.22/0.10/0.10，可由 --config 覆盖）。
    """
    w = CONFIG["weights"]
    return round(100 * (w["R"] * R + w["C"] * C + w["F"] * F + w["Ci"] * Ci + w["A"] * A), 1)


def resolve_company_scores(industry, tier):
    """行业分 / 梯队分「未知」时按默认值计分；返回 (industry, tier, 用了默认值的字段)。"""
    used = []
    ind = _as_num(industry)
    if ind is None:
        ind = DEFAULT_INDUSTRY
        used.append("公司行业分")
    t = _as_num(tier)
    if t is None:
        t = DEFAULT_TIER
        used.append("公司梯队分")
    return ind, t, used


# ---------------------------------------------------------------- 值校验

def _as_num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def _int_1_5(v):
    """把 1–5 的整数（含数字字符串）转成 int；否则返回 None。"""
    if isinstance(v, bool):
        return None
    n = None
    if isinstance(v, int):
        n = v
    elif isinstance(v, float) and v.is_integer():
        n = int(v)
    elif isinstance(v, str) and re.fullmatch(r"[1-5]", v.strip()):
        n = int(v.strip())
    return n if (n is not None and 1 <= n <= 5) else None


def sources_problem(v):
    """检查 全部来源 的「平台|URL；平台|URL」格式。"""
    s = str(v or "").strip()
    if not s:
        return None
    for part in SOURCES_SPLIT_RE.split(s):
        part = part.strip()
        if not part:
            continue
        if "|" not in part:
            return "应为「平台|URL；平台|URL」"
        plat, url = part.split("|", 1)
        if not plat.strip():
            return "来源平台为空"
        if not URL_RE.fullmatch(url.strip()):
            return f"来源链接不是 http(s)：{url.strip()[:40]}"
    return None


def parse_sources(v):
    """把「平台|URL；平台|URL」解析为 [(平台, URL)]；格式异常的部分保留为 ("", 原文)。"""
    out = []
    for part in SOURCES_SPLIT_RE.split(str(v or "").strip()):
        part = part.strip()
        if not part:
            continue
        if "|" in part:
            plat, url = part.split("|", 1)
            out.append((plat.strip(), url.strip()))
        else:
            out.append(("", part))
    return out


def validate_item(it, strict=False):
    """按 FIELD_SPEC 校验一条待入库 JSON；返回错误列表（空 = 通过）。"""
    errs = []
    for name in REQUIRED_INPUTS:
        raw = it.get(name)
        s = "" if raw is None else str(raw).strip()
        if not s or (s == "未知" and name not in UNKNOWN_OK):
            errs.append(f"缺少必填 {name}")
    if strict:
        for name in STRICT_INPUTS:
            if it.get(name) in (None, "") or not str(it.get(name)).strip():
                errs.append(f"strict 缺列 {name}")

    ind = it.get("公司行业分")
    if ind not in (None, "") and str(ind).strip() != "未知":
        n = _as_num(ind)
        if n is None or n not in INDUSTRY_ALLOWED:
            errs.append(f"公司行业分 只能 0.5/0.8/1.0/未知，实际 {ind!r}")

    tier = it.get("公司梯队分")
    if tier not in (None, "") and str(tier).strip() != "未知":
        n = _as_num(tier)
        if n is None or not (0.6 <= n <= 1.0):
            errs.append(f"公司梯队分 应为 0.6–1.0 数字或「未知」，实际 {tier!r}")

    for name in ("契合度(1-5)", "当前可达性(1-5)"):
        raw = it.get(name)
        if raw not in (None, "") and _int_1_5(raw) is None:
            errs.append(f"{name} 应为 1–5 的整数，实际 {raw!r}")

    url = it.get("原文链接")
    if url not in (None, "") and not URL_RE.fullmatch(str(url).strip()):
        errs.append(f"原文链接 必须是 http(s) URL，实际 {str(url)[:60]!r}")

    for name in ("发布日期", "截止日期"):
        raw = it.get(name)
        if raw in (None, ""):
            continue
        prob = date_value_problem(raw, allow_unknown=(name == "发布日期"), allow_month=True)
        if prob:
            errs.append(f"{name} {prob}：{raw!r}")

    prob = sources_problem(it.get("全部来源"))
    if prob:
        errs.append(f"全部来源 {prob}")
    return errs


# ---------------------------------------------------------------- 去重键归一化
#
# 背景（2026-10-14）：严格键 (公司名称, 职位名称, 工作城市) 按「原样字符串」比对，
# 同一岗位因书写差异被重复入册 —— 已实证 6 例：
#   JD262 泡泡玛特      与失效 JD182「泡泡玛特(09992.HK·…)」公司名带标注 → 未认出
#   JD223 三七互娱      与失效 JD196「广州-天河」vs「广州」→ 未认出
#   JD261 中关村科金    与失效 JD181 → 未认出
#   JD259 美团AI Builder 与失效 JD117/159（城市「北京-望京」vs「北京」）
#   JD201 美团 Recruiting 与失效 JD142
#   JD248 / JD258 大参林 库内重复（「广州-荔湾」vs「广州-荔湾区」）
# 归一化只用于「查重」，**不改写表内任何原始值**。
#
# 安全边界：主体名若本身就是占位符（未披露 / 某… / 保密 / 猎头代发），
# 剥离括号会把不同公司的岗误判为同一条 → 此时不归一化，退回原串。

_PLACEHOLDER_BASES = ("未披露", "未公开", "未注明", "保密", "猎头", "某")
_PAREN_TAIL = re.compile(r"[（(][^（()）]*[)）]\s*$")


def norm_company(v):
    """公司名归一化：剥离尾部括号标注（如「泡泡玛特(09992.HK·…)」→「泡泡玛特」）。

    占位符主体不剥离（否则「未披露(猎聘)」与「未披露(智联)」会被误并为同一家）。
    """
    s = str(v or "").strip()
    base = _PAREN_TAIL.sub("", s).strip()
    if not base or len(base) < 2:
        return s
    if any(base.startswith(p) for p in _PLACEHOLDER_BASES):
        return s
    return base


def norm_city(v):
    """城市归一化：取首个分隔符前的城市主体（「北京-望京」→「北京」，「广州-荔湾区」→「广州」）。"""
    s = _PAREN_TAIL.sub("", str(v or "").strip()).strip()
    base = re.split(r"[-－—·/、,，\s]", s)[0].strip()
    return base or s


def key_strict(d):
    """严格去重键：原样字符串。"""
    return (str(d.get("公司名称", "")).strip(),
            str(d.get("职位名称", "")).strip(),
            str(d.get("工作城市", "")).strip())


def key_norm(d):
    """归一化去重键：公司名剥离标注 + 城市取主体（职位名保持原样，括号常含实质区分）。"""
    return (norm_company(d.get("公司名称")),
            str(d.get("职位名称", "")).strip(),
            norm_city(d.get("工作城市")))


# ---------------------------------------------------------------- 工作簿

LOG_SHEET = "迭代日志"
LOG_COLS = ["日期", "变更项", "统计", "原因"]
LOG_STATS_RE = re.compile(r"活跃(\d+)→(\d+)，失效(\d+)→(\d+)")


class Book:
    def __init__(self, path: Path):
        if not path.exists():
            sys.exit(f"找不到工作簿：{path}")
        self.path = path
        self.wb = openpyxl.load_workbook(path)
        self.sheet_active = CONFIG["sheets"]["active"]
        self.sheet_dead = CONFIG["sheets"]["dead"]
        for s in (self.sheet_active, self.sheet_dead):
            if s not in self.wb.sheetnames:
                sys.exit(f"工作簿缺少 Sheet：{s}")
        self.active = self.wb[self.sheet_active]
        self.dead = self.wb[self.sheet_dead]

    @staticmethod
    def _header_index(ws):
        return {ws.cell(1, c).value: c
                for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}

    def hdr_a(self):
        return self._header_index(self.active)

    def hdr_d(self):
        return self._header_index(self.dead)

    @staticmethod
    def rows(ws):
        """返回 [(行号, {表头: 值})]，跳过完全空行。"""
        hdr = Book._header_index(ws)
        out = []
        for r in range(2, ws.max_row + 1):
            vals = {h: ws.cell(r, c).value for h, c in hdr.items()}
            if any(v not in (None, "") for v in vals.values()):
                out.append((r, vals))
        return out

    @staticmethod
    def ensure_column(ws, name):
        """按表头名确保列存在；缺失时追加到末尾并复制表头样式。返回列号。"""
        hdr = Book._header_index(ws)
        if name in hdr:
            return hdr[name]
        col = ws.max_column + 1
        ws.cell(1, col).value = name
        src = ws.cell(1, 1)
        dst = ws.cell(1, col)
        dst.font = _copy.copy(src.font)
        dst.fill = _copy.copy(src.fill)
        dst.border = _copy.copy(src.border)
        dst.alignment = _copy.copy(src.alignment)
        dst.protection = _copy.copy(src.protection)
        dst.number_format = src.number_format
        print(f"  [结构] {ws.title} 已补列「{name}」（第 {col} 列）")
        return col

    def backup(self, why=""):
        # 备份目录随工作簿走（放在同一目录下的「备份/」），保证测试不会污染真实备份
        bdir = self.path.parent / "备份"
        bdir.mkdir(parents=True, exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d-%H%M")
        dst = bdir / f"{self.path.stem}.bak-{stamp}.xlsx"
        n = 1
        while dst.exists():
            dst = bdir / f"{self.path.stem}.bak-{stamp}-{n}.xlsx"
            n += 1
        shutil.copy2(self.path, dst)
        print(f"  [备份] {dst}")
        return dst

    def save(self, do_backup=True):
        """保存工作簿；不再自动写迭代日志（日志由 log 子命令按运行写入）。"""
        if do_backup:
            self.backup()
        self.wb.save(self.path)
        print(f"  [保存] {self.path}")


def ensure_log_sheet(bk: Book):
    """确保「迭代日志」sheet 存在且为 4 列；失败则中止（返回 (ws, hdr)）。"""
    wb = bk.wb
    if LOG_SHEET not in wb.sheetnames:
        try:
            ws = wb.create_sheet(LOG_SHEET)
            for i, name in enumerate(LOG_COLS, start=1):
                ws.cell(1, i).value = name
            for i in range(1, len(LOG_COLS) + 1):
                src = bk.active.cell(1, 1)
                dst = ws.cell(1, i)
                dst.font = _copy.copy(src.font)
                dst.fill = _copy.copy(src.fill)
                dst.border = _copy.copy(src.border)
                dst.alignment = _copy.copy(src.alignment)
                dst.protection = _copy.copy(src.protection)
            print(f"  [结构] 已创建「{LOG_SHEET}」sheet（{len(LOG_COLS)} 列）")
        except Exception as e:  # pragma: no cover
            sys.exit(f"创建「{LOG_SHEET}」失败，已中止：{e}")
    ws = wb[LOG_SHEET]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)
               if ws.cell(1, c).value not in (None, "")]
    if headers != LOG_COLS:
        has_data = any(
            any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1))
            for r in range(2, ws.max_row + 1)
        )
        if has_data:
            sys.exit(f"「{LOG_SHEET}」表头不是 {LOG_COLS}，且已有数据；请人工处理后再写日志")
        try:
            for c in range(1, ws.max_column + 1):
                ws.cell(1, c).value = None
            for i, name in enumerate(LOG_COLS, start=1):
                ws.cell(1, i).value = name
            print(f"  [结构] 「{LOG_SHEET}」表头已重置为 {len(LOG_COLS)} 列")
        except Exception as e:  # pragma: no cover
            sys.exit(f"修正「{LOG_SHEET}」表头失败，已中止：{e}")
    hdr = Book._header_index(ws)
    missing = [h for h in LOG_COLS if h not in hdr]
    if missing:  # pragma: no cover
        sys.exit(f"「{LOG_SHEET}」缺少表头 {missing}，已中止")
    return ws, hdr


def parse_log_stats(v):
    """解析「活跃a→b，失效c→d」，返回 (a,b,c,d) 或 None。"""
    m = LOG_STATS_RE.search(str(v or "").replace(" ", ""))
    return tuple(int(x) for x in m.groups()) if m else None


# ---------------------------------------------------------------- describe

def cmd_describe(args):
    if args.format == "json":
        print(json.dumps(FIELD_SPEC, ensure_ascii=False, indent=2))
        return 0

    def cell(text):
        return str(text).replace("|", "\\|").replace("\n", " ")

    print("| 字段 | 含义 | 分组 | 必填 | 类型/取值 | 写入规则 | 示例 |")
    print("| --- | --- | --- | --- | --- | --- | --- |")
    for f in FIELD_SPEC:
        print(f"| {cell(f['name'])} | {cell(f['meaning'])} | {cell(f['group'])} | "
              f"{cell(f['required'])} | {cell(f['type'])} | {cell(f['write_rule'])} | "
              f"{cell(f['example'])} |")
    return 0


# ---------------------------------------------------------------- audit

def cmd_audit(bk: Book, args):
    problems = 0
    print(f"== 体检：{_rel(bk.path)} ==")

    # 1) 结构
    print("\n[1] 结构")
    for ws in bk.wb.worksheets:
        named = sum(1 for c in range(1, ws.max_column + 1) if ws.cell(1, c).value)
        extra = ws.max_column - named
        print(f"  {ws.title}: {ws.max_row - 1} 数据行 × {ws.max_column} 列"
              f"（有表头 {named}，无表头尾列 {extra}）")

    ha, hd = bk.hdr_a(), bk.hdr_d()
    need_cols = ["去重ID", "最后发现日期", "综合优先级分"] + [
        h for h in REQUIRED_INPUTS if h not in ("最后发现日期", "综合优先级分")]
    missing = [h for h in need_cols if h not in ha]
    if missing:
        print(f"  ✗ 活跃表缺少关键列：{missing}")
        problems += 1

    # 2) 排序
    print("\n[2] 排序（应按 综合优先级分 降序）")
    arows = Book.rows(bk.active)
    seq = [(r, v.get("综合优先级分")) for r, v in arows]
    inv = [(seq[i], seq[i + 1]) for i in range(len(seq) - 1)
           if isinstance(seq[i][1], (int, float)) and isinstance(seq[i + 1][1], (int, float))
           and seq[i][1] < seq[i + 1][1]]
    print(f"  逆序点 {len(inv)} 处")
    if inv:
        problems += 1
        print(f"  首 3 处：{[(a[0], a[1], b[1]) for a, b in inv[:3]]}")
        print("  → 修：python3 jd_tool.py sort --apply")

    # 3) 缺失值
    print("\n[3] 必填字段缺失" + ("（含 --strict 分析三列）" if args.strict else ""))
    miss = []
    check_fields = REQUIRED_INPUTS + ["综合优先级分", "最后发现日期", "首次发现日期"]
    if args.strict:
        check_fields += STRICT_INPUTS
    for r, v in arows:
        for k in check_fields:
            if v.get(k) in (None, ""):
                miss.append(f"{v.get('去重ID')}·{k}")
    print(f"  缺失 {len(miss)} 项" + (f"：{miss[:8]}" if miss else ""))
    if miss:
        problems += 1

    # 3a) 「未知」白名单：只允许 发布日期 / 公司行业分 / 公司梯队分
    unknown_bad = []
    for r, v in arows:
        for k in check_fields:
            if str(v.get(k) or "").strip() == "未知" and k not in UNKNOWN_OK:
                unknown_bad.append(f"{v.get('去重ID')}·{k}")
    print(f"  「未知」越界 {len(unknown_bad)} 项"
          + (f"（仅允许 {'/'.join(sorted(UNKNOWN_OK))}）：{unknown_bad[:8]}" if unknown_bad else ""))
    if unknown_bad:
        problems += 1

    # 3b) 日期格式
    print("\n[3b] 日期格式")
    date_fields_a = [("发布日期", True, True), ("截止日期", False, True),
                     ("首次发现日期", False, False), ("最后发现日期", False, False)]
    date_fields_d = date_fields_a + [("失效日期", False, False)]
    bad_dates = []
    for label, rows, fields in (("活跃", arows, date_fields_a),
                                ("失效", Book.rows(bk.dead), date_fields_d)):
        for r, v in rows:
            for f, allow_unknown, allow_month in fields:
                raw = v.get(f)
                if raw in (None, ""):
                    continue
                prob = date_value_problem(raw, allow_unknown=allow_unknown, allow_month=allow_month)
                if prob:
                    bad_dates.append(f"{label} {v.get('去重ID')}·{f}={raw!r}（{prob}）")
    print(f"  问题 {len(bad_dates)} 处" + (f"：{bad_dates[:6]}" if bad_dates else ""))
    if bad_dates:
        problems += 1

    # 3c) 取值范围
    print("\n[3c] 取值范围（1–5 / 行业分 / 梯队分 / URL）")
    bad_vals = []
    for label, rows in (("活跃", arows), ("失效", Book.rows(bk.dead))):
        for r, v in rows:
            jid = v.get("去重ID")
            for f in ("契合度(1-5)", "当前可达性(1-5)"):
                raw = v.get(f)
                if raw not in (None, "") and _int_1_5(raw) is None:
                    bad_vals.append(f"{label} {jid}·{f}={raw!r}（应为 1–5 整数）")
            ind = v.get("公司行业分")
            if ind not in (None, "") and str(ind).strip() != "未知":
                n = _as_num(ind)
                if n is None or n not in INDUSTRY_ALLOWED:
                    bad_vals.append(f"{label} {jid}·公司行业分={ind!r}（应为 0.5/0.8/1.0/未知）")
            tier = v.get("公司梯队分")
            if tier not in (None, "") and str(tier).strip() != "未知":
                n = _as_num(tier)
                if n is None or not (0.6 <= n <= 1.0):
                    bad_vals.append(f"{label} {jid}·公司梯队分={tier!r}（应为 0.6–1.0/未知）")
            url = v.get("原文链接")
            if url not in (None, "") and not URL_RE.fullmatch(str(url).strip()):
                bad_vals.append(f"{label} {jid}·原文链接={str(url)[:50]!r}（应为 http(s)）")
    print(f"  问题 {len(bad_vals)} 处" + (f"：{bad_vals[:6]}" if bad_vals else ""))
    if bad_vals:
        problems += 1

    # 4) 权重自检（用存储的 R/C/F/Ci/可达性 重算综合分）
    print("\n[4] 权重公式自检（重算 vs 存储）")
    bad = []
    for r, v in arows:
        try:
            ci_server = v.get("城市权重Ci")
            A = float(v["当前可达性(1-5)"]) / 5.0
            expect = total_score(float(v["时效权重R"]), float(v["公司权重C"]),
                                 float(v["契合权重F"]), float(ci_server), A)
            got = float(v["综合优先级分"])
            if abs(expect - got) > 0.15:
                bad.append((v.get("去重ID"), got, expect))
        except (TypeError, ValueError, KeyError):
            continue
    print(f"  与公式不符 {len(bad)} 条" + (f"：{bad[:6]}" if bad else ""))
    if bad:
        problems += 1

    # 5) 跨表重复
    print("\n[5] 活跃 vs 失效 同(公司,职位)（严格原样比对）")
    key = lambda v: (str(v.get("公司名称", "")).strip(), str(v.get("职位名称", "")).strip())
    dead_map = {}
    for r, v in Book.rows(bk.dead):
        dead_map.setdefault(key(v), []).append(v.get("去重ID"))
    ov = [(v.get("去重ID"), v.get("公司名称"), v.get("职位名称"), dead_map[key(v)])
          for r, v in arows if key(v) in dead_map]
    print(f"  {len(ov)} 例" + ("" if not ov else "（岗位重新发布是事实，需人工确认）"))
    for x in ov[:6]:
        print(f"    {x[0]} {x[1]} · {x[2]}  <-> 失效 {x[3]}")
    if ov:
        print("    → 修：重新发布请走 --allow-repost + --repost-evidence；否则人工确认后清理")
        problems += 1

    # 5b) 跨表重复（归一化：公司名剥离标注 + 城市取主体）
    dead_n = {}
    for r, v in Book.rows(bk.dead):
        dead_n.setdefault(key_norm(v), []).append(v.get("去重ID"))
    ov_n = [(v.get("去重ID"), key_norm(v), dead_n[key_norm(v)])
            for r, v in arows
            if key_norm(v) in dead_n and key(v) not in dead_map]
    print(f"\n[5b] 同上，但按「归一化键」（公司名剥离标注 / 城市取主体）：{len(ov_n)} 例")
    for x in ov_n:
        print(f"    {x[0]}  [{x[1][0]} · {x[1][1]} · {x[1][2]}]  <-> 失效 {x[2]}")
    if ov_n:
        print("    → 均为「书写差异导致的重复入册」；add 已能自动识别该形态")
        problems += 1

    # 5c) 活跃表内重复（归一化）
    act_n = {}
    for r, v in arows:
        act_n.setdefault(key_norm(v), []).append(v.get("去重ID"))
    dup_n = {k: v for k, v in act_n.items() if len(v) > 1}
    print(f"\n[5c] 活跃表内重复（归一化键）：{len(dup_n)} 组")
    for k, v in dup_n.items():
        print(f"    {v}  [{k[0]} · {k[1]} · {k[2]}]")
    if dup_n:
        print("    → 同一岗位在活跃表被登记两次，建议人工确认后合并（本工具不自动删）")
        problems += 1

    # 6) 超期未清
    today = args.today
    stale_days = CONFIG["stale_days"]
    stale = [(v.get("去重ID"), v.get("公司名称"), v.get("职位名称"), str(parse_date(v.get("最后发现日期"))))
             for r, v in arows
             if parse_date(v.get("最后发现日期")) and (today - parse_date(v.get("最后发现日期"))).days > stale_days]
    print(f"\n[6] 应移入失效（最后发现 > {stale_days} 天，基准 {today}）：{len(stale)} 条")
    for x in stale[:10]:
        print(f"    {x}")
    if stale:
        print("  → 修：python3 jd_tool.py purge --apply")
        problems += 1

    # 7) 截止日期已过但仍活跃（规则冲突，只报不自动处理）
    conflict = []
    for r, v in arows:
        d = parse_date(v.get("截止日期"))
        if d and d < today:
            conflict.append((v.get("去重ID"), v.get("公司名称"), v.get("职位名称"), str(d)))
    print(f"\n[7] 截止日期已过但仍活跃：{len(conflict)} 条（`jd-maintenance-rules.md` §五 判据之一；截止日期多为估算，不自动清理）")
    for x in conflict:
        print(f"    {x}")
    if conflict:
        print("  判定：以「仍能被检索到」为准，暂不清；截止日期字段多为估算值，不可直接采信")

    # 8) 脏数据
    print("\n[8] 脏数据")
    dirty = 0
    for ws in bk.wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    print(f"    ✗ {ws.title}!{c.coordinate} 以「=」开头 → Excel 会当公式→#NAME?")
                    print(f"      值：{c.value[:80]}")
                    dirty += 1
    for ws, col in ((bk.active, 1), (bk.dead, 1)):
        for r, v in Book.rows(ws):
            s = str(v.get("去重ID") or "")
            if s and not re.fullmatch(r"JD\d+", s.strip()):
                print(f"    ✗ {ws.title} 第{r}行 去重ID 非标准：{s!r}")
                dirty += 1
                break  # 每表只报一行，避免刷屏
    if not dirty:
        print("    无")
    else:
        problems += 1

    # 9) 迭代日志
    print("\n[9] 迭代日志")
    active_now = len(arows)
    dead_now = len(Book.rows(bk.dead))
    if LOG_SHEET not in bk.wb.sheetnames:
        print(f"  ✗ 缺少「{LOG_SHEET}」sheet")
        problems += 1
    else:
        ws = bk.wb[LOG_SHEET]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)
                   if ws.cell(1, c).value not in (None, "")]
        if headers != LOG_COLS:
            print(f"  ✗ 表头应为 {LOG_COLS}，实际 {headers}")
            problems += 1
        log_rows = Book.rows(ws)
        if not log_rows:
            if active_now or dead_now:
                print("  ⚠ 有数据但无迭代日志行；请运行 log 补写本次运行")
        else:
            last = log_rows[-1][1]
            log_bad = []
            prob = date_value_problem(last.get("日期"))
            if prob:
                log_bad.append(f"日期 {last.get('日期')!r}（{prob}）")
            for f in ("变更项", "统计", "原因"):
                if str(last.get(f) or "").strip() == "":
                    log_bad.append(f"{f} 为空")
            stats = parse_log_stats(last.get("统计"))
            if stats is None:
                log_bad.append(f"统计无法解析：{last.get('统计')!r}")
            elif (stats[1], stats[3]) != (active_now, dead_now):
                log_bad.append(f"统计末值 活跃{stats[1]}/失效{stats[3]} ≠ 当前 活跃{active_now}/失效{dead_now}")
            if log_bad:
                print("  ✗ " + "；".join(log_bad))
                problems += 1
            else:
                print(f"  最后一行 OK：{last.get('日期')} | {last.get('统计')}")

    print(f"\n== 汇总：{'全部通过 ✓' if problems == 0 else f'{problems} 类问题需处理'} ==")
    return 1 if (args.strict and problems) else 0


# ---------------------------------------------------------------- add

def _append_note(row, text):
    base = str(row.get("备注/质量说明") or "").strip()
    row["备注/质量说明"] = f"{base} {text}".strip()


def load_repost_evidence(path):
    """载入并校验 --repost-evidence JSON；不合规直接中止。"""
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as e:
        sys.exit(f"重新发布证据读取失败：{e}")
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list) or not data:
        sys.exit("--repost-evidence 必须是非空 JSON 数组")
    need = ["原jd_id", "公司名称", "职位名称", "工作城市", "证据链接", "核实日期", "状态"]
    out = []
    for i, e in enumerate(data, 1):
        if not isinstance(e, dict):
            sys.exit(f"重新发布证据第 {i} 条不是对象")
        miss = [k for k in need if str(e.get(k) or "").strip() == ""]
        if miss:
            sys.exit(f"重新发布证据第 {i} 条缺字段 {miss}")
        if str(e["状态"]).strip() not in ("在招", "可投"):
            sys.exit(f"重新发布证据第 {i} 条状态必须为「在招」或「可投」，实际 {e['状态']!r}")
        if not re.fullmatch(r"JD\d+", str(e["原jd_id"]).strip()):
            sys.exit(f"重新发布证据第 {i} 条原jd_id 必须是 JD+数字，实际 {e['原jd_id']!r}")
        if not URL_RE.fullmatch(str(e["证据链接"]).strip()):
            sys.exit(f"重新发布证据第 {i} 条证据链接必须是 http(s) URL")
        prob = date_value_problem(e["核实日期"], allow_month=False)
        if prob:
            sys.exit(f"重新发布证据第 {i} 条核实日期{prob}：{e['核实日期']!r}")
        out.append(e)
    return out


def match_repost(entries, dead_rows, dead_ids):
    """按 原jd_id 精确匹配失效行（不接受三元组兜底）。

    返回 (证据, 实际失效ID, 错误列表)：
      - 原jd_id 必须存在于失效表；
      - 且必须是当前岗位命中的失效行之一；
      - 证据中的公司/职位/城市必须与该失效行一致（归一化比对）。
    备注只引用实际匹配到的失效 ID，绝不引用输入里可能伪造的 ID。
    """
    errors = []
    for e in entries or []:
        raw_id = str(e.get("原jd_id") or "").strip()
        if raw_id not in dead_rows:
            errors.append(f"原jd_id {raw_id} 不存在于失效表")
            continue
        if raw_id not in dead_ids:
            errors.append(f"原jd_id {raw_id} 与当前岗位不匹配（命中失效行 {dead_ids}）")
            continue
        dead = dead_rows[raw_id]
        if (norm_company(e.get("公司名称")) != norm_company(dead.get("公司名称"))
                or str(e.get("职位名称", "")).strip() != str(dead.get("职位名称", "")).strip()
                or norm_city(e.get("工作城市")) != norm_city(dead.get("工作城市"))):
            errors.append(f"原jd_id {raw_id} 的公司/职位/城市与失效表记录不一致")
            continue
        return e, raw_id, []
    return None, None, errors


def merge_source_into_row(bk: Book, r, it):
    """把来源合并进活跃表第 r 行的「全部来源」；返回 (是否变更, 原因)。"""
    if "全部来源" not in bk.hdr_a():
        Book.ensure_column(bk.active, "全部来源")
    hdr = bk.hdr_a()
    ws = bk.active
    cell = ws.cell(r, hdr["全部来源"])
    cur = "" if cell.value is None else str(cell.value).strip()

    platform = str(it.get("来源平台") or "").strip()
    url = str(it.get("原文链接") or "").strip()
    if not cur:
        # 用库内已有来源初始化；库内缺失时退回本次来源
        exist_plat = str(ws.cell(r, hdr["来源平台"]).value or "").strip() if "来源平台" in hdr else ""
        exist_url = str(ws.cell(r, hdr["原文链接"]).value or "").strip() if "原文链接" in hdr else ""
        init_plat, init_url = exist_plat or platform, exist_url or url
        if not init_plat or not init_url:
            return False, "无来源信息"
        cell.value = f"{init_plat}|{init_url}"
        if init_url == url:
            return True, "init"
    if not url:
        return False, "无来源链接"
    if any(u == url for _, u in parse_sources(cell.value)):
        return False, "已有该来源"
    cell.value = f"{str(cell.value).strip()}；{platform}|{url}"
    return True, "append"


def cmd_add(bk: Book, args):
    src = sys.stdin.read() if args.file == "-" else Path(args.file).read_text(encoding="utf-8")
    try:
        items = json.loads(src)
    except json.JSONDecodeError as e:
        sys.exit(f"JSON 解析失败：{e}")
    if isinstance(items, dict):
        items = [items]
    if not isinstance(items, list):
        sys.exit("JSON 输入必须是对象数组（或单个对象）")

    bad_generated = sorted({k for it in items if isinstance(it, dict)
                            for k in it if k in GENERATED_FIELDS})
    if bad_generated:
        sys.exit(f"输入含系统生成字段 {bad_generated}；请从 JSON 中删除，"
                 f"这些列由 add/recompute 生成")

    if args.allow_repost and not args.repost_evidence:
        sys.exit("--allow-repost 必须同时提供 --repost-evidence <json>（只给开关不给证据 = 无法核实）")
    evidence = load_repost_evidence(args.repost_evidence) if args.repost_evidence else []

    today = args.today
    ha = bk.hdr_a()
    if "全部来源" not in ha:
        Book.ensure_column(bk.active, "全部来源")
        ha = bk.hdr_a()

    missing_cols = [h for h in REQUIRED_INPUTS if h not in ha]
    if missing_cols:
        sys.exit(f"活跃表缺少必填列 {missing_cols}；add 已中止，未写入任何行")

    row_by_strict, row_by_norm = {}, {}
    for r, v in Book.rows(bk.active):
        row_by_strict.setdefault(key_strict(v), r)
        row_by_norm.setdefault(key_norm(v), r)
    dead_by_strict, dead_by_norm = {}, {}
    dead_rows = {}
    for r, v in Book.rows(bk.dead):
        jid = str(v.get("去重ID") or "").strip()
        if jid:
            dead_rows.setdefault(jid, v)
        dead_by_strict.setdefault(key_strict(v), []).append(jid)
        dead_by_norm.setdefault(key_norm(v), []).append(jid)

    # 下一个去重 ID
    mx = 0
    for ws in (bk.active, bk.dead):
        for _, v in Book.rows(ws):
            m = re.fullmatch(r"JD0*(\d+)", str(v.get("去重ID") or "").strip())
            if m:
                mx = max(mx, int(m.group(1)))
    nxt = mx + 1

    added, merged, skipped = [], [], 0
    provenance_errors = 0
    for it in items:
        if not isinstance(it, dict):
            print("  SKIP 输入项不是对象")
            skipped += 1
            continue
        who = f"{it.get('公司名称', '?')} · {it.get('职位名称', '?')}"
        errs = validate_item(it, strict=args.strict)
        if errs:
            print(f"  SKIP {who}：{'；'.join(errs)}")
            skipped += 1
            continue
        if not args.strict:
            lack = [c for c in STRICT_INPUTS
                    if it.get(c) in (None, "") or not str(it.get(c)).strip()]
            if lack:
                print(f"  WARN 缺 strict 列 {lack}：{who}")

        k, kn = key_strict(it), key_norm(it)
        dup_row = row_by_strict.get(k) or row_by_norm.get(kn)
        if dup_row:
            jid = bk.active.cell(dup_row, ha["去重ID"]).value
            if not args.merge_sources:
                print(f"  SKIP 活跃表已有：{jid} · {k}（--no-merge-sources）")
                skipped += 1
                continue
            changed, why = merge_source_into_row(bk, dup_row, it)
            if changed:
                print(f"  合并来源 {jid}：{why} → {it.get('来源平台')}|{it.get('原文链接')}")
                merged.append((jid, it.get("公司名称"), it.get("职位名称"), it.get("来源平台")))
            else:
                print(f"  SKIP {jid} 无新来源（{why}）：{k}")
                skipped += 1
            continue

        dead_ids = dead_by_strict.get(k) or dead_by_norm.get(kn)
        origin = None
        if dead_ids:
            if not args.allow_repost:
                print(f"  SKIP 失效表已有 {dead_ids}"
                      f"（重新发布请加 --allow-repost + --repost-evidence）：{k}")
                skipped += 1
                continue
            match, origin, ev_errs = match_repost(evidence, dead_rows, dead_ids)
            if match is None:
                detail = "；".join(ev_errs) if ev_errs else "未提供可用证据"
                print(f"  ERROR 重新发布证据无效（{detail}）：{k}")
                provenance_errors += 1
                skipped += 1
                continue

        ind, tier, used_default = resolve_company_scores(
            it.get("公司行业分"), it.get("公司梯队分"))
        if used_default:
            print(f"  WARN {who} {'/'.join(used_default)} 为未知 → "
                  f"C 按默认值（行业 {DEFAULT_INDUSTRY} / 梯队 {DEFAULT_TIER}）计算")

        R = r_weight(it.get("发布日期"), today, unknown=args.r_unknown)
        C = c_weight(ind, tier)
        F = float(_int_1_5(it["契合度(1-5)"])) / 5.0
        Ci = ci_weight(it.get("工作城市"))
        A = float(_int_1_5(it["当前可达性(1-5)"])) / 5.0

        row = dict(it)
        row["去重ID"] = f"JD{nxt:03d}"
        row["抓取日期"] = str(today)
        row["首次发现日期"] = str(today)
        row["最后发现日期"] = str(today)
        row["时效权重R"] = R
        row["公司权重C"] = round(C, 3)
        row["契合权重F"] = round(F, 3)
        row["城市权重Ci"] = Ci
        row["综合优先级分"] = total_score(R, C, F, Ci, A)
        if "全部来源" in ha and not str(row.get("全部来源") or "").strip():
            row["全部来源"] = f"{str(it.get('来源平台')).strip()}|{str(it.get('原文链接')).strip()}"
        if origin:
            _append_note(row, f"[重新发布·原 {origin} 已失效]")

        r = bk.active.max_row + 1
        for name, val in row.items():
            if name in ha:
                bk.active.cell(r, ha[name]).value = val
        for name in ha:  # 未提供的列留空（不写 None，保持原样式）
            if name not in row:
                bk.active.cell(r, ha[name]).value = None
        nxt += 1
        row_by_strict.setdefault(k, r)
        row_by_norm.setdefault(kn, r)
        added.append((row["去重ID"], row["公司名称"], row["职位名称"],
                      row["综合优先级分"], R, origin))

    if not added and not merged:
        print(f"  无新增、无合并来源（跳过 {skipped} 条）。")
        return 1 if provenance_errors else 0

    print(f"\n  新增 {len(added)} 条，合并来源 {len(merged)} 条（跳过 {skipped} 条）：")
    for a in added:
        tail = f"  ← 原 {a[5]} 重新发布" if a[5] else ""
        print(f"    {a[0]}  {a[1]} · {a[2]}  综合={a[3]}（R={a[4]}）{tail}")
    for m in merged:
        print(f"    合并来源 {m[0]}  {m[1]} · {m[2]}  +{m[3]}")
    if provenance_errors:
        print(f"  ⚠ 有 {provenance_errors} 条重新发布证据被拒（未入册）")

    if added and any(a[4] == args.r_unknown for a in added) and args.r_unknown != 0.4:
        print(f"  ⚠ 有条目按「发布日期未知 → R={args.r_unknown}」计分")

    if added and not args.no_sort:
        n = do_sort(bk)
        print(f"  已按综合优先级分降序重排（{n} 行）")
    bk.save(do_backup=not args.no_backup)
    return 1 if provenance_errors else 0


# ---------------------------------------------------------------- refresh

REFRESH_STATUS = ("在招", "可投", "仅索引残留", "详情不可访问", "已关闭")
REFRESH_OK = ("在招", "可投")


def load_refresh_evidence(path):
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as e:
        sys.exit(f"刷新证据读取失败：{e}")
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list) or not data:
        sys.exit("--evidence 必须是非空 JSON 数组")
    need = ["jd_id", "平台", "查询词", "证据链接", "核实日期", "状态"]
    out = []
    for i, e in enumerate(data, 1):
        if not isinstance(e, dict):
            sys.exit(f"刷新证据第 {i} 条不是对象")
        miss = [k for k in need if str(e.get(k) or "").strip() == ""]
        if miss:
            sys.exit(f"刷新证据第 {i} 条缺字段 {miss}")
        status = str(e["状态"]).strip()
        if status not in REFRESH_STATUS:
            sys.exit(f"刷新证据第 {i} 条状态必须为 {list(REFRESH_STATUS)} 之一，实际 {status!r}")
        if not URL_RE.fullmatch(str(e["证据链接"]).strip()):
            sys.exit(f"刷新证据第 {i} 条证据链接必须是 http(s) URL")
        prob = date_value_problem(e["核实日期"], allow_month=False)
        if prob:
            sys.exit(f"刷新证据第 {i} 条核实日期{prob}：{e['核实日期']!r}")
        out.append(e)
    return out


def cmd_refresh(bk: Book, args):
    ha = bk.hdr_a()
    if "最后发现日期" not in ha:
        sys.exit("活跃表缺少「最后发现日期」列，无法刷新")
    rows = {}
    for r, v in Book.rows(bk.active):
        rows.setdefault(str(v.get("去重ID") or "").strip(), r)

    if args.evidence:
        entries = load_refresh_evidence(args.evidence)
        hit, nochange, missed = [], [], []
        for e in entries:
            jid = str(e["jd_id"]).strip()
            r = rows.get(jid)
            if r is None:
                missed.append(jid)
                continue
            status = str(e["状态"]).strip()
            if status in REFRESH_OK:
                bk.active.cell(r, ha["最后发现日期"]).value = str(e["核实日期"]).strip()
                hit.append(jid)
            else:
                nochange.append((jid, status))
                print(f"  KEEP {jid} 状态「{status}」→ 不更新最后发现日期"
                      f"（{e['平台']} / {e['查询词']}）")
        print(f"  证据刷新：更新 {len(hit)} 条 {hit[:12]}{' ...' if len(hit) > 12 else ''}")
        if nochange:
            print(f"  未更新 {len(nochange)} 条：{[x[0] for x in nochange][:12]}")
        if missed:
            print(f"  ⚠ 未在活跃表找到：{missed}")
        if not hit:
            print("  无改动。")
            return 0
    else:
        ids = []
        if args.ids:
            ids += [x.strip() for x in args.ids.split(",") if x.strip()]
        if args.ids_file:
            ids += re.findall(r"JD\d+", Path(args.ids_file).read_text(encoding="utf-8"))
        if not ids:
            sys.exit("请用 --evidence <json>（首选），或 --ids JD251,JD256 --confirm-visible --reason \"...\"")
        if not args.confirm_visible:
            sys.exit("手工刷新必须加 --confirm-visible（已再次核实详情页仍在招/可投）")
        if not str(args.reason or "").strip():
            sys.exit("手工刷新必须提供 --reason \"...\"（说明核实依据）")

        date = args.date or args.today
        want = set(ids)
        hit = []
        for jid in want:
            r = rows.get(jid)
            if r is None:
                continue
            bk.active.cell(r, ha["最后发现日期"]).value = str(date)
            hit.append(jid)
        missed = sorted(want - set(hit))
        print(f"  手工刷新「最后发现日期」→ {date}：{len(hit)} 条 "
              f"{sorted(hit)[:12]}{' ...' if len(hit) > 12 else ''}")
        print(f"  理由：{args.reason.strip()}")
        if missed:
            print(f"  ⚠ 未在活跃表找到：{missed}")
        if not hit:
            print("  无改动。")
            return 0

    bk.save(do_backup=not args.no_backup)
    return 0


# ---------------------------------------------------------------- purge

def cmd_purge(bk: Book, args):
    today = args.date or args.today
    stale_days = CONFIG["stale_days"]

    targets = []
    for r, v in Book.rows(bk.active):
        lf = parse_date(v.get("最后发现日期"))
        if lf and (today - lf).days > stale_days:
            targets.append((r, v, f"最后发现 {lf}，距今 >{stale_days} 天未再发现（{today} 扫描）"))

    print(f"== purge（基准 {today}，判据：最后发现 > {stale_days} 天）==")
    print(f"  命中 {len(targets)} 条")
    for r, v, why in targets:
        print(f"    {v.get('去重ID')}  {v.get('公司名称')} · {v.get('职位名称')}  [{why[:24]}…]")
    if not targets:
        print("  无需清理。")
        return 0
    if not args.apply:
        print("\n  [dry-run] 未写入。确认后加 --apply 执行。")
        return 0

    ha = bk.hdr_a()
    hd = bk.hdr_d()
    # 失效表列一律按表头名确保存在（不再硬编码 col33/col34）
    for name in ha:
        if name not in hd:
            Book.ensure_column(bk.dead, name)
    for name in ("全部来源", "失效日期", "失效原因"):
        if name not in hd:
            Book.ensure_column(bk.dead, name)
    hd = bk.hdr_d()

    moved = 0
    for r, v, why in sorted(targets, key=lambda x: -x[0]):   # 自下而上删，行号不失效
        dr = bk.dead.max_row + 1
        for name, c in ha.items():                          # 按表头名搬运，不按列号
            if name in hd:
                bk.dead.cell(dr, hd[name]).value = bk.active.cell(r, c).value
        # 活跃表无表头的尾列（若有游离注记）-> 并入备注，避免丢失
        note = bk.active.cell(r, bk.active.max_column).value
        if (bk.active.max_column > max(ha.values()) and note not in (None, "")
                and "备注/质量说明" in hd):
            base = bk.dead.cell(dr, hd["备注/质量说明"]).value or ""
            bk.dead.cell(dr, hd["备注/质量说明"]).value = f"{base}\n[原活跃表尾注] {note}".strip()
        bk.dead.cell(dr, hd["失效日期"]).value = str(today)
        bk.dead.cell(dr, hd["失效原因"]).value = why
        bk.active.delete_rows(r)
        moved += 1

    print(f"\n  已移入失效 {moved} 条")
    do_sort(bk)
    bk.save(do_backup=not args.no_backup)
    return 0


# ---------------------------------------------------------------- sort

def do_sort(bk: Book):
    """按 综合优先级分 降序重排活跃表（同分保持原相对顺序）。非数值排到最后。"""
    ws = bk.active
    hdr = Book._header_index(ws)
    sc = hdr["综合优先级分"]
    data = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(v not in (None, "") for v in vals):
            data.append(vals)
    def keyf(vals):
        s = vals[sc - 1]
        return (1, 0) if not isinstance(s, (int, float)) else (0, -float(s))
    data.sort(key=keyf)
    for i, vals in enumerate(data):
        for c, val in enumerate(vals, start=1):
            ws.cell(i + 2, c).value = val
    for r in range(len(data) + 2, ws.max_row + 1):   # 清掉多余旧行
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None
    return len(data)


def cmd_sort(bk: Book, args):
    before = [v.get("综合优先级分") for _, v in Book.rows(bk.active)]
    n = do_sort(bk)
    after = [v.get("综合优先级分") for _, v in Book.rows(bk.active)]
    inv_b = sum(1 for i in range(len(before) - 1)
                if isinstance(before[i], (int, float)) and isinstance(before[i + 1], (int, float))
                and before[i] < before[i + 1])
    print(f"  重排 {n} 行：逆序点 {inv_b} → 0")
    print(f"  前 5 分：{before[:5]} → {after[:5]}")
    if not args.apply:
        print("\n  [dry-run] 未写入。确认后加 --apply 执行。")
        return 0
    bk.save(do_backup=not args.no_backup)
    return 0


# ---------------------------------------------------------------- repair

def cmd_repair(bk: Book, args):
    """修三类脏数据（默认 dry-run，--apply 才落盘）：

    ① 以「=」开头的文本单元格 —— Excel 会当公式解析，渲染成 #NAME? 而非文字；
    ② 去重ID 含 BOM / 首尾空白 —— 会让查重与 refresh --ids 静默找不到该条；
    ③ 失效表里混入的「非 JD 说明行」（ID 非 JD#### 且无公司/职位）——
       正文迁入工作簿同目录的旁车文件 `<工作簿名>-说明留档.md` 留档，再删掉该行，让数据从第 2 行开始。

    有正文却 ID 不规范的行**只报不改**（无法安全推断正确 ID）。
    """
    f_fix, id_fix, note_rows, bad_id = [], [], [], []

    # ① 「=」开头的文本
    for ws in bk.wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    f_fix.append((ws.title, c.coordinate, v, v.lstrip("= ")))

    # ②③ 去重ID 检查
    for ws in (bk.active, bk.dead):
        for r, v in Book.rows(ws):
            raw = v.get("去重ID")
            s = "" if raw is None else str(raw)
            s2 = s.replace("\ufeff", "").strip()
            if re.fullmatch(r"JD\d+", s2):
                if s2 != s:
                    id_fix.append((ws.title, r, s, s2))
            elif any(v.get(k) not in (None, "") for k in ("公司名称", "职位名称")):
                bad_id.append((ws.title, r, s))
            elif s2:
                note_rows.append((ws.title, r, s2))

    print(f"== repair：{_rel(bk.path)} ==")
    print(f"\n[①] 「=」开头被当公式的单元格：{len(f_fix)} 处")
    for t, co, old, new in f_fix:
        print(f"    {t}!{co}")
        print(f"      旧：{old[:100]}")
        print(f"      新：{new[:100]}")

    print(f"\n[②] 去重ID 含 BOM/空白：{len(id_fix)} 处")
    for t, r, old, new in id_fix:
        print(f"    {t} 第{r}行：{old!r} → {new!r}")

    print(f"\n[③] 混入的非 JD 说明行：{len(note_rows)} 行")
    for t, r, txt in note_rows:
        print(f"    {t} 第{r}行（将迁入旁车「说明留档.md」后删行）")
        print(f"      内容：{txt[:120]}")

    if bad_id:
        print(f"\n[!] ID 不规范但有正文（只报不改，需人工核）：{len(bad_id)} 行")
        for t, r, s in bad_id:
            print(f"    {t} 第{r}行：{s!r}")

    total = len(f_fix) + len(id_fix) + len(note_rows)
    if total == 0:
        print("\n  无需修复。")
        return 0
    if not args.apply:
        print(f"\n  [dry-run] 共 {total} 处待修，未写入。确认后加 --apply 执行。")
        return 0

    # ---- 落盘 ----
    for t, co, old, new in f_fix:
        bk.wb[t][co].value = new

    for t, r, old, new in id_fix:
        ws = bk.wb[t]
        hdr = Book._header_index(ws)
        ws.cell(r, hdr["去重ID"]).value = new

    if note_rows:
        sidecar = bk.path.parent / f"{bk.path.stem}-说明留档.md"
        with sidecar.open("a", encoding="utf-8") as f:
            for t, r, txt in note_rows:
                f.write(f"【原{t}第{r}行】{txt}\n")
        for t, r, txt in sorted(note_rows, key=lambda x: -x[1]):
            bk.wb[t].delete_rows(r)
        print(f"\n  [③] {len(note_rows)} 行已迁入旁车文件 {sidecar.name} 并删除原行")

    bk.save(do_backup=not args.no_backup)
    return 0


# ---------------------------------------------------------------- set（补历史字段）

# 这些列由计分流程派生，不允许用 set 直接改写（必须走 add / recompute）
SET_BANNED = {"综合优先级分", "时效权重R", "公司权重C", "契合权重F",
              "城市权重Ci", "去重ID", "抓取日期", "首次发现日期", "最后发现日期"}


def cmd_set(bk: Book, args):
    """按 ID 批量**补写空值单元格**（默认 dry-run）。

    用途：修复「字段历史缺失」（如某批记录的「来源平台」漏填）。
    安全约束：
      - 只补**空值**，已有值一律不覆盖（需要改已有值请走专用子命令或人工裁定）；
      - 不允许写评分派生列（SET_BANNED），避免绕过计分流程；
      - 写入前自动备份。
    JSON 形如：[{"id":"JD259","field":"来源平台","value":"猎聘 Liepin(详情页 job/1985262481)"}]
    """
    items = json.load(sys.stdin) if args.file == "-" else json.load(open(args.file, encoding="utf-8"))
    ws = {"active": bk.active, "dead": bk.dead}[args.sheet]
    hdr = Book._header_index(ws)
    rowmap = {}
    for r, v in Book.rows(ws):
        rowmap.setdefault(str(v.get("去重ID") or "").strip(), r)

    want = {}
    for it in items:
        want.setdefault(str(it.get("id", "")).strip(), []).append((it.get("field"), it.get("value")))

    n = 0
    for jid, pairs in want.items():
        r = rowmap.get(jid)
        if r is None:
            print(f"  MISS {jid}（{args.sheet} 表中找不到该 ID）")
            continue
        for field, value in pairs:
            if field in SET_BANNED:
                print(f"  SKIP {jid}·{field}（评分派生列，请走 add / recompute）")
                continue
            if field not in hdr:
                print(f"  SKIP {jid}·{field}（表头无此列）")
                continue
            cur = ws.cell(r, hdr[field]).value
            if cur not in (None, ""):
                print(f"  SKIP {jid}·{field}（已有值 {cur!r}，不覆盖）")
                continue
            print(f"  SET  [{jid}] {field}: (空) → {value!r}")
            if args.apply:
                ws.cell(r, hdr[field]).value = value
            n += 1

    print(f"\n  待写入 {n} 处")
    if not n:
        return 0
    if args.apply:
        bk.save(do_backup=not args.no_backup)
    else:
        print("  [dry-run] 未写入。加 --apply 执行。")
    return 0


def _same_num(a, b):
    x, y = _as_num(a), _as_num(b)
    if x is None or y is None:
        return False
    return abs(x - y) < 1e-6


def cmd_recompute(bk: Book, args):
    """按**存量输入列**重算派生分（默认 dry-run，--apply 才落盘）。

    输入：发布日期 / 公司行业分 / 公司梯队分 / 契合度(1-5) / 当前可达性(1-5) / 工作城市，
          城市权重Ci 优先取存量值（与 add 一致），缺失时按工作城市推导。
    输出：时效权重R / 公司权重C / 契合权重F / 城市权重Ci / 综合优先级分。
    行业分 / 梯队分为「未知」时按默认值计分（行业 0.5、梯队 0.7），与 add 一致；
    发布日期未知时 R 取 --r-unknown（默认 0.4）。
    """
    today = args.today
    ha = bk.hdr_a()
    need = ["发布日期", "公司行业分", "公司梯队分", "契合度(1-5)", "当前可达性(1-5)", "工作城市",
            "时效权重R", "公司权重C", "契合权重F", "城市权重Ci", "综合优先级分"]
    missing = [h for h in need if h not in ha]
    if missing:
        sys.exit(f"活跃表缺少 recompute 所需列：{missing}")

    changes, skipped, fallback_ids = [], [], []
    for r, v in Book.rows(bk.active):
        A = _int_1_5(v.get("当前可达性(1-5)"))
        F_n = _int_1_5(v.get("契合度(1-5)"))
        if A is None or F_n is None:
            skipped.append(v.get("去重ID"))
            continue
        ind, tier, used = resolve_company_scores(v.get("公司行业分"), v.get("公司梯队分"))
        if used:
            fallback_ids.append(v.get("去重ID"))
        R = r_weight(v.get("发布日期"), today, unknown=args.r_unknown)
        C = round(c_weight(ind, tier), 3)
        F = round(F_n / 5.0, 3)
        Ci = ci_weight(v.get("工作城市"), v.get("城市权重Ci"))
        A_f = A / 5.0
        new = {"时效权重R": R, "公司权重C": C, "契合权重F": F,
               "城市权重Ci": Ci, "综合优先级分": total_score(R, C, F, Ci, A_f)}
        diffs = [(k, v.get(k), val) for k, val in new.items() if not _same_num(v.get(k), val)]
        if diffs:
            changes.append((r, v, new, diffs))

    print(f"== recompute（基准 {today}，R 未知={args.r_unknown}）==")
    tail = f"，输入列不可解析 SKIP {len(skipped)} 条：{skipped[:6]}" if skipped else ""
    print(f"  待更新 {len(changes)} 条{tail}")
    if fallback_ids:
        print(f"  WARN 行业分/梯队分为未知，按默认值计分（行业 {DEFAULT_INDUSTRY} / "
              f"梯队 {DEFAULT_TIER}）：{len(fallback_ids)} 条 {fallback_ids[:6]}")
    for r, v, new, diffs in changes[:5]:
        print(f"    {v.get('去重ID')}  {v.get('公司名称')} · {v.get('职位名称')}")
        print("      " + "；".join(f"{k} {old}→{val}" for k, old, val in diffs))
    if not changes:
        print("  无需更新。")
        return 0
    if not args.apply:
        print("\n  [dry-run] 未写入。确认后加 --apply 执行。")
        return 0
    for r, v, new, diffs in changes:
        for k, val in new.items():
            bk.active.cell(r, ha[k]).value = val
    n = do_sort(bk)
    print(f"\n  已重算 {len(changes)} 条，并按综合优先级分重排（{n} 行）")
    bk.save(do_backup=not args.no_backup)
    return 0


# ---------------------------------------------------------------- log（运行级迭代日志）

def cmd_log(bk: Book, args):
    """一次 O2b 运行写一行日志：日期 | 变更项 | 统计 | 原因（--dry-run 只预览）。"""
    change = str(args.change or "").strip()
    reason = str(args.reason or "").strip()
    if not change:
        sys.exit("--change 不能为空（写本次变更摘要，如「新增 JD270–JD272」）")
    if not reason:
        sys.exit("--reason 不能为空（写一句话原因）")

    active_now = len(Book.rows(bk.active))
    dead_now = len(Book.rows(bk.dead))

    before_a, before_d = args.before_active, args.before_dead
    if before_a is None or before_d is None:
        prev = None
        if LOG_SHEET in bk.wb.sheetnames:
            rows = Book.rows(bk.wb[LOG_SHEET])
            if rows:
                prev = rows[-1][1]
        stats = parse_log_stats(prev.get("统计")) if prev else None
        if stats:
            # stats = (活跃前, 活跃后, 失效前, 失效后)；沿用上一行的 after 值
            if before_a is None:
                before_a = stats[1]
            if before_d is None:
                before_d = stats[3]
        else:
            print("  WARN 无上一行统计可用，运行前值按当前数量计（本次统计可能低估变化）")
            if before_a is None:
                before_a = active_now
            if before_d is None:
                before_d = dead_now

    row = [str(args.today), change,
           f"活跃{before_a}→{active_now}，失效{before_d}→{dead_now}", reason]
    print("  日志行：" + " | ".join(row))

    # 先确保日志目标可用（缺表则创建、表头不符则中止），再做备份，避免留下无用备份
    ws, hdr = ensure_log_sheet(bk)
    if args.dry_run:
        print("  [dry-run] 未写入。")
        return 0

    bk.backup("迭代日志")
    r = ws.max_row + 1
    for name, val in zip(LOG_COLS, row):
        ws.cell(r, hdr[name]).value = val
    bk.wb.save(bk.path)
    print(f"  [保存] {bk.path}（迭代日志第 {r} 行）")
    return 0


# ---------------------------------------------------------------- backup

def cmd_backup(bk: Book, args):
    bk.backup("手动备份")
    return 0


# ---------------------------------------------------------------- CLI

def main(argv=None):
    global CONFIG
    p = argparse.ArgumentParser(
        description="JD 活跃库维护工具：字段规格、入库、去重合并、刷新、失效、日志。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n"
               "  jd_tool.py describe [--format markdown|json]\n"
               "  jd_tool.py --workbook <路径> audit [--strict]\n"
               "  jd_tool.py --workbook <路径> add new.json [--strict] [--no-merge-sources]\n"
               "  jd_tool.py --workbook <路径> add new.json --allow-repost --repost-evidence repost.json\n"
               "  jd_tool.py --workbook <路径> refresh --evidence refresh.json\n"
               "  jd_tool.py --workbook <路径> refresh --ids JD251,JD256 --confirm-visible --reason \"...\"\n"
               "  jd_tool.py --workbook <路径> purge [--apply]\n"
               "  jd_tool.py --workbook <路径> log --change \"新增 JD270–JD272\" --reason \"日常扫描\"\n"
               "\n"
               "写入语义：describe / audit 只读；add / refresh 即时写入；"
               "purge / sort / repair / set / recompute 默认预览，需 --apply；\n"
               "log 一次运行写一行（--dry-run 预览）；backup 为手动备份。",
    )
    p.add_argument("--workbook", type=Path, default=Path.cwd() / "JD扫描-活跃库.xlsx",
                   help="工作簿路径（默认当前目录下的 JD扫描-活跃库.xlsx）")
    p.add_argument("--config", type=Path, default=None,
                   help="配置 JSON，可覆盖 sheets/weights/city/city_default/stale_days（默认内置规格）")
    p.add_argument("--today", type=lambda s: dt.date.fromisoformat(s), default=dt.date.today(),
                   help="把「今天」（业务日）固定为某日期（测试用）")
    sub = p.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("describe", help="输出字段规格（FIELD_SPEC，唯一真源）")
    a.add_argument("--format", choices=["markdown", "json"], default="markdown")

    a = sub.add_parser("audit", help="只读体检")
    a.add_argument("--strict", action="store_true", help="有问题时返回非零退出码")

    a = sub.add_parser("add", help="追加当日新 JD（JSON 数组文件，或 - 读 stdin）")
    a.add_argument("file")
    a.add_argument("--strict", action="store_true",
                   help="分析三列（Market Signal / Portfolio Relevance / Portfolio Benchmark）"
                        "任一为空即拒收")
    a.add_argument("--allow-repost", action="store_true",
                   help="允许失效后重新发布（必须同时提供 --repost-evidence）")
    a.add_argument("--repost-evidence",
                   help="重新发布证据 JSON：原jd_id/公司名称/职位名称/工作城市/证据链接/核实日期/状态")
    a.add_argument("--merge-sources", dest="merge_sources", action="store_true", default=True,
                   help="活跃表重复时把新来源合并进「全部来源」（默认开启）")
    a.add_argument("--no-merge-sources", dest="merge_sources", action="store_false",
                   help="活跃表重复时不合并来源，直接跳过")
    a.add_argument("--r-unknown", type=float, default=0.4,
                   help="发布日期未知时的 R（默认 0.4，2026-10-16 已裁定）")
    a.add_argument("--no-sort", action="store_true")
    a.add_argument("--no-backup", action="store_true")

    a = sub.add_parser("refresh", help="刷新「最后发现日期」（首选 --evidence）")
    a.add_argument("--evidence", help="刷新证据 JSON：jd_id/平台/查询词/证据链接/核实日期/状态")
    a.add_argument("--ids", help="逗号分隔，如 JD251,JD256（需 --confirm-visible --reason）")
    a.add_argument("--ids-file", help="从文件里抓取所有 JDxxx")
    a.add_argument("--confirm-visible", action="store_true",
                   help="手工路径：确认已再次核实详情页在招/可投")
    a.add_argument("--reason", help="手工路径：核实依据说明")
    a.add_argument("--date", type=lambda s: dt.date.fromisoformat(s), default=None,
                   help="手工刷新写入的日期（默认同 --today）")
    a.add_argument("--no-backup", action="store_true")

    a = sub.add_parser("purge", help="超期移入失效（默认 dry-run）")
    a.add_argument("--apply", action="store_true")
    a.add_argument("--date", type=lambda s: dt.date.fromisoformat(s), default=None,
                   help="基准日（默认同 --today）")
    a.add_argument("--no-backup", action="store_true")

    a = sub.add_parser("sort", help="按综合优先级分降序重排（默认 dry-run）")
    a.add_argument("--apply", action="store_true")
    a.add_argument("--no-backup", action="store_true")

    a = sub.add_parser("repair", help="修脏数据：「=」开头公式 / ID 含 BOM / 非 JD 说明行（默认 dry-run）")
    a.add_argument("--apply", action="store_true")
    a.add_argument("--no-backup", action="store_true")

    a = sub.add_parser("set", help="按 ID 批量补写空值单元格（默认 dry-run）")
    a.add_argument("--file", required=True, help="JSON 数组文件，或 - 读 stdin")
    a.add_argument("--sheet", choices=["active", "dead"], default="active")
    a.add_argument("--apply", action="store_true")
    a.add_argument("--no-backup", action="store_true")

    a = sub.add_parser("recompute", help="按存量输入列重算分值（默认 dry-run）")
    a.add_argument("--apply", action="store_true")
    a.add_argument("--r-unknown", type=float, default=0.4,
                   help="发布日期未知时的 R（默认 0.4）")
    a.add_argument("--no-backup", action="store_true")

    a = sub.add_parser("log", help="写一行运行级迭代日志（4 列）")
    a.add_argument("--change", required=True, help="变更项摘要，如「新增 JD270–JD272」")
    a.add_argument("--reason", required=True, help="一句话原因")
    a.add_argument("--before-active", type=int, default=None, help="运行前活跃条数")
    a.add_argument("--before-dead", type=int, default=None, help="运行前失效条数")
    a.add_argument("--dry-run", action="store_true", help="只预览，不写入")

    sub.add_parser("backup", help="手动备份")

    args = p.parse_args(argv)
    CONFIG = load_config(args.config)
    if args.cmd == "describe":
        return cmd_describe(args)
    bk = Book(args.workbook)
    return {
        "audit": cmd_audit, "add": cmd_add, "refresh": cmd_refresh,
        "purge": cmd_purge, "sort": cmd_sort, "repair": cmd_repair,
        "set": cmd_set, "recompute": cmd_recompute, "log": cmd_log,
        "backup": cmd_backup,
    }[args.cmd](bk, args)


if __name__ == "__main__":
    sys.exit(main())
