#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
jd_tool.py — JD 活跃库维护工具（通用）

把原先「每天手写一个新脚本」的做法收敛成一个可复用工具：同一套权重公式、
同一套去重 / 失效 / 备份 / 重排逻辑，只写一次，不再每天重抄。

用法：
    python3 jd_tool.py --workbook <路径> audit                 # 只读体检，不改数据
    python3 jd_tool.py --workbook <路径> add new.json          # 追加当日新 JD（JSON 数组）
    python3 jd_tool.py --workbook <路径> add new.json --strict # strict 四列缺任一即拒收
    python3 jd_tool.py --workbook <路径> add new.json --r-unknown 0.4  # 覆盖「发布日期未知」时的 R
    python3 jd_tool.py --workbook <路径> refresh --ids JD251,JD256   # 刷新「最后发现日期」
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

    --workbook 默认当前目录下的「JD扫描-活跃库.xlsx」；--config 可覆盖默认规格（见下）。

规则真源：见本技能 references/jd-maintenance-rules.md（权重/失效/文件策略等默认）。
本工具只是它的可执行实现，不新增规则。
每次写操作保存前自动向「迭代日志」sheet 追加一行（日期 / 变更项 / 旧值 / 新值 / 原因）；
工作簿无该 sheet 时只打印警告，不阻断保存。

R 时效「未知」口径（2026-10-16 用户裁定）：
    未知取 0.4（与 >90 天同级，不把「发布日期不明」当成新鲜）；
    存量数据按需执行 recompute 重算。--r-unknown 可临时覆盖该默认值。

设计要点
- 列一律按「表头名」定位，绝不按列号硬编码（2026-09-22 曾因列号错位漏了 10 条）。
- 任何写操作前自动备份到工作簿同目录下的 `备份/`。
- purge / sort / repair / set / recompute 默认 dry-run，必须显式 --apply 才落盘。
- add 默认只校验 5 个必填输入列；加 --strict 时还要求 strict 四列全部非空。
- 失效表的 col33/col34 是「失效日期 / 失效原因」（无表头），迁移时写入这两列。
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

# 综合优先级分必填的输入列（用于 add 校验与 audit 重算）
REQUIRED_INPUTS = ["公司名称", "职位名称", "工作城市", "契合度(1-5)", "当前可达性(1-5)"]

# add --strict 要求全部非空的四列（规则 §二；空则拒收）
STRICT_COLS = ["岗位/职能专有能力要求", "Market Signal", "Portfolio Relevance", "Portfolio Benchmark"]


def _rel(p: Path) -> str:
    try:
        return str(Path(p).relative_to(Path.cwd()))
    except ValueError:
        return str(p)


# ---------------------------------------------------------------- 日期

def parse_date(v):
    """容错解析：datetime / date / '2026-10-07' / '2026-07-31(更新)' / '2026-09' / '2026'。"""
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if v is None:
        return None
    s = str(v).strip()
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)          # 2026-10-07（可带括号后缀）
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.match(r"(\d{4})-(\d{1,2})(?!\d)", s)              # 2026-09
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), 1)
    m = re.match(r"(\d{4})(?!\d)", s)                        # 2026
    if m:
        return dt.date(int(m.group(1)), 1, 1)
    return None


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

    LOG_SHEET = "迭代日志"
    LOG_COLS = ["日期", "变更项", "旧值", "新值", "原因"]

    def append_iter_log(self, why=""):
        """保存前向「迭代日志」追加一行；sheet/表头缺失只警告，不报错。"""
        if self.LOG_SHEET not in self.wb.sheetnames:
            print(f"  WARN 无「{self.LOG_SHEET}」sheet，跳过日志写入")
            return False
        ws = self.wb[self.LOG_SHEET]
        hdr = self._header_index(ws)
        missing = [h for h in self.LOG_COLS if h not in hdr]
        if missing:
            print(f"  WARN 「{self.LOG_SHEET}」缺表头 {missing}，跳过日志写入")
            return False
        vals = [str(dt.date.today()), why or "update", "", "", "auto-by jd_tool"]
        r = ws.max_row + 1
        for name, val in zip(self.LOG_COLS, vals):
            ws.cell(r, hdr[name]).value = val
        return True

    def save(self, do_backup=True, why=""):
        if do_backup:
            self.backup(why)
        self.append_iter_log(why)
        self.wb.save(self.path)
        print(f"  [保存] {self.path}")


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
    missing = [h for h in ["去重ID", "公司名称", "职位名称", "工作城市", "最后发现日期", "综合优先级分"]
               if h not in ha]
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
    print("\n[3] 必填字段缺失")
    miss = []
    for r, v in arows:
        for k in ["综合优先级分", "最后发现日期", "首次发现日期", "契合度(1-5)", "当前可达性(1-5)"]:
            if v.get(k) in (None, ""):
                miss.append(f"{v.get('去重ID')}·{k}")
    print(f"  缺失 {len(miss)} 项" + (f"：{miss[:8]}" if miss else ""))
    if miss:
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

    print(f"\n== 汇总：{'全部通过 ✓' if problems == 0 else f'{problems} 类问题需处理'} ==")
    return 1 if (args.strict and problems) else 0


# ---------------------------------------------------------------- add

def cmd_add(bk: Book, args):
    src = sys.stdin.read() if args.file == "-" else Path(args.file).read_text(encoding="utf-8")
    try:
        items = json.loads(src)
    except json.JSONDecodeError as e:
        sys.exit(f"JSON 解析失败：{e}")
    if isinstance(items, dict):
        items = [items]

    today = args.today
    ha = bk.hdr_a()
    hd = bk.hdr_d()

    key3 = lambda d: (str(d.get("公司名称", "")).strip(),
                      str(d.get("职位名称", "")).strip(),
                      str(d.get("工作城市", "")).strip())

    have = {key_strict(v) for _, v in Book.rows(bk.active)}
    have_n = {}
    for _, v in Book.rows(bk.active):
        have_n.setdefault(key_norm(v), []).append(v.get("去重ID"))
    dead = {}
    dead_n = {}
    for _, v in Book.rows(bk.dead):
        dead.setdefault(key_strict(v), []).append(v.get("去重ID"))
        dead_n.setdefault(key_norm(v), []).append(v.get("去重ID"))

    # 下一个去重 ID
    mx = 0
    for ws in (bk.active, bk.dead):
        for _, v in Book.rows(ws):
            m = re.fullmatch(r"JD0*(\d+)", str(v.get("去重ID") or "").strip())
            if m:
                mx = max(mx, int(m.group(1)))
    nxt = mx + 1

    added = []
    for it in items:
        miss = [k for k in REQUIRED_INPUTS if it.get(k) in (None, "")]
        if miss:
            print(f"  SKIP 缺少必填 {miss}：{it.get('公司名称', '?')} · {it.get('职位名称', '?')}")
            continue
        lack = [c for c in STRICT_COLS if it.get(c) in (None, "")]
        if lack:
            if args.strict:
                print(f"  SKIP strict 缺列 {lack}：{it.get('公司名称', '?')} · {it.get('职位名称', '?')}")
                continue
            print(f"  WARN 缺 strict 列 {lack}：{it.get('公司名称', '?')} · {it.get('职位名称', '?')}")
        k = key3(it)
        kn = key_norm(it)
        if k in have:
            print(f"  SKIP 活跃表已有：{k}")
            continue
        if kn in have_n and kn != k:
            print(f"  SKIP 活跃表已有（归一化命中·书写差异）：{kn} ≈ {have_n[kn]}")
            continue
        if k in dead and not args.allow_repost:
            print(f"  SKIP 失效表已有 {dead[k]}（重新发布请加 --allow-repost）：{k}")
            continue
        if kn in dead_n and kn != k and not args.allow_repost:
            print(f"  SKIP 失效表已有 {dead_n[kn]}（归一化命中·书写差异；重新发布请加 --allow-repost）：{kn}")
            continue

        R = r_weight(it.get("发布日期"), today, unknown=args.r_unknown)
        C = c_weight(it.get("公司行业分", 0.5), it.get("公司梯队分", 0.7))
        F = float(it["契合度(1-5)"]) / 5.0
        Ci = ci_weight(it.get("工作城市"), it.get("城市权重Ci"))
        A = float(it["当前可达性(1-5)"]) / 5.0

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
        row["权重明细"] = f"R{R}|C{round(C,2)}|F{F:.2f}|Ci{Ci}|A{A:.2f}"
        orig = dead.get(k) or dead_n.get(kn)
        if orig:  # allow_repost
            row["备注/质量说明"] = (str(row.get("备注/质量说明", ""))
                                    + f" [重新发布·原 {'/'.join(map(str, orig))} 已失效]").strip()

        r = bk.active.max_row + 1
        for name, val in row.items():
            if name in ha:
                bk.active.cell(r, ha[name]).value = val
        for name in ha:  # 未提供的列留空（不写 None，保持原样式）
            if name not in row:
                bk.active.cell(r, ha[name]).value = None
        nxt += 1
        have.add(k)
        have_n.setdefault(kn, []).append(row["去重ID"])
        added.append((row["去重ID"], row["公司名称"], row["职位名称"], row["综合优先级分"], R))

    if not added:
        print("  无新增。")
        return 0

    print(f"\n  新增 {len(added)} 条：")
    for a in added:
        print(f"    {a[0]}  {a[1]} · {a[2]}  综合={a[3]}（R={a[4]}）")
    if any(a[4] == args.r_unknown for a in added) and args.r_unknown != 0.4:
        print(f"  ⚠ 有条目按「发布日期未知 → R={args.r_unknown}」计分")

    if not args.no_sort:
        n = do_sort(bk)
        print(f"  已按综合优先级分降序重排（{n} 行）")
    bk.save(do_backup=not args.no_backup, why=f"add +{len(added)}")
    return 0


# ---------------------------------------------------------------- refresh

def cmd_refresh(bk: Book, args):
    ids = []
    if args.ids:
        ids += [x.strip() for x in args.ids.split(",") if x.strip()]
    if args.ids_file:
        raw = Path(args.ids_file).read_text(encoding="utf-8")
        ids += re.findall(r"JD\d+", raw)
    if not ids:
        sys.exit("请用 --ids JD251,JD256 或 --ids-file <文件> 指定要刷新的去重ID")

    want = set(ids)
    hit, missed = [], []
    ha = bk.hdr_a()
    for r, v in Book.rows(bk.active):
        i = str(v.get("去重ID") or "").strip()
        if i in want:
            bk.active.cell(r, ha["最后发现日期"]).value = str(args.date)
            hit.append(i)
            want.discard(i)
    missed = sorted(want)

    print(f"  刷新「最后发现日期」→ {args.date}：{len(hit)} 条 {hit[:12]}{' ...' if len(hit) > 12 else ''}")
    if missed:
        print(f"  ⚠ 未在活跃表找到：{missed}")
    if not hit:
        print("  无改动。")
        return 0
    bk.save(do_backup=not args.no_backup, why=f"refresh {len(hit)}")
    return 0


# ---------------------------------------------------------------- purge

def cmd_purge(bk: Book, args):
    today = args.date or args.today
    hd = bk.hdr_d()
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

    # 失效表 col33/col34 是「失效日期 / 失效原因」（原无表头）——只在真正写入时补表头
    if hd.get("失效日期") is None:
        bk.dead.cell(1, 33).value = "失效日期"
        bk.dead.cell(1, 34).value = "失效原因"
        bk.dead.cell(1, 33).font = _copy.copy(bk.dead.cell(1, 1).font)
        bk.dead.cell(1, 34).font = _copy.copy(bk.dead.cell(1, 1).font)
        hd = bk.hdr_d()
        print("  [结构] 失效表 col33/col34 已补表头：失效日期 / 失效原因")

    ha = bk.hdr_a()
    moved = 0
    for r, v, why in sorted(targets, key=lambda x: -x[0]):   # 自下而上删，行号不失效
        dr = bk.dead.max_row + 1
        for name, c in ha.items():                          # 按表头名搬运，不按列号
            bk.dead.cell(dr, hd[name]).value = bk.active.cell(r, c).value
        # 活跃表无表头的尾列（col33 若有游离注记）-> 并入备注，避免丢失
        note = bk.active.cell(r, bk.active.max_column).value
        if bk.active.max_column > max(ha.values()) and note not in (None, ""):
            base = bk.dead.cell(dr, hd["备注/质量说明"]).value or ""
            bk.dead.cell(dr, hd["备注/质量说明"]).value = f"{base}\n[原活跃表尾注] {note}".strip()
        bk.dead.cell(dr, hd["失效日期"]).value = str(today)
        bk.dead.cell(dr, hd["失效原因"]).value = why
        bk.active.delete_rows(r)
        moved += 1

    print(f"\n  已移入失效 {moved} 条")
    do_sort(bk)
    bk.save(do_backup=not args.no_backup, why=f"purge -{moved}")
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
    bk.save(do_backup=not args.no_backup, why="sort")
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

    bk.save(do_backup=not args.no_backup, why=f"repair {total}")
    return 0


# ---------------------------------------------------------------- set（补历史字段）

# 这些列由计分流程派生，不允许用 set 直接改写（必须走 add / recompute）
SET_BANNED = {"综合优先级分", "时效权重R", "公司权重C", "契合权重F",
              "城市权重Ci", "权重明细"}


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
        bk.save(do_backup=not args.no_backup, why=f"set {n} 处")
    else:
        print("  [dry-run] 未写入。加 --apply 执行。")
    return 0


def _as_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def _same_num(a, b):
    x, y = _as_num(a), _as_num(b)
    if x is None or y is None:
        return False
    return abs(x - y) < 1e-6


def cmd_recompute(bk: Book, args):
    """按**存量输入列**重算派生分（默认 dry-run，--apply 才落盘）。

    输入：发布日期 / 公司行业分 / 公司梯队分 / 契合度(1-5) / 当前可达性(1-5) / 工作城市，
          城市权重Ci 优先取存量值（与 add 一致），缺失时按工作城市推导。
    输出：时效权重R / 公司权重C / 契合权重F / 城市权重Ci / 综合优先级分，
          以及「权重明细」（活跃表存在该列时一并同步）。
    发布日期未知时 R 取 --r-unknown（默认 0.4）。
    """
    today = args.today
    ha = bk.hdr_a()
    need = ["发布日期", "公司行业分", "公司梯队分", "契合度(1-5)", "当前可达性(1-5)", "工作城市",
            "时效权重R", "公司权重C", "契合权重F", "城市权重Ci", "综合优先级分"]
    missing = [h for h in need if h not in ha]
    if missing:
        sys.exit(f"活跃表缺少 recompute 所需列：{missing}")

    changes, skipped = [], []
    for r, v in Book.rows(bk.active):
        try:
            R = r_weight(v.get("发布日期"), today, unknown=args.r_unknown)
            C = round(c_weight(v.get("公司行业分", 0.5), v.get("公司梯队分", 0.7)), 3)
            F = round(float(v["契合度(1-5)"]) / 5.0, 3)
            Ci = ci_weight(v.get("工作城市"), v.get("城市权重Ci"))
            A = float(v["当前可达性(1-5)"]) / 5.0
        except (TypeError, ValueError):
            skipped.append(v.get("去重ID"))
            continue
        new = {"时效权重R": R, "公司权重C": C, "契合权重F": F,
               "城市权重Ci": Ci, "综合优先级分": total_score(R, C, F, Ci, A)}
        diffs = [(k, v.get(k), val) for k, val in new.items() if not _same_num(v.get(k), val)]
        if "权重明细" in ha:
            new["权重明细"] = f"R{R}|C{round(C,2)}|F{F:.2f}|Ci{Ci}|A{A:.2f}"
            if str(v.get("权重明细") or "") != new["权重明细"]:
                diffs.append(("权重明细", v.get("权重明细"), new["权重明细"]))
        if diffs:
            changes.append((r, v, new, diffs))

    print(f"== recompute（基准 {today}，R 未知={args.r_unknown}）==")
    tail = f"，输入列不可解析 SKIP {len(skipped)} 条：{skipped[:6]}" if skipped else ""
    print(f"  待更新 {len(changes)} 条{tail}")
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
    bk.save(do_backup=not args.no_backup, why="recompute")
    return 0


# ---------------------------------------------------------------- backup

def cmd_backup(bk: Book, args):
    bk.backup("手动备份")
    return 0


# ---------------------------------------------------------------- CLI

def main(argv=None):
    global CONFIG
    p = argparse.ArgumentParser(
        description="JD 活跃库维护工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n"
               "  jd_tool.py --workbook <路径> add new.json --strict\n"
               "  jd_tool.py --workbook <路径> add new.json --r-unknown 0.4\n"
               "  jd_tool.py --workbook <路径> recompute [--apply]",
    )
    p.add_argument("--workbook", type=Path, default=Path.cwd() / "JD扫描-活跃库.xlsx",
                   help="工作簿路径（默认当前目录下的 JD扫描-活跃库.xlsx）")
    p.add_argument("--config", type=Path, default=None,
                   help="配置 JSON，可覆盖 sheets/weights/city/city_default/stale_days（默认内置规格）")
    p.add_argument("--today", type=lambda s: dt.date.fromisoformat(s), default=dt.date.today(),
                   help="把「今天」固定为某日期（测试用）")
    sub = p.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("audit", help="只读体检")
    a.add_argument("--strict", action="store_true", help="有问题时返回非零退出码")

    a = sub.add_parser("add", help="追加当日新 JD（JSON 数组文件，或 - 读 stdin）")
    a.add_argument("file")
    a.add_argument("--allow-repost", action="store_true", help="允许失效后重新发布的岗位再入册")
    a.add_argument("--strict", action="store_true",
                   help="strict 四列（岗位/职能专有能力要求 / Market Signal / Portfolio Relevance / "
                        "Portfolio Benchmark）任一为空即拒收")
    a.add_argument("--r-unknown", type=float, default=0.4,
                   help="发布日期未知时的 R（默认 0.4，2026-10-16 已裁定）")
    a.add_argument("--no-sort", action="store_true")
    a.add_argument("--no-backup", action="store_true")

    a = sub.add_parser("refresh", help="刷新「最后发现日期」")
    a.add_argument("--ids", help="逗号分隔，如 JD251,JD256")
    a.add_argument("--ids-file", help="从文件里抓取所有 JDxxx")
    a.add_argument("--date", type=lambda s: dt.date.fromisoformat(s), default=dt.date.today())
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

    sub.add_parser("backup", help="手动备份")

    args = p.parse_args(argv)
    CONFIG = load_config(args.config)
    bk = Book(args.workbook)
    return {
        "audit": cmd_audit, "add": cmd_add, "refresh": cmd_refresh,
        "purge": cmd_purge, "sort": cmd_sort, "repair": cmd_repair,
        "set": cmd_set, "recompute": cmd_recompute, "backup": cmd_backup,
    }[args.cmd](bk, args)


if __name__ == "__main__":
    sys.exit(main())
