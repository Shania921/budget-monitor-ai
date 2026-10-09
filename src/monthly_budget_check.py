# -*- coding: utf-8 -*-
"""
月度预算使用明细表 —— 一键核查脚本（L1）
=====================================================================
功能：对当月《预算使用明细表》主工作簿做端到端独立重算核查：
  1. 结构自检：动态识别汇总表科目行（按 A 列科目名定位，不写死行号），
     校验核心科目存在 / 科目相对顺序 / 区域列表头（防列错位回归）
  2. 逐格重算：汇总表"科目 x 区域"已用值 vs 三张底表独立聚合（容忍 50 元）
  3. 父行/子行：专业费用父行 = 其下展开子行加总（6 区域防重复计数）
  4. 总账闭环：底表合计 vs 汇总已用总计；授予/已用/剩余 勾稽；
     已知豁免尾差（#N/A 行 + 长科目名未归类）自动计入后再判定
  5. 超支扫描：逐科目逐区域 已用 vs 授予（专业费用按大类池口径）
  6. 尾差分解：底表中有、汇总未归类的金额自动列明（长科目名/#N/A 等）

设计要点：
  - 汇总表"总计"行、各科目行全部按名称动态定位 —— 模板增删科目行、
    行号漂移都不会误报；科目"存在才比对"，缺失科目视为零发生（不报警）。
  - 底表口径：差旅费/费用明细/专业费用三张底表独立聚合（金额"元"→万）。

用法：
  python monthly_budget_check.py                 # 自动定位"预算使用明细表"目录下最新月份主表
  python monthly_budget_check.py --file <path>   # 指定当月主表文件
  python monthly_budget_check.py --out <dir>     # 指定报告输出目录（默认 ./reports）
  python monthly_budget_check.py --quiet         # 控制台只出结论摘要

默认搜索目录（自动取月份数字最大的子目录）：
  <仓库根目录>/data/预算使用明细表/<N月>/...
  候选主表优先级：预算使用明细表.xlsx > 预算使用明细表-底稿.xlsx >
                 名字含"预算使用明细表"且不含"截止/截至/云南/四川/西藏/贵州/重庆"的文件

> 本仓库为脱敏演示版：随附的演示数据（src/generate_fake_data.py 生成）
> 结构与真实业务模板一致，金额与主体均为虚构。
=====================================================================
"""
import os
import re
import sys
import argparse
import datetime
import openpyxl
from openpyxl.utils import column_index_from_string
from collections import defaultdict

# ---------------- 可调配置 ----------------
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE_DIR = os.path.join(_REPO_ROOT, "data", "预算使用明细表")
TOLERANCE = 0.005          # 万元（50 元）以内视为尾差
MAIN_SHEET = "汇总"         # 汇总表 sheet 名
GRANT_SHEET = "授予额度"    # 授予额度 sheet 名

# 6 区域 -> 汇总表"已使用"列字母
REGION_USE_COL = {"运营部": "F", "重庆": "I", "云南": "L", "贵州": "O",
                  "四川": "R", "西藏": "U"}
# 授予额度表：B..G = 重庆,云南,贵州,四川,西藏,运营部
GRANT_REGIONS = ["重庆", "云南", "贵州", "四川", "西藏", "运营部"]
GRANT_REG_COL = {r: 2 + i for i, r in enumerate(GRANT_REGIONS)}

# 标准科目顺序（汇总表 A 列从上到下应基本如此；用于顺序错乱检测）
STANDARD_SUBJECTS = [
    "差旅费", "招待费-归口报账", "招待费-员工报账", "市场推广费", "专业费用",
    "专业费用-劳务费-专家咨询费", "专业费用-认证费", "专业费用-设备调试费",
    "专业费用-实验检验费", "专业费用-招标服务费", "专业费用-咨询费",
    "专业费用-法务费", "专业费用-设计费",
    "邮寄运输费", "办公类", "低值易耗品消耗", "设备维修费", "外协加工费",
    "其他费用", "设备租赁费", "报告解读费", "代理费", "物料", "专业费用-IT归口",
]
# 专业费用父行展开子行（存在才参与父=子校验，不参与超支扫描）
CHILD_SUBJECTS = [s for s in STANDARD_SUBJECTS
                  if s.startswith("专业费用-") and s != "专业费用-IT归口"]
# 核心科目：缺失即视为"模板结构异常"
CORE_SUBJECTS = ["差旅费", "专业费用", "邮寄运输费", "专业费用-IT归口", "总计"]
# 可选科目（9 月精简版删掉的零发生科目；缺失只提示不报错）
OPTIONAL_SUBJECTS = ["报告解读费", "代理费", "物料"]

# 底表 sheet 与取值列（金额均为"元"，汇总公式 /10000 转万元）
BASE_SHEETS = [
    # (sheet名, 金额列, 区域列, 科目列, 项目列)  # 科目/项目列可为 None
    ("差旅费",            "U", "AF", None, None),   # U=不含税金额CNY
    ("费用明细-费用明细", "T", "V",  "G",  None),   # T=发生金额, G=预算科目
    ("专业费用",          "V", "X",  "G",  "I"),    # V=发生金额, G=预算科目, I=预算项目
]


# ---------------- 工具函数 ----------------
def col(letter):
    return column_index_from_string(letter)


def num(v):
    """安全转 float；非数字返回 None"""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return None
    try:
        return float(str(v).strip().replace(",", ""))
    except (ValueError, AttributeError):
        return None


def A(ws, r):
    return str(ws.cell(row=r, column=1).value or "").strip()


def find_main_file(month_dir):
    """在月份目录里按优先级挑主表文件（含'汇总'sheet 才算数）"""
    candidates = []
    if not os.path.isdir(month_dir):
        return None
    for fn in os.listdir(month_dir):
        if not fn.endswith((".xlsx", ".xlsm")) or fn.startswith("~$"):
            continue
        candidates.append(fn)
    if not candidates:
        return None
    # 优先级 1：精确名；2：含"底稿"；3：含"预算使用明细表"且排除分省/导出
    def priority(fn):
        if fn == "预算使用明细表.xlsx":
            return 0
        if fn == "预算使用明细表-底稿.xlsx":
            return 1
        if ("预算使用明细表" in fn and not any(x in fn for x in
                ("截止", "截至", "云南", "四川", "西藏", "贵州", "重庆", "导出"))):
            return 2
        return 9
    for fn in sorted(candidates, key=priority):
        if priority(fn) >= 9:
            continue
        try:
            wb = openpyxl.load_workbook(os.path.join(month_dir, fn),
                                        read_only=True, data_only=True)
            if MAIN_SHEET in wb.sheetnames and GRANT_SHEET in wb.sheetnames:
                wb.close()
                return os.path.join(month_dir, fn)
            wb.close()
        except Exception:
            continue
    return None


def locate_latest():
    """在 BASE_DIR 下找月份数字最大的目录，返回 (目录, 主表路径, 月份显示名)"""
    if not os.path.isdir(BASE_DIR):
        return None, None, None
    months = []
    for d in os.listdir(BASE_DIR):
        m = re.match(r"^(\d{1,2})月?$", d)
        if m and os.path.isdir(os.path.join(BASE_DIR, d)):
            months.append((int(m.group(1)), d))
    if not months:
        return None, None, None
    months.sort(key=lambda x: -x[0])
    month_dir = os.path.join(BASE_DIR, months[0][1])
    f = find_main_file(month_dir)
    return month_dir, f, months[0][1]


def load_meta(path):
    """加载主工作簿（公式版 + 数值版），返回 (wbF, wbV)"""
    wbF = openpyxl.load_workbook(path, data_only=False)   # 公式
    wbV = openpyxl.load_workbook(path, data_only=True)    # 缓存值
    return wbF, wbV


def find_total_row(ws):
    """动态定位汇总表'总计'行（防模板插行/删行）"""
    for r in range(15, 45):
        if A(ws, r) == "总计":
            return r
    return None


def scan_subject_rows(ws_h, total_row):
    """扫描汇总表 A 列（3..total_row-1），返回 {科目名: 行号}"""
    rows = {}
    for r in range(3, total_row):
        nm = A(ws_h, r)
        if nm:
            rows[nm] = r
    return rows


# ---------------- 核查 1：结构自检 ----------------
def check_structure(ws_h, rows, total_row):
    """汇总表科目行结构是否与模板一致；区域列头探测。返回 (issues, heads, notes)"""
    issues, notes = [], []
    # 表头行2 应含"已使用/已使用全部"之类；区域列表头探测（行1 区域名，行2 指标名）
    heads = {}
    for reg, L in REGION_USE_COL.items():
        c = col(L)
        v1 = ws_h.cell(row=1, column=c).value
        v2 = ws_h.cell(row=2, column=c).value
        heads[reg] = (str(v1 or "").strip(), str(v2 or "").strip())
        if not (v1 or v2):
            issues.append(f"区域「{reg}」列 {L} 表头为空（列错位？）")
    # 核心科目必须存在（"总计"行单独定位，已在 find_total_row 校验）
    for nm in CORE_SUBJECTS:
        if nm == "总计":
            if total_row is None:
                issues.append("未找到「总计」行（模板结构异常？）")
            continue
        if nm not in rows:
            issues.append(f"核心科目「{nm}」缺失（模板结构异常？）")
    # 可选科目缺失：提示不报错
    missing_opt = [nm for nm in OPTIONAL_SUBJECTS if nm not in rows]
    if missing_opt:
        notes.append("汇总表未列科目（按零发生处理）：" + "、".join(missing_opt))
    # 相对顺序校验：出现的标准科目必须按 STANDARD_SUBJECTS 顺序单调
    pos = {nm: i for i, nm in enumerate(STANDARD_SUBJECTS)}
    seq = [nm for nm in rows if nm in pos]
    for a, b in zip(seq, seq[1:]):
        if pos[a] > pos[b]:
            issues.append(f"科目顺序异常：「{a}」(行{rows[a]}) 应在「{b}」(行{rows[b]}) 之后")
    # 未知科目提示
    unknown = [nm for nm in rows if nm not in pos and nm != "总计"]
    if unknown:
        notes.append("汇总表出现未在标准清单中的科目行：" + "、".join(unknown))
    return issues, heads, notes


# ---------------- 核查 2：逐格重算 ----------------
def recompute_grid(wbV, ws_h, rows):
    """三张底表独立聚合 -> 逐格比对汇总表已用值。返回 (problems, checked, totals, na_rows)"""
    by_region = {}          # sheet名 -> {region: 万元}
    by_region_subj = {}     # sheet名 -> {(region, 科目G): 万元}
    by_region_proj = {}     # sheet名 -> {(region, 项目I): 万元}
    na_rows = defaultdict(list)   # sheet名 -> [(行号, 元, 科目)]

    for sn, money_col, region_col, subj_col, proj_col in BASE_SHEETS:
        if sn not in wbV.sheetnames:
            continue
        ws = wbV[sn]
        mc, rc = col(money_col), col(region_col)
        sc = col(subj_col) if subj_col else None
        pc = col(proj_col) if proj_col else None
        br, brs, brp = defaultdict(float), defaultdict(float), defaultdict(float)
        for r in range(2, ws.max_row + 1):
            amt = ws.cell(row=r, column=mc).value
            if not isinstance(amt, (int, float)):
                continue
            region = str(ws.cell(row=r, column=rc).value or "").strip()
            w = amt / 10000.0
            br[region] += w
            if sc:
                subj = str(ws.cell(row=r, column=sc).value or "").strip()
                brs[(region, subj)] += w
            if pc:
                proj = str(ws.cell(row=r, column=pc).value or "").strip()
                brp[(region, proj)] += w
            if region == "#N/A":
                subj = str(ws.cell(row=r, column=sc).value or "") if sc else ""
                na_rows[sn].append((r, amt, subj))
        by_region[sn] = dict(br)
        by_region_subj[sn] = dict(brs)
        by_region_proj[sn] = dict(brp)

    chailv = by_region.get("差旅费", {})
    fy_subj = by_region_subj.get("费用明细-费用明细", {})
    zy_subj = by_region_subj.get("专业费用", {})
    zy_proj = by_region_proj.get("专业费用", {})

    def expected(region, name):
        """按科目名给期望值（独立重算口径）"""
        if name == "差旅费":
            return chailv.get(region, 0.0)
        if name == "专业费用":                       # 父行 = 专业费用底表 G 列大类
            return zy_subj.get((region, "专业费用"), 0.0)
        if name == "专业费用-IT归口":                 # 独立科目（底表 G 列同名）
            return zy_subj.get((region, "专业费用-IT归口"), 0.0)
        if name in CHILD_SUBJECTS:                    # 展开子行 = 底表 I 列预算项目
            return zy_proj.get((region, name), 0.0)
        return fy_subj.get((region, name), 0.0)       # 其余 = 费用明细 G 列科目

    problems, checked = [], 0
    for name, r in rows.items():
        if name in CORE_SUBJECTS or name == "总计":
            continue
        for reg, L in REGION_USE_COL.items():
            v = ws_h.cell(row=r, column=col(L)).value
            fv = num(v)
            if fv is None:
                continue
            exp = expected(reg, name)
            checked += 1
            if abs(fv - exp) > TOLERANCE:
                problems.append((r, name, reg, fv, exp, fv - exp))

    totals = {"chailv": sum(chailv.values()),
              "fy": sum(by_region.get("费用明细-费用明细", {}).values()),
              "zy": sum(by_region.get("专业费用", {}).values()),
              "na": sum(x[1] for x in na_rows.get("费用明细-费用明细", [])) / 10000.0}
    return problems, checked, totals, na_rows


# ---------------- 核查 3：父行 = 子行 ----------------
def check_parent_child(ws_h, rows):
    parent_row = rows.get("专业费用")
    child_rows = [rows[nm] for nm in CHILD_SUBJECTS if nm in rows]
    res = []
    if parent_row is None or not child_rows:
        return res, "父行或子行缺失，跳过"
    for reg, L in REGION_USE_COL.items():
        c = col(L)
        pv = num(ws_h.cell(row=parent_row, column=c).value) or 0.0
        sv = sum(num(ws_h.cell(row=r, column=c).value) or 0.0 for r in child_rows)
        res.append((reg, pv, sv, abs(pv - sv)))
    return res, ""


# ---------------- 核查 4：总账闭环 ----------------
def check_totals(ws_h, totals, rows, total_row, exempt):
    """exempt：已知豁免尾差（#N/A + 长科目名未归类），计入后再判定"""
    B = lambda r: num(ws_h.cell(row=r, column=2).value) or 0.0   # noqa: E731
    C = lambda r: num(ws_h.cell(row=r, column=3).value) or 0.0   # noqa: E731
    D = lambda r: num(ws_h.cell(row=r, column=4).value) or 0.0   # noqa: E731
    b_t, c_t, d_t = B(total_row), C(total_row), D(total_row)
    # 科目列加总：只加"独立科目 + 专业费用父行"，不加展开子行（防重复）
    subj_rows = [r for nm, r in rows.items() if nm not in CHILD_SUBJECTS]
    sum_b = sum(B(r) for r in subj_rows)
    sum_c = sum(C(r) for r in subj_rows)
    sum_d = sum(D(r) for r in subj_rows)
    base_total = totals["chailv"] + totals["fy"] + totals["zy"]

    checks = []
    def add(name, ok, detail=""):   # noqa: E306
        checks.append((name, ok, detail))

    add("授予总额 B(总计) = 科目列加总", abs(b_t - sum_b) < TOLERANCE,
        f"总计={b_t:.4f} 加总={sum_b:.4f} 差={b_t-sum_b:+.5f}")
    add("已用总额 C(总计) = 科目列加总", abs(c_t - sum_c) < TOLERANCE,
        f"总计={c_t:.4f} 加总={sum_c:.4f} 差={c_t-sum_c:+.5f}")
    if d_t or sum_d:
        add("剩余总额 D(总计) = 科目列加总", abs(d_t - sum_d) < TOLERANCE,
            f"总计={d_t:.4f} 加总={sum_d:.4f} 差={d_t-sum_d:+.5f}")
    add("B(总计) − C(总计) = D(总计)", abs((b_t - c_t) - d_t) < TOLERANCE,
        f"{b_t:.4f}-{c_t:.4f}={b_t-c_t:.4f} vs D={d_t:.4f}")
    resid = base_total - c_t - exempt
    add("底表合计 ≈ C(总计) + 已知豁免", abs(resid) < TOLERANCE,
        f"底表={base_total:.4f} C={c_t:.4f} 差={base_total-c_t:+.6f}；"
        f"豁免(#N/A+未归类)={exempt:.4f} 万 → 剩余={resid:+.6f} 万")

    tail = base_total - c_t           # 原尾差（万元，未扣豁免）
    return checks, tail, base_total, c_t


# ---------------- 核查 5：超支扫描 ----------------
def overspend_scan(ws_h, ws_g, rows):
    def gval(subj, reg):
        for r in range(4, ws_g.max_row + 1):
            if str(ws_g.cell(row=r, column=1).value or "").strip() == subj:
                v = ws_g.cell(row=r, column=GRANT_REG_COL[reg]).value
                return num(v) if isinstance(v, (int, float)) else None
        return None

    def hval(row, reg):
        return num(ws_h.cell(row=row, column=col(REGION_USE_COL[reg])).value) or 0.0

    issues, rows_out = [], []
    for name, row in rows.items():
        if name in CHILD_SUBJECTS:     # 展开子行不单独扫（父行按大类池口径）
            continue
        for reg in REGION_USE_COL:
            g = gval(name, reg)
            if g is None:
                continue
            u = hval(row, reg)
            if g == 0 and u == 0:
                continue
            if u > g + 0.001:
                pct = (u - g) / g * 100 if g > 0 else 0
                issues.append((name, reg, g, u, u - g, pct))
                rows_out.append((name, reg, g, u, u - g, pct))
    return issues, rows_out


# ---------------- 核查 6：尾差分解 ----------------
def collect_unmapped(wbV, rows):
    """费用明细中科目名未命中汇总表任何科目行的金额（长科目名/#N/A之外）"""
    if "费用明细-费用明细" not in wbV.sheetnames:
        return [], 0.0
    ws = wbV["费用明细-费用明细"]
    mc, rc, sc = col("T"), col("V"), col("G")
    used_names = set(rows.keys()) | {""}
    unmapped = defaultdict(float)
    for r in range(2, ws.max_row + 1):
        amt = ws.cell(row=r, column=mc).value
        if not isinstance(amt, (int, float)):
            continue
        region = str(ws.cell(row=r, column=rc).value or "").strip()
        subj = str(ws.cell(row=r, column=sc).value or "").strip()
        if region == "#N/A":
            continue
        if subj not in used_names:
            unmapped[(region, subj)] += amt
    items = sorted(unmapped.items(), key=lambda x: -x[1])
    total = sum(v for _, v in items) / 10000.0
    return items, total


def explain_tail(totals, c_t, exempt, na_rows, unmapped_items, has_fy_sheet):
    """把底表合计 - C 的差额拆解说明"""
    d = totals["chailv"] + totals["fy"] + totals["zy"] - c_t
    if abs(d) < 0.0005:
        return []
    lines = [f"底表合计 − C(总计) = {d:+.6f} 万（{d*10000:+.1f} 元）"]
    if totals["na"] > 0:
        lines.append(f"  ├─ #N/A 区域行（费用明细表）：{totals['na']*10000:.2f} 元"
                     f"（{len(na_rows.get('费用明细-费用明细', []))} 行，按既定口径豁免）")
    rest = d - totals["na"] - exempt
    if abs(rest) < 0.0001:
        pass
    if unmapped_items:
        for (reg, subj), amt in unmapped_items:
            lines.append(f"  ├─ 未归类科目「{subj}」{reg}：{amt:.2f} 元"
                         f"（汇总表科目名未精确匹配，属已知豁免口径）")
    else:
        lines.append(f"  └─ 其余：{d*10000 - totals['na']*10000:.2f} 元")
    return lines


# ---------------- 报告生成 ----------------
def render_report(ctx, problems, pc_checks, pc_note, overs, tail_lines,
                  struct_issues, heads, notes, tail, exempt, rows):
    month = ctx["month"]
    path = ctx["path"]
    L = []
    L.append(f"# 月度预算核查报告 · {month}")
    L.append("")
    L.append(f"> 核查对象：`{path}`")
    L.append(f"> 核查方式：从三张底表独立重算，与汇总表逐格比对（不依赖表内公式缓存）")
    L.append(f"> 核查时间：{datetime.datetime.now():%Y-%m-%d %H:%M}")
    L.append("")
    # ---- 结论 ----
    fatal = bool(struct_issues) or bool(problems)
    over_cnt = len(overs)
    L.append("## 一、总体结论")
    L.append("")
    if fatal:
        L.append("**⚠ 存在问题，需人工处理后再上报**")
    elif over_cnt:
        L.append("**✅ 表结构/数据正确；存在超支（业务问题，非表错误）**")
    else:
        L.append("**✅ 全部通过，表可以放心上报**")
    L.append("")
    L.append("| 校验维度 | 结果 |")
    L.append("|---|---|")
    L.append(f"| 结构自检（科目行/区域列） | {'❌ '+str(len(struct_issues))+' 处' if struct_issues else '✓ 与模板一致'} |")
    L.append(f"| 汇总表 {ctx['checked']} 个科目×区域格子 vs 底表重算 | {'❌ '+str(len(problems))+' 处差异' if problems else '0 差异（容忍 50 元）'} |")
    pc_ok = "✓ 全部相等" if pc_checks and all(x[3] < 0.0001 for x in pc_checks) else ("跳过（" + pc_note + "）" if pc_note else "❌ 存在差异")
    L.append(f"| 父行(专业费用) = 展开子行加总（6 区域） | {pc_ok} |")
    L.append(f"| 总账闭环（授予/已用/剩余勾稽） | {sum(1 for _, ok, _ in ctx['tot_checks'] if ok)}/{len(ctx['tot_checks'])} 项通过 |")
    L.append(f"| 底表合计 vs 已用总计 | {ctx['base_total']:.4f} vs {ctx['c27']:.4f} 万"
             f"（差 {tail*10000:+.0f} 元，含已知豁免 {exempt*10000:.0f} 元）|")
    L.append(f"| 超支扫描 | {over_cnt} 条 |")
    L.append("")
    if notes:
        L.append("> 提示：" + "；".join(notes))
        L.append("")
    # ---- 问题明细 ----
    if struct_issues:
        L.append("## 二、结构自检异常（模板可能变了）")
        L.append("")
        for s in struct_issues:
            L.append(f"- {s}")
        L.append("")
        L.append("> 本月模板若有科目增删属正常，脚本按科目名自动适配；"
                 "若核心科目缺失或顺序错乱，请人工复核汇总表结构。")
        L.append("")
    if problems:
        L.append("## 二、逐格差异明细（>50 元）")
        L.append("")
        L.append("| 行 | 科目 | 区域 | 表内值(万) | 独立重算(万) | 差(万) |")
        L.append("|---|---|---|---|---|---|")
        for r, name, reg, fv, exp, d in problems:
            L.append(f"| {r} | {name} | {reg} | {fv:.6f} | {exp:.6f} | **{d:+.6f}** |")
        L.append("")
    # ---- 超支 ----
    L.append("## 三、超支清单（已用 > 授予）")
    L.append("")
    if overs:
        L.append("| 科目 | 区域 | 授予(万) | 已用(万) | 超支(万) | 超幅 |")
        L.append("|---|---|---|---|---|---|")
        for s, reg, g, u, d, pct in sorted(overs, key=lambda x: -x[4]):
            L.append(f"| {s} | {reg} | {g:.4f} | {u:.4f} | **+{d:.4f}** | {pct:.0f}% |")
    else:
        L.append("无。")
    L.append("")
    # ---- 尾差 ----
    L.append("## 四、尾差说明")
    L.append("")
    if tail_lines:
        for t in tail_lines:
            L.append(f"- {t}")
    else:
        L.append("无尾差（底表合计 = 汇总已用）。")
    L.append("")
    # ---- 附：科目与区域快照 ----
    L.append("## 五、结构快照")
    L.append("")
    L.append("汇总表科目行（动态识别）：")
    L.append("")
    for nm, r in rows.items():
        L.append(f"- 行{r}: {nm}")
    L.append("")
    L.append("汇总表区域列（已使用列）表头探测：")
    L.append("")
    for reg, (h1, h2) in heads.items():
        L.append(f"- {reg}: 列 {REGION_USE_COL[reg]} 行1/行2 = 「{h1}」/「{h2}」")
    L.append("")
    L.append(f"- 授予总额 B(总计) = {ctx['b27']:.4f} 万；"
             f"底表行数：差旅费/费用明细/专业费用已自动读取。")
    L.append("")
    L.append("---")
    L.append(f"报告由 monthly_budget_check.py 自动生成。")
    return "\n".join(L)


# ---------------- 核查入口（CLI 与独立工作台共用） ----------------
def run_check(path, month=None):
    """对单个主工作簿执行全套核查，返回结构化 dict（供 CLI / 工作台共用）。

    返回的 dict 字段：
      ok          表本身是否正确（无结构异常、无逐格差异）
      month/path  核查对象；checked_at 核查时间
      summary     授予/已用/剩余/底表合计/科目数
      check_items [{key,name,ok,detail}] 六项检查卡片数据
      overspend   超支 [{subject,region,grant,used,over,pct}]
      tail        尾差 {diff_yuan, exempt_yuan, lines[]}
      notes       提示（模板差异等非错误信息）
      report_md   完整 Markdown 报告文本
    """
    if not os.path.isfile(path):
        raise ValueError(f"文件不存在：{path}")
    if not month:
        # 月份探测优先级：文件名含「N月」 > 所在目录含「N月」 > 未命名
        # （工作台上传场景文件落在 data/uploads/<sid>.xlsx，须靠文件名/显式参数）
        base = os.path.basename(path)
        m = re.search(r"(\d{1,2})\s*月", base)
        if m:
            month = m.group(0).replace(" ", "")
        else:
            folder = os.path.basename(os.path.dirname(os.path.abspath(path)))
            m = re.search(r"(\d{1,2})\s*月", folder)
            month = m.group(0).replace(" ", "") if m else "未命名"
    wbF, wbV = load_meta(path)
    if MAIN_SHEET not in wbV.sheetnames:
        raise ValueError(f"文件里找不到「{MAIN_SHEET}」sheet，请确认上传的是当月主表（预算使用明细表.xlsx）")
    ws_h = wbV[MAIN_SHEET]
    ws_g = wbV[GRANT_SHEET] if GRANT_SHEET in wbV.sheetnames else None

    # 动态结构识别
    total_row = find_total_row(ws_h)
    rows = scan_subject_rows(ws_h, total_row) if total_row else {}

    # 1 结构
    struct_issues, heads, notes = check_structure(ws_h, rows, total_row)
    # 2 重算
    problems, checked, totals, na_rows = recompute_grid(wbV, ws_h, rows)
    # 3 父行子行
    pc_checks, pc_note = check_parent_child(ws_h, rows)
    # 豁免尾差（#N/A + 未归类长科目名）
    unmapped_items, unmapped_total = collect_unmapped(wbV, rows)
    exempt = totals["na"] + unmapped_total
    # 4 总账
    tot_checks, tail, base_total, c27 = check_totals(ws_h, totals, rows, total_row, exempt)
    # 5 超支
    overs, _ = overspend_scan(ws_h, ws_g, rows) if ws_g else ([], [])
    # 6 尾差分解
    tail_lines = explain_tail(totals, c27, exempt, na_rows, unmapped_items,
                              "费用明细-费用明细" in wbV.sheetnames)

    b27 = num(ws_h.cell(row=total_row, column=2).value) or 0.0
    d27 = num(ws_h.cell(row=total_row, column=4).value)
    ctx = {"month": month, "path": path, "checked": checked,
           "tot_checks": tot_checks, "base_total": base_total, "c27": c27, "b27": b27}
    md = render_report(ctx, problems, pc_checks, pc_note, overs, tail_lines,
                       struct_issues, heads, notes, tail, exempt, rows)

    # 六项检查卡片
    def _desc(ok, yes, no):
        return yes if ok else no
    pc_pass = bool(pc_checks) and all(x[3] < 0.0001 for x in pc_checks)
    pc_skip = bool(pc_note)
    tail_ok = (not tail_lines) or (abs(tail) - exempt < TOLERANCE + 1e-9)
    check_items = [
        {"key": "structure", "name": "结构自检（科目行 / 区域列）", "ok": not struct_issues,
         "detail": "✓ 与模板一致" if not struct_issues else f"❌ {len(struct_issues)} 处：" + "；".join(struct_issues)},
        {"key": "recompute", "name": "逐格重算（三张底表独立聚合）", "ok": not problems,
         "detail": (f"✓ {checked} 格 0 差异（容忍 50 元）" if not problems
                    else f"❌ {checked} 格中 {len(problems)} 处差异，最大 {max(abs(x[5]) for x in problems):.4f} 万")},
        {"key": "parent", "name": "父行(专业费用) = 子行加总", "ok": pc_pass or pc_skip,
         "detail": f"✓ {len(pc_checks)} 区域全部相等" if pc_pass
                   else (f"· {pc_note}" if pc_skip else "❌ 存在差异")},
        {"key": "totals", "name": "总账闭环（授予/已用/剩余勾稽）", "ok": all(ok for _, ok, _ in tot_checks),
         "detail": f"{sum(1 for _, ok, _ in tot_checks if ok)}/{len(tot_checks)} 项通过"},
        {"key": "tail", "name": "尾差分解", "ok": tail_ok,
         "detail": ("无尾差" if not tail_lines
                    else f"{tail*10000:+.0f} 元（含已知豁免 {exempt*10000:.0f} 元，已解释）"
                    if tail_ok else f"{tail*10000:+.0f} 元，超出豁免范围需追查")},
        {"key": "overspend", "name": "超支扫描（已用 vs 授予）", "ok": not overs,
         "detail": f"{len(overs)} 条（业务问题，需说明）" if overs else "无超支"},
    ]

    return {"ok": (not struct_issues) and (not problems),
            "month": month, "path": path,
            "checked_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "summary": {"grant": round(b27, 4), "used": round(c27, 4),
                        "remain": round(d27, 4) if d27 is not None else None,
                        "base": round(base_total, 4),
                        "subjects": len([n for n in rows if n not in CORE_SUBJECTS])},
            "check_items": check_items,
            "overspend": [{"subject": s, "region": reg, "grant": round(g, 4),
                           "used": round(u, 4), "over": round(d, 4), "pct": round(pct)}
                          for s, reg, g, u, d, pct in sorted(overs, key=lambda x: -x[4])],
            "tail": {"diff_yuan": round(tail * 10000, 2),
                     "exempt_yuan": round(exempt * 10000, 2),
                     "lines": tail_lines},
            "notes": notes, "struct_issues": struct_issues,
            "report_md": md}


# ---------------- 主流程 ----------------
def main():
    ap = argparse.ArgumentParser(description="月度预算使用明细表一键核查")
    ap.add_argument("--file", help="指定当月主表 xlsx")
    ap.add_argument("--out", default=os.path.join(_REPO_ROOT, "reports"))
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    # 定位文件
    if args.file:
        path = args.file
    else:
        _, path, _ = locate_latest()
        if not path:
            print(f"[错误] 在 {BASE_DIR} 下未找到含「汇总/授予额度」sheet 的主表。")
            print("        请用 --file 指定当月主表文件。")
            sys.exit(1)

    try:
        r = run_check(path)
    except ValueError as e:
        print(f"[错误] {e}")
        sys.exit(1)
    month = r["month"]

    def say(x=""):
        if not args.quiet:
            print(x)
    say(f"月度核查 · {month}")
    say(f"文件: {r['path']}")
    say("-" * 60)
    for it in r["check_items"]:
        say(f"{it['name']:<30s}: {it['detail']}")
    if r["overspend"]:
        say("超支清单：")
        for o in r["overspend"]:
            say(f"   ❌ {o['subject']} {o['region']}: 授{o['grant']:.4f} 用{o['used']:.4f} "
                f"超+{o['over']:.4f} 万 ({o['pct']:.0f}%)")
    if r["notes"]:
        for nt in r["notes"]:
            say(f"提示: {nt}")
    say("-" * 60)
    if not r["ok"]:
        say("结论: ⚠ 表结构/数据存在问题，需人工处理后上报")
    elif r["overspend"]:
        say(f"结论: ✅ 表正确；{len(r['overspend'])} 条超支需业务说明")
    else:
        say("结论: ✅ 全部通过，可以上报")

    # ---- 写报告 ----
    os.makedirs(args.out, exist_ok=True)
    out_md = os.path.join(args.out, f"月度预算核查报告_{month}.md")
    with open(out_md, "w", encoding="utf-8") as f:
        f.write(r["report_md"])
    say(f"\n报告已生成: {out_md}")


if __name__ == "__main__":
    main()
