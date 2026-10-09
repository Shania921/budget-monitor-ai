# -*- coding: utf-8 -*-
"""
AI 归因简报生成器
=====================================================================
输入：一键核查脚本的运行结果（已算好的结构化指标 —— AI 不做任何计算）
输出：超支科目归因分析 + 管理简报初稿（Markdown），供人工复核后发布

两种工作模式：
  1. LLM 模式：仓库根目录存在 config.json 且已填 api_key 时，
     调用 OpenAI 兼容接口（/chat/completions）生成归因简报
  2. 演示模式：无 config.json 或未填 key 时，用规则模板生成
     （归因维度菜单与 LLM 提示词一致，保证离线也能完整跑通流程）

防幻觉设计：
  - 只把 Python 算好的指标（科目/区域/授予/已用/超额/超幅/执行率）
    交给模型，原始明细不进提示词，模型只做解释、不做计算
  - 归因维度限定固定菜单：进度集中(时点错配)/业务量增加/单价上升/
    一次性项目/统计口径差异，不允许自由发挥
  - 输出必须带"待人工复核"标记
"""
import os
import sys
import json
import datetime
import urllib.request
import urllib.error

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(REPO_ROOT, "config.json")

REASON_MENU = ["进度集中（时点错配）", "业务量增加", "单价上升",
               "一次性项目", "统计口径差异"]

SYSTEM_PROMPT = (
    "你是一家企业区域财务BP的预算分析助理。你会收到一份当月预算执行的结构化指标"
    "（JSON，全部由程序计算，你不做任何计算）。请对每一条超支科目生成归因分析初稿。\n"
    "要求：\n"
    "1. 归因必须且只能从以下维度菜单中选择并组合："
    + "、".join(REASON_MENU) + "；\n"
    "2. 每条超支给出：归因判断、支撑逻辑（基于给定数字）、建议动作（问询/管控/调剂）；\n"
    "3. 语气克制、面向管理层，不夸大；数字必须引用给定的值，不得自行推算新数字；\n"
    "4. 输出为 Markdown，含三个部分：## 一、总体执行情况；## 二、超支归因（逐项）；"
    "## 三、建议动作；\n"
    "5. 文末注明：本简报由 AI 生成初稿，待人工复核后发布。"
)


def load_config():
    if not os.path.isfile(CONFIG_PATH):
        return None
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        if cfg.get("api_key"):
            return cfg
    except (json.JSONDecodeError, OSError):
        pass
    return None


def build_metrics(result):
    """从核查结果提取喂给模型/模板的指标（全部程序算好）"""
    s = result["summary"]
    rate = (s["used"] / s["grant"] * 100) if s["grant"] else 0
    return {
        "月份": result["month"],
        "授予总额_万": s["grant"],
        "已使用总额_万": s["used"],
        "执行率_百分比": round(rate, 1),
        "超支条目": [
            {"科目": o["subject"], "区域": o["region"], "授予_万": o["grant"],
             "已用_万": o["used"], "超额_万": o["over"], "超幅_百分比": o["pct"]}
            for o in result["overspend"]
        ],
        "尾差_元": result["tail"]["diff_yuan"],
        "已解释豁免_元": result["tail"]["exempt_yuan"],
    }


def call_llm(cfg, metrics):
    """调用 OpenAI 兼容 /chat/completions；失败返回 None"""
    url = cfg.get("base_url", "https://api.openai.com/v1").rstrip("/") + "/chat/completions"
    body = json.dumps({
        "model": cfg.get("model", "gpt-4o-mini"),
        "temperature": 0.2,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": "当月预算执行指标如下，请生成归因简报初稿：\n"
             + json.dumps(metrics, ensure_ascii=False, indent=2)},
        ],
    }).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, method="POST",
        headers={"Content-Type": "application/json",
                 "Authorization": "Bearer " + cfg["api_key"]})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data["choices"][0]["message"]["content"].strip()
    except (urllib.error.URLError, KeyError, IndexError, TimeoutError) as e:
        print(f"[提示] LLM 调用失败（{e}），改用规则模板生成。")
        return None


def rule_based_report(metrics):
    """演示模式：规则模板归因（维度菜单与 LLM 提示词一致）"""
    m = metrics
    L = []
    L.append("## 一、总体执行情况")
    L.append("")
    L.append(f"{m['月份']}授予总额 {m['授予总额_万']:.1f} 万元，已使用 {m['已使用总额_万']:.1f} 万元，"
             f"执行率 {m['执行率_百分比']:.1f}%。")
    if m["超支条目"]:
        L.append(f"共 {len(m['超支条目'])} 条科目×区域超支，合计超额 "
                 f"{sum(x['超额_万'] for x in m['超支条目']):.1f} 万元，明细如下：")
        L.append("")
        L.append("| 科目 | 区域 | 授予(万) | 已用(万) | 超额(万) | 超幅 |")
        L.append("|---|---|---|---|---|---|")
        for x in m["超支条目"]:
            L.append(f"| {x['科目']} | {x['区域']} | {x['授予_万']:.1f} | "
                     f"{x['已用_万']:.1f} | +{x['超额_万']:.1f} | {x['超幅_百分比']:.0f}% |")
    else:
        L.append("本月无超支科目。")
    L.append("")
    L.append("## 二、超支归因（逐项，规则模板生成）")
    L.append("")
    for x in m["超支条目"]:
        pct = x["超幅_百分比"]
        subj = x["科目"]
        if "招待" in subj:
            guess = "业务量增加（拜访/接待场次上升）或进度集中（时点错配：活动集中在当月验收）"
            action = "要求省区提供当月招待明细台账，超支部分补办审批或调增额度"
        elif "市场推广" in subj:
            guess = "一次性项目（当季市场活动集中投放）或单价上升"
            action = "核对活动结算单价与合同，评估剩余季度额度是否需调剂"
        elif "专业费用" in subj:
            guess = "业务量增加（在途项目集中结算）或统计口径差异（跨期费用当月入账）"
            action = "按预算项目拆分对账，确认是否有跨期项目应摊入后续月份"
        else:
            guess = "业务量增加或单价上升，需省区补充说明"
            action = "发问询单，要求 3 个工作日内回复原因及改善措施"
        L.append(f"### {subj} @ {x['区域']}（超 +{x['超额_万']:.1f} 万，超幅 {pct:.0f}%）")
        L.append("")
        L.append(f"- **归因初判**：{guess}。")
        L.append(f"- **支撑逻辑**：授予 {x['授予_万']:.1f} 万、已用 {x['已用_万']:.1f} 万，"
                 f"超出 {x['超额_万']:.1f} 万；超幅 {pct:.0f}% "
                 f"{'偏高，优先核查业务量' if pct >= 25 else '温和，优先核对口径与跨期'}。")
        L.append(f"- **建议动作**：{action}。")
        L.append("")
    L.append("## 三、建议动作汇总")
    L.append("")
    L.append("1. 上述超支项由责任省区在 3 个工作日内反馈归因说明；")
    L.append("2. 超幅 ≥ 25% 的项目同步抄送区域负责人；")
    L.append("3. 下月滚动预测时复核相关科目额度是否需调剂。")
    L.append("")
    L.append("---")
    L.append("本简报由规则模板生成（演示模式）。配置 config.json 中的 api_key 后，"
             "相同指标将交由大模型生成归因初稿。**所有结论均待人工复核后发布。**")
    return "\n".join(L)


def generate(result, out_dir=None):
    """入口：传入 monthly_budget_check.run_check 的结果 dict，生成简报文件"""
    metrics = build_metrics(result)
    cfg = load_config()
    mode = "LLM"
    if cfg:
        content = call_llm(cfg, metrics)
        if content is None:
            mode = "规则模板"
    else:
        mode = "规则模板"
    if mode == "规则模板":
        content = rule_based_report(metrics)

    out_dir = out_dir or os.path.join(REPO_ROOT, "reports")
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"AI归因简报_{result['month']}.md")
    header = (f"# AI 归因简报 · {result['month']}\n\n"
              f"> 生成模式：{mode} | 生成时间：{datetime.datetime.now():%Y-%m-%d %H:%M}\n"
              f"> 指标来源：monthly_budget_check.py 独立重算（AI 不参与计算）\n\n")
    with open(path, "w", encoding="utf-8") as f:
        f.write(header + content + "\n")
    print(f"[OK] AI 归因简报已生成（{mode}模式）: {path}")
    return path


if __name__ == "__main__":
    from monthly_budget_check import run_check, locate_latest
    _, path, _ = locate_latest()
    if not path or len(sys.argv) > 1:
        path = sys.argv[1] if len(sys.argv) > 1 else None
    if not path:
        print("[错误] 未找到主表，用法: python ai_report.py <主表.xlsx>")
        sys.exit(1)
    generate(run_check(path))
