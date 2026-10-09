# -*- coding: utf-8 -*-
"""
一键运行入口（run.bat 调用）
=====================================================================
流程：确保演示数据存在 -> 预算核查 -> AI 归因简报
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

from generate_fake_data import DATA_DIR, main as gen_main          # noqa: E402
from monthly_budget_check import run_check, locate_latest          # noqa: E402
import ai_report                                                   # noqa: E402


def ensure_demo_data():
    _, path, _ = locate_latest()
    if path:
        return path
    print("[初始化] 未找到主表，生成演示数据 ...")
    gen_main()
    _, path, _ = locate_latest()
    return path


def main():
    print("=" * 62)
    print(" 区域预算执行监控与 AI 归因分析助手（脱敏演示版）")
    print("=" * 62)
    path = ensure_demo_data()
    if not path:
        print("[错误] 未找到可核查的主表文件。")
        sys.exit(1)

    print(f"\n[1/2] 预算核查（独立重算，不信任表内公式缓存）...")
    result = run_check(path)
    print(f"      月份: {result['month']}")
    for it in result["check_items"]:
        print(f"      {it['name']}: {it['detail']}")
    if result["overspend"]:
        print("      超支清单:")
        for o in result["overspend"]:
            print(f"         - {o['subject']} @ {o['region']}: "
                  f"授 {o['grant']:.1f} / 用 {o['used']:.1f} / "
                  f"超 +{o['over']:.1f} 万 ({o['pct']:.0f}%)")

    print(f"\n[2/2] AI 归因简报 ...")
    ai_report.generate(result)

    print("\n" + "=" * 62)
    if not result["ok"]:
        print(" 结论: 表结构/数据存在问题，需人工处理后上报")
    elif result["overspend"]:
        print(f" 结论: 表正确；{len(result['overspend'])} 条超支待业务说明（见 AI 归因简报）")
    else:
        print(" 结论: 全部通过，可以上报")
    print(" 报告目录: reports/")
    print("=" * 62)


if __name__ == "__main__":
    main()
