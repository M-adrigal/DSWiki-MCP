#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pptx_quality_gate.py — PPT 准入评分器

Word 的质量门看的是「样式用没用对」，PPT 的问题不一样：
PPT 最常见、也最致命的坏样本是「每页贴一整张图片」——
视觉上看不出毛病，但形状树里只有一个 picture，架构图怎么画的、说明文字怎么排的，
一个字都提取不到。所以 PPT 的质量门核心指标是「原生可编辑率」。

评分维度（满分 100）：
  原生可编辑率 40   原生形状/文本框/表格/连接线/组合 占全部形状的比例
  文本量       25   全部页面的字符总数
  页数         15   样本量
  版式多样性   10   用了多少种版式（全用同一个版式说明是套壳）
  主题完整性   10   是否有配色方案与字体方案

另设前置门槛：页数 < 5 直接判「样本不足（不可用）」。

用法：
  python3 pptx_quality_gate.py a.pptx b.pptx --json out.json
"""

import sys
import os
import json
import argparse

from pptx_probe import extract_pptx

EDITABLE_KINDS = ("原生形状", "文本框", "表格", "连接线/线条", "组合", "占位符", "SmartArt")


def evaluate(path):
    r = extract_pptx(path)
    shapes = r.get("形状类型 Total", {})
    total = sum(shapes.values())
    if total == 0:
        return None

    pics = shapes.get("图片", 0)
    editable = sum(v for k, v in shapes.items() if k in EDITABLE_KINDS)
    ratio = editable / total

    text_total = sum(p.get("文本量", 0) for p in r.get("页面明细", []))
    pages = r["页数"]
    layouts = len(r.get("版式分布", {}))
    colors = r.get("主题", {}).get("colors", {})
    fonts = r.get("主题", {}).get("fonts", {})
    has_theme = bool(colors) and bool(fonts)
    has_ladder = bool(r.get("母版字号阶梯"))

    insufficient = []          # 用列表保序，末尾统一去重

    deductions = []

    # 原生可编辑率 40：>=80% 不扣，0% 扣满
    d1 = 0 if ratio >= 0.8 else round((0.8 - ratio) / 0.8 * 40)
    if d1:
        deductions.append((d1, f"原生可编辑率仅 {ratio:.0%}"
                               f"（{editable}/{total} 个形状是原生内容，{pics} 个是图片）"))

    # 文本量 25：>=500 字不扣
    # 样本不足时（页数或文本量任一不达标）分数本来就不给，扣分只写进不足原因，
    # 否则「3 页」会被页数、文本量、版式各罚一次，同一个事实罚三遍。
    sample_ok = pages >= 5 and text_total >= 500
    d2 = 0 if text_total >= 500 else round((500 - text_total) / 500 * 25)
    if d2:
        if sample_ok:
            deductions.append((d2, f"全部页面合计仅 {text_total} 字（下限 500）"))
        else:
            insufficient.append(f"全部页面合计仅 {text_total} 字（下限 500）")
            d2 = 0

    # 页数 15
    d3 = 15 if pages < 5 else (8 if pages < 10 else 0)
    if d3:
        if pages < 5:
            insufficient.append(f"仅 {pages} 页（下限 5 页）")
            d3 = 0                     # 已进「样本不足」，不重复扣
        else:
            deductions.append((d3, f"仅 {pages} 页，样本量偏少"))

    # 版式多样性 10
    # 页数不足时版式必然少，再扣一次等于同一个事实罚两遍（3 页 1 版式会被扣 25 分）。
    # 版式是「设计手法」，页数是「样本量」，样本不足时只在不足原因里提示，不重复扣分。
    d4 = 0
    if pages >= 5:
        d4 = 10 if layouts <= 1 else (5 if layouts <= 2 else 0)
        if d4:
            deductions.append((d4, f"只用了 {layouts} 种版式"
                                   + ("，像是套壳而非设计" if layouts <= 1 else "，版式偏单一")))
    elif layouts <= 1:
        # 不改分数，只在不足原因里说清楚，避免把"少"读成"差"
        insufficient.append(f"{pages} 页只用了 {layouts} 种版式；"
                            "页数太少时版式必然单一，这一项不参与评分")

    # 主题完整性 10
    d5 = 0
    if not colors:
        d5 += 5
        deductions.append((5, "无主题配色方案"))
    if not fonts:
        d5 += 5
        deductions.append((5, "无主题字体方案"))

    score = 100 - sum(d for d, _ in deductions)

    if insufficient:
        verdict = "样本不足（不可用）"
    elif score >= 75:
        verdict = "准入"
    elif score >= 55:
        verdict = "降权准入（不计入『一致项』）"
    else:
        verdict = "拒收（进待修复队列）"

    # 去重保序：同一件事（如页数不足）只出现一次
    insufficient = list(dict.fromkeys(insufficient))

    return {
        "file": os.path.basename(path),
        "score": (None if insufficient else score),   # 样本不足时分数无意义
        "verdict": verdict,
        "页数": pages,
        "形状总数": total,
        "图片数": pics,
        "原生可编辑率": round(ratio, 3),
        "文本量": text_total,
        "版式数": layouts,
        "有主题": has_theme,
        "有字号阶梯": has_ladder,
        "不足原因": insufficient,
        "扣分明细": [{"扣分": d, "原因": x} for d, x in deductions],
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--json", dest="json_out")
    args = ap.parse_args()

    rows = []
    for f in args.files:
        if not os.path.exists(f) or not f.lower().endswith(".pptx"):
            continue
        try:
            r = evaluate(f)
            if r:
                rows.append(r)
                print(f"[完成] {r['file']}  {r['score']} 分  {r['verdict']}", file=sys.stderr)
        except Exception as e:  # noqa: BLE001
            print(f"[失败] {os.path.basename(f)}：{type(e).__name__} {e}", file=sys.stderr)

    if not rows:
        return 1

    print(f"\n{'文档':34s}{'分数':>5s}{'原生率':>7s}{'页数':>5s}{'文本量':>7s}  判定")
    print("-" * 92)
    for r in sorted(rows, key=lambda x: -x["score"]):
        print(f"{r['file'][:34]:34s}{r['score']:>5d}{r['原生可编辑率']*100:>6.0f}%"
              f"{r['页数']:>5d}{r['文本量']:>7d}  {r['verdict']}")

    print("\n扣分明细")
    print("-" * 92)
    for r in sorted(rows, key=lambda x: -x["score"]):
        print(f"\n■ {r['file']}（{r['score']} 分）")
        if r.get("不足原因"):
            print("  ⛔ " + "；".join(r["不足原因"]))
        if not r["扣分明细"]:
            print("  无扣分")
        for d in r["扣分明细"]:
            print(f"  -{d['扣分']:>2d}  {d['原因']}")

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as fh:
            json.dump(rows, fh, ensure_ascii=False, indent=2)
        print(f"\n[JSON] {args.json_out}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
