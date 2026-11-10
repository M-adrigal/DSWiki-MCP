#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
batch_compare.py — 批量对比：同一批文档里，哪些规格是一致的，哪些是跑偏的

输入：docx_probe.py 的 --json 与 style_probe.py 的 --json
输出：按文档类型分组的「一致项 / 不一致项」清单

用途：在建画像之前先做体检。一致项才能进画像，不一致项要进 conflicts 等人工裁决。
"""

import sys
import re
import json
import argparse
from collections import defaultdict, Counter


CANON = {
    "需求说明书": "需求说明书",
    "需求规格说明书": "需求说明书",   # 同一文档类型的不同叫法，归并
    "产品功能说明书": "产品功能说明书",
    "详细设计说明书": "详细设计说明书",
    "产品概要设计": "产品概要设计",
    "操作手册": "操作手册",
}


def doc_type_of(name):
    """从文件名推断文档类型：去掉节点前缀与扩展名，再按同义词归一。"""
    n = re.sub(r"\.docx$", "", name, flags=re.I)
    n = re.sub(r"^(区域节点|业务节点|运行状态监测中心|数据空间)_?", "", n)
    n = re.sub(r"(_备份.*|_V?\d+(\.\d+)*$|V\d+(\.\d+)*$)", "", n)
    n = n.strip("_")
    return CANON.get(n, n or name)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--format-json", required=True)
    ap.add_argument("--style-json")
    args = ap.parse_args()

    fmt = json.load(open(args.format_json, encoding="utf-8"))
    sty = {}
    if args.style_json:
        for s in json.load(open(args.style_json, encoding="utf-8")):
            sty[s["file"]] = s["data"]

    groups = defaultdict(list)
    for d in fmt["docs"]:
        groups[doc_type_of(d["file"])].append(d)

    print(f"共 {len(fmt['docs'])} 份文档，归为 {len(groups)} 类\n")
    for g in sorted(groups, key=lambda k: -len(groups[k])):
        print(f"{'='*70}\n【{g}】{len(groups[g])} 份")
        for d in groups[g]:
            print(f"   - {d['file']}（{d['paragraphs']} 段）")
        print()

        # ---- 页面
        papers = {d["page"].get("纸张") for d in groups[g]}
        margins = {d["page"].get("页边距") for d in groups[g]}
        print(f"  纸张：{'一致 ' + list(papers)[0] if len(papers)==1 else '不一致 → ' + ' ／ '.join(map(str,papers))}")
        print(f"  页边距：{'一致' if len(margins)==1 else '不一致 → ' + ' ／ '.join(map(str,margins))}")

        # ---- 核心样式
        for key in ("Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4"):
            vals = []
            for d in groups[g]:
                s = d["styles"].get(key)
                if not s:
                    vals.append(None)
                else:
                    sp = s["spec"]
                    vals.append((sp["font_eastasia"], sp["size_cn"], sp["line_spacing"], sp["indent"]))
            present = [v for v in vals if v]
            if not present:
                continue
            uniq = Counter(present)
            if len(uniq) == 1:
                f, sz, ln, ind = present[0]
                print(f"  {key}：一致 → {f} {sz}，{ln}，{ind}")
            else:
                desc = "；".join(
                    f"{f} {sz}/{ln}/{ind}" + (f"（{n}份）" if n > 1 else "")
                    for (f, sz, ln, ind), n in uniq.most_common())
                missing = len(vals) - len(present)
                extra = f"，另有 {missing} 份无此样式" if missing else ""
                print(f"  {key}：不一致 → {desc}{extra}")

        # ---- 章节深度与标题数
        depths = [max([h["level"] for h in d["outline"]], default=0) for d in groups[g]]
        hcounts = [len(d["outline"]) for d in groups[g]]
        print(f"  标题层级深度：{'一致 ' + str(depths[0]) + ' 级' if len(set(depths))==1 else '不一致 → ' + str(depths)}")
        print(f"  标题数量：{hcounts}")

        # ---- 编号体系
        schemes = []
        for d in groups[g]:
            ns = d.get("numbering_scheme", {})
            schemes.append(tuple(sorted((k, v["pattern"]) for k, v in ns.items())))
        print(f"  编号体系：{'一致' if len(set(schemes))==1 else '不一致（' + str(len(set(schemes))) + ' 种）'}")

        # ---- 文风
        st = [sty[d["file"]] for d in groups[g] if d["file"] in sty]
        if st:
            def rng(k, fmt_str="{:.1f}"):
                v = [x[k] for x in st]
                return f"{fmt_str.format(min(v))}–{fmt_str.format(max(v))}"
            print(f"  平均句长：{rng('平均句长')} 字 ｜ 截尾均值：{rng('截尾平均句长')} 字"
                  f" ｜ 技术词占比：{rng('技术词占比','{:.1%}')}"
                  f" ｜ 硬性要求占比：{rng('硬性要求占比','{:.1%}')}")
            print(f"  列表化：{rng('列表化段落占比','{:.0%}')} ｜ "
                  f"面向人群：{'一致' if len({x['面向人群判定'] for x in st})==1 else '不一致'}")
            bad = [d["file"] for d, x in zip(groups[g], st) if x["指标可信度"] != "高"]
            if bad:
                print(f"  ⚠️ 指标可信度低：{'、'.join(bad)}（存在超长无标点块，均值不可信）")
        print()


if __name__ == "__main__":
    sys.exit(main())
