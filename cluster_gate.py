#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cluster_gate.py — 闸门③ 聚类分流

判定「这一堆文档到底是几类」，决定拆几个画像。

十维特征（只看客观指标，不看文件名）：
  排版 5 维：纸张、正文字体、正文字号、行距、首行缩进
  结构 3 维：标题字号序列、层级深度、章节构成（标题语义类别 Jaccard 相似度）
  文风 2 维：术语暴露度、列表化程度（分档比较）

关键设计：
  - 缺失 ≠ 不同。一方没有某层级，属于完整性差异（闸门①已处理），不计为模板差异。
  - 章节构成必须扫全部层级，不能只看 H1——交维这几份文档都只有 1 个一级标题，
    只看 H1 会让相似度恒为 1.0。
  - 差异 2 维不自动合并。错合不可逆，错分只是多一次人工确认。

判定规则（看差异跨越几个维度，不看差异大小）：
  ≤1 维 → 噪声，自动强合并（同一子型）
  =2 维 → 存疑，不合并，挂起交人工裁决
  ≥3 维 → 另一套模板，拆子画像

文件名只用于事后核对，不参与聚类。

聚类方式：全链接（complete linkage）——只有簇间任意两份的差异都不超阈值才合并，
避免「A像B、B像C、A不像C」的链式误合并。

用法：
  python3 cluster_gate.py --format-json all-format.json --style-json all-style.json \
                          --quality-json quality.json --json clusters.json
  不给 --quality-json 时，拒收文档也会参与聚类（不建议）。
"""

import sys
import re
import json
import argparse
import itertools

# 文档类型同义词表（仅用于事后核对，不参与聚类）
CANON = {
    "需求说明书": "需求说明书", "需求规格说明书": "需求说明书",
    "产品功能说明书": "产品功能说明书", "详细设计说明书": "详细设计说明书",
    "产品概要设计": "产品概要设计", "操作手册": "操作手册",
}

DIM_NAMES = ["纸张", "正文字体", "正文字号", "行距", "首行缩进", "标题字号序列",
             "术语暴露度", "列表化程度", "层级深度", "章节构成"]

# 章节语义关键词表：把一级标题归到语义类别，用于比较「文档结构像不像」
SEC_KEYS = {
    "前言": ["前言", "概述", "引言", "编制目的", "编写目的", "背景"],
    "范围": ["范围", "适用"],
    "功能需求": ["功能", "需求"],
    "非功能": ["非功能", "性能", "安全", "可靠", "并发"],
    "接口": ["接口", "API", "对接", "联调"],
    "数据": ["数据", "表结构", "字段", "模型", "字典"],
    "操作": ["流程", "步骤", "操作", "使用", "指引", "手册"],
    "术语": ["术语", "缩略", "定义", "名词"],
    "附录": ["附录", "参考", "引用"],
    "管理": ["管理", "配置", "维护", "运维"],
    "设计": ["设计", "架构", "方案", "部署"],
}


def section_keys(outline):
    """
    标题 → 语义类别集合。
    注意：不能只看一级标题。交维这几份文档都只有 1 个一级标题，结构信息全在
    H2–H5 里，只看 H1 会让所有文档的相似度恒为 1.0，失去区分度。
    """
    keys = set()
    for h in outline or []:
        t = h.get("text", "")
        for k, words in SEC_KEYS.items():
            if any(w in t for w in words):
                keys.add(k)
                break
    return keys


def doc_type_hint(name):
    n = re.sub(r"\.docx$", "", name, flags=re.I)
    n = re.sub(r"^(区域节点|业务节点|运行状态监测中心|数据空间)_?", "", n)
    n = re.sub(r"(_备份.*|_V?\d+(\.\d+)*$|V\d+(\.\d+)*$)", "", n).strip("_")
    return CANON.get(n, n)


def parse_multiple(s):
    """'1.15 倍' → 1.15；'继承' / None → None"""
    if not s or "倍" not in s:
        return None
    m = re.search(r"([\d.]+)", s)
    return float(m.group(1)) if m else None


def parse_indent(s):
    """'首行缩进 2 字符' → 2.0；'无缩进' → 0.0；其它 → None"""
    if not s:
        return None
    if "无缩进" in s:
        return 0.0
    m = re.search(r"([\d.]+)\s*字符", s)
    return float(m.group(1)) if m else None


def bucket_tech(v):
    if v is None:
        return None
    if v < 0.15:
        return "业务向(<15%)"
    if v < 0.5:
        return "中性(15-50%)"
    return "技术向(>50%)"


def bucket_list(v):
    if v is None:
        return None
    if v < 0.10:
        return "低(<10%)"
    if v < 0.30:
        return "中(10-30%)"
    return "高(>30%)"


def extract_features(doc, style):
    st = doc["styles"]
    body = st.get("Normal") or st.get("正文")
    spec = body["spec"] if body else {}

    heads = {}
    for i in range(1, 6):
        k = f"Heading {i}"
        if k in st:
            heads[i] = st[k]["spec"]["size_pt"]

    return {
        "file": doc["file"],
        "纸张": doc["page"].get("纸张"),
        "正文字体": spec.get("font_eastasia"),
        "正文字号": spec.get("size_pt"),
        "行距": parse_multiple(spec.get("line_spacing")),
        "首行缩进": parse_indent(spec.get("indent")),
        "标题字号序列": heads,
        "技术词占比": style.get("技术词占比") if style else None,
        "列表化程度": style.get("列表化段落占比") if style else None,
        "技术词档": bucket_tech(style.get("技术词占比") if style else None),
        "列表化档": bucket_list(style.get("列表化段落占比") if style else None),
        "层级深度": max([h["level"] for h in doc.get("outline", [])], default=0),
        "章节构成": section_keys(doc.get("outline", [])),
        "章节构成列表": sorted(section_keys(doc.get("outline", []))),
    }


def compare(a, b):
    """返回 (差异维度列表, 缺失维度列表)"""
    diffs, missing = [], []

    if a["纸张"] and b["纸张"]:
        if a["纸张"] != b["纸张"]:
            diffs.append(f"纸张（{a['纸张']} vs {b['纸张']}）")
    else:
        missing.append("纸张")

    for key in ("正文字体",):
        if a[key] and b[key]:
            if a[key] != b[key]:
                diffs.append(f"{key}（{a[key]} vs {b[key]}）")
        else:
            missing.append(key)

    if a["正文字号"] and b["正文字号"]:
        if abs(a["正文字号"] - b["正文字号"]) > 0.6:
            diffs.append(f"正文字号（{a['正文字号']:g}pt vs {b['正文字号']:g}pt）")
    else:
        missing.append("正文字号")

    # 行距：任一方为「继承」时无法判定，记为缺失而不是差异
    if a["行距"] is not None and b["行距"] is not None:
        if abs(a["行距"] - b["行距"]) > 0.05:
            diffs.append(f"行距（{a['行距']:g} 倍 vs {b['行距']:g} 倍）")
    else:
        missing.append("行距")

    if a["首行缩进"] is not None and b["首行缩进"] is not None:
        if abs(a["首行缩进"] - b["首行缩进"]) > 0.3:
            diffs.append(f"首行缩进（{a['首行缩进']:g} vs {b['首行缩进']:g} 字符）")
    else:
        missing.append("首行缩进")

    # 标题字号：只比双方都有的层级；一方缺某级属于完整性差异，不算模板差异
    common = set(a["标题字号序列"]) & set(b["标题字号序列"])
    if common:
        bad = [lv for lv in sorted(common)
               if abs(a["标题字号序列"][lv] - b["标题字号序列"][lv]) > 0.6]
        if bad:
            lv = bad[0]
            diffs.append(f"标题字号（H{lv} {a['标题字号序列'][lv]:g}pt vs "
                         f"{b['标题字号序列'][lv]:g}pt）")
        if set(a["标题字号序列"]) != set(b["标题字号序列"]):
            missing.append("标题层级完整性")
    else:
        missing.append("标题字号序列")

    if a["技术词档"] and b["技术词档"]:
        if a["技术词档"] != b["技术词档"]:
            diffs.append(f"术语暴露度（{a['技术词档']} vs {b['技术词档']}）")
    else:
        missing.append("术语暴露度")

    if a["列表化档"] and b["列表化档"]:
        if a["列表化档"] != b["列表化档"]:
            diffs.append(f"列表化程度（{a['列表化档']} vs {b['列表化档']}）")
    else:
        missing.append("列表化程度")

    # 层级深度：缺失层级（深度 0）不是「不同」而是「结构不可读」，单独记缺失
    if a["层级深度"] and b["层级深度"]:
        if a["层级深度"] != b["层级深度"]:
            diffs.append(f"层级深度（{a['层级深度']} 级 vs {b['层级深度']} 级）")
    else:
        missing.append("层级深度（一方无标题）")

    # 章节构成：一级标题的语义类别 Jaccard 相似度
    sa, sb = a["章节构成"], b["章节构成"]
    if sa and sb:
        jac = len(sa & sb) / len(sa | sb)
        if jac < 0.34:
            diffs.append(f"章节构成（相似度 {jac:.0%}：{'+'.join(sorted(sa))} vs "
                         f"{'+'.join(sorted(sb))}）")
    else:
        missing.append("章节构成")

    return diffs, missing


def cluster(features, strong=1, weak=2):
    """全链接凝聚聚类。strong=强合并阈值，weak=弱合并阈值。"""
    n = len(features)
    pairs = {}
    for i, j in itertools.combinations(range(n), 2):
        d, m = compare(features[i], features[j])
        pairs[(i, j)] = (d, m)

    def max_diff(a_idx, b_idx):
        """两簇之间的最大差异维度数（全链接）"""
        best = 0
        detail = None
        for i in a_idx:
            for j in b_idx:
                key = (i, j) if i < j else (j, i)
                d = pairs[key][0]
                if len(d) > best:
                    best = len(d)
                    detail = d
        return best, detail

    clusters = [[i] for i in range(n)]
    merge_log = []
    while True:
        best = None
        for ci, cj in itertools.combinations(range(len(clusters)), 2):
            d, detail = max_diff(clusters[ci], clusters[cj])
            if d <= strong and (best is None or d < best[0]):
                best = (d, ci, cj, detail)
        if best is None:
            break
        d, ci, cj, detail = best
        kind = "强合并（差异≤1维，判为噪声）"
        merge_log.append({
            "类型": kind, "维度数": d, "差异": detail,
            "并入": [features[i]["file"] for i in clusters[cj]],
            "到": [features[i]["file"] for i in clusters[ci]],
        })
        clusters[ci] = clusters[ci] + clusters[cj]
        clusters.pop(cj)

    # 未合并但差异只有 2 维的 → 存疑对，交人工裁决（不自动合并）
    pending = []
    for ci, cj in itertools.combinations(range(len(clusters)), 2):
        d, detail = max_diff(clusters[ci], clusters[cj])
        if d <= weak:
            pending.append({
                "维度数": d, "差异": detail,
                "A": [features[i]["file"] for i in clusters[ci]],
                "B": [features[i]["file"] for i in clusters[cj]],
            })

    # 每组内最大差异（判断组内是否存疑）
    out = []
    for c in clusters:
        worst = 0
        worst_detail = None
        for i, j in itertools.combinations(c, 2):
            key = (i, j) if i < j else (j, i)
            d = pairs[key][0]
            if len(d) > worst:
                worst = len(d)
                worst_detail = d
        out.append({
            "docs": [features[i]["file"] for i in c],
            "idx": c,
            "max_diff": worst,
            "max_detail": worst_detail,
            # 单份文档没有可比对象，说它「同型」是无意义的——组内一致性没法判断。
            # 这类组的真实含义是「样本不足，建出来的画像会过拟合」。
            "status": ("单份（样本不足，无法判断一致性）" if len(c) == 1 else
                       "同型（差异≤1维）" if worst <= strong else
                       "存疑（组内差异较大）"),
        })
    return out, pairs, merge_log, pending


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--format-json", required=True)
    ap.add_argument("--style-json", required=True)
    ap.add_argument("--quality-json", help="闸门① 的评分结果；拒收文档将被排除，不参与聚类")
    ap.add_argument("--json", dest="json_out")
    args = ap.parse_args()

    fmt = json.load(open(args.format_json, encoding="utf-8"))
    sty = {s["file"]: s["data"] for s in json.load(open(args.style_json, encoding="utf-8"))}

    docs = fmt["docs"]
    excluded = []
    if args.quality_json:
        quality = {q["file"]: q for q in json.load(open(args.quality_json, encoding="utf-8"))}
        keep = []
        for d in docs:
            q = quality.get(d["file"])
            if q and ("拒收" in q.get("verdict", "") or "不可用" in q.get("verdict", "")):
                excluded.append((d["file"], q["score"], q["verdict"]))
            else:
                keep.append(d)
        docs = keep
    if excluded:
        print("=" * 96)
        print("已被闸门① 拒收，不参与聚类")
        print("=" * 96)
        for f, s, v in excluded:
            print(f"  {f}  {s} 分  {v}")
        print()

    features = [extract_features(d, sty.get(d["file"])) for d in docs]

    print("=" * 96)
    print("特征向量（不参与聚类的原始指标）")
    print("=" * 96)
    print(f"{'文档':34s}{'纸张':16s}{'字体':8s}{'正文':8s}{'行距':6s}{'缩进':6s}{'技术词':9s}{'列表化'}")
    print("-" * 96)
    for f in features:
        size = f"{f['正文字号']:g}pt" if f["正文字号"] else "-"
        line = f"{f['行距']:g}倍" if f["行距"] is not None else "-"
        ind = f"{f['首行缩进']:g}字" if f["首行缩进"] is not None else "-"
        tech = f"{f['技术词占比']*100:.1f}%" if f["技术词占比"] is not None else "-"
        lst = f"{f['列表化程度']*100:.0f}%" if f["列表化程度"] is not None else "-"
        print(f"{f['file'][:34]:34s}{(f['纸张'] or '-')[:15]:16s}"
              f"{(f['正文字体'] or '-')[:7]:8s}{size:8s}{line:6s}{ind:6s}{tech:9s}{lst}")

    groups, pairs, merges, pending = cluster(features)

    print("\n" + "=" * 96)
    print("两两差异维度数（≥3 即判为另一套模板）")
    print("=" * 96)
    names = [f["file"][:14] for f in features]
    print(" " * 16 + "".join(f"{n:16s}" for n in names))
    for i, f in enumerate(features):
        row = f"{names[i]:16s}"
        for j in range(len(features)):
            if i == j:
                row += f"{'—':16s}"
            else:
                key = (i, j) if i < j else (j, i)
                row += f"{len(pairs[key][0]):<16d}"
        print(row)

    print("\n" + "=" * 96)
    print(f"聚类结果：{len(groups)} 组")
    print("=" * 96)
    for gi, g in enumerate(sorted(groups, key=lambda x: -len(x["docs"])), 1):
        hints = {doc_type_hint(d) for d in g["docs"]}
        print(f"\n组 {gi}（{len(g['docs'])} 份）— 组内最大差异 {g['max_diff']} 维 · {g['status']}")
        for d in g["docs"]:
            print(f"    - {d}   [文件名推断类型：{doc_type_hint(d)}]")
        if g["max_detail"]:
            print(f"    交错明细：{'；'.join(g['max_detail'])}")
        if len(hints) > 1:
            print(f"    ⚠️ 组内含 {len(hints)} 种文件名类型：{'、'.join(hints)} → 需人工确认")

    print("\n" + "=" * 96)
    print("合并日志（谁并到了谁，为什么）")
    print("=" * 96)
    for m in merges:
        print(f"  [{m['类型']}] {m['维度数']} 维差异 → 将 {len(m['并入'])} 份并入")
        if m["差异"]:
            print(f"       差异：{'；'.join(m['差异'])}")

    print("\n" + "=" * 96)
    print("未合并但差异仅 2 维的「存疑对」——不自动合并，交人工裁决")
    print("=" * 96)
    if not pending:
        print("  无")
    for p in pending:
        print(f"\n  {p['维度数']} 维差异：")
        print(f"    A：{'、'.join(p['A'])}")
        print(f"    B：{'、'.join(p['B'])}")
        for x in p["差异"]:
            print(f"    · {x}")

    print("\n" + "=" * 96)
    print("核对：同名文档类型内部，模板是否统一（拆开≠聚类出错，而是该类型真的有两套模板）")
    print("=" * 96)
    by_hint = {}
    for gi, g in enumerate(sorted(groups, key=lambda x: -len(x["docs"])), 1):
        for d in g["docs"]:
            by_hint.setdefault(doc_type_hint(d), []).append(gi)
    for h, gs in by_hint.items():
        n = len(set(gs))
        mark = "✅ 内部统一" if n == 1 else f"⚠️ 内部有 {n} 套模板，不统一"
        print(f"  {h:16s} {mark}（组 {sorted(set(gs))}）")

    if args.json_out:
        clean_features = []
        for f in features:
            g = dict(f)
            g.pop("章节构成", None)          # set 不可序列化，保留列表版
            clean_features.append(g)
        with open(args.json_out, "w", encoding="utf-8") as fh:
            json.dump({"features": clean_features,
                       "groups": [{k: v for k, v in g.items() if k != "idx"} for g in groups],
                       "pending": pending,
                       "merge_log": merges}, fh, ensure_ascii=False, indent=2)
        print(f"\n[JSON] {args.json_out}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
