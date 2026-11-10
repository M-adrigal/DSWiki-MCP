#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pptx_probe.py — PPT 视觉风格与版式结构提取器

提取三层信息：
  A. 主题层：配色方案（含十六进制色值）、字体方案（中英文）、母版字号阶梯
  B. 页面层：页面尺寸比例、版式使用分布、页面类型识别（封面/目录/章节/内容/结尾）
  C. 图形层：**架构图/流程图是怎么画的** —— 区分三种来源
       1) 原生形状：圆角矩形/矩形/箭头/连接线，可完整还原画法
       2) SmartArt：可读层级与文字，视觉呈现不可精确还原
       3) 图片：只能记录位置、尺寸、占版面比例（画法不可提取，需视觉识别）

用法：
  python3 pptx_probe.py a.pptx
  python3 pptx_probe.py a.pptx b.pptx --out report.md --json out.json
"""

import sys
import os
import re
import json
import zipfile
import argparse
from collections import Counter
from datetime import datetime

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def q(tag):
    pfx, local = tag.split(":")
    return "{%s}%s" % (NS[pfx], local)


# 常见预设形状 → 中文
PRST_CN = {
    "roundRect": "圆角矩形", "rect": "矩形", "ellipse": "椭圆", "flowChartProcess": "流程-处理",
    "flowChartDecision": "流程-判断", "flowChartTerminator": "流程-起止", "flowChartData": "流程-数据",
    "flowChartDocument": "流程-文档", "flowChartPredefinedProcess": "流程-子流程",
    "diamond": "菱形", "chevron": "箭头形", "pentagon": "五边形", "rightArrow": "右箭头",
    "leftArrow": "左箭头", "upArrow": "上箭头", "downArrow": "下箭头",
    "leftRightArrow": "左右箭头", "bentArrow": "折线箭头", "curvedRightArrow": "弧形箭头",
    "cylinder": "圆柱（数据库）", "cloud": "云", "hexagon": "六边形", "triangle": "三角形",
    "line": "直线", "straightConnector1": "直线连接符",
    "bentConnector3": "折线连接符", "curvedConnector3": "弧线连接符",
    "frame": "框架", "plaque": "圆角标签", "can": "圆柱", "cube": "立方体",
    "donut": "环形", "arc": "弧形", "blockArc": "块弧", "chord": "弦形",
    "moon": "月牙", "sun": "太阳", "heart": "心形", "star5": "五角星",
}


def emu_cm(v):
    return round(Emu(v).cm, 2) if v is not None else None


# ---------------------------------------------------------------- 主题

def read_theme(pptx_path):
    """从 zip 内直接读 theme1.xml，拿配色与字体方案。"""
    out = {"colors": {}, "fonts": {}, "name": None}
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = [n for n in z.namelist() if re.match(r"ppt/theme/theme\d+\.xml$", n)]
            if not names:
                return out
            root = etree.fromstring(z.read(sorted(names)[0]))
    except (zipfile.BadZipFile, etree.XMLSyntaxError, KeyError):
        return out

    nm = root.find(q("a:theme") + "/" + q("a:name"))
    if nm is not None:
        out["name"] = nm.get("name")

    cs = root.find(".//" + q("a:clrScheme"))
    if cs is not None:
        label = {"dk1": "深色1（正文）", "lt1": "浅色1（背景）", "dk2": "深色2（标题）",
                 "lt2": "浅色2", "accent1": "主色1", "accent2": "辅色2", "accent3": "辅色3",
                 "accent4": "辅色4", "accent5": "辅色5", "accent6": "辅色6",
                 "hlink": "超链接", "folHlink": "已访问链接"}
        for i, child in enumerate(cs):
            key = etree.QName(child).localname
            val = None
            for sub in child:
                tag = etree.QName(sub).localname
                if tag == "srgbClr":
                    val = "#" + sub.get("val", "").upper()
                elif tag == "sysClr":
                    val = "#" + (sub.get("lastClr") or "").upper()
            out["colors"][label.get(key, key)] = val

    fs = root.find(".//" + q("a:fontScheme"))
    if fs is not None:
        for which, cn in (("majorFont", "标题字体"), ("minorFont", "正文字体")):
            node = fs.find(q("a:" + which))
            if node is None:
                continue
            latin = node.find(q("a:latin"))
            ea = node.find(q("a:ea"))
            out["fonts"][cn] = {
                "西文": latin.get("typeface") if latin is not None else None,
                "中文": ea.get("typeface") if ea is not None else None,
            }
    return out


def read_master_font_sizes(prs):
    """母版 txStyles 的字号阶梯（标题/正文各级）。"""
    out = {}
    try:
        master = prs.slide_masters[0]
        el = master.element
    except (IndexError, AttributeError):
        return out
    for tag, cn in (("titleStyle", "标题样式"), ("bodyStyle", "正文样式"), ("otherStyle", "其他样式")):
        node = el.find(".//" + q("p:txStyles") + "/" + q("p:" + tag))
        if node is None:
            continue
        lvls = []
        for lvl in node.findall(q("a:lvl1pPr")) + node.findall(q("a:lvl2pPr")) + \
                node.findall(q("a:lvl3pPr")):
            dr = lvl.find(q("a:defRPr"))
            if dr is not None and dr.get("sz"):
                lvls.append(int(dr.get("sz")) / 100.0)
        if lvls:
            out[cn] = lvls
    return out


# ---------------------------------------------------------------- 形状分类

def shape_kind(shape):
    """把形状归类到可解释的类别。"""
    raw = shape._element.xml
    if "diagram" in raw and "graphicData" in raw:
        return "SmartArt"
    try:
        st = shape.shape_type
    except (NotImplementedError, AttributeError):
        st = None
    name = (st and str(st)) or ""

    if "PICTURE" in name:
        return "图片"
    if "TABLE" in name:
        return "表格"
    if "CHART" in name:
        return "图表"
    if "GROUP" in name:
        return "组合"
    if "LINE" in name or "FREEFORM" in name:
        return "连接线/线条"
    if "TEXT_BOX" in name:
        return "文本框"
    if "AUTO_SHAPE" in name:
        return "原生形状"
    if "PLACEHOLDER" in name:
        return "占位符"
    return name.split(" ")[0] if name else "其他"


def prst_of(shape):
    """读预设几何名，如 roundRect。"""
    el = shape._element
    for g in el.iter(q("a:prstGeom")):
        return g.get("prst")
    return None


def fill_color(shape):
    try:
        f = shape.fill
        if f.type is not None and str(f.type) == "MSO_FILL_TYPE.SOLID (1)":
            return "#" + str(f.fore_color.rgb)
    except (AttributeError, TypeError, ValueError):
        pass
    # 兜底：从 XML 读 srgbClr
    m = re.search(r'<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"', shape._element.xml)
    return "#" + m.group(1).upper() if m else None


def has_arrow(shape):
    """连接线是否带箭头端点。"""
    x = shape._element.xml
    return ("headEnd" in x and 'type="triangle"' in x) or \
           ("tailEnd" in x and 'type="triangle"' in x)


def text_props(shape):
    """收集形状内所有 run 的字号（pt）与字体名。"""
    sizes, fonts = [], []
    if not getattr(shape, "has_text_frame", False) or shape.has_text_frame is False:
        return sizes, fonts
    try:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size is not None:
                    sizes.append(round(r.font.size.pt, 1))
                # 中文字体在 a:ea，python-pptx 的 font.name 只给 a:latin
                rpr = r._r.find(q("a:rPr"))
                ea = latin = None
                if rpr is not None:
                    e = rpr.find(q("a:ea"))
                    l = rpr.find(q("a:latin"))
                    ea = e.get("typeface") if e is not None else None
                    latin = l.get("typeface") if l is not None else None
                if ea and ea != "+mn-ea" and ea != "+mj-ea":
                    fonts.append(ea)
                elif latin and latin not in ("+mn-lt", "+mj-lt"):
                    fonts.append(latin)
    except (AttributeError, ValueError):
        pass
    return sizes, fonts


def classify_page(idx, total, layout_name, kinds, text_all, prsts, sizes):
    """启发式判断页面类型。"""
    ln = (layout_name or "").lower()
    if idx == 1 and total > 3:
        return "封面"
    if "title slide" in ln or "标题幻灯片" in ln:
        return "封面"
    if idx == total and total > 3:
        return "结尾页"
    if re.match(r"^\s*(目录|agenda|contents)", text_all or "", re.I) and len(text_all) < 500:
        return "目录/导航页"
    if re.fullmatch(r"\s*[\d一二三四五六七八九十]+\s*[、.]?\s*", text_all or ""):
        return "章节过渡页"
    if len(text_all or "") < 40 and (prsts or kinds.get("图片")):
        return "图形/示意页"
    if kinds.get("表格"):
        return "表格页"
    if "section" in ln or "节" in ln:
        return "章节过渡页"
    return "内容页"


DIAGRAM_KINDS = ("原生形状", "连接线/线条", "SmartArt", "图片", "组合")


# ---------------------------------------------------------------- 主提取

def extract_pptx(path):
    prs = Presentation(path)
    theme = read_theme(path)
    master_sizes = read_master_font_sizes(prs)

    w_cm, h_cm = emu_cm(prs.slide_width), emu_cm(prs.slide_height)
    ratio = round(w_cm / h_cm, 3) if w_cm and h_cm else None
    ratio_label = "16:9 宽屏" if ratio and abs(ratio - 16 / 9) < 0.03 else \
                  ("4:3" if ratio and abs(ratio - 4 / 3) < 0.03 else f"自定义（{ratio}）")

    layout_counter = Counter()
    shape_kinds_all = Counter()
    prst_counter = Counter()
    size_counter = Counter()
    font_counter = Counter()
    fill_counter = Counter()
    pages = []

    slides = list(prs.slides)
    for idx, s in enumerate(slides, 1):
        layout = s.slide_layout.name if s.slide_layout is not None else "(无)"
        layout_counter[layout] += 1
        kinds = Counter()
        prsts = Counter()
        sizes = []
        texts = []
        diagram_shapes = []

        for sh in s.shapes:
            k = shape_kind(sh)
            kinds[k] += 1
            shape_kinds_all[k] += 1

            if getattr(sh, "has_text_frame", False):
                try:
                    if sh.text_frame.text:
                        texts.append(sh.text_frame.text)
                except (AttributeError, ValueError):
                    pass
            szs, fnts = text_props(sh)
            for sz in szs:
                sizes.append(sz)
                size_counter[sz] += 1
            for fn in fnts:
                font_counter[fn] += 1

            if k in ("原生形状", "连接线/线条"):
                p = prst_of(sh)
                if p:
                    prsts[p] += 1
                    prst_counter[p] += 1

            fc = fill_color(sh) if k in ("原生形状", "文本框", "占位符") else None
            if fc:
                fill_counter[fc] += 1

            if k in DIAGRAM_KINDS:
                area = 0
                try:
                    if sh.width and sh.height:
                        area = sh.width.cm * sh.height.cm
                except (AttributeError, TypeError):
                    pass
                diagram_shapes.append({
                    "kind": k,
                    "name": getattr(sh, "name", ""),
                    "prst": prst_of(sh),
                    "pos": (emu_cm(sh.left), emu_cm(sh.top)),
                    "size": (emu_cm(sh.width), emu_cm(sh.height)),
                    "area_ratio": round(area / (w_cm * h_cm) * 100, 1) if w_cm and h_cm and area else None,
                    "arrow": has_arrow(sh) if k == "连接线/线条" else None,
                })

        text_all = " ".join(texts)
        page_type = classify_page(idx, len(slides), layout, kinds, text_all, prsts, sizes)

        pages.append({
            "页": idx,
            "版式": layout,
            "类型": page_type,
            "形状构成": dict(kinds),
            "字号范围": (f"{min(sizes):g}–{max(sizes):g}pt" if sizes else "继承母版"),
            "文本量": len(text_all),
            "图形元素": diagram_shapes,
            "标题": (texts[0][:40] if texts else ""),
        })

    # 图形页统计：原生绘制 vs 图片
    native_pages, image_pages, smartart_pages = [], [], []
    for p in pages:
        ks = p["形状构成"]
        if ks.get("SmartArt"):
            smartart_pages.append(p["页"])
        if ks.get("原生形状", 0) >= 3 or ks.get("连接线/线条", 0) >= 2:
            native_pages.append(p["页"])
        if ks.get("图片", 0) >= 1:
            image_pages.append(p["页"])

    return {
        "file": os.path.basename(path),
        "path": path,
        "页数": len(slides),
        "页面尺寸": f"{w_cm} × {h_cm} cm（{ratio_label}）",
        "主题": theme,
        "母版字号阶梯": master_sizes,
        "版式分布": dict(layout_counter.most_common()),
        "页面类型分布": dict(Counter(p["类型"] for p in pages).most_common()),
        "形状类型 Total": dict(shape_kinds_all.most_common()),
        "预设形状 Top": dict(prst_counter.most_common(15)),
        "字号分布 Top": dict(size_counter.most_common(12)),
        "实际字体 Top": dict(font_counter.most_common(10)),
        "填充色 Top": dict(fill_counter.most_common(10)),
        "图形页归类": {
            "原生形状绘制页": native_pages,
            "图片型图形页": image_pages,
            "SmartArt 页": smartart_pages,
        },
        "页面明细": pages,
    }


# ---------------------------------------------------------------- 报告

def render_markdown(results):
    L = []
    for r in results:
        L.append(f"# PPT 视觉风格与版式提取报告：{r['file']}")
        L.append("")
        L.append(f"> 生成时间 {datetime.now():%Y-%m-%d %H:%M} ｜ 共 {r['页数']} 页")
        L.append("")
        L.append("## 一、画布与主题")
        L.append("")
        L.append(f"- **页面尺寸**：{r['页面尺寸']}")
        if r["主题"].get("name"):
            L.append(f"- **主题名**：{r['主题']['name']}")
        L.append("")
        L.append("**配色方案**")
        L.append("")
        L.append("| 角色 | 色值 |")
        L.append("|---|---|")
        for k, v in r["主题"].get("colors", {}).items():
            L.append(f"| {k} | {v or '-'} |")
        L.append("")
        L.append("**字体方案**")
        L.append("")
        L.append("| 用途 | 西文 | 中文 |")
        L.append("|---|---|---|")
        for k, v in r["主题"].get("fonts", {}).items():
            L.append(f"| {k} | {v.get('西文') or '-'} | {v.get('中文') or '-'} |")
        L.append("")
        if r["母版字号阶梯"]:
            L.append("**母版字号阶梯**")
            L.append("")
            for k, v in r["母版字号阶梯"].items():
                L.append(f"- {k}：" + " / ".join(f"{s:g}pt" for s in v))
            L.append("")

        L.append("## 二、页面类型与版式分布")
        L.append("")
        L.append("| 页面类型 | 页数 |")
        L.append("|---|---|")
        for k, v in r["页面类型分布"].items():
            L.append(f"| {k} | {v} |")
        L.append("")
        L.append("| 使用的版式 | 次数 |")
        L.append("|---|---|")
        for k, v in r["版式分布"].items():
            L.append(f"| {k} | {v} |")
        L.append("")

        L.append("## 三、图形元素构成（架构图/流程图怎么画的）")
        L.append("")
        g = r["图形页归类"]
        L.append(f"- **原生形状绘制页**（可完整还原画法）：{len(g['原生形状绘制页'])} 页 "
                 f"→ 第 {', '.join(map(str, g['原生形状绘制页'][:20])) or '无'} 页")
        L.append(f"- **图片型图形页**（画法无法提取，只能记住位置尺寸）：{len(g['图片型图形页'])} 页 "
                 f"→ 第 {', '.join(map(str, g['图片型图形页'][:20])) or '无'} 页")
        L.append(f"- **SmartArt 页**（层级可读、视觉不可精确还原）：{len(g['SmartArt 页'])} 页 "
                 f"→ 第 {', '.join(map(str, g['SmartArt 页'][:20])) or '无'} 页")
        L.append("")
        if r["预设形状 Top"]:
            L.append("**原生形状的几何类型（推断画图习惯）**")
            L.append("")
            L.append("| 几何类型 | 含义 | 出现次数 |")
            L.append("|---|---|---|")
            for k, v in r["预设形状 Top"].items():
                L.append(f"| {k} | {PRST_CN.get(k, '—')} | {v} |")
            L.append("")
        if r["填充色 Top"]:
            L.append("**形状填充色 Top**（图形配色习惯）")
            L.append("")
            L.append("、".join(f"`{k}`×{v}" for k, v in r["填充色 Top"].items()))
            L.append("")

        L.append("## 四、字号与元素分布")
        L.append("")
        L.append("| 形状类型 | 出现次数 |")
        L.append("|---|---|")
        for k, v in r["形状类型 Total"].items():
            L.append(f"| {k} | {v} |")
        L.append("")
        L.append("**正文字号分布**：" + "、".join(f"{k:g}pt×{v}" for k, v in r["字号分布 Top"].items()))
        L.append("")
        if r.get("实际字体 Top"):
            L.append("**页面内实际使用字体**（run 级统计，比主题字体方案更准）")
            L.append("")
            L.append("、".join(f"`{k}`×{v}" for k, v in r["实际字体 Top"].items()))
            L.append("")

        L.append("## 五、逐页明细")
        L.append("")
        L.append("| 页 | 类型 | 版式 | 形状构成 | 字号范围 | 标题 |")
        L.append("|---|---|---|---|---|---|")
        for p in r["页面明细"]:
            comp = "、".join(f"{k}×{v}" for k, v in sorted(p["形状构成"].items(), key=lambda x: -x[1]))
            L.append(f"| {p['页']} | {p['类型']} | {p['版式']} | {comp} | {p['字号范围']} | "
                     f"{(p['标题'] or '').replace('|', '/')} |")
        L.append("")
        L.append("---")
        L.append("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="PPT 视觉风格与版式提取器")
    ap.add_argument("files", nargs="+")
    ap.add_argument("--out")
    ap.add_argument("--json", dest="json_out")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    results = []
    for f in args.files:
        if not f.lower().endswith(".pptx"):
            print(f"[跳过] 非 pptx：{f}", file=sys.stderr)
            continue
        if not os.path.exists(f):
            print(f"[跳过] 不存在：{f}", file=sys.stderr)
            continue
        try:
            results.append(extract_pptx(f))
            if not args.quiet:
                print(f"[完成] {os.path.basename(f)}", file=sys.stderr)
        except Exception as e:  # noqa: BLE001
            print(f"[失败] {os.path.basename(f)}：{type(e).__name__} {e}", file=sys.stderr)

    if not results:
        print("没有可处理的文件", file=sys.stderr)
        return 1

    md = render_markdown(results)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(md)
        print(f"[报告] {args.out}", file=sys.stderr)
    else:
        print(md)

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as fh:
            json.dump(results, fh, ensure_ascii=False, indent=2, default=str)
        print(f"[JSON] {args.json_out}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
