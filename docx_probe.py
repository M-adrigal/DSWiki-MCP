#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx_probe.py — Word 排版格式提取器

从 .docx 中提取「可复用的排版规格」，输出 Markdown 规格表 + JSON。

核心能力：
  1. 样式继承链解析（basedOn → docDefaults），否则读不到"正文是宋体小四"
  2. 中文东亚字体（w:eastAsia）单独读取，python-docx 的 font.name 只给西文字体
  3. 首行缩进按「字符」计量（w:firstLineChars），而不是磅值
  4. 标题序号：自动编号（numbering.xml）+ 手动编号（文本正则）双路识别
  5. 章节树（文档骨架）
  6. 表格规格
  7. 多文档合并：同一规格出现在几份文档里（取交集，避免单份过拟合）

用法：
  python3 docx_probe.py a.docx
  python3 docx_probe.py a.docx b.docx c.docx --out report.md --json out.json
"""

import sys
import os
import re
import json
import argparse
from collections import Counter, OrderedDict
from datetime import datetime

from docx import Document
from docx.shared import Length
from docx.oxml.ns import qn

# ---------------------------------------------------------------- 常量映射

# 磅值 → 中文字号名
PT_TO_CN = {
    42: "初号", 36: "小初", 26: "一号", 24: "小一", 22: "二号", 18: "小二",
    16: "三号", 15: "小三", 14: "四号", 12: "小四", 10.5: "五号", 9: "小五",
    7.5: "六号", 6.5: "小六",
}

ALIGN_CN = {
    0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐",
    4: "分散对齐", None: "继承",
}

# OOXML w:jc 的字符串值 → 中文
JC_CN = {
    "left": "左对齐", "start": "左对齐", "center": "居中", "right": "右对齐",
    "end": "右对齐", "both": "两端对齐", "distribute": "分散对齐",
}

# 手动编号：1.1 / 1、 / （一） / 第一章 / 一、
RE_NUM_DOTTED = re.compile(r"^\s*(\d+(?:\.\d+)*)[\.\s、]")
RE_NUM_CN_PAREN = re.compile(r"^\s*[（(][一二三四五六七八九十]+[)）]")
RE_NUM_CN_CHAPTER = re.compile(r"^\s*第[一二三四五六七八九十百零\d]+[章节篇部分]")
RE_NUM_CN_SIMPLE = re.compile(r"^\s*[一二三四五六七八九十]+、")


def size_cn(pt):
    """磅值 → '小四 (12pt)' 这种可读形式"""
    if pt is None:
        return "继承"
    key = round(pt * 2) / 2
    name = PT_TO_CN.get(key)
    val = ("%g" % pt)
    return f"{name} {val}pt" if name else f"{val}pt"


# ---------------------------------------------------------------- 样式继承链

class StyleResolver:
    """解析 styles.xml，把每个样式的有效属性沿 basedOn 链合并出来。"""

    def __init__(self, doc):
        self.styles_el = doc.styles.element
        self._raw = {}       # styleId -> element
        self._cache = {}
        self._doc_default = self._read_doc_defaults()

        for st in self.styles_el.findall(qn("w:style")):
            sid = st.get(qn("w:styleId"))
            if sid:
                self._raw[sid] = st

    def _read_doc_defaults(self):
        out = {"size": None, "eastasia": None, "ascii": None, "align": None}
        base = qn("w:docDefaults")
        rpr = self.styles_el.find(base + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr"))
        if rpr is not None:
            rf = rpr.find(qn("w:rFonts"))
            if rf is not None:
                out["eastasia"] = rf.get(qn("w:eastAsia"))
                out["ascii"] = rf.get(qn("w:ascii"))
            sz = rpr.find(qn("w:sz"))
            if sz is not None:
                try:
                    out["size"] = int(sz.get(qn("w:val"))) / 2.0
                except (TypeError, ValueError):
                    pass
        ppr = self.styles_el.find(base + "/" + qn("w:pPrDefault") + "/" + qn("w:pPr"))
        if ppr is not None:
            j = ppr.find(qn("w:jc"))
            if j is not None:
                out["align"] = j.get(qn("w:val"))
        return out

    def effective(self, style_id):
        """返回该样式合并后的有效属性字典。"""
        if not style_id:
            return dict(self._doc_default, bold=None, italic=None, color=None,
                        align=None, line_spacing=None, first_line_chars=None,
                        first_line_twips=None, space_before=None, space_after=None,
                        outline=None)
        if style_id in self._cache:
            return self._cache[style_id]

        el = self._raw.get(style_id)
        if el is None:
            return self.effective(None)

        base = el.find(qn("w:basedOn"))
        base_id = base.get(qn("w:val")) if base is not None else None
        acc = self.effective(base_id) if base_id and base_id != style_id else dict(
            self._doc_default, bold=None, italic=None, color=None, align=None,
            line_spacing=None, first_line_chars=None, first_line_twips=None,
            space_before=None, space_after=None, outline=None)
        acc = dict(acc)

        rpr = el.find(qn("w:rPr"))
        if rpr is not None:
            rf = rpr.find(qn("w:rFonts"))
            if rf is not None:
                if rf.get(qn("w:eastAsia")):
                    acc["eastasia"] = rf.get(qn("w:eastAsia"))
                if rf.get(qn("w:ascii")):
                    acc["ascii"] = rf.get(qn("w:ascii"))
            sz = rpr.find(qn("w:sz"))
            if sz is not None:
                try:
                    acc["size"] = int(sz.get(qn("w:val"))) / 2.0
                except (TypeError, ValueError):
                    pass
            b = rpr.find(qn("w:b"))
            if b is not None:
                acc["bold"] = b.get(qn("w:val")) not in ("0", "false")
            i = rpr.find(qn("w:i"))
            if i is not None:
                acc["italic"] = i.get(qn("w:val")) not in ("0", "false")
            c = rpr.find(qn("w:color"))
            if c is not None:
                acc["color"] = c.get(qn("w:val"))

        ppr = el.find(qn("w:pPr"))
        if ppr is not None:
            j = ppr.find(qn("w:jc"))
            if j is not None:
                acc["align"] = j.get(qn("w:val"))
            sp = ppr.find(qn("w:spacing"))
            if sp is not None:
                ln = sp.get(qn("w:line"))
                rule = sp.get(qn("w:lineRule"))
                if ln:
                    try:
                        lnv = int(ln)
                        if rule == "auto":
                            acc["line_spacing"] = round(lnv / 240.0, 2)
                        elif rule in ("exact", "atLeast"):
                            acc["line_spacing"] = f"固定 {lnv/20.0:g}pt"
                    except ValueError:
                        pass
                for k, tag in (("space_before", "w:before"), ("space_after", "w:after")):
                    v = sp.get(qn(tag))
                    if v:
                        try:
                            acc[k] = int(v) / 20.0
                        except ValueError:
                            pass
            ind = ppr.find(qn("w:ind"))
            if ind is not None:
                flc = ind.get(qn("w:firstLineChars"))
                if flc:
                    acc["first_line_chars"] = int(flc) / 100.0
                fl = ind.get(qn("w:firstLine"))
                if fl:
                    try:
                        acc["first_line_twips"] = int(fl) / 20.0   # twips → pt
                    except ValueError:
                        pass
            ol = ppr.find(qn("w:outlineLvl"))
            if ol is not None:
                acc["outline"] = int(ol.get(qn("w:val")))

        self._cache[style_id] = acc
        return acc

    def outline_of(self, style_id):
        return self.effective(style_id).get("outline")


# ---------------------------------------------------------------- 编号

def numbering_map(doc):
    """numId -> ilvl -> {fmt, text, start}"""
    out = {}
    try:
        numbering = doc.part.numbering_part.element
    except (KeyError, AttributeError, NotImplementedError):
        return out

    abstract = {}
    for an in numbering.findall(qn("w:abstractNum")):
        aid = an.get(qn("w:abstractNumId"))
        lvls = {}
        for lvl in an.findall(qn("w:lvl")):
            il = lvl.get(qn("w:ilvl"))
            fmt = lvl.find(qn("w:numFmt"))
            txt = lvl.find(qn("w:lvlText"))
            st = lvl.find(qn("w:start"))
            lvls[il] = {
                "fmt": fmt.get(qn("w:val")) if fmt is not None else None,
                "text": txt.get(qn("w:val")) if txt is not None else None,
                "start": st.get(qn("w:val")) if st is not None else None,
            }
        abstract[aid] = lvls

    for num in numbering.findall(qn("w:num")):
        nid = num.get(qn("w:numId"))
        ref = num.find(qn("w:abstractNumId"))
        if ref is not None:
            out[nid] = abstract.get(ref.get(qn("w:val")), {})
    return out


FMT_CN = {
    "decimal": "阿拉伯数字", "upperRoman": "大写罗马", "lowerRoman": "小写罗马",
    "upperLetter": "大写字母", "lowerLetter": "小写字母",
    "chineseCounting": "中文数字", "chineseCountingThousand": "中文数字",
    "ideographTraditional": "天干", "decimalEnclosedCircle": "圈码",
    "none": "无",
}


def detect_numbering(par, para_style_id, num_map, text):
    """返回 (编号形式描述, 层级)。优先自动编号，其次手动文本编号。"""
    ppr = par._p.find(qn("w:pPr"))
    num_id = ilvl = None
    if ppr is not None:
        npr = ppr.find(qn("w:numPr"))
        if npr is not None:
            nid = npr.find(qn("w:numId"))
            il = npr.find(qn("w:ilvl"))
            num_id = nid.get(qn("w:val")) if nid is not None else None
            ilvl = il.get(qn("w:val")) if il is not None else None

    if num_id and num_id != "0":
        lvls = num_map.get(num_id, {})
        lv = lvls.get(ilvl or "0", {})
        fmt = FMT_CN.get(lv.get("fmt") or "", lv.get("fmt") or "")
        txt = lv.get("text") or ""
        if txt:
            return f"自动编号 {txt}（{fmt}）", (int(ilvl) + 1 if ilvl else 1)
        return f"自动编号（{fmt}）", (int(ilvl) + 1 if ilvl else 1)

    if not text:
        return None, None
    if RE_NUM_DOTTED.match(text):
        m = RE_NUM_DOTTED.match(text)
        return f"手动编号 1.1 式（层级 {m.group(1).count('.') + 1}）", m.group(1).count(".") + 1
    if RE_NUM_CN_CHAPTER.match(text):
        return "手动编号 第X章/节", 1
    if RE_NUM_CN_PAREN.match(text):
        return "手动编号 （一）式", 2
    if RE_NUM_CN_SIMPLE.match(text):
        return "手动编号 一、式", 1
    return None, None


# ---------------------------------------------------------------- 段落分析

def para_signature(par, resolver, num_map):
    """把一个段落归纳成一条规格指纹。"""
    style_id = par.style.style_id if par.style is not None else None
    style_name = par.style.name if par.style is not None else "(无样式)"
    seff = resolver.effective(style_id)

    # run 级：取出现最多的 run 属性（正文段落通常一致）
    run_props = Counter()
    for r in par.runs:
        if not r.text.strip():
            continue
        rpr = r._r.find(qn("w:rPr"))
        size = None
        ea = None
        ascii_f = None
        bold = None
        if rpr is not None:
            rf = rpr.find(qn("w:rFonts"))
            if rf is not None:
                ea = rf.get(qn("w:eastAsia"))
                ascii_f = rf.get(qn("w:ascii"))
            sz = rpr.find(qn("w:sz"))
            if sz is not None:
                try:
                    size = int(sz.get(qn("w:val"))) / 2.0
                except (TypeError, ValueError):
                    pass
            b = rpr.find(qn("w:b"))
            if b is not None:
                bold = b.get(qn("w:val")) not in ("0", "false")
        run_props[(size, ea, ascii_f, bold)] += 1

    if run_props:
        size, ea, ascii_f, bold = run_props.most_common(1)[0][0]
    else:
        size, ea, ascii_f, bold = None, None, None, None

    size = size if size is not None else seff.get("size")
    ea = ea or seff.get("eastasia")
    ascii_f = ascii_f or seff.get("ascii")
    bold = bold if bold is not None else seff.get("bold")

    pf = par.paragraph_format
    # 行距
    ls = pf.line_spacing
    if isinstance(ls, Length):
        line = f"固定 {ls.pt:g}pt"
    elif isinstance(ls, float):
        line = f"{ls:g} 倍"
    elif ls is not None:
        try:
            line = f"{float(ls):g} 倍"
        except (TypeError, ValueError):
            line = str(ls)
    else:
        line = seff.get("line_spacing")
        line = (f"{line:g} 倍" if isinstance(line, float) else line) or "继承"

    # 首行缩进（优先「字符」单位，其次 twips，再次 python-docx 折算）
    flc = seff.get("first_line_chars")
    ppr = par._p.find(qn("w:pPr"))
    if ppr is not None:
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            if ind.get(qn("w:firstLineChars")):
                flc = int(ind.get(qn("w:firstLineChars"))) / 100.0
            elif ind.get(qn("w:firstLine")):
                try:
                    tw = int(ind.get(qn("w:firstLine"))) / 20.0
                    flc = round(tw / (size or 12), 2)
                except ValueError:
                    pass
    if flc is None and seff.get("first_line_twips"):
        flc = round(seff["first_line_twips"] / (size or 12), 2)
    if flc is None and pf.first_line_indent is not None:
        try:
            flc = round(pf.first_line_indent.pt / (size or 12), 2)
        except (AttributeError, TypeError, ZeroDivisionError):
            flc = None
    indent = f"首行缩进 {flc:g} 字符" if flc else "无缩进"

    if pf.alignment is not None:
        align = ALIGN_CN.get(int(pf.alignment), str(pf.alignment))
    else:
        align = JC_CN.get(seff.get("align") or "", "继承")

    sb = pf.space_before.pt if pf.space_before is not None else seff.get("space_before")
    sa = pf.space_after.pt if pf.space_after is not None else seff.get("space_after")

    num_desc, num_level = detect_numbering(par, style_id, num_map, par.text)

    sig = {
        "style": style_name,
        "font_eastasia": ea or "继承",
        "font_ascii": ascii_f or "继承",
        "size_pt": size,
        "size_cn": size_cn(size),
        "bold": bool(bold),
        "align": align,
        "line_spacing": line,
        "indent": indent,
        "space_before": sb,
        "space_after": sa,
        "numbering": num_desc,
        "numbering_level": num_level,
    }
    return sig, style_name, seff


def extract_docx(path):
    doc = Document(path)
    resolver = StyleResolver(doc)
    num_map = numbering_map(doc)

    # 页面设置（字段可能缺失，不能直接取 .cm）
    page = {}
    for sect in doc.sections:
        def _c(v, nd=2):
            return f"{v.cm:.{nd}f}" if v is not None else "未设置"
        page = {
            "纸张": f"{_c(sect.page_width, 1)} × {_c(sect.page_height, 1)} cm",
            "页边距": (f"上 {_c(sect.top_margin)} / 下 {_c(sect.bottom_margin)} / "
                     f"左 {_c(sect.left_margin)} / 右 {_c(sect.right_margin)} cm"),
        }
        break

    # 段落规格统计
    style_stats = OrderedDict()
    headings = []
    total_paras = 0

    for par in doc.paragraphs:
        text = par.text.strip()
        if not text:
            continue
        total_paras += 1
        sig, sname, seff = para_signature(par, resolver, num_map)

        key = sname
        if key not in style_stats:
            style_stats[key] = {"count": 0, "sigs": Counter(), "samples": []}
        style_stats[key]["count"] += 1
        sig_key = json.dumps({k: v for k, v in sig.items() if k != "style"},
                             ensure_ascii=False, sort_keys=True)
        style_stats[key]["sigs"][sig_key] += 1
        if len(style_stats[key]["samples"]) < 2:
            style_stats[key]["samples"].append(text[:60])

        # 标题判定
        outline = resolver.outline_of(par.style.style_id) if par.style is not None else None
        is_heading = (outline is not None) or sname.lower().startswith(("heading", "标题")) \
                     or sname in ("Title", "Subtitle")
        if is_heading and text:
            lvl = (outline + 1) if outline is not None else _heading_level_from_name(sname)
            headings.append({"level": lvl, "text": text, "style": sname})

    # 表格规格
    tables = []
    for i, t in enumerate(doc.tables):
        info = {"序号": i + 1, "行数": len(t.rows), "列数": len(t.columns),
                "样式": t.style.name if t.style is not None else "(无)",
                "首行加粗": None, "首行字号": None, "首行底纹": None,
                "边框": _table_borders(t)}
        try:
            hdr = t.rows[0]
            bolds, sizes, shades = [], [], []
            for c in hdr.cells:
                for p in c.paragraphs:
                    for r in p.runs:
                        if r.text.strip():
                            if r.font.bold is not None:
                                bolds.append(r.font.bold)
                            if r.font.size is not None:
                                sizes.append(r.font.size.pt)
                tcpr = c._tc.find(qn("w:tcPr"))
                if tcpr is not None:
                    sh = tcpr.find(qn("w:shd"))
                    if sh is not None and sh.get(qn("w:fill")) not in (None, "auto", "FFFFFF"):
                        shades.append(sh.get(qn("w:fill")))
            info["首行加粗"] = (all(bolds) if bolds else None)
            info["首行字号"] = (size_cn(Counter(sizes).most_common(1)[0][0]) if sizes else None)
            info["首行底纹"] = (Counter(shades).most_common(1)[0][0] if shades else "无")
        except (IndexError, AttributeError):
            pass

        # 正文单元格字号（首行之外）：表头 10.5pt、正文 9pt 这种阶梯很常见，
        # 只报首行会让人以为整表都是 10.5pt
        body_sizes = []
        try:
            for row in t.rows[1:]:
                for c in row.cells:
                    for p in c.paragraphs:
                        for r in p.runs:
                            if r.text.strip() and r.font.size is not None:
                                body_sizes.append(r.font.size.pt)
        except (IndexError, AttributeError):
            pass
        info["正文字号"] = (size_cn(Counter(body_sizes).most_common(1)[0][0])
                          if body_sizes else None)

        # 列宽（cm）：None 表示按内容自动
        widths = []
        try:
            for col in t.columns:
                widths.append(round(col.width.cm, 2) if col.width else None)
        except (AttributeError, ValueError):
            pass
        info["列宽(cm)"] = widths

        # 跨页重复表头：w:trPr/w:tblHeader
        repeated = None
        try:
            trpr = t.rows[0]._tr.find(qn("w:trPr"))
            th = trpr.find(qn("w:tblHeader")) if trpr is not None else None
            if th is not None:
                repeated = th.get(qn("w:val")) not in ("0", "false")
        except (IndexError, AttributeError):
            pass
        info["跨页重复表头"] = repeated

        tables.append(info)

    return {
        "file": os.path.basename(path),
        "path": path,
        "页面设置": page,
        "有效段落数": total_paras,
        "样式规格": style_stats,
        "章节树": headings,
        "表格": tables,
    }


def _heading_level_from_name(name):
    m = re.search(r"(\d+)", name)
    return int(m.group(1)) if m else 1


def _table_borders(t):
    tblpr = t._tbl.find(qn("w:tblPr"))
    if tblpr is None:
        return "继承"
    b = tblpr.find(qn("w:tblBorders"))
    if b is None:
        return "继承（可能来自表格样式）"
    parts = []
    for tag, cn in (("w:top", "上"), ("w:bottom", "下"), ("w:left", "左"),
                    ("w:right", "右"), ("w:insideH", "内横"), ("w:insideV", "内竖")):
        e = b.find(qn(tag))
        if e is None:
            continue
        val = e.get(qn("w:val"))
        if val in ("none", "nil"):
            parts.append(f"{cn}无")
        else:
            sz = e.get(qn("w:sz"))
            color = e.get(qn("w:color"))
            parts.append(f"{cn}{sz and int(sz)/8:g}pt{color and '/' + color}")
    return "、".join(parts) if parts else "继承"


# ---------------------------------------------------------------- 报告输出

def heading_numbering_scheme(headings):
    """按层级归纳标题的编号体系：第 N 级用「1.1」还是「一、」还是无编号。"""
    out = OrderedDict()
    for h in headings:
        lvl = h["level"]
        pat, _ = _numbering_pattern(h["text"])
        d = out.setdefault(lvl, {"patterns": Counter(), "samples": [], "count": 0})
        d["patterns"][pat] += 1
        d["count"] += 1
        if len(d["samples"]) < 2:
            d["samples"].append(h["text"][:40])
    for lvl, d in out.items():
        d["pattern"] = d["patterns"].most_common(1)[0][0]
        d["styles"] = ", ".join(sorted({h["style"] for h in headings if h["level"] == lvl}))
    return out


def _numbering_pattern(text):
    m = RE_NUM_DOTTED.match(text)
    if m:
        depth = m.group(1).count(".") + 1
        return f"阿拉伯数字「1.1」式（{depth} 段）", depth
    if RE_NUM_CN_CHAPTER.match(text):
        return "「第X章/节」式", 1
    if RE_NUM_CN_PAREN.match(text):
        return "「（一）」式", 2
    if RE_NUM_CN_SIMPLE.match(text):
        return "「一、」式", 1
    return "无编号", 0


def representative(stats_entry):
    """取该样式下出现最多的那套规格作为代表。"""
    sig_json, n = stats_entry["sigs"].most_common(1)[0]
    sig = json.loads(sig_json)
    sig["_share"] = f"{n}/{stats_entry['count']}"
    return sig


def render_markdown(results, merged=None):
    L = []
    for r in results:
        L.append(f"# 排版规格提取报告：{r['file']}")
        L.append("")
        L.append(f"> 生成时间 {datetime.now():%Y-%m-%d %H:%M} ｜ 有效段落 {r['有效段落数']} 个")
        L.append("")
        L.append("## 一、页面设置")
        L.append("")
        for k, v in r["页面设置"].items():
            L.append(f"- **{k}**：{v}")
        L.append("")
        L.append("## 二、样式规格表")
        L.append("")
        L.append("| 样式名 | 段落数 | 中文字体 | 西文字体 | 字号 | 加粗 | 对齐 | 行距 | 缩进 | 段前/后 | 编号形式 |")
        L.append("|---|---|---|---|---|---|---|---|---|---|---|")
        for sname, entry in r["样式规格"].items():
            s = representative(entry)
            sb = f"{s['space_before']:g}" if isinstance(s["space_before"], (int, float)) else "-"
            sa = f"{s['space_after']:g}" if isinstance(s["space_after"], (int, float)) else "-"
            L.append(
                f"| {sname} | {entry['count']} | {s['font_eastasia']} | {s['font_ascii']} | "
                f"{s['size_cn']} | {'是' if s['bold'] else '否'} | {s['align']} | "
                f"{s['line_spacing']} | {s['indent']} | {sb}/{sa} | {s['numbering'] or '-'} |"
            )
        L.append("")
        L.append("## 三、文档骨架（章节树）")
        L.append("")
        if r["章节树"]:
            for h in r["章节树"]:
                L.append(f"{'  ' * (h['level'] - 1)}- {'#' * min(h['level'], 6)} {h['text']}  `{h['style']}`")
        else:
            L.append("_未识别到标题层级（可能全部使用手动加粗而非标题样式）_")
        L.append("")
        L.append("## 四、标题序号体系")
        L.append("")
        scheme = heading_numbering_scheme(r["章节树"])
        if scheme:
            L.append("| 层级 | 编号形式 | 该层标题数 | 使用样式 | 样例 |")
            L.append("|---|---|---|---|---|")
            for lvl, d in scheme.items():
                L.append(f"| 第 {lvl} 级 | {d['pattern']} | {d['count']} | {d['styles']} | "
                         + " ／ ".join(d["samples"]) + " |")
        else:
            L.append("_未识别到标题_")
        L.append("")
        L.append("## 五、表格规格")
        L.append("")
        if r["表格"]:
            L.append("| # | 行×列 | 表格样式 | 首行加粗 | 首行字号 | 正文字号 | 边框 |")
            L.append("|---|---|---|---|---|---|---|")
            for t in r["表格"]:
                L.append(f"| {t['序号']} | {t['行数']}×{t['列数']} | {t['样式']} | "
                         f"{t['首行加粗']} | {t['首行字号'] or '-'} | {t.get('正文字号') or '-'} | {t['边框']} |")
        else:
            L.append("_无表格_")
        L.append("")
        L.append("## 六、样式明细（含样例）")
        L.append("")
        for sname, entry in r["样式规格"].items():
            L.append(f"### {sname}（{entry['count']} 段）")
            for samp in entry["samples"]:
                L.append(f"- 样例：{samp}")
            for sig_json, n in entry["sigs"].most_common(3):
                s = json.loads(sig_json)
                L.append(f"  - {n} 段：{s['font_eastasia']} {s['size_cn']}，{s['line_spacing']}，"
                         f"{s['indent']}，{s['align']}")
            L.append("")
        L.append("---")
        L.append("")

    if merged:
        L.append("# 多文档合并：跨文档一致的规格（取交集）")
        L.append("")
        L.append("| 样式名 | 出现文档数 | 一致的中文字体 | 一致的字号 | 一致的行距 | 一致的缩进 |")
        L.append("|---|---|---|---|---|---|")
        for sname, m in merged.items():
            L.append(f"| {sname} | {m['files']} | {m['font']} | {m['size']} | {m['line']} | {m['indent']} |")
        L.append("")
    return "\n".join(L)


def merge_results(results):
    """跨文档聚合：同一样式名在多份文档中的规格是否一致。"""
    agg = {}
    n_files = len(results)
    for r in results:
        for sname, entry in r["样式规格"].items():
            s = representative(entry)
            a = agg.setdefault(sname, {"files": set(), "font": Counter(),
                                       "size": Counter(), "line": Counter(),
                                       "indent": Counter()})
            a["files"].add(r["file"])
            a["font"][s["font_eastasia"]] += 1
            a["size"][s["size_cn"]] += 1
            a["line"][s["line_spacing"]] += 1
            a["indent"][s["indent"]] += 1

    out = {}
    for sname, a in agg.items():
        def pick(c):
            v, n = c.most_common(1)[0]
            return v if n == len(a["files"]) else f"{v}（不一致，{n}/{len(a['files'])}）"
        out[sname] = {"files": f"{len(a['files'])}/{n_files}",
                      "font": pick(a["font"]), "size": pick(a["size"]),
                      "line": pick(a["line"]), "indent": pick(a["indent"])}
    return dict(sorted(out.items(), key=lambda kv: -int(kv[1]["files"].split("/")[0])))


def main():
    ap = argparse.ArgumentParser(description="Word 排版格式提取器")
    ap.add_argument("files", nargs="+", help="一个或多个 .docx 路径")
    ap.add_argument("--out", help="Markdown 报告输出路径")
    ap.add_argument("--json", dest="json_out", help="JSON 输出路径")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    results = []
    for f in args.files:
        if not f.lower().endswith(".docx"):
            print(f"[跳过] 非 docx：{f}", file=sys.stderr)
            continue
        if not os.path.exists(f):
            print(f"[跳过] 不存在：{f}", file=sys.stderr)
            continue
        try:
            results.append(extract_docx(f))
            if not args.quiet:
                print(f"[完成] {os.path.basename(f)}", file=sys.stderr)
        except Exception as e:  # noqa: BLE001
            print(f"[失败] {os.path.basename(f)}：{type(e).__name__} {e}", file=sys.stderr)

    if not results:
        print("没有可处理的文件", file=sys.stderr)
        return 1

    merged = merge_results(results) if len(results) > 1 else None
    md = render_markdown(results, merged)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(md)
        print(f"[报告] {args.out}", file=sys.stderr)
    else:
        print(md)

    if args.json_out:
        payload = []
        for r in results:
            payload.append({
                "file": r["file"], "path": r["path"],
                "page": r["页面设置"], "paragraphs": r["有效段落数"],
                "styles": {k: {"count": v["count"], "spec": representative(v)}
                           for k, v in r["样式规格"].items()},
                "outline": r["章节树"],
                "numbering_scheme": {str(k): {"pattern": v["pattern"], "count": v["count"],
                                              "styles": v["styles"]}
                                     for k, v in heading_numbering_scheme(r["章节树"]).items()},
                "tables": r["表格"],
            })
        with open(args.json_out, "w", encoding="utf-8") as fh:
            json.dump({"docs": payload, "merged": merged}, fh, ensure_ascii=False, indent=2)
        print(f"[JSON] {args.json_out}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
