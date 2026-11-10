#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
quality_gate.py — 文档准入评分器

回答一个问题：这份文档**够不够格**进入画像？

注意：质量分只判断「能不能用」，不判断「是不是同一类」。
两件事必须分开：
  - 质量差 → 清洗/修复，或用但降权
  - 类型不同 → 拆画像，不是清洗能解决的

评分维度（满分 100）：
  标题规范性 25  伪标题（Normal + 手工放大字号）占比
  样式覆盖率 20  有多少段落是靠套样式而非手工刷格式
  层级连贯性 15  标题层级是否跳级（H1 直接跳 H4）
  编号一致性 15  是否混用手打编号与自动编号 / 多种编号形式
  内容完整度 15  是否残留「待补充」「TODO」「暂无」等占位符
  结构完整度 10  是否有目录、有前言/概述

用法：
  python3 quality_gate.py a.docx b.docx
  python3 quality_gate.py *.docx --json out.json
"""

import sys
import os
import re
import json
import argparse
from collections import Counter

from docx import Document
from docx.oxml.ns import qn

import service_config

# 图注、表注、数据来源这类短文字，字号通常小于正文，属正常排版而非「手刷」
CAPTION_RE = re.compile(r"^\s*(图|表|Figure|Table|注|来源|数据来源|备注)[\s\d:：]")

PLACEHOLDERS = ["待补充", "待完善", "待确定", "TODO", "TBD", "暂无", "XXX", "xxx"]
NUM_STYLES = {
    "阿拉伯点分": re.compile(r"^\s*\d+(\.\d+)*[\.\s、]"),
    "中文顿号": re.compile(r"^\s*[一二三四五六七八九十]+、"),
    "第X章": re.compile(r"^\s*第[一二三四五六七八九十百零\d]+[章节]"),
    "括号中文": re.compile(r"^\s*[（(][一二三四五六七八九十]+[)）]"),
}


def run_max_size(p):
    szs = []
    for r in p.runs:
        if r.font.size is not None:
            szs.append(r.font.size.pt)
    return max(szs) if szs else None


def is_toc(p):
    return (p.style.name or "").lower().startswith("toc")


def style_body_size(doc, paras):
    """正文基准字号，按继承链逐级回退。返回 (字号, 来源)。

    顺序：Normal 样式 → docDefaults → 正文段落显式字号的众数 → 12pt 兜底。

    不能取「Normal 段落里最常见的显式字号」当第一选择：正文段落通常不写字号
    （继承样式），写上显式字号的往往是封面大标题。实测一份技术文档只有 3 个
    Normal 段落写了字号（36/28/18pt，全是封面），基准被判成 36pt，
    连带所有「字号异常」提示和手刷检测一起失效。

    **兜底必须标出来**：12pt 是猜的，不是文档事实。用猜的值去判「字号是否异常」，
    判出来的结论没法追溯。调用方看到来源=「兜底」就该知道这一项不可信。
    """
    # 1 Normal 样式
    try:
        st = doc.styles["Normal"]
        if st.font.size:
            return st.font.size.pt, "Normal 样式"
    except (KeyError, AttributeError):
        pass

    # 2 docDefaults（w:sz 单位是半磅）
    try:
        dd = doc.styles.element.find(qn("w:docDefaults"))
        if dd is not None:
            rpd = dd.find(qn("w:rPrDefault"))
            rpr = rpd.find(qn("w:rPr")) if rpd is not None else None
            sz = rpr.find(qn("w:sz")) if rpr is not None else None
            if sz is not None and sz.get(qn("w:val")):
                return int(sz.get(qn("w:val"))) / 2, "docDefaults"
    except (AttributeError, ValueError, TypeError):
        pass

    # 3 正文段落显式字号的众数（排除封面量级的大字）
    sizes = [s for p in paras
             if p.style.name in ("Normal", "正文") and (s := run_max_size(p))]
    if sizes:
        med = sorted(sizes)[len(sizes) // 2]
        near = [s for s in sizes if 0.5 * med <= s <= 2 * med] or sizes
        return Counter(near).most_common(1)[0][0], "正文段落众数"

    return 12.0, "兜底（文档未定义字号，12pt 是默认猜测）"


def cover_span(paras, base_size):
    """封面区：第一个真标题之前、且字号显著大于正文的那些段落。

    封面用大字号是正常做法，不该按「正文手工改字号」或「伪标题」扣分。
    返回索引集合。没有真标题、或文档开头就是标题时，视为无封面。
    """
    first_head = next((i for i, p in enumerate(paras)
                       if (p.style.name or "").lower().startswith(("heading", "标题"))), None)
    if first_head is None:
        return set()
    cover = set()
    for i in range(first_head):
        s = run_max_size(paras[i])
        # 1.5 倍起算：封面上「v2.0.0 基线版」这类 18pt（正文 11pt 的 1.64 倍）
        # 也要算封面；只放在第一个真标题之前，不会误伤正文里的小标题
        if s and s >= base_size * 1.5:
            cover.add(i)
    return cover


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# 图注/表题的常见形态：「图1 …」「图 1-1 …」「Figure 3 …」「表2-1 …」
CAPTION_LINE_RE = re.compile(r"^\s*(图|表|Figure|Table)\s*[\d一二三四五六七八九十]+([-—.．]\d+)*")
CAPTION_WORD_RE = re.compile(r"^\s*(图注|表题|注[:：]|资料来源|数据来源)")


def _is_heading_para(el) -> bool:
    ppr = el.find(W_NS + "pPr")
    if ppr is None:
        return False
    ps = ppr.find(W_NS + "pStyle")
    if ps is None:
        return False
    val = (ps.get(W_NS + "val") or "").lower()
    return val.startswith(("heading", "title", "subtitle")) or val.startswith("标题")


def _para_text(el) -> str:
    return "".join(t.text or "" for t in el.iter(W_NS + "t")).strip()


def _block_seq(doc):
    """按文档顺序返回 [('p', 元素, 文本, 是否含图), ('tbl', 元素)]。"""
    seq = []
    for b in doc.element.body.iterchildren():
        tag = b.tag.replace(W_NS, "")
        if tag == "p":
            seq.append(("p", b, _para_text(b), b.find(".//" + W_NS + "drawing") is not None))
        elif tag == "tbl":
            seq.append(("tbl", b, "", False))
    return seq


def _style_num_id(style, depth=0):
    """沿 basedOn 链找样式上的 numPr，返回 (numId, 样式名)。

    项目符号列表通常靠样式带编号（List Bullet 的 pPr 里有 numPr），
    段落本身没有 numPr——只查段落级会把这批全漏掉。
    """
    while style is not None and depth < 6:
        try:
            ppr = style.element.find(qn("w:pPr"))
            if ppr is not None:
                npr = ppr.find(qn("w:numPr"))
                if npr is not None:
                    nid = npr.find(qn("w:numId"))
                    val = nid.get(qn("w:val")) if nid is not None else None
                    if val not in (None, "0"):
                        return val, style.name
        except (AttributeError, KeyError):
            return None, None
        style = style.base_style
        depth += 1
    return None, None


def _abstract_formats(doc):
    """numId → 该编号定义里出现的 numFmt 集合（判断是不是层级式 %1.%2）。"""
    out = {}
    try:
        numbering = doc.part.numbering_part.element
    except (KeyError, AttributeError, NotImplementedError):
        return out
    abs_map = {}
    for an in numbering.findall(qn("w:abstractNum")):
        aid = an.get(qn("w:abstractNumId"))
        fmts, txts = set(), []
        for lvl in an.findall(qn("w:lvl")):
            f = lvl.find(qn("w:numFmt"))
            t = lvl.find(qn("w:lvlText"))
            if f is not None and f.get(qn("w:val")):
                fmts.add(f.get(qn("w:val")))
            if t is not None and t.get(qn("w:val")):
                txts.append(t.get(qn("w:val")))
        abs_map[aid] = (fmts, txts)
    for num in numbering.findall(qn("w:num")):
        nid = num.get(qn("w:numId"))
        an = num.find(qn("w:abstractNumId"))
        aid = an.get(qn("w:val")) if an is not None else None
        fmts, txts = abs_map.get(aid, (set(), []))
        # 层级式：lvlText 含「%1.%2」这类多级引用
        hier = any(re.search(r"%1[.．]", t) for t in txts)
        out[nid] = {"格式": fmts, "层级式": hier}
    return out


def count_auto_numbering(doc):
    """数自动编号。返回 (总数, 层级式数量)。

    必须 покрыт两种来源：
      段落级 w:pPr/w:numPr —— 手工应用的编号
      样式级 w:pPr/w:numPr —— List Bullet / 编号样式带的编号（原来全漏）
    """
    fmt = _abstract_formats(doc)
    total = hier = 0
    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        nid = None
        ppr = p._p.find(qn("w:pPr"))
        if ppr is not None:
            npr = ppr.find(qn("w:numPr"))
            if npr is not None:
                n = npr.find(qn("w:numId"))
                nid = n.get(qn("w:val")) if n is not None else None
                if nid in (None, "0"):
                    nid = None
        if nid is None and p.style is not None:
            nid, _ = _style_num_id(p.style)
        if nid is None:
            continue
        total += 1
        if fmt.get(nid, {}).get("层级式"):
            hier += 1
    return total, hier


def count_captions(doc):
    """按「与图片/表格相邻」识别图注表题，而不是只数正文里的编号引用。

    相邻判据比数引用可靠：正文写「如下图」不引编号时，数引用会误判缺失；
    反过来正文引用了编号但图注没写，数引用又会误判通过。
    规则：图片段落的下一段（跳过空段与标题）若以「图N / Figure N / 图注 / 来源」开头，
    即认为有图注；表格看前后各一段。
    """
    seq = _block_seq(doc)
    fig_n = sum(1 for k, _, _, img in seq if k == "p" and img)
    tbl_n = sum(1 for k, _, _, _ in seq if k == "tbl")

    def caption_after(idx, limit=3):
        seen = 0
        for j in range(idx + 1, min(idx + 1 + limit * 3, len(seq))):
            k, el, txt, img = seq[j]
            if k != "p" or not txt or img:
                continue
            if _is_heading_para(el):
                return False       # 紧跟标题说明没有图注
            if CAPTION_LINE_RE.match(txt) or CAPTION_WORD_RE.match(txt):
                return True
            seen += 1
            if seen >= limit:
                return False
        return False

    def caption_before(idx, limit=2):
        seen = 0
        for j in range(idx - 1, max(idx - 1 - limit * 3, -1), -1):
            k, el, txt, img = seq[j]
            if k != "p" or not txt or img:
                continue
            if _is_heading_para(el):
                return False
            if CAPTION_LINE_RE.match(txt) or CAPTION_WORD_RE.match(txt):
                return True
            seen += 1
            if seen >= limit:
                return False
        return False

    fig_caps = 0
    for i, (k, _, _, img) in enumerate(seq):
        if k == "p" and img and caption_after(i):
            fig_caps += 1
    tbl_caps = 0
    for i, (k, _, _, _) in enumerate(seq):
        # 表题可能在表格上方或下方：向下查要从表格自身的位置起算。
        # 原来写 caption_after(i + 1) 把「表格后第一段」整段跳过了，
        # 于是「表格 + 表题在下」永远判不出来——实现与公开判据（前或后一段）不一致。
        if k == "tbl" and (caption_before(i) or caption_after(i, limit=2)):
            tbl_caps += 1
    return fig_caps, fig_n, tbl_caps, tbl_n


def delivery_checks(doc, paras):
    """交付前必查项：图注表题、页码页眉、修订记录、格式残留。

    这些不计分——完成度与排版规范性是两回事。混进分数里会让评分口径说不清
    （无目录扣 5、无页码扣 3，谁定的？）。

    三档结果：通过 / 缺失 / 不适用。缺第三档时「文档没图」会显示成「图注通过」，
    看起来像图注做得好——把「不存在」说成「合格」是误导。
    """
    body = "\n".join(p.text for p in paras)
    items = []

    # 「结果」是给人读的措辞，会随语义调整；「状态」是给机器筛的三态。
    # 两者必须同步维护：把「格式残留」改成「无残留/存在残留」时没在这里登记，
    # 「无残留」就会落到兜底 fail——正常项被下游误报成缺陷。改措辞先改这里。
    PASS_STATES = {"通过", "无残留"}
    NA_STATES = {"不适用"}

    def add(name, state, detail):
        """state 是给人读的措辞，必须语义准确；status 是给机器筛的。

        「格式残留」有残留时若写「缺失」，字面会读成「缺了残留」——反了。
        所以它用「无残留 / 存在残留」；其余项「缺失」= 该有的没有，语义是对的。
        status 统一三态，调用方按它筛，不用猜中文。
        """
        status = ("pass" if state in PASS_STATES else
                  "n/a" if state in NA_STATES else "fail")
        items.append({"项": name, "结果": state, "状态": status, "说明": detail})

    # ---- 图注/表题：以「与图片相邻」为主要判据，编号引用只作辅助说明
    fig_caps, fig_n, tbl_caps, tbl_n = count_captions(doc)
    fig_refs = len(re.findall(r"(图|Figure)\s*[\d一二三四五六七八九十]+([-—.．]\d+)*", body))
    tbl_refs = len(re.findall(r"(表|Table)\s*[\d一二三四五六七八九十]+([-—.．]\d+)*", body))
    fig_loose = len(re.findall(r"(如图|见图|见下图|参见图|下图所示|如图所示)", body))

    if fig_n == 0:
        add("图注编号", "不适用", "无内嵌图片，本项不适用（不是「通过」）")
    elif fig_caps >= fig_n:
        add("图注编号", "通过",
            f"{fig_n} 张图都检出了图注（按图片相邻段落识别）")
    else:
        add("图注编号", "缺失",
            f"{fig_n} 张图中只有 {fig_caps} 张带图注；正文「图N」引用 {fig_refs} 处"
            + (f"、无编号引用 {fig_loose} 处" if fig_loose else "")
            + "。判据：图片相邻段落以「图N/Figure N/图注/来源」开头")

    if tbl_n == 0:
        add("表题编号", "不适用", "无表格，本项不适用（不是「通过」）")
    elif tbl_caps >= tbl_n:
        add("表题编号", "通过", f"{tbl_n} 张表都检出了表题（按表格相邻段落识别）")
    else:
        add("表题编号", "缺失",
            f"{tbl_n} 张表中只有 {tbl_caps} 张带表题；正文「表N」引用 {tbl_refs} 处"
            "。判据：表格前（后）一段以「表N/Table N/表题/来源」开头")

    # ---- 页码：页脚里有 PAGE 域
    has_page = False
    for sec in doc.sections:
        try:
            has_page = has_page or "PAGE" in sec.footer._element.xml
        except (AttributeError, ValueError):
            pass
    add("页码", "通过" if has_page else "缺失",
        "页脚含页码域" if has_page else "页脚未检出 PAGE 域")

    # ---- 页眉
    header = ""
    for sec in doc.sections:
        try:
            header += "".join(p.text for p in sec.header.paragraphs)
        except (AttributeError, ValueError):
            pass
    add("页眉", "通过" if header.strip() else "缺失",
        f"页眉：{header.strip()[:24]}" if header.strip() else "页眉为空")

    # ---- 修订记录 / 文档信息表
    hit = re.search(r"(修订记录|版本记录|变更记录|文档信息|文档修订|修改记录)", body)
    add("修订记录", "通过" if hit else "缺失",
        f"检出「{hit.group(0)}」" if hit else "未检出修订记录/文档信息表")

    # ---- 格式残留
    md = re.findall(r"\*\*[^*\n]{1,40}\*\*", body)
    todo = re.findall(r"(TODO|TBD|待补充|待确认)", body)
    add("格式残留", "无残留" if not (md or todo) else "存在残留",
        "无" if not (md or todo) else
        f"Markdown 粗体 {len(md)} 处、待办标记 {len(todo)} 处（{', '.join((md + todo)[:3])}）")

    return items


DOC_TYPES = {"auto", "公文", "技术文档", "网页", "报告"}


def evaluate(path, doc_type: str = "auto"):
    """doc_type: auto / 公文 / 技术文档 / 网页 / 报告。

    文类影响的是「按什么标准判」——公文必须首行缩进、必须有目录；
    技术文档和网页风格不缩进是常态，短文档没有目录也正常。
    不传就等于按通用标准判，会把技术文档的常规做法记成缺陷。
    """
    if doc_type not in DOC_TYPES:
        raise ValueError(f"doc_type 只支持 {'/'.join(sorted(DOC_TYPES))}，收到 {doc_type}")
    doc = Document(path)
    # 目录条目不是正文：算进段落数会让「样本充分度」和「手刷率」分母虚高，
    # 也会让质量门与文风层报出两个不同的段落数（实测 503 vs 474，差 29 条目录）。
    paras = [p for p in doc.paragraphs if p.text.strip() and not is_toc(p)]
    if not paras:
        return None

    # ---- 0. 样本充分度（前置门槛，不参与打分）
    #      理由：段落太少的文档即使格式规范，也提不出可复用的画像。
    #      这类问题是「能不能用」，不是「有多好」，所以不能只扣分，要一票否决。
    _q = service_config.load()["quality"]
    char_count = sum(len(re.sub(r"\s", "", p.text)) for p in paras)
    insufficient = []
    if len(paras) < _q["min_paragraphs"]:
        insufficient.append(f"有效段落仅 {len(paras)} 个（下限 {_q['min_paragraphs']}）")
    if char_count < _q["min_chars"]:
        insufficient.append(f"正文仅 {char_count} 字（下限 {_q['min_chars']}）")

    # ---- 基准字号：取 Normal 样式定义值，不取封面大标题
    base_size, base_src = style_body_size(doc, paras)
    cover = cover_span(paras, base_size)

    # ---- 1. 伪标题：Normal 样式 + 字号明显大于正文 + 段落短（封面区不算）
    fake_headings = []
    for i, p in enumerate(paras):
        if i in cover or p.style.name not in ("Normal", "正文"):
            continue
        t = p.text.strip()
        # 「目录」本身常是 Normal + 放大字号，但它不是「伪标题」——否则会与
        # 「有目录 = true」自相矛盾
        if re.fullmatch(r"(目\s*录|contents|CONTENTS)", t):
            continue
        s = run_max_size(p)
        # 形态一：放大字号。原来卡 len<=40，把「20pt + 45 字」的长伪标题漏了；
        # 标题跨两行是常态，放宽到 60。
        if s and s >= base_size * 1.25 and len(t) <= 60:
            fake_headings.append((t[:24], s))
            continue
        # 形态二：只加粗不放大。这类最容易漏，但加粗在正文里也常用于强调，
        # 所以判据要保守：整段加粗 + 短 + 句末无标点 + 不是列表/图注。
        all_bold = all(r.font.bold for r in p.runs if r.text.strip())
        if (all_bold and len(t) <= 24
                and not t.endswith(("。", "；", ":", "：", ",", "，"))
                and not CAPTION_RE.match(t)
                and not re.match(r"^\s*([0-9]+[.、)]|[（(])", t)):
            fake_headings.append((t[:24], None))

    # ---- 2. 手刷格式检测：正文段落里字号偏离样式继承值的比例（封面区不算）
    #      （不能用「非 Normal 段落占比」——长文档正文本来就全是 Normal）
    #     只统计「像正文的段落」：图注/表注/来源/引文常刻意用小字号，
    #     把它们算成「手刷」会让真缺陷被噪声淹没（实测三份文档误判 22 段、真缺陷 0）。
    # 保留正文索引，清单里的「第N段」才会指向文档里的真实位置
    normal_paras = [(i, p) for i, p in enumerate(paras)
                    if i not in cover and p.style.name in ("Normal", "正文")]
    brushed = 0
    brushed_samples = []
    for idx, p in normal_paras:
        t = p.text.strip()
        if len(t) < 10 or CAPTION_RE.match(t):
            continue
        sz = run_max_size(p)
        if sz is None:
            continue
        if sz < base_size * 0.85:      # 明显小字：图注、引文、脚注，不按手刷算
            continue
        if abs(sz - base_size) > 0.6:
            brushed += 1
            if len(brushed_samples) < 11:
                brushed_samples.append(
                    f"正文第{idx + 1}段「{t[:16]}」{sz:g}pt（基准 {base_size:g}pt）")
    brush_rate = brushed / len(normal_paras) if normal_paras else 0

    # ---- 3. 层级连贯性：标题层级序列跳级检测
    levels = []
    for p in paras:
        sn = (p.style.name or "")
        m = re.search(r"(\d+)", sn)
        if sn.lower().startswith(("heading", "标题")) and m:
            levels.append(int(m.group(1)))
    jumps = []
    for a, b in zip(levels, levels[1:]):
        if b - a > 1:
            jumps.append(f"H{a}→H{b}")

    # ---- 4. 编号一致性
    manual_nums = Counter()
    for p in paras:
        t = p.text.strip()
        if is_toc(p) or len(t) > 60:
            continue
        for name, rx in NUM_STYLES.items():
            if rx.match(t):
                manual_nums[name] += 1
                break
    auto_num, auto_hier = count_auto_numbering(doc)

    # ---- 5. 内容完整度：占位符
    # 必须按「独立占位」判断，不能子串匹配：正文里「暂无自动轮转策略」是正常表述，
    # 用 full.count("暂无") 会把它算成残留占位符，凭空扣分。
    holes = Counter()
    for p in paras:
        t = p.text.strip()
        if not t:
            continue
        for ph in PLACEHOLDERS:
            if (t == ph
                    or re.fullmatch(rf"{re.escape(ph)}\s*[:：]?\s*", t)
                    or (len(t) <= 8 and t.endswith(ph))):
                holes[ph] += 1
                break

    # ---- 6. 结构完整度（目录要兼容 TOC 域：域未生成内容时不会有 toc 样式段落）
    has_toc = any(is_toc(p) for p in paras)
    if not has_toc:
        try:
            has_toc = bool(re.search(r"TOC\s*\\", doc.element.xml))
        except (AttributeError, TypeError):
            has_toc = False
    # 只认「前言」这样的固定叫法会误判：技术文档常叫「导读」「产品整体介绍」「背景」。
    # 判据是「第一个一级标题承担了引言作用」，不是名字必须叫前言。
    PREFACE_WORDS = (r"前言|概述|引言|导读|简介|介绍|背景|编制目的|编写目的|"
                     r"文档说明|关于本文档|本文档")
    has_preface = any(re.search(PREFACE_WORDS, p.text) and
                      (p.style.name or "").lower().startswith(("heading", "标题"))
                      for p in paras)

    # ---- 打分
    deductions = []
    # 标题规范性 25（伪标题率 0→0 分，>=20%→25 分）
    fake_ratio = len(fake_headings) / len(paras)
    d1 = min(25, round(fake_ratio / 0.20 * 25)) if fake_ratio else 0
    if d1:
        deductions.append((d1, f"伪标题 {len(fake_headings)} 处（占 {fake_ratio:.1%}），"
                              f"如「{fake_headings[0][0]}」"
                        + (f"{fake_headings[0][1]:g}pt" if fake_headings[0][1]
                           else "仅加粗未放大")))
    # 手刷格式 20（手刷率 0→0 分，>=30%→20 分）
    d2 = min(20, round(brush_rate / 0.30 * 20))
    if d2:
        deductions.append((d2, f"正文有 {brushed} 段手工改过字号（占正文 {brush_rate:.1%}），"
                              "未走样式"))
    # 层级连贯性 15（每次跳级扣 5）
    d3 = min(15, len(jumps) * 5)
    if d3:
        deductions.append((d3, f"层级跳级 {len(jumps)} 次（{jumps[0]}）"))
    # 编号一致性 15
    # 「第X章」是章级编号、「1.1」是节级编号，两者共存属于正常层级体系，不算混用。
    # 真正算缺陷的是出现层级体系之外的形式（如「一、」「（一）」与「1.1」并用）。
    HIER_OK = {"第X章", "阿拉伯点分"}
    forms = set(manual_nums)
    others = forms - HIER_OK
    d4 = 0
    if others:
        n = len(forms)
        d4 = min(15, 7 * max(1, n - 1))
        detail = "、".join(f"{k}×{manual_nums[k]}" for k in sorted(forms))
        deductions.append((d4, f"混用 {n} 种手打编号：{detail}"))
    # 自动编号与手打编号并存，只有在「同一层级」才算缺陷。
    # 章号手打（第1章）+ 列表自动编号（numPr）是完全正常的组合：
    # 章标题通常手打、列表用 Word 自动编号，两者不在一个层级上。
    # 只有手打的是层级编号（阿拉伯点分/第X章）且自动编号也用于层级时才扣。
    # 只有「层级式自动编号」（%1.%2 这种）与手打层级编号撞在同一层才算缺陷。
    # 项目符号列表（List Bullet 之类，numFmt=bullet）和章号并存是正常排版。
    hier_manual = {"第X章", "阿拉伯点分"} & set(manual_nums)
    if auto_hier and hier_manual:
        d4 = min(15, d4 + 8)
        deductions.append((d4, f"层级式自动编号（{auto_hier} 处）与手打的层级编号"
                               f"（{'、'.join(sorted(hier_manual))}）并存，同一层级两套编号"))
    # 内容完整度 15
    d5 = min(15, sum(holes.values()) * 3)
    if d5:
        deductions.append((d5, "残留占位符：" + "、".join(f"{k}×{v}" for k, v in holes.items())))
    # 结构完整度 10
    d6 = 0
    # 技术文档/网页在短篇幅下没有目录是常态（靠标题层级导航），不按公文标准扣分
    toc_required = not (doc_type in ("技术文档", "网页") and len(paras) < 150)
    if not has_toc and toc_required:
        d6 += 5
        deductions.append((5, "无目录"))
    if not has_preface:
        d6 += 5
        deductions.append((5, "无前言/概述章节"))

    score = 100 - sum(d for d, _ in deductions)

    if insufficient:
        verdict = "样本不足（不可用）"
    elif score >= _q["pass"]:
        verdict = "准入"
    elif score >= _q["weighted"]:
        verdict = "降权准入（其规则不计入『一致项』）"
    else:
        verdict = "拒收（进待修复队列）"

    return {
        "file": os.path.basename(path),
        # 样本不足时分数没有意义：3 段文档算出的 90 分和 88 段文档的 90 分不是一回事。
        # 同时给分数和「不可用」两个信号，使用者必然去引用那个数字。
        "score": (None if insufficient else score),
        "verdict": verdict,
        "正文字数": char_count,
        "不足原因": insufficient,
        "段落数": len(paras),
        "基准字号": base_size,
        "基准字号来源": base_src,
        "伪标题数": len(fake_headings),
        "伪标题样例": fake_headings[:3],
        "手刷段落数": brushed,
        "手刷率": round(brush_rate, 3),
        # 只给数量没法核对。给清单（最多 10 条）：段落开头 + 字号，
        # 让人能直接翻到那一段确认是不是真手改。
        "手刷清单": brushed_samples[:10],
        "手刷清单是否完整": len(brushed_samples) <= 10,
        "标题数": len(levels),
        "跳级次数": len(jumps),
        "手打编号": dict(manual_nums),
        "自动编号数": auto_num,
        "层级式自动编号数": auto_hier,
        "占位符": dict(holes),
        "文档类型": doc_type,
        "有目录": has_toc,
        "有前言": has_preface,
        "扣分明细": [{"扣分": d, "原因": r} for d, r in deductions],
        "交付检查": delivery_checks(doc, paras),
        "评分说明": (
            "满分 100，六项权重：标题规范 25 / 手刷格式 20 / 层级连贯 15 / "
            "编号一致 15 / 内容完整 15 / 结构完整 10。扣分权重是经验值，不是行业标准，"
            "因此**不同文档之间的分数不可横向比较**，只能看同一份文档的纵向变化。"
            "「交付检查」不计分——完成度与排版规范性是两回事。"
        ),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--json", dest="json_out")
    args = ap.parse_args()

    rows = []
    for f in args.files:
        if not os.path.exists(f) or not f.lower().endswith(".docx"):
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

    def rank(r):
        bad = ("不可用" in r["verdict"]) or ("拒收" in r["verdict"])
        return (1 if bad else 0, -r["score"])

    print(f"\n{'文档':38s} {'分数':>4s}  判定")
    print("-" * 84)
    for r in sorted(rows, key=rank):
        print(f"{r['file'][:38]:38s} {r['score']:>4d}  {r['verdict']}")

    print("\n扣分明细")
    print("-" * 84)
    for r in sorted(rows, key=rank):
        print(f"\n■ {r['file']}（{r['score']} 分，{r['正文字数']} 字）")
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
