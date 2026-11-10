#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
style_probe.py — 写作风格量化提取器

设计原则：**不输出形容词，只输出可核对的数字与可执行的规则**。
"语言专业严谨"毫无用处；"情态动词'应'出现 128 次，占总情态词的 62%"才能变成规则。

提取维度：
  1. 句法节奏：句长均值/分布/长句占比
  2. 人称与视角：我们/我/系统/用户/贵方 使用频次
  3. 情态强度：应/须/必须/需/宜/可/不得 的分布（决定"要求硬度"）
  4. 面向人群：技术词密度 vs 业务词密度 → 判定技术向/业务向
  5. 结构习惯：段落长度、列表化程度、表格密度、分号并列
  6. 标题句式：动宾式 vs 名词短语式
  7. 篇幅分布

用法：
  python3 style_probe.py a.docx b.docx --out style.md --json style.json
"""

import sys
import os
import re
import json
import argparse
import statistics
from collections import Counter
from datetime import datetime

from docx import Document

# ---------------------------------------------------------------- 词表

# 只放中文术语。英文术语全部在 TECH_TERMS_EN 里——
# 两边都放同一批词（原来 API/JSON/SQL/SDK 都在两边），统计时会各算一次，
# 实测「API」被算成 8 次（真实 4 次）。
TECH_TERMS = [
    "接口", "数据库", "字段", "表结构", "索引", "缓存", "中间件", "服务", "微服务",
    "部署", "集群", "线程", "并发", "日志", "报文",
    "鉴权", "加密", "网关", "容器", "消息队列", "前端", "后端",
    "算法", "模型", "存储", "带宽", "负载", "事务", "回滚", "幂等", "超时", "重试",
]

BIZ_TERMS = [
    "用户", "客户", "业务", "流程", "审批", "订单", "登记", "查询", "统计", "报表",
    "权限", "角色", "场景", "需求", "产品", "运营", "交付", "对账", "结算", "合同",
    "目录", "资源", "节点", "认证", "溯源", "监测", "大屏", "门户", "工作台", "待办",
]

# ---------------------------------------------------------------- 英文术语
# 中文文档夹英文是常态，但英文术语永远追不上词表（新框架每天在出）。
# 所以分两层：
#   词表命中 —— 已知术语，可精确计数
#   形态识别 —— 驼峰/全大写/带下划线点号的标识符、以及常见技术后缀，
#              不依赖词表也能认出「这是个技术标识」
# 词表命中的数字可以直接引用；形态识别只能说明「存在技术标识」，不能当计数用。

TECH_TERMS_EN = [
    # 架构与协议
    "Agent", "MCP", "API", "SDK", "REST", "GraphQL", "RPC", "gRPC", "HTTP", "HTTPS",
    "WebSocket", "SSE", "OAuth", "JWT", "CORS", "CDN", "DNS", "TCP", "UDP", "TLS",
    # 数据与存储
    "SQL", "NoSQL", "SQLite", "PostgreSQL", "MySQL", "Redis", "MongoDB", "Kafka",
    "RabbitMQ", "ETL", "OLAP", "OLTP", "Schema", "Index", "Cache", "Buffer",
    # 运行时与部署
    "Docker", "Kubernetes", "K8s", "microVM", "VM", "Container", "Pod", "Helm",
    "CI/CD", "DevOps", "SRE", "Sandbox", "Runtime", "Daemon", "Cron",
    # AI / Agent
    "LLM", "Prompt", "Token", "Embedding", "RAG", "Fine-tune", "Inference",
    "Context", "Tool", "Skill", "Workflow", "Pipeline", "Orchestrator",
    # 工程通用
    "Config", "Deploy", "Debug", "Log", "Metric", "Trace", "Monitor", "Alert",
    "Test", "Mock", "Stub", "Benchmark", "Latency", "Throughput", "QPS", "TPS",
    "JSON", "XML", "YAML", "TOML", "CSV", "Markdown", "HTML", "CSS", "JS",
]

# 常见技术后缀：靠它判断「这是个技术名词」而不是普通英文单词
TECH_SUFFIX_RE = re.compile(
    r"\b[A-Za-z][A-Za-z0-9_.\-]*?(?:API|SDK|DB|SQL|VM|IO|UI|UX|ML|AI|LLM|"
    r"Service|Server|Client|Manager|Handler|Loader|Runner|Worker|Adapter|"
    r"Provider|Factory|Builder|Controller|Repository|Gateway|Proxy|"
    r"Config|Context|Session|Token|Schema|Index|Cache|Queue|Pool|Lock)\b",
    re.I,
)

# 驼峰 / 全大写缩写 / 带下划线或点号的标识符
IDENT_RE = re.compile(r"\b(?:[a-z]+[A-Z][A-Za-z0-9]*|[A-Z]{2,}[A-Za-z0-9]*|"
                      r"[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+|"
                      r"[A-Za-z][A-Za-z0-9]*\.[a-z]{2,4})\b")

# 整段英文：连续 3 个以上纯 ASCII 单词
EN_SENTENCE_RE = re.compile(r"(?:\b[A-Za-z][A-Za-z'\-]*\b[\s,.;:!?]+){3,}")

# 情态动词 → 强度分级
MODAL_STRONG = ["必须", "不得", "禁止", "不允许", "严禁"]
# 「需要」不在硬性要求里：它同时是「需」的复合词，两套规则并存必然自相矛盾
# （实测 100 次「需要」被计成硬性要求 100 次，而「需」的明细显示真情态 0）。
# 统一规则：多字词只作为单字词的「排除项」，不单独计入情态统计。
MODAL_REQUIRED = ["应", "须", "需"]
MODAL_SOFT = ["宜", "建议", "可考虑", "优先"]
MODAL_ALLOW = ["可以", "可", "支持", "允许"]

PRONOUNS = ["我们", "我司", "我方", "本公司", "本平台", "贵方", "贵司", "您", "你",
            "本系统", "系统", "平台", "用户", "管理员", "第三方", "甲方", "乙方"]

# 单字情态词的复合词黑名单。
# 必须排除：直接 text.count("需") 会把「需求」「需要」里的「需」全部算成情态动词，
# 实测交维需求说明书「需」原始 152 次里有 141 次是复合词，硬性要求占比直接从
# 24% 虚高到 60%，结论完全翻转。
MODAL_EXCLUDE = {
    "需": ["需求", "需要", "必需", "无需", "所需", "供需", "急需", "内需",
           "需方", "需方", "需求方"],          # 「需方」是业务名词（我是需方），不是情态动词
    "应": ["对应", "响应", "适应", "供应", "应用", "相应", "反应", "应答", "应缴", "应标"],
    "可": ["可能", "可见", "可视", "可靠", "可行", "可控", "可用", "可扩展", "可配置",
           "可维护", "可读", "可追溯", "认可", "许可", "可以", "可否", "可比"],
    "须": ["必须"],
    "宜": ["适宜", "便宜"],
}


def count_modal(text, word):
    """
    统计情态词出现次数。单字情态词需排除复合词里的同字。
    两字及以上（支持/需要/必须/不得）直接计数即可。
    """
    if len(word) > 1:
        return text.count(word)
    comps = MODAL_EXCLUDE.get(word, [])
    spans = []
    for c in comps:
        for m in re.finditer(re.escape(c), text):
            spans.append((m.start(), m.end()))
    return sum(1 for m in re.finditer(re.escape(word), text)
               if not any(s <= m.start() < e for s, e in spans))

SENT_SPLIT = re.compile(r"[。！？!?]+")
CLAUSE_SPLIT = re.compile(r"[，,；;]")


def load_parts(path):
    """把文档拆成 (正文段落, 表格单元格文本) 两部分。

    必须拆开：表格单元格不是段落。混在一起算，平均段长/平均句长/短句占比全部失真
    ——单元格普遍十几个字，会把「正文长段」拉成「短句成段」，结论方向都反了。
    实测一份 88 段正文 + 307 格的文档：混算后段长 47.8 字，实际 165.6 字。
    """
    doc = Document(path)
    paras, cells = [], []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        sname = (p.style.name or "").lower()
        if sname.startswith("toc") or sname in ("toc 1", "toc 2", "toc 3"):
            continue
        paras.append(t)
    for tb in doc.tables:
        for row in tb.rows:
            for cell in row.cells:
                for line in cell.text.split("\n"):
                    t = line.strip()
                    if t:
                        cells.append(t)
    return paras, cells


def load_text(path):
    """只要正文段落（不含表格）。表格文本请用 load_parts 单独取。"""
    return load_parts(path)[0]


def analyze(paras, headings=None, table_text=""):
    """headings: [(level, text)] 用于标题句式分析。

    table_text: 表格单元格文本（可选）。**只参与词汇统计，不参与段落/句子统计**。
    表格里没有"段落"和"句子"的概念，混进来会把段长句长算错。
    """
    body = [p for p in paras for p in [p.strip()] if p]
    full_body = "\n".join(body)                    # 正文：句法/段落指标只用它
    full = (full_body + "\n" + table_text).strip()  # 全文：词汇指标用它

    body_chars = len(re.sub(r"\s", "", full_body))
    table_chars = len(re.sub(r"\s", "", table_text))
    total_chars = body_chars + table_chars

    # 句子切分
    sentences = []
    for p in body:
        for s in SENT_SPLIT.split(p):
            s = s.strip()
            if len(s) >= 2:
                sentences.append(s)
    lens = [len(s) for s in sentences]

    # 句内逗号数（判断"长链句"）
    commas = [len(CLAUSE_SPLIT.findall(s)) for s in sentences]

    tech_hits = Counter()
    for t in TECH_TERMS:
        n = full.count(t)
        if n:
            tech_hits[t] = n
    # 英文术语：大小写不敏感计数（文档里常写 agent / Agent / AGENT），
    # 但词表按「首字母大写的标准写法」记录，输出时统一成词表写法
    # 英文必须按词边界匹配，不能用子串：中文可以子串（无词形变化），
    # 英文子串会把 testing 算成 test、login 算成 log、toolbar 算成 tool。
    # 实测一份纯英文文档因 "testing" 被判成「技术向」、技术词占比 100%。
    for t in TECH_TERMS_EN:
        n = len(re.findall(rf"(?<![A-Za-z]){re.escape(t)}(?![A-Za-z])", full, re.I))
        if n:
            tech_hits[t] += n

    # 形态识别：不依赖词表，找「看起来像技术标识」的英文 token
    # 只算词表没覆盖到的，避免重复计数
    known_lower = {t.lower() for t in TECH_TERMS + TECH_TERMS_EN}
    ident_hits = Counter()
    for m in IDENT_RE.finditer(full):
        tok = m.group(0)
        if tok.lower() in known_lower:
            continue
        if len(tok) < 3 or tok.isdigit():
            continue
        ident_hits[tok] += 1

    en_chars = sum(1 for ch in full if ch.isascii() and ch.isalpha())
    en_ratio = round(en_chars / total_chars, 3) if total_chars else 0

    # 整段英文：连续 3 个以上 ASCII 单词的段落。中英混排里这类段落通常是
    # 引用原文、代码注释、接口说明——它们的技术词会拉高密度，但要单独可查。
    en_paras = [p for p in body if EN_SENTENCE_RE.search(p)
                and not re.search(r"[\u4e00-\u9fff]", p)]
    en_para_ratio = round(len(en_paras) / len(body), 3) if body else 0

    # 仅含中文的段落里切出的句长中位数：混排时用它做对照，避免被英文拉偏
    zh_lens = []
    for p in body:
        if EN_SENTENCE_RE.search(p) and not re.search(r"[\u4e00-\u9fff]", p):
            continue
        for s_ in SENT_SPLIT.split(p):
            s_ = s_.strip()
            if len(s_) >= 2:
                zh_lens.append(len(s_))
    zh_sent_median = round(statistics.median(zh_lens), 1) if zh_lens else None
    biz_hits = Counter()
    for t in BIZ_TERMS:
        n = full.count(t)
        if n:
            biz_hits[t] = n

    tech_total = sum(tech_hits.values()) + sum(ident_hits.values())
    biz_total = sum(biz_hits.values())
    denom = tech_total + biz_total
    tech_ratio = round(tech_total / denom, 3) if denom else None
    # 分母为 0 时不能判人群：一份完全没命中技术词/业务词的文档，
    # 说它「业务向（技术实现被隐藏）」是编的——那是「没测到」，不是「隐藏」。
    if tech_ratio is None:
        if en_ratio > 0.3:
            audience = (f"无法判定（英文字母占 {en_ratio:.0%}，词表以中文为主，"
                        "建议按技术向理解，但本工具不做明确判定）")
        else:
            audience = "无法判定（技术词与业务词均未命中，词表未覆盖该领域用词）"
    elif tech_ratio >= 0.5:
        audience = "技术向（面向开发/实施，可直接出现技术名词）"
    elif tech_ratio >= 0.3:
        audience = "技术偏中性（业务人员能读懂，但夹带技术名词）"
    else:
        audience = "业务向（面向业务人员，技术实现被隐藏）"

    # 情态动词
    modals = {}
    for label, words in (("强禁令（必须/不得）", MODAL_STRONG),
                         ("硬性要求（应/须/需）", MODAL_REQUIRED),
                         ("建议性（宜/建议）", MODAL_SOFT),
                         ("可选（可/支持）", MODAL_ALLOW)):
        c = Counter()
        for w in words:
            n = count_modal(full, w)
            if n:
                c[w] = n
        modals[label] = dict(c.most_common())

    modal_total = sum(sum(v.values()) for v in modals.values())
    strong = sum(modals["强禁令（必须/不得）"].values())
    required = sum(modals["硬性要求（应/须/需）"].values())
    modal_hard_ratio = round((strong + required) / modal_total, 3) if modal_total else 0

    # ---- 口径自检：单字情态词被复合词稀释的程度
    # 这是最容易出错的地方。「需」原始 152 次里 141 次是「需求/需要」，
    # 不报出来，读的人会以为硬性要求占比 60%，实际只有 25%。
    modal_excl = {}
    for w in ("需", "应", "可", "须", "宜"):
        raw = full.count(w)
        if not raw:
            continue
        real = count_modal(full, w)
        comps = [c for c in MODAL_EXCLUDE.get(w, []) if full.count(c)]
        top_comp = max(comps, key=lambda c: full.count(c)) if comps else None
        modal_excl[w] = {"原始出现": raw, "真情态": real,
                         "复合词排除率": round(1 - real / raw, 3),
                         "主要复合词": (f"{top_comp}×{full.count(top_comp)}" if top_comp else None)}

    # 人称视角
    pron = {}
    for w in PRONOUNS:
        n = full.count(w)
        if n:
            pron[w] = n

    # 结构习惯
    list_like = sum(1 for p in body if re.match(
        r"^\s*([0-9]+[\.、)）]|[a-zA-Z][\.\)]|[ivxIVX]+[\.\)]"
        r"|[（(][一二三四五六七八九十\d]+[)）]|[•·\-*])", p))
    table_cells = sum(1 for p in body if "\t" in p)
    para_lens = [len(p) for p in body]

    # 标题句式
    verb_obj = 0
    noun_phrase = 0
    verb_starts = ("实现", "支持", "提供", "完成", "建立", "设计", "优化", "新增", "改造",
                   "对接", "配置", "查询", "查看", "导出", "导入", "登记", "发布")
    heads = [h for _, h in (headings or [])]
    for h in heads:
        if any(h.startswith(v) for v in verb_starts):
            verb_obj += 1
        else:
            noun_phrase += 1

    # 只分「动宾/名词短语」两档会得出「48:0 完全统一」这种没信息量的结论
    # （「产品定位」「限制维度一览」「循环防护机制 (loop_guard)」都被归成一类）。
    # 补一组可核对的形态特征，用来看真实差异在哪。
    head_feat = {
        "标题数": len(heads),
        "平均字数": round(statistics.mean([len(h) for h in heads]), 1) if heads else 0,
        "最短字数": min((len(h) for h in heads), default=0),
        "最长字数": max((len(h) for h in heads), default=0),
        "含英文或型号": sum(1 for h in heads if re.search(r"[A-Za-z][A-Za-z0-9_.\-]*", h)),
        "带括号补充": sum(1 for h in heads if re.search(r"[（(]", h)),
        "疑问式": sum(1 for h in heads if h.rstrip().endswith(("？", "?"))),
        "短语式(>4字)": sum(1 for h in heads if len(h) > 4),
        "单词式(≤4字)": sum(1 for h in heads if len(h) <= 4),
        "分类口径": "动宾/名词两档只按首词是否强动词判断，粒度粗；"
                    "「管理」「配置」这类既是动词也是名词，会误分。请看上面这组形态特征。",
    }

    # total_chars 已在开头按「正文 + 表格」算好，这里不要再覆盖

    # 离群块检测：单「句」超长通常不是句子，而是表格转文本/代码/无标点罗列
    outlier_blocks = [n for n in lens if n > 500]
    # 截尾均值：去掉最长 1% 后重算，避免离群值把均值拉偏
    trimmed_mean = None
    if lens:
        srt = sorted(lens)
        cut = max(1, int(len(srt) * 0.01))
        core = srt[:-cut] if len(srt) > cut else srt
        trimmed_mean = round(statistics.mean(core), 1)

    mean_len = round(statistics.mean(lens), 1) if lens else 0
    sd = round(statistics.pstdev(lens), 1) if len(lens) > 1 else 0
    reliability = "高"
    if sd > 2 * mean_len or outlier_blocks:
        reliability = "低（存在超长无标点块，均值不可信，请看截尾均值与中位数）"

    # ---- 自检告警：让失真的指标自己说出来，而不是等人发现
    checks = []
    for w, d in modal_excl.items():
        if d["原始出现"] >= 5 and d["复合词排除率"] >= 0.6:
            checks.append({
                "级别": "高", "指标": f"情态词「{w}」",
                "说明": f"原始 {d['原始出现']} 次中 {d['原始出现'] - d['真情态']} 次属复合词"
                        f"（主要是「{d['主要复合词'] or '复合词'}」），"
                        f"仅 {d['真情态']} 次是真情态动词。"
                        f"硬性要求占比已按真情态计算，但不要引用原始词频。",
            })
    if total_chars > 5000 and tech_total / total_chars * 1000 < 1:
        checks.append({
            "级别": "中", "指标": "术语暴露度",
            "说明": f"{total_chars} 字文档技术词密度仅 {tech_total / total_chars * 1000:.1f}/千字，"
                    "词表可能与该领域不匹配，实际暴露度被低估（只会低估，不会高估）",
        })
    # 句长按字符算：英文单词平均 5 个字符才等于 1 个信息单位，
    # 中文 1.5 个字符就是 1 个字。混排文档的句长均值会被英文段落拉高。
    if en_para_ratio > 0.1:
        checks.append({
            "级别": "中", "指标": "句长（中英混排）",
            "说明": f"有 {len(en_paras)} 段纯英文（占 {en_para_ratio:.0%}）。"
                    "句长按字符统计，英文单词的字符数远多于中文，"
                    "平均句长会被拉高，请对照「中文段落句长中位数」判断，不要直接用均值",
        })
    if not pron:
        checks.append({"级别": "中", "指标": "人称",
                       "说明": "未检出任何人称词，可能正文提取不全，或完全去人称化"})
    if modal_total == 0:
        checks.append({"级别": "中", "指标": "情态",
                       "说明": "未检出情态动词，该文档可能不是规范性文体（说明/汇报/分析类）"})
    if mean_len and sd > 2 * mean_len:
        checks.append({"级别": "高", "指标": "句长",
                       "说明": "标准差超过均值 2 倍，平均值不可信，请用截尾均值与中位数"})

    return {
        "总字数": total_chars,
        "正文字数": body_chars,
        "表格字数": table_chars,
        "段落数": len(body),                    # 只数正文段落
        "表格单元格数": len([c for c in table_text.split("\n") if c.strip()]),
        "句子数": len(sentences),
        "平均句长": mean_len,
        "截尾平均句长": trimmed_mean,
        "离群块数量(>500字)": len(outlier_blocks),
        "指标可信度": reliability,
        "句长中位数": round(statistics.median(lens), 1) if lens else 0,
        "句长标准差": round(statistics.pstdev(lens), 1) if len(lens) > 1 else 0,
        "超长句占比(>60字)": round(sum(1 for n in lens if n > 60) / len(lens), 3) if lens else 0,
        "短句占比(<15字)": round(sum(1 for n in lens if n < 15) / len(lens), 3) if lens else 0,
        "平均逗号链": round(statistics.mean(commas), 2) if commas else 0,
        "平均段长": round(statistics.mean(para_lens), 1) if para_lens else 0,
        "段落长度标准差": round(statistics.pstdev(para_lens), 1) if len(para_lens) > 1 else 0,
        "列表化段落占比": round(list_like / len(body), 3) if body else 0,
        "含制表符段落数": table_cells,
        "情态词": modals,
        "情态词总数": modal_total,
        "复合词排除明细": modal_excl,
        "自检告警": checks,
        "口径说明": {
            "情态词": "单字情态词（需/应/可/须/宜）已排除复合词（需求/对应/可见/可以等）内的同字，"
                     "只统计独立成词的情态用法。硬性要求占比按排除后的真情态计算。",
            "句长": "按 。！？ 切分；截尾均值去掉最长 1%，用于抵抗表格转文本/代码等超长无标点块。",
            "技术词密度": "中文按内置词表匹配，英文按词表 + 形态识别（驼峰/全大写/"
                          "含下划线或点号）。词表仍然是穷举的，新框架、自造名词会漏计，"
                          "但形态识别能兜住一部分；具体漏了哪些见「英文混入」里的"
                          "「词表未覆盖的技术标识」。",
            "中英混排": "英文术语分两层：词表命中可精确计数；形态识别只说明存在技术标识，"
                        "不能当计数引用。句长按字符算，英文段落会拉高均值，"
                        "混排文档请看「中文段落句长中位数」。",
            "列表化段落占比": "按段落是否以数字/字母/罗马数字/圆点/括号编号开头判定。",
            "段落与句子": "只用正文段落，不含表格单元格。表格没有段落/句子概念，"
                         "混算会把段长句长算错（实测 88 段正文 + 307 格混算后段长从 165.6 掉到 47.8）。",
            "词汇统计范围": "技术词/业务词/情态词/人称按「正文 + 表格」全文统计，"
                           "密度分母用总字数；表格文本不计入段落与句子。",
        },
        "硬性要求占比": modal_hard_ratio,
        "情态词说明": "「情态词」里各项已是排除复合词后的真情态，可直接引用。"
                     "「复合词排除明细」里的「原始出现」仅供核对，**不要引用**："
                     "如「应」原始 8 次可能只有 2 次是情态动词（其余是「应用/对应」），"
                     "「须」常被「必须」占满，而「必须」已单独计入强禁令。",
        "人称使用": pron,
        "技术词密度(每千字)": round(tech_total / total_chars * 1000, 2) if total_chars else 0,
        "业务词密度(每千字)": round(biz_total / total_chars * 1000, 2) if total_chars else 0,
        "技术词占比": tech_ratio,   # None 表示未测到任何词，不是 0%
        "面向人群判定": audience,
        "技术词Top": dict(tech_hits.most_common(12)),
        "中文段落句长中位数": zh_sent_median,
        "英文混入": {
            "英文字母占比": en_ratio,
            "整段英文段落数": len(en_paras),
            "整段英文占比": en_para_ratio,
            "词表未覆盖的技术标识": dict(ident_hits.most_common(10)),
            "说明": "英文术语分两层统计：词表命中可精确计数；"
                    "「词表未覆盖的技术标识」按形态识别（驼峰/全大写/含下划线或点号），"
                    "只能说明存在技术标识，不能当精确计数引用。",
        },
        "业务词Top": dict(biz_hits.most_common(12)),
        "标题动宾式": verb_obj,
        "标题名词短语式": noun_phrase,
        "标题特征": head_feat,
    }


def render(results):
    L = ["# 写作风格量化指纹", ""]
    L.append(f"> 生成时间 {datetime.now():%Y-%m-%d %H:%M}")
    L.append("")
    L.append("> 说明：本表只给可核对的数字。形容词层面的判断请结合数字自行确认，"
             "或交由模型按本表生成规则。")
    L.append("")

    if len(results) > 1:
        L.append("## 跨文档对比\n")
        keys = ["总字数", "平均句长", "截尾平均句长", "指标可信度", "句长标准差",
                "超长句占比(>60字)", "平均段长",
                "列表化段落占比", "硬性要求占比", "技术词密度(每千字)", "面向人群判定"]
        L.append("| 指标 | " + " | ".join(r["file"] for r in results) + " |")
        L.append("|---" * (len(results) + 1) + "|")
        for k in keys:
            L.append(f"| {k} | " + " | ".join(str(r["data"][k]) for r in results) + " |")
        L.append("")

    for r in results:
        d = r["data"]
        L.append(f"## {r['file']}")
        L.append("")
        L.append("### 1. 句法节奏")
        L.append("")
        if d.get("自检告警"):
            L.append("> ⚠️ **自检告警**")
            for c in d["自检告警"]:
                L.append(f"> - [{c['级别']}] **{c['指标']}**：{c['说明']}")
            L.append("")
        L.append(f"- 平均句长 **{d['平均句长']} 字**（中位 {d['句长中位数']}，标准差 {d['句长标准差']}）")
        L.append(f"- 截尾平均句长（去最长 1%）**{d['截尾平均句长']} 字** ｜ 指标可信度：{d['指标可信度']}")
        if d["离群块数量(>500字)"]:
            L.append(f"- ⚠️ 检出 **{d['离群块数量(>500字)']} 个超长无标点块**（>500字，通常是表格转文本、代码或罗列），"
                     "均值已被拉偏，请以截尾均值/中位数为准")
        L.append(f"- 超长句（>60字）占 **{d['超长句占比(>60字)']:.1%}**，短句（<15字）占 **{d['短句占比(<15字)']:.1%}**")
        L.append(f"- 平均每句 {d['平均逗号链']} 个逗号（数值高说明爱用长链句、少用句号断句）")
        L.append(f"- 平均段长 {d['平均段长']} 字（标准差 {d['段落长度标准差']}）")
        L.append("")
        L.append("### 2. 面向人群")
        L.append("")
        L.append(f"- 判定：**{d['面向人群判定']}**")
        _tr = d.get("技术词占比")
        L.append(f"- 技术词密度 {d['技术词密度(每千字)']}/千字，业务词密度 {d['业务词密度(每千字)']}/千字，"
                 f"技术词占比 {'未测到（词表未覆盖该领域用词）' if _tr is None else f'{_tr:.1%}'}")
        _en = d.get("英文混入") or {}
        if _en.get("英文字母占比"):
            L.append(f"- 中英混排：英文字母占 {_en['英文字母占比']:.0%}，"
                     f"整段英文 {_en.get('整段英文段落数', 0)} 段"
                     f"（{_en.get('整段英文占比', 0):.0%}）；"
                     f"中文段落句长中位数 {d.get('中文段落句长中位数')} 字")
            _unk = _en.get("词表未覆盖的技术标识") or {}
            if _unk:
                L.append("- 词表未覆盖的技术标识：" +
                         "、".join(f"{k}×{v}" for k, v in list(_unk.items())[:8]))
        if d["技术词Top"]:
            L.append("- 技术词 Top：" + "、".join(f"{k}×{v}" for k, v in d["技术词Top"].items()))
        if d["业务词Top"]:
            L.append("- 业务词 Top：" + "、".join(f"{k}×{v}" for k, v in d["业务词Top"].items()))
        L.append("")
        L.append("### 3. 要求硬度（情态动词）")
        L.append("")
        L.append(f"- 情态词总数 {d['情态词总数']}，其中硬性要求（应/须/需/必须/不得）占 **{d['硬性要求占比']:.1%}**")
        for label, c in d["情态词"].items():
            if c:
                L.append(f"- {label}：" + "、".join(f"`{k}`×{v}" for k, v in c.items()))
        if d.get("复合词排除明细"):
            L.append("")
            L.append("| 单字情态词 | 原始出现 | 真情态 | 复合词排除率 |")
            L.append("|---|---|---|---|")
            for w, x in d["复合词排除明细"].items():
                L.append(f"| {w} | {x['原始出现']} | {x['真情态']} | {x['复合词排除率']:.0%} |")
            L.append("")
            L.append("> 上表说明：原始词频里含大量「需求」「对应」「可见」这类复合词，"
                     "**硬性要求占比已按排除后的真情态计算**。")
        L.append("")
        L.append("### 4. 视角与人称")
        L.append("")
        if d["人称使用"]:
            L.append("、".join(f"`{k}`×{v}" for k, v in sorted(d["人称使用"].items(), key=lambda x: -x[1])))
        else:
            L.append("_无人称词出现，属完全去人称化写作_")
        L.append("")
        L.append("### 5. 结构习惯")
        L.append("")
        L.append(f"- 列表化段落占比 **{d['列表化段落占比']:.1%}**（越高越爱用编号/项目符号而非连续叙述）")
        L.append(f"- 标题中动宾式 {d['标题动宾式']} 个 / 名词短语式 {d['标题名词短语式']} 个")
        L.append("")
        L.append("---")
        L.append("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="写作风格量化提取器")
    ap.add_argument("files", nargs="+")
    ap.add_argument("--out")
    ap.add_argument("--json", dest="json_out")
    args = ap.parse_args()

    results = []
    for f in args.files:
        if not f.lower().endswith(".docx"):
            print(f"[跳过] 仅支持 docx（PPT 文本提取待接入）：{f}", file=sys.stderr)
            continue
        if not os.path.exists(f):
            print(f"[跳过] 不存在：{f}", file=sys.stderr)
            continue
        try:
            doc = Document(f)
            headings = []
            for p in doc.paragraphs:
                sn = (p.style.name or "").lower()
                if sn.startswith("heading") or sn.startswith("标题"):
                    if p.text.strip():
                        headings.append((sn, p.text.strip()))
            data = analyze(load_text(f), headings)
            results.append({"file": os.path.basename(f), "data": data})
            print(f"[完成] {os.path.basename(f)}", file=sys.stderr)
        except Exception as e:  # noqa: BLE001
            print(f"[失败] {os.path.basename(f)}：{type(e).__name__} {e}", file=sys.stderr)

    if not results:
        return 1
    md = render(results)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(md)
        print(f"[报告] {args.out}", file=sys.stderr)
    else:
        print(md)
    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as fh:
            json.dump(results, fh, ensure_ascii=False, indent=2)
        print(f"[JSON] {args.json_out}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
