#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DSWiki-MCP — 文档风格/格式画像提取 MCP 服务

设计分工（关键）：
  本服务只提供「硬指标」——可数、可复现、可核对的数字。
  语义判断（规则命名、内容套路、禁区识别）由调用的 agent 完成。
  服务不生成任何风格形容词，也不替 agent 下结论。

  理由：LLM 数句长两次结果不一样；代码算不出「这是一套条件句叙述法」。
  所以代码是标尺，LLM 是写评语的人，标尺不能被评语的人修改。

对外 9 个工具（按任务动作分组，不是按对象的增删改查拆）：
  analyze_document    一份文档：质量门 + 排版/视觉 + 文风（按扩展名分流 Word/PPT）
  analyze_batch       一批文档：质量门筛选 → 聚类判几类 →（可选）一致性对比
  get_profile         传 id 读一份，不传列清单
  save_profile        画像落库（强制每条规则带证据，没证据拒收）
  edit_rules          改单条规则：add / update / delete
  manage_profile      状态与元信息：promote / retire / clone / update_meta /
                      resolve_conflict / export
  merge_profiles      合并拆错的画像
  check_profiles      画像库体检（九项）
  service_control     运维：status / reload / config / set

内部的 extract_format / extract_voice / check_quality / add_rule / promote_profile …
仍然保留，但不再注册为工具，只作为上面这些入口的实现。

启动：python3 server.py（stdio）
"""

import os
import re
import sys
import json
import time
import functools
import importlib
from collections import Counter
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent                      # 项目根目录
sys.path.insert(0, str(ROOT))

import docx_probe            # noqa: E402
import style_probe           # noqa: E402
import quality_gate          # noqa: E402
import pptx_probe            # noqa: E402
import cluster_gate          # noqa: E402
import batch_compare         # noqa: E402
import version as engine_version  # noqa: E402
import profile_doctor        # noqa: E402
import service_config        # noqa: E402
from difflib import SequenceMatcher  # noqa: E402
from docx import Document    # noqa: E402

from mcp.server.mcpserver import MCPServer   # noqa: E402

# 画像库放在稳定的用户目录，不放工作区。
# 原来绑在 ROOT/profiles（即某个会话工作区 2026-11-02-14-06-04 下），
# 换一个会话、或清理旧工作区，画像就整体消失——数据不能跟着会话走。
# 需要改位置时用环境变量 DOC_STYLE_HOME 覆盖。
DATA_HOME = Path(os.environ.get("DOC_STYLE_HOME",
                                Path.home() / ".workbuddy" / "doc-style")).expanduser()
PROFILE_ROOT = service_config.PROFILE_ROOT   # 单一来源，别再各写一份

# 画像的作用域。不同 scope 的寿命和门槛不一样：
#   通用 —— 跨项目复用，几乎不过期，所以要求更多规则、必须写适用描述
#   项目 —— 只对某个项目有效，项目结束就该废弃，按默认门槛
#   临时 —— 试验用，不允许转正，避免"临时画像"混进生成侧
SCOPES = ("通用", "项目", "临时")
if not PROFILE_ROOT.exists():
    PROFILE_ROOT.mkdir(parents=True)   # 此处 _ensure_dir 还没定义，直接建

# 旧位置（工作区内）的画像一次性迁移过来，不覆盖已有的
_LEGACY_ROOT = ROOT / "profiles"
if _LEGACY_ROOT.exists():
    try:
        for _sub in _LEGACY_ROOT.iterdir():
            _target = PROFILE_ROOT / _sub.name
            if _sub.is_dir() and not _sub.name.startswith(".") and not _target.exists():
                import shutil as _shutil
                _shutil.move(str(_sub), str(_target))
        _legacy_audit = ROOT / "audit.jsonl"
        if _legacy_audit.exists() and not (DATA_HOME / "audit.jsonl").exists():
            import shutil as _shutil
            _shutil.move(str(_legacy_audit), str(DATA_HOME / "audit.jsonl"))
    except OSError:
        pass    # 迁移失败不能挡住服务启动


_ENGINE_MODULES = ("docx_probe", "style_probe", "quality_gate", "pptx_probe",
                   "pptx_quality_gate", "cluster_gate", "batch_compare",
                   "profile_doctor", "version", "service_config")
_MTIME: dict = {}


def _maybe_reload() -> list:
    """
    引擎模块改了 mtime 就重新加载，省掉「改代码→重启 MCP」这一轮。

    背景：MCP 服务是宿主持有的常驻子进程，不会随代码改动更新。
    之前每次改统计口径都要用户去连接器页 toggle，而且重启 app 都未必生效。
    这里在每个工具调用前比一次 mtime，变了就 importlib.reload。

    只 reload 引擎模块，不动 server.py 自己——自己改了还是得重启一次。
    """
    changed = []
    for name in _ENGINE_MODULES:
        mod = sys.modules.get(name)
        if mod is None:
            continue
        f = getattr(mod, "__file__", None)
        if not f:
            continue
        try:
            m = os.path.getmtime(f)
        except OSError:
            continue
        if name in _MTIME and m > _MTIME[name] + 0.5:
            changed.append(name)
        _MTIME[name] = m
    for name in changed:
        try:
            importlib.reload(sys.modules[name])
        except Exception:  # noqa: BLE001
            # 重载失败不能影响工具本身，保留旧版本继续跑
            _MTIME.pop(name, None)
    return changed


def _ensure_dir(p: Path) -> Path:
    """某些沙箱环境下 mkdir(exist_ok=True) 仍会抛 EEXIST，这里显式判存在。"""
    if not p.exists():
        p.mkdir(parents=True, exist_ok=True)
    return p


# 运行指标：audit.jsonl 记「谁改了什么」，这里记「跑得怎么样」。
# 两者分开，混在一起会让审计日志变得没法读（调用噪声淹没变更记录）。
# 只统计对外暴露的工具。内部函数（check_quality / extract_format 等）也被 _safe
# 包装，一并记录会让计数虚高——analyze_document 一次调用会连带记 3 条。
# 新增工具时记得同步这里；self_test 里有断言兜底。
_TOOL_NAMES = {"analyze_document", "analyze_batch", "get_profile", "save_profile",
               "edit_rules", "manage_profile", "merge_profiles", "check_profiles",
               "service_control"}

CALL_LOG = service_config.DATA_HOME / "calls.jsonl"
_CALL_MAX = 5000          # 滚动上限：超过就从旧到新截断
_CALL_SINCE_TRIM = 0


def _record_call(tool: str, ms: float, status: str) -> None:
    global _CALL_SINCE_TRIM
    try:
        with CALL_LOG.open("a", encoding="utf-8") as f:
            f.write(json.dumps({"t": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "tool": tool, "ms": round(ms, 1),
                                "s": status}, ensure_ascii=False) + "\n")
        _CALL_SINCE_TRIM += 1
        # 每积累 200 条才检查一次行数，避免每次调用都读整个文件
        if _CALL_SINCE_TRIM >= 200:
            _CALL_SINCE_TRIM = 0
            lines = CALL_LOG.read_text(encoding="utf-8").splitlines()
            if len(lines) > _CALL_MAX:
                CALL_LOG.write_text("\n".join(lines[-_CALL_MAX:]) + "\n",
                                    encoding="utf-8")
    except OSError:
        pass    # 采不到指标不能影响工具本身


def _verdict_of(result) -> str:
    """把「调用成功」和「业务上被拒」分开。

    两者都返回 200，但含义完全不同：前者说明服务在干活，
    后者说明调用方用错了。混成一个成功率会掩盖真实问题。
    """
    if not isinstance(result, dict):
        return "ok"
    if result.get("error"):
        # 这个分支是「函数正常返回、内容是错误」——说明服务预料到了这类问题，
        # 属于调用方输入不对（看文档能解决）。真正没预料到的异常走 except 分支，
        # 在那里标 bug。两者混在一起，stats 就分不出「该修代码」和「该改用法」。
        return "bad_input"
    for k in ("saved", "promoted", "updated", "added", "deleted", "exported",
              "backup", "merged", "retired", "imported"):
        if result.get(k) is False:
            return "rejected"
    return "ok"


def _safe(fn):
    """
    把所有异常转成结构化错误，并在每次调用前做一次引擎热加载。

    没有异常包装，MCP 客户端只会收到一句 "Error executing tool"，
    调用方分不清是文件不存在、扩展名不对，还是解析失败——等于没有错误信息。
    """
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        _maybe_reload()
        t0 = time.perf_counter()
        # 初始 None：区分「没走到判定」和「判定结果就是 error」。
        # 原来初始值用 "error"，和 _verdict_of 的返回值撞名，
        # 导致业务拒绝被一律记成 exception。
        status = None
        try:
            result = fn(*args, **kwargs)
            status = _verdict_of(result)
            return result
        except FileNotFoundError as e:
            status = "bad_input"
            return {"error": str(e), "错误类型": "文件不存在",
                    "建议": "确认路径；文件名含空格、全角括号时须原样传入"}
        except ValueError as e:
            msg = str(e)
            # 按错误内容给对应建议，不要所有 ValueError 都套「检查参数」——
            # 文件损坏时提示「文件不存在/扩展名不对」等于没说，还会误导去改路径。
            if "docx" in msg.lower() or "pptx" in msg.lower() or "扩展名" in msg:
                tip = ("只支持 .docx / .pptx。.doc / .wps / .pdf 请先在 Word/WPS 里"
                       "另存为 .docx；确认文件本身不是改了扩展名的旧格式")
            elif "profile_id" in msg:
                tip = "profile_id 只允许字母、数字、下划线、连字符，长度 2–64"
            elif "doc_type" in msg:
                tip = "doc_type 可选：auto / 公文 / 技术文档 / 网页 / 报告"
            else:
                tip = "检查该参数取值；错误信息里已说明具体哪一项不合法"
            status = "bad_input"
            return {"error": msg, "错误类型": "参数不合法", "建议": tip}
        except PermissionError as e:
            status = "bad_input"
            return {"error": str(e), "错误类型": "没有读取权限", "建议": "检查文件权限"}
        except TypeError as e:
            # 区分「调用方传错」和「服务内部写错」：两者的处置完全不同。
            # 内部错误却提示「检查参数形状」，会让人反复改参数——这次的
            # unsupported format string 就是内部 None 格式化，被误报成参数问题。
            internal = e.__traceback__ is not None and e.__traceback__.tb_next is not None
            if internal:
                status = "bug"
                return {"error": str(e), "错误类型": "内部错误（TypeError）",
                        "建议": "这是服务端缺陷，不是你的参数问题。请把 error 与 _审计 里的"
                                "口径版本、规则指纹一并反馈，不要反复重试或改参数"}
            status = "bad_input"
            return {"error": f"参数类型不对：{e}", "错误类型": "参数类型错误",
                    "建议": "检查参数形状：路径是字符串、rules 是对象数组、patch 是对象"}
        except Exception as e:  # noqa: BLE001
            name = type(e).__name__
            if name in ("PackageNotFoundError", "BadZipFile", "KeyError"):
                status = "bad_input"
                return {"error": str(e), "错误类型": name,
                        "建议": "文件打不开：可能不是有效的 Office 文件、已损坏，"
                                "或是改了扩展名的旧格式（.doc 改名成 .docx）。"
                                "请用 Word/WPS 打开确认，另存为 .docx 后重试"}
            status = "bug"
            return {"error": str(e), "错误类型": name,
                    "建议": "先核对参数形状是否符合该工具 schema；"
                            "参数无误仍失败时，把这条错误原样反馈给维护者，不要反复重试"}
        finally:
            # 放在 finally 里：except 分支提前 return 也能记到。
            # 只记对外工具，内部调用不单独计数。
            #
            # 但 stats 本身不记：它是「读指标」的观测动作，记了就等于观测者
            # 污染观测对象——删掉 calls.jsonl 后调 stats，会先写进一条自己的
            # 记录，于是永远看不到「暂无调用记录」这个状态。
            _is_stats = (fn.__name__ == "service_control"
                         and (args[:1] or [kwargs.get("action")])[0] == "stats")
            if fn.__name__ in _TOOL_NAMES and not _is_stats:
                _record_call(fn.__name__, (time.perf_counter() - t0) * 1000,
                             status or "bug")
    return wrapper


_ensure_dir(PROFILE_ROOT)

INSTRUCTIONS = """\
# 角色

你是文档画像分析师。代码负责统计，你负责判断。
你的产出不是「风格描述」，是**可被第三人按规则验证的写作规格**。

# 分工（不可越界）

代码给你的：字号、行距、缩进、编号、层级深度、句长分布、词频、情态分布、列表化程度。
你要做的：归纳套路、判断成因、识别禁区、把数字翻译成可执行规则。

**禁止**：
- 自己统计数字（你数两遍结果不一样，用工具给的）
- 引用未排除复合词的原始词频（「需」可能在「需求」里，看自检告警）
- 把 analysis_brief 的问题原样抄进画像（那是要你回答的，不是内容）
- 在没有原文证据时下结论

# 口径与版本（改东西前先看这里）

「口径」是四层，别混成一件：
1. **统计口径** —— 数字怎么算。口径版本 + 规则指纹标识。改词表/切分/算法才动它，
   动了旧画像数字作废，必须全量重算。
2. **判定口径** —— 按哪套标准判。`doc_type` 决定，默认 auto＝通用/公文标准；
   技术文档不缩进是常态、段落 <150 缺目录不扣分。改它**不改版本**，但同一份文档分数会不同。
3. **准入阈值** —— 多少分、多少样本才配进画像。config.json 里（pass=80 等），可改，不改版本。
4. **呈现口径** —— 返回多少内容。`preset` / `detail`，只影响体积，不影响任何数字。

`service_control(action="status")` 的「口径说明」列出四层当前值、本版变更与是否需重算。
看到规则指纹变了别直接重算：先比输入指纹，再比同一份文档的 score，
分数没变说明只动了展示层。判据永远是变更记录里的「需重算」，不是版本号本身。

# 工作流

0. 动手前先 `check_profiles`。已有画像里如果「口径版本」或「数字一致性」报错，
   说明那份画像的数字是旧口径算的，不能当参考，更不要沿用它的结论
1. 一份文档：`analyze_document(preset="build_profile")`。
   不要用 preset="full"——那是排查用的，14 KB 里有一半（表格明细、样式表）写画像用不上。
   验收已成稿文档用 `preset="audit"`，只要 1–2 KB。
   低于 60 分或「样本不足」的文档不进画像
2. 多份文档：`analyze_batch` 判该建几个画像——差异 ≥3 维拆开；=2 维挂起，问人
   技术文档要传 `doc_type=技术文档`，否则批量质量门按公文标准判、会被整批剔除
3. **先读「先读这个 / 自检告警」和「复合词排除明细」**，再引用任何数字
4. 读原文，逐条回答 `analysis_brief`。每条都要引原文
5. **先把 B1–B6 的答案列出来给人看**，再写 DNA。不要跳过直接写画像
6. `save_profile` → 传 `spec_from_docx`，证据会被校验，不合规直接拒收。
   meta 必须带「文档类型」，并尽量带「适用描述」和 `scope`：
     适用描述  一句话说清适用于什么文档。**这不是可选项**——库里画像多了以后，
              别人靠这句话选画像，没有它只能看 profile_id 猜
     scope     通用 / 项目 / 临时。通用跨项目复用（门槛更高、必须写描述），
              临时不允许转正
7. 改规则用 `edit_rules`，改状态/元信息用 `manage_profile`，别整份重写
8. 发现两个画像其实是同一类（聚类拆错了）→ `merge_profiles`，冲突会留在 conflicts.md
   等你裁决（`manage_profile(action="resolve_conflict")`）
9. 已废弃的画像不能直接 promote；要复活就 `manage_profile(action="clone")` 复制成新 ID

# 工具速查（9 个）

- `analyze_document`  一份文档。**按任务选 preset**：build_profile（建画像，~10 KB）/
                      audit（验收，~1.5 KB）/ full（排查，~14 KB）。
                      按扩展名自动分流 Word/PPT
- `analyze_batch`     一批文档：质量门筛选 → 判该建几个画像 →（with_compare）一致性对比
                      doc_type 默认 auto＝通用/公文口径；技术文档要显式指定，否则按公文标准判成拒收
                      返回里的「质量门口径」写明本次用的是哪套
- `get_profile`       传 profile_id 读一份（含 DNA/规则/硬指标），不传列清单
- `save_profile`      落库（证据会被逐条校验）
- `edit_rules`        action=add / update / delete，改单条规则不用重写 DNA
- `manage_profile`    action=promote / retire / clone / update_meta /
                      resolve_conflict / export
- `merge_profiles`    合并拆错的画像（多来源 + 可带新 DNA）
- `check_profiles`    体检：完整性/证据/数字/口径/源文档/元信息/重复…
- `service_control`   action=status / stats / reload / config / set / backup
                      stats 看运行指标（调用量、成功率、耗时、库分布）

PPT 和 Word 走同一个入口（`analyze_document` 自动分流），只是 PPT 没有文风量化，
写作风格要你自己读文字判断。

`save_profile` 的 meta **必须带「文档类型」**：它是按类型检索和「重复画像」检测的唯一依据，
缺了检测会静默失效，`manage_profile(action="promote")` 也会拒绝。

**覆盖「生效」中的画像会被拒绝**，要显式 `overwrite=True`。遇到这个拒绝先想清楚：
多数情况该用 `edit_rules` 增量改，而不是整份替换。候选/废弃状态可以直接覆盖。

PPT 的流程和 Word 一致：`analyze_document` 会自动识别 .pptx，一次给出质量门和视觉版式。
但 PPT 没有文风量化，写作风格需要你自己读文字判断。

# 输出契约

DNA.md 必须含八节：
元信息 / 结构骨架 / 排版规格 / 写作风格 / 内容套路 / 缺陷诊断 / 冲突与待定 / 变更记录

rules 每条四个字段：
- assertion 一条可执行的句子
- judge 第三人怎么验这条（如「抽查10条需求，统计『需』占比」）
- evidence 只接受 `spec.xxx`（须真实存在）或 `原文：<片段>`
- confidence 高 / 中 / 低

# 规则怎么写（这是最容易做错的地方）

✅ 好规则：「单句不超过 26 字」
   判据：统计句长，超过即扣
   证据：spec.voice.截尾平均句长=18.2

✅ 好规则：「需求条目统一用『需 + 动词』」
   判据：抽查 10 条，统计「需」占比
   证据：原文：数据资源标识码：支持填写多个，逗号隔开

❌ 坏规则：「语言简洁专业」——无法验证，不是规则，是形容词
❌ 坏规则：「整体风格偏商务」——不可执行
❌ 坏规则：「需出现 152 次」——原始词频含「需求」，真实情态只有 12 次

# 证据不足时怎么办

写「无法判定：样本中未出现可支撑该判断的段落」，并记进「冲突与待定」。
**不要编一个看起来合理的。** 编出来的规则比没有规则更糟——它会被当成标准执行。

# 交付前自检（逐条过）

1. 每条规则的 judge，一个没读过这份文档的人能照着验吗？
2. 有没有把推断写成事实？推断必须标「推断」并进「冲突与待定」
3. 缺陷诊断如实写了吗？样本本身可能是错的，别美化
4. 引用的是排除复合词后的数字，还是原始词频？
5. analysis_brief 的每个问题都回答了吗？都引原文了吗？
"""

mcp = MCPServer(name="DSWiki-MCP", instructions=INSTRUCTIONS)


# ---------------------------------------------------------------- 内部工具

def _check_path(path: str, suffix: str) -> Path:
    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(f"文件不存在：{p}")
    if not p.name.lower().endswith(suffix):
        raise ValueError(f"需要 {suffix} 文件，收到：{p.name}")
    return p


def _raw_to_json_doc(raw: dict) -> dict:
    """把 docx_probe.extract_docx 的原始返回，转成与 --json 一致的结构。"""
    styles = {}
    for k, v in raw["样式规格"].items():
        styles[k] = {"count": v["count"], "spec": docx_probe.representative(v)}
    return {
        "file": raw["file"],
        "path": raw["path"],
        "page": raw["页面设置"],
        "paragraphs": raw["有效段落数"],
        "styles": styles,
        "numbering_scheme": {
            str(k): {"pattern": v["pattern"], "count": v["count"], "styles": v["styles"]}
            for k, v in docx_probe.heading_numbering_scheme(raw["章节树"]).items()
        },
        "outline": raw["章节树"],
        "tables": raw["表格"],
    }


def _voice_metrics(path: Path) -> dict:
    doc = Document(str(path))
    headings = []
    for p in doc.paragraphs:
        sn = (p.style.name or "").lower()
        if (sn.startswith("heading") or sn.startswith("标题")) and p.text.strip():
            headings.append((sn, p.text.strip()))
    # 正文与表格必须分开传：表格单元格不是段落，混进去会把段长/句长算错
    body, cells = style_probe.load_parts(str(path))
    return style_probe.analyze(body, headings, table_text="\n".join(cells))


def _pct(v, digits: int = 1) -> str:
    """把比率格式化成百分比。None 表示「没测到」，既不能当 0% 也不能崩。

    这里踩过一次坑：技术词占比在「词表双零」时返回 None（表示未覆盖该领域用词），
    而 brief 里还写着 :.1%，于是纯叙述类文档直接把整个文风阶段打崩。
    凡是可能为 None 的比率，一律走这个函数。
    """
    if v is None:
        return "未测到（词表未覆盖该领域用词）"
    try:
        return f"{float(v) * 100:.{digits}f}%"
    except (TypeError, ValueError):
        return "未测到"


def _voice_brief(m: dict, heads: list | None = None) -> list:
    """
    把「代码算不出来、必须 LLM 判断」的任务列清楚。
    每条都带：已知数据（让它有据可依）、具体动作、输出形态、不合格示例。
    不给输出形态和反例，模型会写成「语言简洁专业」这种没法验证的形容词。
    """
    modals = m.get("情态词", {})
    hard = modals.get("硬性要求（应/须/需）", {})
    allow = modals.get("可选（可/支持）", {})
    excl = m.get("复合词排除明细", {})
    top_tech = list(m.get("技术词Top", {}).items())[:8]
    heads = heads or []

    return [
        {
            "id": "B1",
            "任务": "判断这批文档是「提要求」还是「做描述 / 转述标准」",
            "已知": {
                "硬性要求占比": _pct(m.get("硬性要求占比")),
                "硬性情态词分布": hard,
                "可选情态词分布": allow,
                "复合词排除（仅供理解差异，禁止引用原始数）": {
                    k: f"原始{v['原始出现']}→真情态{v['真情态']}"
                    for k, v in excl.items()},
            },
            "要你做的事": "从原文找 2–3 处「真情态」的实际用法（不是『需求』『可以』这类复合词里的字），"
                       "看这个情态词修饰的是「自己的要求」还是「引用别人的规定」。",
            "输出形态": "一句话结论（提要求 / 做描述 / 转述标准），后跟 2–3 条原文引用，每条不超过 40 字。",
            "不合格示例": "「用语规范严谨」——没回答是提要求还是做描述",
        },
        {
            "id": "B2",
            "任务": "归纳「描述一件事」的固定套路（要素顺序）",
            "已知": {
                "平均段长": f"{m.get('平均段长', 0)} 字",
                "列表化段落占比": _pct(m.get("列表化段落占比")),
                "句长中位数": f"{m.get('句长中位数', 0)} 字",
            },
            "要你做的事": "用 analyze_document(preset='build_profile') 返回的 outline 定位 2–3 个同级章节，逐段读，"
                       "看它们是不是同一套写法：先说什么、后说什么、有没有固定的引导句式。",
            "输出形态": "一个编号的要素顺序（如 ①标题 ②规则条目 ③例外说明），"
                     "每个要素后附 1 段原文摘录（≤40 字）。",
            "不合格示例": "「层次分明、条理清晰」——没有给出要素顺序",
        },
        {
            "id": "B3",
            "任务": "找出词表漏掉的领域专有名词",
            "已知": {
                "词表已命中 Top": dict(top_tech),
                "技术词密度": f"{m.get('技术词密度(每千字)', 0)}/千字",
                "英文混入": {
                    "英文字母占比": _pct((m.get("英文混入") or {}).get("英文字母占比"), 0),
                    "整段英文段落数": (m.get("英文混入") or {}).get("整段英文段落数"),
                    "词表未覆盖的技术标识": (m.get("英文混入") or {}).get("词表未覆盖的技术标识"),
                },
                "中文段落句长中位数": m.get("中文段落句长中位数"),
                "技术词占比": _pct(m.get("技术词占比")),
                "说明": "中文按词表匹配、英文按词表+形态识别；仍是穷举法，"
                        "新框架与自造名词会漏计，实际暴露度只会更高。"
                        "具体漏了什么见「词表未覆盖的技术标识」",
            },
            "要你做的事": "在原文里找**不在上面 Top 里、但明显属于本领域**的名词，"
                       "逐个判断：面向业务人员写文档时，这个词是必须用、还是应该翻译成业务动作。",
            "输出形态": "表格：词 / 量级（少<10、中10–50、多>50）/ 是否列入禁区 / 理由。"
                     "量级你目测即可，不要精确统计。",
            "不合格示例": "「技术术语较多」——没有具体词",
        },
        {
            "id": "B4",
            "任务": "判断标题句式是否统一，给统一方案",
            "已知": {
                "动宾式数量": m.get("标题动宾式", 0),
                "名词短语式数量": m.get("标题名词短语式", 0),
                "标题样例": heads[:10],
            },
            "要你做的事": "看上面的标题样例，判断是不是清一色一种句式。"
                       "若统一，指出有没有个别别扭的；若不统一，给出该定哪一种。",
            "输出形态": "结论（统一 / 不统一）+ 3 个具体标题作为例证。"
                     "**如果没有反例，直接写「无反例」**，不要硬凑。",
            "不合格示例": "硬凑三个反例；或只写「标题简洁」",
        },
        {
            "id": "B5",
            "任务": "识别禁区——样本里从来不出现、但同类文档常写的写法",
            "已知": {
                "人称分布": m.get("人称使用", {}),
                "列表化占比": _pct(m.get("列表化段落占比")),
                "情态词总数": m.get("情态词总数", 0),
            },
            "要你做的事": "通读时留意「本来该有但没有」的东西：该有的章节有没有、"
                       "常见的表述方式有没有出现、有没有量化指标、有没有行动建议。",
            "输出形态": "每条格式：「禁止 X ｜ 依据：全样本 0 次 / 某章节缺失」。至少 3 条。",
            "不合格示例": "「避免口语化」——没有依据",
        },
        {
            "id": "B6",
            "任务": "如实记录样本自身的缺陷（不是评价好坏，是记录事实）",
            "已知": {
                "自检告警": m.get("自检告警", []),
                "指标可信度": m.get("指标可信度", "高"),
            },
            "要你做的事": "把排版和结构上的不一致、缺失、可疑之处逐条列出，"
                       "注明依据（哪个样式、哪个章节、哪个数字）。",
            "输出形态": "表格：# / 现象 / 依据 / 影响。**不要写改进建议**，是否改由人决定。",
            "不合格示例": "把缺陷写成「可进一步优化」这种模糊表述",
        },
    ]


def _profile_dir(profile_id: str) -> Path:
    if not re.fullmatch(r"[a-zA-Z0-9_\-]{2,64}", profile_id):
        raise ValueError("profile_id 只能含字母、数字、下划线、连字符，长度 2–64")
    # 下划线/点开头是内部保留前缀：list_profiles、_library_stats、backup_profiles
    # 都会跳过这类目录（用于放临时与元数据）。以前这里放行、那边跳过，
    # 结果是「存得上、读得到、能转正，但清单里 0 份、备份里 0 份」——
    # 数据份数承诺静默失守，比直接报错危险得多。
    if profile_id[0] in ("_", "."):
        raise ValueError(
            f"profile_id 不能以「{profile_id[0]}」开头：下划线与点是内部保留前缀，"
            "这类画像不会出现在清单、统计与整库备份里。请改用字母或数字开头")
    return PROFILE_ROOT / profile_id


# ---------------------------------------------------------------- 工具

@_safe
def extract_format(docx_path: str) -> dict:
    """提取 Word 排版规格：页面设置、样式表（中文字体/字号/加粗/对齐/行距/缩进/段前后）、
    标题编号体系、章节树、表格规格。这是纯代码提取，结果可复现。

    返回 spec 供后续 checklist 与 save_profile 的 evidence 引用（用 spec.xxx 路径）。
    """
    p = _check_path(docx_path, ".docx")
    raw = docx_probe.extract_docx(str(p))
    doc = _raw_to_json_doc(raw)
    notes = []
    for k, v in doc["styles"].items():
        sp = v["spec"]
        if sp.get("size_pt") and round(sp["size_pt"] * 2) / 2 not in docx_probe.PT_TO_CN:
            notes.append(f"样式「{k}」字号 {sp['size_pt']:g}pt 不是中文标准字号"
                         f"（常用：小四 12 / 五号 10.5 / 四号 14）。"
                         "这是样式定义值，属于设计选择；只有段落里再单独覆写字号才算手改")
        if (sp.get("indent") or "").startswith("无") and k in ("Normal", "正文"):
            # 中性措辞：技术/网页风格文档本来就不用首行缩进，按公文标准说它"不规范"是错的
            notes.append("正文无首行缩进（技术文档与网页风格常见；中文公文一般要求缩进 2 字符，"
                         "视文档类型而定）")
    if not doc["outline"]:
        notes.append("未识别到任何标题样式，章节结构对机器不可读——该文档不适合建画像")
    return {"spec": doc, "提示": notes}


@_safe
def extract_voice(docx_path: str) -> dict:
    """提取 Word 文风硬指标：句长分布（含截尾均值与可信度）、人称、情态硬度、
    术语暴露度、列表化程度、标题句式。

    同时在 analysis_brief 里返回「需要你判断」的开放问题——这些是代码算不出来的，
    必须由你读原文回答并引原文。不要把 analysis_brief 原样抄进画像。
    """
    p = _check_path(docx_path, ".docx")
    m = _voice_metrics(p)
    doc = Document(str(p))
    heads = [x.text.strip() for x in doc.paragraphs
             if (x.style.name or "").lower().startswith(("heading", "标题")) and x.text.strip()]
    warnings = list(m.get("自检告警", []))
    if m["指标可信度"] != "高":
        warnings.append({"级别": "高", "指标": "句长",
                         "说明": f"{m['指标可信度']}。请用截尾均值 {m['截尾平均句长']} 字"
                                 f"和中位数 {m['句长中位数']} 字，不要用平均值。"})
    return {
        "file": p.name,
        "指标": m,
        "自检告警": warnings,
        "标题样例": heads[:10],
        "analysis_brief": _voice_brief(m, heads),
        "使用提示": "先看「自检告警」，再看 analysis_brief。每条 brief 都给了已知数据、"
                "要做的动作、输出形态和反例，照着做即可，不要自由发挥。",
    }


@_safe
def check_quality(docx_path: str, doc_type: str = "auto") -> dict:
    """Word 准入评分（满分 100）：标题规范性 25 / 手刷格式 20 / 层级连贯 15 /
    编号一致性 15 / 内容完整 15 / 结构完整 10，另设样本充分度前置门槛
    （段落 <20 或字数 <500 判「样本不足」）。

    verdict 为「拒收」或「样本不足」的文档不要用来建画像。

    「交付检查」单独列出图注表题、页码页眉、修订记录、格式残留，**不计分**。
    分数只能纵向比（同一份文档改前改后），不能横向比——扣分权重是经验值，不是标准。
    """
    p = _check_path(docx_path, ".docx")
    r = quality_gate.evaluate(str(p), doc_type=doc_type)
    if r is None:
        # 不是「文件不存在」，而是有文件但没有可提取的正文段落（例如整页只有表格/图片）。
        # 保留「不可用」三个字：cluster 的排除条件依赖它。
        return {"file": p.name, "verdict": "无正文段落（可能只有表格/图片，不可用）",
                "score": None, "不足原因": ["没有任何正文段落"]}
    return r


@_safe
def extract_pptx_style(pptx_path: str) -> dict:
    """提取 PPT 视觉与版式：页面尺寸、主题配色与字体方案、母版字号阶梯、
    版式分布、页面类型识别、图形元素来源（原生形状 / SmartArt / 图片）。

    图形来源决定「架构图/流程图能不能还原」：
    原生形状可完整还原画法，图片型只能记录位置尺寸，SmartArt 视觉不可精确还原。
    """
    p = _check_path(pptx_path, ".pptx")
    return pptx_probe.extract_pptx(str(p))


@_safe
def check_quality_pptx(pptx_path: str) -> dict:
    """PPT 准入评分（满分 100）：原生可编辑率 40 / 文本量 25 / 页数 15 /
    版式多样性 10 / 主题完整性 10。页数 <5 判「样本不足」。

    PPT 最常见的废样本是「每页贴一整张图」——视觉没毛病，但形状树里只有一个
    picture，画法一个字都提取不到。看 native_edit_ratio 即可识别。
    """
    p = _check_path(pptx_path, ".pptx")
    import pptx_quality_gate
    return pptx_quality_gate.evaluate(str(p))


@_safe
def cluster_documents(docx_paths: list[str], exclude_rejected: bool = True,
                      doc_type: str = "auto") -> dict:
    """多份文档聚类分流：用十维特征（排版5 + 结构3 + 文风2）判断它们到底是几类，
    决定应该拆成几个画像。

    判定：差异 ≤1 维自动合并 / =2 维挂起交人工裁决 / ≥3 维判为另一套模板。
    差异 2 维不自动合并，因为错合不可逆（会把两套模板揉成一个画像），
    错分只是多一次人工确认。

    exclude_rejected=True 时先跑质量门，拒收与样本不足的文档不参与聚类。
    doc_type 透传给质量门：以前批量链路固定按 auto（通用/公文）判，
    技术文档会被判成拒收、整批不进画像，而单文档工具指定类型却能准入——
    同一份文档两个工具结论不同，剩下的样本还会偏向公文风格。
    """
    docs, voices, skipped = [], {}, []
    for path in docx_paths:
        # 逐份隔离：一份损坏/路径失效不该让整批失败。
        # 批量导入的真实场景就是「一堆杂乱文档」，其中必然有几份打不开。
        try:
            p = _check_path(path, ".docx")
            if exclude_rejected:
                q = quality_gate.evaluate(str(p), doc_type=doc_type)
                # 空文档 evaluate 返回 None，不能只判 verdict 字符串，
                # 否则空文档会漏进聚类——它所有维度都"缺失"，会误并到任意一组
                bad = (q is None) or any(k in q.get("verdict", "")
                                         for k in ("拒收", "不可用", "空"))
                if bad:
                    skipped.append({
                        "file": p.name,
                        "score": q["score"] if q else None,
                        "verdict": q["verdict"] if q else "无正文段落（不可用）",
                    })
                    continue
            doc = _raw_to_json_doc(docx_probe.extract_docx(str(p)))
            voice = _voice_metrics(p)
        except Exception as e:                      # noqa: BLE001 - 逐份隔离，不能中断整批
            skipped.append({
                "file": Path(path).name,
                "score": None,
                "verdict": f"无法解析（{type(e).__name__}）",
                "原因": str(e)[:100],
            })
            continue
        docs.append(doc)
        voices[p.name] = voice

    if len(docs) < 2:
        return {"error": "可用文档少于 2 份，无法聚类", "已跳过": skipped}

    feats = [cluster_gate.extract_features(d, voices.get(d["file"])) for d in docs]
    _cc = service_config.load()["cluster"]
    groups, pairs, merges, pending = cluster_gate.cluster(
        feats, strong=_cc["strong"], weak=_cc["weak"])

    out_groups = []
    for i, g in enumerate(sorted(groups, key=lambda x: -len(x["docs"])), 1):
        out_groups.append({
            "组": i,
            "文档": g["docs"],
            "数量": len(g["docs"]),
            "组内最大差异": g["max_diff"],
            "状态": g["status"],
            "差异明细": g["max_detail"],
        })
    return {
        "已跳过": skipped,
        "组数": len(out_groups),
        "分组": out_groups,
        "存疑对（差异2维，需人工裁决）": pending,
        "解释": "组数即该建几个画像。组内只有 1 份说明样本不足，画像会过拟合。",
    }


@_safe
def compare_documents(docx_paths: list[str], cluster_groups=None) -> dict:
    """多文档一致性对比：按组逐项比对哪些规格一致、哪些不一致。

    传了 cluster_groups 就按**聚类组**分（analyze_batch 走这条），与「判该建几个
    画像」用同一套规则；没传才退回按文件名推断（独立调用时保持兼容）。

    以前无论谁调都按文件名猜类型，而聚类是按十维特征分组——两套规则互不相干：
    a/b 聚类判同型，对比却因文件名不同各成一「类型」，每组只有 1 份，
    「不一致项」恒为空。看着像比过，其实没比。
    """
    results = []
    voices = {}
    for path in docx_paths:
        p = _check_path(path, ".docx")
        results.append(_raw_to_json_doc(docx_probe.extract_docx(str(p))))
        voices[p.name] = _voice_metrics(p)

    if cluster_groups:
        name_to_group = {}
        for g in cluster_groups:
            for n in g.get("文档", []):
                name_to_group[n] = g.get("组")
        groups = {}
        for d in results:
            key = (f"聚类组{name_to_group[d['file']]}"
                   if d["file"] in name_to_group else "未分组")
            groups.setdefault(key, []).append(d)
        basis = "十维特征聚类（与本次 analyze_batch 的分组一致）"
    else:
        groups = {}
        for d in results:
            groups.setdefault(batch_compare.doc_type_of(d["file"]), []).append(d)
        basis = "文件名推断（独立调用，未关联聚类结果）"

    out = []
    for gname, gdocs in groups.items():
        item = {"类型": gname, "文档": [d["file"] for d in gdocs], "一致项": [],
                "不一致项": [], "分组依据": basis}
        if len(gdocs) < 2:
            # 组内只有 1 份时，「一致项」是自己跟自己比出来的，
            # 不能让人读成「已完成跨文档比对且没有差异」。
            item["说明"] = ("组内只有 1 份，无从比较；「一致项」是这份文档自身的"
                            "规格，不是跨文档比对结论")
        papers = {d["page"].get("纸张") for d in gdocs}
        if len(papers) == 1:
            item["一致项"].append(f"纸张 {list(papers)[0]}")
        else:
            item["不一致项"].append("纸张：" + " ／ ".join(map(str, papers)))
        for style in ("Normal", "Heading 1", "Heading 2", "Heading 3"):
            vals = []
            for d in gdocs:
                s = d["styles"].get(style)
                vals.append((s["spec"]["font_eastasia"], s["spec"]["size_cn"],
                             s["spec"]["line_spacing"], s["spec"]["indent"]) if s else None)
            present = [v for v in vals if v]
            if not present:
                continue
            uniq = set(present)
            if len(uniq) == 1:
                f, sz, ln, ind = present[0]
                item["一致项"].append(f"{style}：{f} {sz} / {ln} / {ind}")
            else:
                item["不一致项"].append(
                    f"{style}：" + " ／ ".join(f"{f} {sz}/{ln}/{ind}" for f, sz, ln, ind in uniq))
        out.append(item)
    ret = {"对比": out}
    if not cluster_groups:
        ret["提示"] = ("本次按文件名推断类型分组，与聚类分组可能不一致。"
                      "要跟聚类结论对齐，请用 analyze_batch(with_compare=True)，"
                      "它会把聚类组传进来")
    return ret


def _resolve_strict(spec, path: str):
    """返回 (是否完整解析, 段数, 解析到的值)。支持中文键与含空格的键（Heading 1）。"""
    if not path.startswith("spec."):
        return False, 0, None
    rest, cur, n = path[5:], spec, 0
    while isinstance(cur, dict) and rest:
        keys = sorted(cur.keys(), key=len, reverse=True)
        hit = next((k for k in keys if rest == k or rest.startswith(k + ".")), None)
        if hit is None:
            return False, n, None
        cur = cur[hit]
        n += 1
        rest = rest[len(hit):].lstrip(".")
    return rest == "" and n >= 1, n, cur


def _spec_path_ok(spec: dict, evidence: str) -> tuple:
    """
    校验 evidence 里的 spec.xxx 是否真的能在 spec.json 里解析出来。

    不能用正则按点号/空格截断：spec 的键大量是中文（列表化段落占比），
    还有含空格的（Heading 1）。所以从长到短试前缀。

    但光试前缀会误放行——「spec.voice.不存在」会一路退到合法短前缀「spec.voice」。
    所以加一条：解析到的路径必须至少 2 段，或者整串本身就是完整路径。
    """
    if spec is None:
        return False, "未提供 spec，无法校验 spec. 路径是否真实存在"
    if not evidence.startswith("spec."):
        return False, "不是 spec. 路径"
    bare = evidence.strip()
    limit = min(len(evidence), 120)
    for end in range(limit, 4, -1):
        cand = evidence[:end].rstrip("：:，。、,（）() \t")
        ok, seg, val = _resolve_strict(spec, cand)
        if not ok:
            continue
        # 至少 2 段；或单段但指向具体值（列表/标量）。单段指向整个字典没有意义。
        if seg >= 2 or not isinstance(val, dict):
            return True, "ok"
    return False, f"路径无法解析或层级不足（试到「{evidence[:40]}」）"


@mcp.tool()
@_safe
def save_profile(profile_id: str, dna_markdown: str, rules: list[dict],
                 spec: dict | None = None, spec_from_docx: str | None = None,
                 meta: dict | None = None, overwrite: bool = False) -> dict:
    """画像落库。会做四件事：校验规则证据、写文件、记版本。

    rules 是对象数组，每条含四个字段：
      assertion  规则正文，必须是可执行的句子（不要写「语言专业」这种形容词）
      judge      怎么验这条规则（如「抽查10条需求，统计『需』占比」）
      evidence   证据。只接受两种形式：
                   spec.voice.平均句长  —— 指向 spec.json 的字段路径（会校验是否真实存在）
                   原文：<原文片段>      —— 原文引用
      confidence 高 / 中 / 低；缺省按「中」处理，但会在返回里提示

    强烈建议传 spec_from_docx（源 docx 路径）或 spec，否则 spec. 形式的证据无法校验
    是否指向真实存在的字段，模型可以引用一个不存在的字段而蒙混过关。

    证据为空、格式不合规、或指向不存在的字段，整个保存操作都会失败。
    这条约束是硬的：没有证据的规则等于模型编的，不能进画像库污染后续所有输出。

    meta 里建议带三个字段：
      文档类型   按类型检索、重复画像检测的唯一依据。缺了后续 promote 会被拒。
      适用描述   一句话说明这份画像适用于什么文档，例如
                 「适用于含接口清单与字段表的中文需求规格书」。
                 库里画像多了以后，别人靠这句话选画像；没有它只能看 profile_id 猜。
      scope      通用 / 项目 / 临时（默认「项目」）。通用画像跨项目复用，
                 门槛更高；临时画像不允许转正。

    **覆盖一份「生效」中的画像会被拒绝**，要显式传 overwrite=True 才写。
    生效画像是别人正在用的资产，不能静默替换。候选/废弃状态可以直接覆盖
    （旧 DNA 会归档到 history/，可回滚）。
    """
    if not isinstance(rules, list) or any(not isinstance(r, dict) for r in rules):
        return {"saved": False,
                "拒绝原因": "rules 必须是对象数组，每项含 assertion / judge / evidence",
                "示例": [{"assertion": "单句不超过 26 字", "judge": "统计句长",
                          "evidence": "spec.voice.截尾平均句长=18.2", "confidence": "高"}]}

    missing_conf = [r.get("id") or f"第{i}条" for i, r in enumerate(rules, 1)
                    if not (r.get("confidence") or "").strip()]

    problems = []
    if not dna_markdown or len(dna_markdown.strip()) < 200:
        problems.append("DNA 内容过短（少于 200 字），不像是完整画像")
    if not rules:
        problems.append("rules 为空，画像至少要有一条可执行规则")
    for r in rules:
        r.setdefault("confidence", "中")

    if spec is None and spec_from_docx:
        sp = _check_path(spec_from_docx, ".docx")
        raw = docx_probe.extract_docx(str(sp))
        fmt = _raw_to_json_doc(raw)
        st = sp.stat()
        spec = {"source": sp.name,
                "source_path": str(sp),
                "source_mtime": st.st_mtime,
                "source_size": st.st_size,
                "engine_version": engine_version.ENGINE_VERSION,
                "styles": fmt["styles"], "outline": fmt["outline"],
                "page": fmt["page"], "numbering_scheme": fmt["numbering_scheme"],
                "tables": fmt["tables"], "voice": _voice_metrics(sp)}
    elif spec is not None:
        # 外部传入的 spec 也要打上口径版本，否则事后无法判断是否过期
        spec = dict(spec)
        spec.setdefault("engine_version", engine_version.ENGINE_VERSION)

    bad = []
    for i, r in enumerate(rules or [], 1):
        rid = r.get("id") or f"R{i}"
        miss = [k for k in ("assertion", "judge", "evidence") if not (r.get(k) or "").strip()]
        if miss:
            bad.append(f"{rid} 缺字段：{'、'.join(miss)}")
            continue
        ev = r["evidence"].strip()
        if ev.startswith("spec."):
            ok, why = _spec_path_ok(spec, ev)
            if not ok:
                bad.append(f"{rid} 的证据路径无法解析：{why}")
        elif ev.startswith(("原文：", "原文:")):
            continue
        else:
            bad.append(f"{rid} 的 evidence 格式不合规（须以 spec. 或「原文：」开头）：{ev[:40]}")
    if bad:
        problems.append("以下规则证据不合规 → " + "；".join(bad))

    if problems:
        return {"saved": False, "拒绝原因": problems,
                "提示": "修正后重试。evidence 要么指向 spec 里真实存在的字段，要么是原文引用。"}

    # scope 校验：取值错了会一路传到体检/promote 的判据里，必须在入口拦住
    scope = ((meta or {}).get("scope") or "项目").strip()
    if scope not in SCOPES:
        return {"saved": False,
                "拒绝原因": f"scope 只支持 {'/'.join(SCOPES)}，收到「{scope}」",
                "说明": "通用=跨项目复用（门槛更高）；项目=只对某个项目有效；"
                        "临时=试验用，不允许转正"}
    if meta is not None:
        meta = {**meta, "scope": scope}

    d = _profile_dir(profile_id)
    existed = (d / "DNA.md").exists()
    prev_status, prev_scope = "", ""
    if existed:
        try:
            prev = json.loads((d / "meta.json").read_text(encoding="utf-8"))
            prev_status = prev.get("status", "")
            prev_scope = prev.get("scope", "项目")
        except (OSError, json.JSONDecodeError):
            prev = {}
        # 生效中的画像是别人正在用的资产，不能静默替换。
        # 候选/废弃可以直接覆盖——还没人依赖它，旧 DNA 也会归档。
        if "生效" in str(prev_status) and not overwrite:
            return {"saved": False,
                    "拒绝原因": f"画像 {profile_id} 当前是「{prev_status}」状态，"
                               "覆盖会替换掉正在被使用的版本",
                    "现状": {"规则数": prev.get("rule_count"),
                             "文档类型": prev.get("文档类型"),
                             "适用描述": prev.get("适用描述")},
                    "可选做法": [
                        "改用 edit_rules 增量修改规则（推荐，保留历史与状态）",
                        "确认要整份替换：带 overwrite=True 重试（旧版会归档到 history/）",
                        "另起一个 profile_id，用 merge_profiles 合并"]}

    _ensure_dir(d)
    _ensure_dir(d / "history")

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    if existed:
        old = (d / "DNA.md").read_text(encoding="utf-8")
        (d / "history" / f"DNA-{ts}.md").write_text(old, encoding="utf-8")

    (d / "DNA.md").write_text(dna_markdown, encoding="utf-8")

    # 落库即编号。以前这里直接写 rules.json，id 全是 null——
    # 结果是「按 R1 改」报「找不到 R1」，而「按编号删」因为另一条路径
    # （_save_rules 会重编号）反而能成功，两条路行为不一致。
    # 规则没有稳定编号，DNA.md 里写的「见 R3」也就无从指认。
    for i, r in enumerate(rules, 1):
        if r.get("id") not in (None, "", f"R{i}"):
            r.setdefault("原id", r["id"])
        r["id"] = f"R{i}"
    (d / "rules.json").write_text(
        json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
    if spec is not None:
        (d / "spec.json").write_text(
            json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
    meta_full = {"profile_id": profile_id, "saved_at": ts,
                 "rule_count": len(rules), "status": "候选",
                 **(meta or {})}
    # 文档类型统一到「文档类型」这一个键上：它是按类型检索与「重复画像」检测的唯一依据。
    # 缺了它，profile_doctor 只能拿 profile_id 兜底，同类型两个画像永远查不出重复——静默失效。
    doc_type = (meta_full.get("文档类型") or meta_full.get("doc_type") or "").strip()
    if doc_type:
        meta_full["文档类型"] = doc_type
    (d / "meta.json").write_text(
        json.dumps(meta_full, ensure_ascii=False, indent=2), encoding="utf-8")

    ret = {"saved": True, "path": str(d), "规则数": len(rules), "状态": "候选",
           "规则编号": [f"R{i}" for i in range(1, len(rules) + 1)],
           "scope": scope,
           "下一步": "画像为「候选」状态，需人工确认后才能被生成侧使用"}
    if existed:
        ret["已覆盖"] = (f"该画像原为「{prev_status or '未知'}」状态，旧版已归档到 history/"
                        f"DNA-{ts}.md，可回滚")
    if not doc_type:
        ret["警告"] = ("未提供「文档类型」，重复画像检测与按类型检索会失效，"
                       'manage_profile(action="promote") 会因此拒绝。请在 meta 里补 文档类型。')
    desc = (meta_full.get("适用描述") or "").strip()
    if not desc:
        ret["建议"] = ("meta 里补一句「适用描述」，比如「适用于含接口清单与字段表的中文"
                       "需求规格书」。画像多了以后，别人靠这句话选画像，"
                       "没有它只能看 profile_id 猜")
    if scope == "通用" and not desc:
        ret["警告"] = ((ret.get("警告") + "；" if ret.get("警告") else "")
                       + "scope=通用 的画像跨项目复用，必须写「适用描述」，否则 promote 会被拒")
    if missing_conf:
        ret["提示"] = (f"这些规则没写 confidence，已按「中」处理："
                       f"{'、'.join(missing_conf)}。置信度影响规则可信度判断，建议补上。")
    return ret


def _touch_usage(profile_id: str, purpose: str | None) -> dict:
    """记一次使用：use_count +1、更新 last_used_at。用于识别「建了没人用」的死画像。"""
    mf = _profile_dir(profile_id) / "meta.json"
    if not mf.exists():
        return {}
    try:
        meta = json.loads(mf.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    meta["use_count"] = int(meta.get("use_count", 0)) + 1
    meta["last_used_at"] = datetime.now().strftime("%Y%m%d-%H%M%S")
    if purpose:
        meta["last_used_for"] = purpose[:80]
    mf.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    return meta


@_safe
def load_profile(profile_id: str, purpose: str | None = None, track: bool = True) -> dict:
    """读回画像：DNA.md 正文、规则清单、元信息、硬指标。

    每次调用记一次使用（use_count +1、更新 last_used_at），用于事后判断哪些画像
    建了却从来没人用。track=False 可关闭打点（仅查看用）。
    """
    d = _profile_dir(profile_id)
    if not (d / "DNA.md").exists():
        return {"error": f"画像 {profile_id} 不存在", "可用": list_profiles()["画像"]}
    meta = _touch_usage(profile_id, purpose) if track else {}
    out = {"profile_id": profile_id, "DNA": (d / "DNA.md").read_text(encoding="utf-8")}
    for name in ("rules.json", "spec.json", "meta.json"):
        f = d / name
        if f.exists():
            out[name.replace(".json", "")] = json.loads(f.read_text(encoding="utf-8"))
    if meta:
        out["本次已记使用"] = {"use_count": meta.get("use_count"),
                            "last_used_at": meta.get("last_used_at")}
    return out


@_safe
def list_profiles() -> dict:
    """列出画像库：适用描述、状态、scope、文档类型、规则数、使用次数。

    选画像先看「适用描述」——它就是给这一步用的。没有描述只能看 profile_id 猜，
    库里超过三五份就选不准了。
    use_count 长期为 0 的画像应考虑废弃，别让它占着名字。
    """
    items = []
    skipped = []
    for sub in sorted(PROFILE_ROOT.iterdir()) if PROFILE_ROOT.exists() else []:
        if not sub.is_dir():
            continue
        if sub.name.startswith(("_", ".")):
            # 列不出来也要说一声，否则「库里 3 份、清单 0 份」会让人以为数据丢了
            skipped.append(sub.name)
            continue
        meta = {}
        mf = sub / "meta.json"
        if mf.exists():
            try:
                meta = json.loads(mf.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                meta = {}
        items.append({
            "profile_id": sub.name,
            # 适用描述放在前面：选画像时先看「适用什么文档」，再看 id
            "适用描述": meta.get("适用描述") or "（未填，只能靠 profile_id 判断）",
            "状态": meta.get("status", "未知"),
            "scope": meta.get("scope") or "项目",
            "文档类型": meta.get("文档类型") or meta.get("doc_type"),
            "保存时间": meta.get("saved_at"),
            "规则数": meta.get("rule_count"),
            "使用次数": meta.get("use_count", 0),
            "最后使用": meta.get("last_used_at"),
            "已并入": meta.get("merged_into"),
        })
    ret = {"画像": items, "根目录": str(PROFILE_ROOT),
           "口径版本": engine_version.ENGINE_VERSION}
    if skipped:
        ret["已跳过"] = skipped
        ret["提示"] = (f"{len(skipped)} 个目录以下划线或点开头，是内部保留前缀，"
                      "不计入画像库。正常情况下不会有这类目录；"
                      "若见到，说明存放方式不对，其内容不会被备份")
    return ret


@mcp.tool()
@_safe
def backup_profiles(out_path: str) -> dict:
    """把整个画像库打成单个 JSON。用于换机器、换工作区、清理前留底。"""
    if not PROFILE_ROOT.exists():
        return {"backup": False, "拒绝原因": "画像库不存在"}
    items = []
    skipped = []
    for sub in sorted(PROFILE_ROOT.iterdir()):
        if not sub.is_dir():
            continue
        # 下划线/点开头与缺 DNA.md 的目录都会跳过，但**必须记下来**：
        # 备份少了几份却不吭声，等于数据份数承诺失守（曾经踩过：
        # _test_xxx 存得上、能转正，但备份里一份没有）。
        if sub.name.startswith(("_", ".")):
            skipped.append({"目录": sub.name, "原因": "内部保留前缀，不纳入画像库"})
            continue
        if not (sub / "DNA.md").exists():
            skipped.append({"目录": sub.name, "原因": "缺 DNA.md，不是完整画像"})
            continue
        items.append({
            "profile_id": sub.name,
            "DNA": (sub / "DNA.md").read_text(encoding="utf-8"),
            "rules": _load_json(sub / "rules.json", []),
            "spec": _load_json(sub / "spec.json"),
            "meta": _load_json(sub / "meta.json", {}),
        })
    out = Path(out_path).expanduser()
    try:
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps({
            "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "engine_version": engine_version.ENGINE_VERSION,
            "画像库": str(PROFILE_ROOT),
            "画像数": len(items),
            "profiles": items,
        }, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError as e:
        return {"backup": False, "拒绝原因": f"无法写入 {out}：{e.strerror or e}"}
    _audit("backup_profiles", "-", {"out": str(out), "count": len(items),
                                   "skipped": len(skipped)})
    ret = {"backup": True, "path": str(out), "画像数": len(items),
           "大小": f"{out.stat().st_size} 字节"}
    if skipped:
        ret["已跳过"] = skipped
        ret["警告"] = (f"库里有 {len(skipped)} 个目录未被备份。"
                      "如果其中有你以为已保存的画像，说明它放在了保留前缀下，"
                      "请改用字母/数字开头的 profile_id 重新落库")
    return ret


@_safe
def check_profiles(profile_id: str | None = None) -> dict:
    """画像库体检，九项：完整性 / 状态匹配 / 证据可解析 / 数字一致性 /
    口径版本 / 源文档 / 规则厚度 / 元信息 / 重复画像。

    最该关注两项互补检查：
      数字一致性 —— DNA.md 写的数与 spec.json 实测值是否自洽
      口径版本   —— spec 记录的口径与当前引擎是否同步
    内部自洽但口径过期是最危险的状态：文件完整、引用通畅，看起来毫无问题，
    但数字全部来自旧口径。

    不传 profile_id 则体检全部，并额外检测「同类型多个可用画像」（可能拆错了）。
    """
    pids = [p.name for p in sorted(PROFILE_ROOT.iterdir())
            if p.is_dir() and not p.name.startswith("_")] if PROFILE_ROOT.exists() else []
    if profile_id:
        if profile_id not in pids:
            return {"error": f"画像 {profile_id} 不存在", "可用": pids}
        pids = [profile_id]
    if not pids:
        return {"画像": [], "说明": "画像库为空"}

    results = [profile_doctor.inspect(p) for p in pids]

    dup = {}
    untagged = []
    if not profile_id:
        by_type = {}
        for r in results:
            if "过期" in r["status"] or "废弃" in r["status"]:
                continue
            if not r.get("doc_type"):
                untagged.append(r["profile_id"])
                continue
            by_type.setdefault(r["doc_type"], []).append(r["profile_id"])
        dup = {k: v for k, v in by_type.items() if len(v) > 1}

    need = [r["profile_id"] for r in results if r["rebuild"]]
    todo = [{"profile_id": r["profile_id"], "动作": a}
            for r in results for a in r["actions"]]
    return {
        "口径版本": engine_version.ENGINE_VERSION,
        "画像数": len(results),
        "需重算": need,
        "重复画像": dup,
        "未标文档类型": untagged,
        "画像": results,
        "建议动作": todo,
        "结论": ("全部正常" if not need and not dup and not untagged else
                 "；".join(filter(None, [
                     f"{len(need)} 个需重算" if need else "",
                     f"{len(dup)} 组重复画像需合并" if dup else "",
                     f"{len(untagged)} 个未标文档类型（重复检测失效）" if untagged else ""]))),
    }


@mcp.tool()
@_safe
def merge_profiles(target_id: str, source_ids: list[str],
                   dna_markdown: str | None = None, reason: str = "") -> dict:
    """把多个画像合并进 target。用于「聚类拆错了，其实是同一类」的场景。

    行为：
      - 规则按 assertion 去重，同一条保留 confidence 更高的
      - 相似但不同的规则（相似度 0.6–1.0）不自动合并，记进 conflicts.md 等人工裁决
      - 源画像标记「废弃（已并入 target）」，文件保留可追溯，不删除
      - 合并后 target 状态回到「候选」，必须重新确认才能用

    合并是不可逆的结构性操作，所以不做自动冲突裁决——宁可留给人看。
    """
    if not source_ids:
        return {"merged": False, "拒绝原因": "source_ids 为空"}
    tdir = _profile_dir(target_id)
    if not (tdir / "rules.json").exists():
        return {"merged": False, "拒绝原因": f"目标画像 {target_id} 不存在或无 rules.json"}

    sdirs = []
    for sid in source_ids:
        if sid == target_id:
            return {"merged": False, "拒绝原因": "source_ids 不能包含 target_id 自己"}
        sd = _profile_dir(sid)
        if not (sd / "rules.json").exists():
            return {"merged": False, "拒绝原因": f"源画像 {sid} 不存在或无 rules.json"}
        sdirs.append((sid, sd))

    def _rules(d):
        return json.loads((d / "rules.json").read_text(encoding="utf-8"))

    pool = list(_rules(tdir))
    for sid, sd in sdirs:
        for r in _rules(sd):
            r = dict(r)
            r.setdefault("来源画像", sid)
            pool.append(r)

    merged, dropped, conflicts = [], [], []
    rank = {"高": 3, "中": 2, "低": 1}
    for r in pool:
        a = (r.get("assertion") or "").strip()
        if not a:
            continue
        same = next((m for m in merged if m["a"] == a), None)
        if same:
            if rank.get(r.get("confidence"), 0) > rank.get(same["rule"].get("confidence"), 0):
                dropped.append(same["rule"].get("id"))
                merged.remove(same)
                merged.append({"rule": r, "a": a})
            else:
                dropped.append(r.get("id"))
            continue
        near = None
        for m in merged:
            ratio = SequenceMatcher(None, a, m["a"]).ratio()
            if 0.6 <= ratio < 1.0:
                near = (m, ratio)
                break
        if near:
            conflicts.append({"新规则": a[:70], "已有相近规则": near[0]["a"][:70],
                              "相似度": round(near[1], 2),
                              "新规则来源": r.get("来源画像", target_id)})
        merged.append({"rule": r, "a": a})

    final = [m["rule"] for m in merged]
    # 规则 ID 必须唯一，否则跨画像引用会指错。合并后整体重编号，原 ID 留档可追溯。
    for i, r in enumerate(final, 1):
        old = r.get("id")
        if old and old != f"R{i}":
            r.setdefault("原id", old)
        r["id"] = f"R{i}"
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    _ensure_dir(tdir / "history")
    for f in ("DNA.md", "rules.json", "meta.json"):
        src = tdir / f
        if src.exists():
            (tdir / "history" / f"{src.stem}-{ts}{src.suffix}").write_text(
                src.read_text(encoding="utf-8"), encoding="utf-8")

    (tdir / "rules.json").write_text(
        json.dumps(final, ensure_ascii=False, indent=2), encoding="utf-8")

    if dna_markdown:
        (tdir / "DNA.md").write_text(dna_markdown, encoding="utf-8")

    # 冲突记录必须独立处理：
    # 原实现写成 `if dna_markdown: ... elif conflicts:`，一旦调用方传了 DNA，
    # 冲突就永远不会落盘——而冲突恰恰是最需要留痕的东西。
    # 另外无冲突时也要归档旧记录，否则上一轮的冲突会一直挂在那里冒充新的。
    cf = tdir / "conflicts.md"
    if conflicts:
        lines = [f"# 合并冲突（{ts}）", "", f"来源：{'、'.join(source_ids)}", ""]
        for c in conflicts:
            lines += [f"## {c['新规则']}",
                      f"- 与已有规则相似度 {c['相似度']}：{c['已有相近规则']}",
                      f"- 来源画像：{c['新规则来源']}", "- 处理：**待人工裁决**", ""]
        cf.write_text("\n".join(lines), encoding="utf-8")
    elif cf.exists():
        (tdir / "history" / f"conflicts-{ts}.md").write_text(
            cf.read_text(encoding="utf-8"), encoding="utf-8")
        cf.write_text(
            f"# 合并冲突（{ts}）\n\n本次合并（来源：{'、'.join(source_ids)}）未发现疑似冲突。\n"
            "上一次的冲突记录已归档到 history/。\n", encoding="utf-8")

    mf = tdir / "meta.json"
    meta = json.loads(mf.read_text(encoding="utf-8")) if mf.exists() else {}
    meta.update({
        "status": "候选（合并后待确认）",
        "rule_count": len(final),
        "merged_from": list(meta.get("merged_from", [])) + source_ids,
        "merged_at": ts,
        "merge_reason": reason[:200] if reason else "",
    })
    mf.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    for sid, sd in sdirs:
        smf = sd / "meta.json"
        sm = json.loads(smf.read_text(encoding="utf-8")) if smf.exists() else {}
        sm["status"] = "废弃（已并入 %s）" % target_id
        sm["merged_into"] = target_id
        sm["merged_at"] = ts
        smf.write_text(json.dumps(sm, ensure_ascii=False, indent=2), encoding="utf-8")

    return {
        "merged": True,
        "target": target_id,
        "并入": source_ids,
        "规则总数": len(final),
        "去重丢弃": len(dropped),
        "疑似冲突": conflicts,
        "状态": "候选（合并后待确认）",
        "下一步": "人工确认合并结果；有疑似冲突请看 conflicts.md",
    }


# ================================================================
# 以下为「完全操控」补齐：服务自述 / 配置 / 规则级编辑 / 生命周期
# ================================================================

AUDIT_PATH = DATA_HOME / "audit.jsonl"


def _audit(action: str, profile_id: str, detail: dict) -> None:
    """
    所有写操作留痕。出问题时要能回答「谁在什么时候改了什么」。

    注意字段顺序：detail 展开必须在 action / ts / profile_id 之前，
    否则 detail 里同名字段（比如 resolve_conflict 的 action=keep_b）会把动作名覆盖掉。
    """
    rec = {**detail, "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "action": action, "profile_id": profile_id, "actor": "agent"}
    with open(AUDIT_PATH, "a", encoding="utf-8") as fh:
        fh.write(json.dumps(rec, ensure_ascii=False) + "\n")


def _load_json(p: Path, default=None):
    if not p.exists():
        return default
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return default


def _save_rules(d: Path, rules: list, action: str, pid: str,
                detail: dict, reset_status: bool = True) -> None:
    """写规则：先归档旧版，再重编号（保证 ID 唯一），最后同步 meta。"""
    _ensure_dir(d / "history")
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    src = d / "rules.json"
    if src.exists():
        (d / "history" / f"rules-{ts}.json").write_text(
            src.read_text(encoding="utf-8"), encoding="utf-8")
    for i, r in enumerate(rules, 1):
        old = r.get("id")
        if old and old != f"R{i}":
            r.setdefault("原id", old)
        r["id"] = f"R{i}"
    src.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")

    meta = _load_json(d / "meta.json", {}) or {}
    meta["rule_count"] = len(rules)
    if reset_status and "生效" in str(meta.get("status", "")):
        meta["status"] = "候选（规则变更，待重新确认）"
    (d / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    _audit(action, pid, detail)


def _check_evidence(spec, rule: dict) -> str:
    """返回空串表示通过，否则返回拒绝原因。"""
    miss = [k for k in ("assertion", "judge", "evidence")
            if not (rule.get(k) or "").strip()]
    if miss:
        return f"缺字段：{'、'.join(miss)}"
    ev = rule["evidence"].strip()
    if ev.startswith("spec."):
        ok, why = _spec_path_ok(spec, ev)
        return "" if ok else f"证据路径无法解析：{why}"
    if ev.startswith(("原文：", "原文:")):
        return ""
    return "evidence 须以 spec. 或「原文：」开头"


@_safe
def service_status() -> dict:
    """服务自述：口径版本与四层口径、引擎模块加载状态、画像库统计、配置、热加载基线。

    排查「改了代码为什么没生效」时先看这里——能看到每个模块的 mtime 与热加载基线，
    以及 server.py 自身是否需要重启（它无法热加载）。
    「口径说明」答的是「现在跑的是哪一版、哪层口径动过、旧画像要不要重算」。
    """
    mods = []
    for name in _ENGINE_MODULES:
        mod = sys.modules.get(name)
        f = getattr(mod, "__file__", None) if mod else None
        mt = None
        if f:
            try:
                mt = datetime.fromtimestamp(os.path.getmtime(f)).strftime("%m-%d %H:%M:%S")
            except OSError:
                mt = None
        base = _MTIME.get(name)
        mods.append({
            "模块": name,
            "已加载": mod is not None,
            "文件mtime": mt,
            "热加载基线": (datetime.fromtimestamp(base).strftime("%m-%d %H:%M:%S")
                       if base else "未建立（首调用后建立）"),
            "待重载": bool(base and f and os.path.getmtime(f) > base + 0.5),
        })

    items = list_profiles()["画像"]
    by_status: dict = {}
    for it in items:
        by_status[it["状态"]] = by_status.get(it["状态"], 0) + 1

    return {
        "口径版本": engine_version.ENGINE_VERSION,
        "口径说明": engine_version.version_summary(),
        "画像库": {
            "根目录": str(PROFILE_ROOT),
            "画像数": len(items),
            "按状态统计": by_status,
            "使用次数为0的画像": [it["profile_id"] for it in items
                            if not it.get("使用次数")],
        },
        "引擎模块": mods,
        "配置": service_config.load(),
        "规则指纹": engine_version.rules_fingerprint(),
        "工具层指纹": engine_version.server_fingerprint(),
        "配置文件": str(service_config.CONFIG_PATH),
        "审计日志": str(AUDIT_PATH),
        "指令行数": len(INSTRUCTIONS.splitlines()),
        "提示": "server.py 自身的改动无法热加载；新增工具也必须重启才能出现在工具列表",
    }


@_safe
def reload_engine(force: bool = False) -> dict:
    """手动触发引擎热加载。force=True 时无视 mtime 强制重载全部引擎模块。

    只对引擎模块（探针、判定、聚类、配置、版本）有效。
    server.py 自己改不了自己——新增工具或改工具签名仍需重启。
    """
    if force:
        changed, failed = [], []
        for name in _ENGINE_MODULES:
            mod = sys.modules.get(name)
            if not mod:
                continue
            try:
                importlib.reload(mod)
                changed.append(name)
                f = getattr(mod, "__file__", None)
                if f:
                    _MTIME[name] = os.path.getmtime(f)
            except Exception as e:  # noqa: BLE001
                failed.append(f"{name}: {type(e).__name__}")
    else:
        changed, failed = _maybe_reload(), []
    _audit("reload_engine", "-", {"force": force, "changed": changed})
    return {"已重载": changed or "无变化（mtime 未变）",
            "失败": failed or "无",
            "当前口径版本": engine_version.ENGINE_VERSION,
            "提醒": "server.py 自身不会被重载"}


@_safe
def get_config() -> dict:
    """读取服务配置：质量门槛、聚类阈值、反馈与陈旧阈值、句长上限、提升校验条件。"""
    return {
        "当前配置": service_config.load(),
        "配置文件": str(service_config.CONFIG_PATH),
        "可改范围": "只放判定阈值；统计口径不在这里，口径变更必须走 version 升级 + 全量重算",
    }


@_safe
def set_config(patch: dict, confirm: bool = False) -> dict:
    """修改配置，深合并。patch 形如 {"quality": {"pass": 85}}。

    默认只预览不写入，必须 confirm=True 才落盘——防止模型手滑把门槛改坏。
    写入前校验三件事：未知键（会静默失效，必须拦）、类型、跨字段矛盾。
    """
    problems = service_config.validate(patch, service_config.load())
    if problems:
        return {"已写入": False, "拒绝原因": problems}

    merged = service_config._merge(service_config.load(), patch)

    if not confirm:
        return {"已写入": False, "预览": merged,
                "提示": "确认无误后带 confirm=True 再调一次"}

    new = service_config.patch(patch)
    _audit("set_config", "-", {"patch": patch})
    return {"已写入": True, "当前配置": new}


@_safe
def add_rule(profile_id: str, rule: dict) -> dict:
    """给画像追加一条规则。自动做证据校验、重编号、归档旧版、写审计。

    规则必须含 assertion（可执行句）/ judge（怎么验）/ evidence（spec.字段 或 原文：）。
    confidence 缺省按「中」处理，并在返回里提示。
    若画像原本是「生效」状态，改完会自动退回「候选」，需重新确认。
    """
    if not isinstance(rule, dict):
        return {"added": False, "拒绝原因": "rule 必须是对象",
                "示例": {"assertion": "…", "judge": "…", "evidence": "原文：…",
                         "confidence": "高"}}
    rule = {**rule}
    no_conf = not (rule.get("confidence") or "").strip()
    rule.setdefault("confidence", "中")

    d = _profile_dir(profile_id)
    if not (d / "rules.json").exists():
        return {"added": False, "拒绝原因": f"画像 {profile_id} 不存在或无 rules.json"}
    spec = _load_json(d / "spec.json")
    why = _check_evidence(spec, rule)
    if why:
        return {"added": False, "拒绝原因": why}

    rules = _load_json(d / "rules.json", []) or []
    rules.append(dict(rule))
    _save_rules(d, rules, "add_rule", profile_id, {"rule": rule.get("assertion", "")[:60]})
    out = {"added": True, "规则总数": len(rules),
           "提示": "规则变更后画像退回候选状态，需 manage_profile(action=\"promote\") 重新确认"}
    if no_conf:
        out["警告"] = "没写 confidence，已按「中」处理；置信度影响后续规则可信度判断，建议补上"
    return out


@_safe
def update_rule(profile_id: str, rule_id: str, patch: dict) -> dict:
    """修改指定规则。只允许改 assertion / judge / evidence / confidence，
    改完会重新校验证据并归档旧版。"""
    d = _profile_dir(profile_id)
    if not (d / "rules.json").exists():
        return {"updated": False, "拒绝原因": f"画像 {profile_id} 不存在"}
    allowed = {"assertion", "judge", "evidence", "confidence"}
    bad = [k for k in patch if k not in allowed]
    if bad:
        return {"updated": False, "拒绝原因": f"不允许修改字段：{'、'.join(bad)}"}

    rules = _load_json(d / "rules.json", []) or []
    target = next((r for r in rules if r.get("id") == rule_id), None)
    if target is None:
        return {"updated": False, "拒绝原因": f"找不到规则 {rule_id}",
                "可用": [r.get("id") for r in rules]}

    merged = {**target, **patch}
    why = _check_evidence(_load_json(d / "spec.json"), merged)
    if why:
        return {"updated": False, "拒绝原因": why}

    idx = rules.index(target)
    rules[idx] = merged
    _save_rules(d, rules, "update_rule", profile_id,
                {"rule_id": rule_id, "patch": patch})
    return {"updated": True, "规则": merged}


@_safe
def delete_rule(profile_id: str, rule_id: str, reason: str) -> dict:
    """删除一条规则。必须写 reason（进审计日志）。删完退回候选状态。"""
    if not reason.strip():
        return {"deleted": False, "拒绝原因": "必须填写 reason，删除操作要留痕"}
    d = _profile_dir(profile_id)
    if not (d / "rules.json").exists():
        return {"deleted": False, "拒绝原因": f"画像 {profile_id} 不存在"}
    rules = _load_json(d / "rules.json", []) or []
    target = next((r for r in rules if r.get("id") == rule_id), None)
    if target is None:
        return {"deleted": False, "拒绝原因": f"找不到规则 {rule_id}",
                "可用": [r.get("id") for r in rules]}
    if len(rules) <= 1:
        return {"deleted": False, "拒绝原因": "画像至少要保留一条规则，否则等于空画像"}

    rules.remove(target)
    _save_rules(d, rules, "delete_rule", profile_id,
                {"rule_id": rule_id, "assertion": target.get("assertion", "")[:60],
                 "reason": reason})
    return {"deleted": True, "规则总数": len(rules), "被删规则": target.get("assertion")}


@_safe
def update_profile_meta(profile_id: str, patch: dict, reason: str = "") -> dict:
    """改元信息，不用重写整份 DNA。典型用途：补/改「文档类型」「面向人群」「置信度」。

    不能通过这里改 status、confirmed_by、use_count —— 状态流转必须走
    promote_profile / retire_profile，否则审计链会断。
    """
    d = _profile_dir(profile_id)
    mf = d / "meta.json"
    if not mf.exists():
        return {"updated": False, "拒绝原因": f"画像 {profile_id} 的 meta.json 不存在"}

    locked = {"status", "confirmed_by", "promoted_at", "use_count", "last_used_at"}
    hit = locked & set(patch or {})
    if hit:
        return {"updated": False,
                "拒绝原因": f"这些字段不能直接改：{'、'.join(sorted(hit))}",
                "说明": "状态走 manage_profile(action=\"promote\"/\"retire\")；使用统计由 get_profile 自动记"}

    meta = json.loads(mf.read_text(encoding="utf-8"))
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    hist = d / "history"
    _ensure_dir(hist)
    (hist / f"meta-{ts}.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    before = {k: meta.get(k) for k in (patch or {})}
    meta.update(patch or {})
    if meta.get("doc_type") and not meta.get("文档类型"):
        meta["文档类型"] = meta["doc_type"]
    meta["meta_updated_at"] = ts
    mf.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    _audit("update_profile_meta", profile_id,
           {"reason": reason,
            "changed": {k: [before.get(k), v] for k, v in (patch or {}).items()}})
    return {"updated": True,
            "变更": {k: {"原值": before.get(k), "新值": v} for k, v in (patch or {}).items()},
            "旧版": f"history/meta-{ts}.json"}


@_safe
def promote_profile(profile_id: str, confirmed_by: str, note: str = "") -> dict:
    """把画像从「候选」提升为「生效」，生效后生成侧才能拉取。

    **必须填 confirmed_by**（谁确认的），写进审计日志。这是刻意保留的唯一人工关卡：
    画像一旦生效会影响后续所有产出，不能让模型自己拍板。

    提升前自动校验四项：文件完整、口径未过期、无未裁决冲突、规则数达标。
    """
    if not (confirmed_by or "").strip():
        return {"promoted": False,
                "拒绝原因": "必须填写 confirmed_by（谁确认的）。这一步需要人明确授权，模型不能自签。"}

    d = _profile_dir(profile_id)
    if not (d / "DNA.md").exists():
        return {"promoted": False, "拒绝原因": f"画像 {profile_id} 不存在"}

    cfg = service_config.load()["promotion"]
    rules = _load_json(d / "rules.json", []) or []
    spec = _load_json(d / "spec.json", {}) or {}
    meta = _load_json(d / "meta.json", {}) or {}

    blocks, warns = [], []

    # 废弃是人工明确决定，不能被一次 promote 悄悄撤销。
    # 想复活就 clone 一份重新确认——这样「废弃」这个动作才有意义。
    status = str(meta.get("status") or "")
    if "废弃" in status:
        return {"promoted": False,
                "拒绝原因": f"画像已废弃（{status}），不能直接重新生效",
                "建议": 'manage_profile(action="clone") 复制成新 ID 再重新确认'}

    # scope 决定门槛，不只是一个标签：
    #   通用 —— 跨项目复用、几乎不过期，要求规则更多、必须写适用描述
    #   临时 —— 试验用，不允许转正
    scope = str(meta.get("scope") or "项目")
    if scope == "临时":
        return {"promoted": False,
                "拒绝原因": "scope=临时 的画像不允许转正",
                "建议": "确认要长期使用，就先 manage_profile(action=\"update_meta\") "
                        "把 scope 改成「项目」或「通用」，再 promote"}
    need = cfg.get("min_rules_shared", 8) if scope == "通用" else cfg["min_rules"]
    if scope == "通用" and not (meta.get("适用描述") or "").strip():
        blocks.append("scope=通用 的画像必须写「适用描述」——"
                      "跨项目复用时别人靠这句话判断适不适用")

    if not rules:
        blocks.append("无规则")
    elif len(rules) < need:
        warns.append(f"规则仅 {len(rules)} 条，低于{'通用画像' if scope == '通用' else ''}"
                     f"建议下限 {need}")

    # 生效画像必须带文档类型：生成侧是按类型找画像的，没有类型等于找不到；
    # 同时「重复画像」检测也依赖它，缺失会让两个同类画像悄悄共存。
    if not (meta.get("文档类型") or meta.get("doc_type") or "").strip():
        blocks.append("meta 缺「文档类型」，生效后无法按类型被检索、也无法参与重复检测。"
                      "请先补 meta.文档类型（或调用 manage_profile(action=\"update_meta\"))")

    need, why = engine_version.needs_rebuild(spec.get("engine_version"))
    if need and cfg["block_on_stale"]:
        blocks.append(f"口径问题：{why}")

    cf = d / "conflicts.md"
    if cf.exists():
        txt = cf.read_text(encoding="utf-8")
        n = txt.count("待人工裁决")
        if n and cfg["block_on_conflict"]:
            blocks.append(f"有 {n} 条未裁决冲突，先 resolve_conflict")

    if blocks:
        return {"promoted": False, "拒绝原因": blocks, "警告": warns}

    meta["status"] = "生效"
    meta["promoted_at"] = datetime.now().strftime("%Y%m%d-%H%M%S")
    meta["confirmed_by"] = confirmed_by
    if note:
        meta["promote_note"] = note[:200]
    (d / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    _audit("promote_profile", profile_id,
           {"confirmed_by": confirmed_by, "note": note, "rule_count": len(rules)})
    return {"promoted": True, "状态": "生效", "确认人": confirmed_by,
            "规则数": len(rules), "警告": warns or "无"}


@_safe
def retire_profile(profile_id: str, reason: str) -> dict:
    """废弃画像：状态改为「废弃」，文件全部保留可追溯，不删除。必须写 reason。"""
    if not reason.strip():
        return {"retired": False, "拒绝原因": "必须填写 reason"}
    d = _profile_dir(profile_id)
    if not (d / "meta.json").exists():
        return {"retired": False, "拒绝原因": f"画像 {profile_id} 不存在"}
    meta = _load_json(d / "meta.json", {}) or {}
    meta["status"] = "废弃"
    meta["retired_at"] = datetime.now().strftime("%Y%m%d-%H%M%S")
    meta["retire_reason"] = reason[:200]
    (d / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    _audit("retire_profile", profile_id, {"reason": reason})
    return {"retired": True, "状态": "废弃",
            "提示": "文件保留，可用 manage_profile(action=\"clone\") 复活"}


@_safe
def clone_profile(profile_id: str, new_id: str, note: str = "") -> dict:
    """复制画像为新 ID。用于「想试试改规则但不想动原画像」。

    副本状态一律为「候选」，并清空使用统计——它是新画像，使用次数从 0 算。
    """
    src = _profile_dir(profile_id)
    if not (src / "DNA.md").exists():
        return {"cloned": False, "拒绝原因": f"源画像 {profile_id} 不存在"}
    dst = _profile_dir(new_id)
    if dst.exists():
        return {"cloned": False, "拒绝原因": f"目标 {new_id} 已存在，不能覆盖"}

    _ensure_dir(dst)
    for f in ("DNA.md", "rules.json", "spec.json", "conflicts.md"):
        if (src / f).exists():
            (dst / f).write_text((src / f).read_text(encoding="utf-8"), encoding="utf-8")
    meta = _load_json(src / "meta.json", {}) or {}
    meta.pop("use_count", None)
    meta.pop("last_used_at", None)
    meta["status"] = "候选"
    meta["cloned_from"] = profile_id
    meta["cloned_at"] = datetime.now().strftime("%Y%m%d-%H%M%S")
    if note:
        meta["clone_note"] = note[:200]
    (dst / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    _audit("clone_profile", new_id, {"from": profile_id, "note": note})
    return {"cloned": True, "新画像": new_id, "状态": "候选"}


@_safe
def resolve_conflict(profile_id: str, index: int, action: str,
                     note: str = "") -> dict:
    """裁决 conflicts.md 里的第 index 条冲突。**序号从 1 开始**（第一条是 1，不是 0）。

    先看 manage_profile(action="check_conflicts") 或画像目录下 conflicts.md 的编号，
    再把对应序号传进来。传 0 或超出总数会返回可用范围。

    action：
      keep_both —— 两条都保留，只在文件里标记已裁决
      keep_a    —— 保留「已有」那条，删掉「新规则」
      keep_b    —— 保留「新规则」，删掉「已有」那条
      dismiss   —— 误判，两条都留，标记为非冲突
    """
    d = _profile_dir(profile_id)
    cf = d / "conflicts.md"
    if not cf.exists():
        return {"resolved": False, "拒绝原因": "该画像没有 conflicts.md"}
    if action not in ("keep_both", "keep_a", "keep_b", "dismiss"):
        return {"resolved": False,
                "拒绝原因": "action 只能是 keep_both / keep_a / keep_b / dismiss"}

    txt = cf.read_text(encoding="utf-8")
    blocks = txt.split("\n## ")
    # blocks[0] 是文件头，其余每块一条冲突
    total_cf = len(blocks) - 1
    if index < 1 or index > total_cf:
        ret = {"resolved": False,
               "拒绝原因": f"conflict_index 超范围：当前共 {total_cf} 条冲突",
               "说明": f"序号**从 1 开始**，可用范围是 1–{total_cf}；收到的是 {index}"}
        if index == 0:
            ret["提示"] = "传 0 是常见的 off-by-one 误用。第一条冲突的序号是 1，不是 0"
        return ret

    block = blocks[index]
    title = block.splitlines()[0].strip()
    m_new = re.search(r"相似度\s*[\d.]+：(.+)", block)
    existing = m_new.group(1).strip() if m_new else None

    deleted = None
    if action in ("keep_a", "keep_b"):
        target_assertion = existing if action == "keep_b" else title
        rules = _load_json(d / "rules.json", []) or []
        hit = next((r for r in rules if (r.get("assertion") or "").strip() == target_assertion), None)
        if hit is None:
            return {"resolved": False,
                    "拒绝原因": f"找不到 assertion 为「{target_assertion}」的规则，"
                              "可能已被修改，请用 delete_rule 手动处理"}
        if len(rules) <= 1:
            return {"resolved": False, "拒绝原因": "只剩一条规则，不能删"}
        rules.remove(hit)
        _save_rules(d, rules, "resolve_conflict", profile_id,
                    {"index": index, "action": action, "deleted": target_assertion[:50]})
        deleted = target_assertion

    label = {"keep_both": "已裁决：两条都保留", "keep_a": "已裁决：保留已有规则",
             "keep_b": "已裁决：保留新规则", "dismiss": "已裁决：误判，非冲突"}[action]
    block = block.replace("**待人工裁决**", f"**{label}**" + (f"（{note}）" if note else ""))
    blocks[index] = block
    cf.write_text("\n## ".join(blocks), encoding="utf-8")
    _audit("resolve_conflict", profile_id,
           {"index": index, "action": action, "note": note, "deleted": deleted})

    rest = cf.read_text(encoding="utf-8").count("待人工裁决")
    return {"resolved": True, "裁决": label, "删除的规则": deleted or "无",
            "剩余待裁决": rest}


@_safe
def export_profile(profile_id: str, out_path: str) -> dict:
    """把画像打包成单个 JSON 文件（DNA + 规则 + 硬指标 + 元信息），用于备份或迁移。"""
    d = _profile_dir(profile_id)
    if not (d / "DNA.md").exists():
        return {"exported": False, "拒绝原因": f"画像 {profile_id} 不存在"}
    bundle = {
        "profile_id": profile_id,
        "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "engine_version": engine_version.ENGINE_VERSION,
        "DNA": (d / "DNA.md").read_text(encoding="utf-8"),
        "rules": _load_json(d / "rules.json", []),
        "spec": _load_json(d / "spec.json"),
        "meta": _load_json(d / "meta.json", {}),
    }
    out = Path(out_path).expanduser()
    try:
        out.parent.mkdir(parents=True, exist_ok=True)   # 父目录不存在就建，别假装兜底
        out.write_text(json.dumps(bundle, ensure_ascii=False, indent=2),
                       encoding="utf-8")
    except OSError as e:
        return {"exported": False,
                "拒绝原因": f"无法写入 {out}：{e.strerror or e}",
                "建议": "确认目标目录可写；不要导出到 /System 等受保护目录"}
    _audit("export_profile", profile_id, {"out": str(out)})
    return {"exported": True, "path": str(out),
            "大小": f"{out.stat().st_size} 字节"}


# ============================================================================
# 对外只暴露 9 个工具，按「任务动作」分组，而不是按对象的增删改查拆。
#
# 25 个细粒度工具的代价：模型在选工具上犯错（同一件事有 3 个入口）、
# 每个工具的 schema 和描述白占上下文、调用方要自己拼流程。
# 上面那些实现（extract_format / add_rule / promote_profile …）仍然保留，
# 但不再注册，只作为这里的内部实现。
# ============================================================================


def _stamp(result: dict, path: Path) -> dict:
    """给结果盖审计戳：口径版本 + 规则指纹 + 输入指纹。

    为什么必须有：同一份文件 13 分钟内从 93 分变 100 分，如果没有戳，
    唯一的解释是「代码被改过」这句话——不可验证。有了戳就能说清：
    口径变了 / 规则变了 / 文件变了，三者必居其一。
    """
    import hashlib
    try:
        st = path.stat()
        # 内容摘要而非 mtime：拷贝/还原会改 mtime 但内容没变
        digest = hashlib.sha1(path.read_bytes()).hexdigest()[:12]
        result["_审计"] = {
            "口径版本": engine_version.ENGINE_VERSION,
            "规则指纹": engine_version.rules_fingerprint(),      # 数字怎么算
            "工具层指纹": engine_version.server_fingerprint(),   # 接口长什么样
            "输入指纹": digest,
            "输入大小": st.st_size,
            "输入修改时间": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            "生成时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "指纹说明": "规则指纹只覆盖 6 个统计引擎模块；改接口/加参数/改返回结构看"
                        "「工具层指纹」。两个都不变而结果变了 = 文件本身变了，比输入指纹。",
        }
    except OSError:
        result["_审计"] = {"口径版本": engine_version.ENGINE_VERSION,
                          "规则指纹": engine_version.rules_fingerprint(),
                          "工具层指纹": engine_version.server_fingerprint(),
                          "输入指纹": None, "说明": "无法读取文件，未计算输入指纹"}
    return result


def _raise_sub_error(result: dict) -> dict:
    """把子步骤里的错误提到顶层。

    内部实现都套了 @_safe，出错时返回 {"error": ...} 而不是抛异常。如果直接塞进
    结果里，调用方会看到「质量门」这个键存在，就以为分析成功了——实际里面只有一条
    报错。这种「看起来成功」比直接失败更危险，所以统一提到顶层。
    """
    for stage, val in result.items():
        if isinstance(val, dict) and val.get("error"):
            return {"file": result.get("file"), "类型": result.get("类型"),
                    "错误阶段": stage, "error": val["error"],
                    "错误类型": val.get("错误类型"), "建议": val.get("建议"),
                    "提示": "该阶段失败，整份分析结果不可用；修正后重试"}
    return result


# 按「你要干什么」预设，而不是按「数据长什么样」。
#
# 原来只有 want（要哪几块）和 detail（要多少），这两个都是按数据的形状切的：
# 建画像时真正要的是「质量门 + 文风 brief + 章节结构」，而不是「全量 spec 的 7.6 KB」。
# 实测全量 14.4 KB，其中表格明细 2.8 KB 在建画像时用不上。
# 这借的是 skill 的思路：先用 100 字的描述路由，正文按需加载。
_PRESETS = {
    "audit": {
        "用途": "验收已成稿文档：质量门（含交付检查）+ 排版提示",
        "want": ["quality", "format"],
        "spec": "none",          # 排版只要「提示」，不要 7 KB 的 spec
        "brief": "none",
        "约": "1–2 KB",
    },
    "build_profile": {
        "用途": "建画像：质量门 + 文风（含 6 条 brief）+ 章节结构",
        "want": ["quality", "voice", "format"],
        "spec": ["file", "path", "page", "paragraphs", "styles",
                 "numbering_scheme", "outline"],   # 不含 tables
        "brief": "full",
        "约": "10 KB",
    },
    "full": {
        "用途": "全量，排查问题用",
        "want": None,
        "spec": None,
        "brief": "full",
        "约": "14 KB",
    },
}


def _apply_preset(out: dict, cfg: dict) -> dict:
    """按 preset 裁剪返回。只裁「排版」的 spec 和文风的 brief，指标一个不动。

    指标不能裁：那是可核对的事实，裁掉之后就无法回溯某个结论是怎么来的。
    """
    fmt = out.get("排版")
    if isinstance(fmt, dict):
        keep = cfg.get("spec")
        if keep == "none":
            fmt.pop("spec", None)
            fmt["spec_已省略"] = "preset=audit 不含排版规格明细；需要时用 preset=full"
        elif keep is not None:
            spec = fmt.get("spec", {}) or {}
            dropped = [k for k in spec if k not in keep]
            fmt["spec"] = {k: spec[k] for k in keep if k in spec}
            if dropped:
                fmt["spec_已省略"] = (f"preset 未包含：{'、'.join(dropped)}；"
                                      "需要时用 preset=full")

    voice = out.get("文风")
    if cfg.get("brief") == "none" and isinstance(voice, dict):
        if voice.pop("analysis_brief", None):
            voice["brief_已省略"] = "preset=audit 不需要 brief；建画像请用 preset=build_profile"
    return out


@mcp.tool()
@_safe
def analyze_document(path: str, want: list[str] | None = None,
                     detail: str = "full", doc_type: str = "auto",
                     preset: str = "full") -> dict:
    """一份文档的全部分析：准入质量 + 规格（排版/视觉）+ 文风硬指标。

    扩展名自动分流：`.docx` 走 Word 链路，`.pptx` 走 PPT 链路，不需要你判断该调哪个。

    **按任务选 preset，比按字段挑 want 更省事**：
      preset="audit"          验收已成稿文档。质量门（含交付检查）+ 排版提示，约 1–2 KB
      preset="build_profile"  建画像。质量门 + 文风（含 6 条 brief）+ 章节结构，约 10 KB
                              （不含表格明细——那是排查用的）
      preset="full"           全量，约 14 KB。排查问题、要看表格/样式明细时用

    高级用法：want 可只取某几块
      Word  ["quality", "format", "voice"]（默认全给）
      PPT   ["quality", "style"]（默认全给）
    detail="compact" 会把 outline 砍到 15 条、brief 只留标题，一般不用，
    需要按字段做比 preset 更细的裁剪时才用。

    **doc_type 决定按什么标准判**，默认 "auto"（通用标准）：
      公文      要求首行缩进 2 字符、必须有目录
      技术文档  不缩进是常态；段落 <150 时缺目录不扣分
      网页      同上
      报告      通用标准
    不声明就等于按公文口径判——技术文档的常规做法会被记成缺陷。
    这是**唯一**能让评分贴合文档实际类型的方式，不要省。

    返回里若出现「先读这个」，说明文风指标报了自检告警——某个数字不可信
    （情态词被「需求/需方」这类复合词撑大、句长被表格转文本拉偏等）。
    这时不要拿那个数字去写规则，先读告警说明。

    所有返回都带 `_审计`（口径版本 / 规则指纹 / 输入指纹）。两次结果不一样时，
    先比这三个值，能立刻分清是口径变了、规则改了，还是文件本身换了。
    """
    if preset not in _PRESETS:
        raise ValueError(f"preset 只支持 {'/'.join(_PRESETS)}，收到 {preset}")
    cfg = _PRESETS[preset]

    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"文件不存在：{path}")
    suffix = p.suffix.lower()
    want = want or []
    # preset 提供默认的 want；显式传了 want 就用显式的
    if not want and cfg.get("want"):
        want = list(cfg["want"])

    if suffix == ".docx":
        keys = want or ["quality", "format", "voice"]
        out = {"file": p.name, "类型": "Word"}
        if "quality" in keys:
            out["质量门"] = check_quality(str(p), doc_type=doc_type)
        if "format" in keys:
            f = extract_format(str(p))
            if detail == "compact":
                spec = f.get("spec", {})
                outline = spec.get("outline", []) or []
                tables = spec.get("tables", []) or []
                spec["outline"] = {"总数": len(outline),
                                   "前 15 条": outline[:15],
                                   "说明": "compact 模式只给前 15 条；完整结构用 detail=full"}
                spec["tables"] = {"总数": len(tables),
                                  "首表": tables[0] if tables else None,
                                  "说明": "compact 模式只给首表；列宽/正文字号等明细用 detail=full"}
            out["排版"] = f
        if "voice" in keys:
            v = extract_voice(str(p))
            if detail == "compact":
                v["analysis_brief"] = [{"id": b["id"], "任务": b["任务"]}
                                       for b in v.get("analysis_brief", [])]
            out["文风"] = v
            warn = v.get("自检告警") or []
            if warn:
                out["先读这个"] = (f"文风提取有 {len(warn)} 条自检告警（详情见 文风.自检告警），"
                                   "涉及指标的数字不可直接引用")
        _apply_preset(out, cfg)
        return _stamp(_raise_sub_error(out), p)

    if suffix == ".pptx":
        keys = want or ["quality", "style"]
        if preset == "build_profile" and not want:
            keys = ["quality", "style"]      # PPT 没有文风层，建画像同样要看版式
        if preset == "audit" and not want:
            keys = ["quality"]
        out = {"file": p.name, "类型": "PPT"}
        if "quality" in keys:
            out["质量门"] = check_quality_pptx(str(p))
        if "style" in keys:
            out["视觉与版式"] = extract_pptx_style(str(p))
        return _stamp(_raise_sub_error(out), p)

    raise ValueError(f"只支持 .docx / .pptx，收到 {p.suffix}")


@mcp.tool()
@_safe
def analyze_batch(paths: list[str], exclude_rejected: bool = True,
                  with_compare: bool = False, doc_type: str = "auto") -> dict:
    """一批文档：质量门筛选 → 判该建几个画像 →（可选）逐项一致性对比。

    cluster 用十维特征（排版 5 + 结构 3 + 文风 2）：
      差异 ≤1 维  自动合并
      差异 =2 维  挂起，进「存疑对」等你裁决（错合不可逆，错分只是多确认一次）
      差异 ≥3 维  判为另一套模板，拆成两个画像

    **doc_type 决定批量质量门按什么标准判**，默认 "auto"（通用/公文标准）：
      公文      要求首行缩进 2 字符、必须有目录
      技术文档  不缩进是常态；段落 <150 时缺目录不扣分
      网页      同上；报告 = 通用标准
    不声明就等于按公文口径判。技术文档若不显式传 doc_type=技术文档，
    会被判成拒收、不参与建画像——与 analyze_document(doc_type=技术文档)
    的结论不一致。返回值里的「质量门口径」会写明本次用的是哪套。

    with_compare=True 时按**聚类组**给出「组内哪些规格一致、哪些不一致」。
    组内只有 1 份说明样本不足，建出来的画像会过拟合。
    """
    names = [Path(x).name for x in paths]
    dup = sorted({n for n in names if names.count(n) > 1})
    out = cluster_documents(paths, exclude_rejected=exclude_rejected,
                            doc_type=doc_type)
    # 口径必须随结果一起返回：以前调用方看不到批量质量门用的是哪套标准，
    # 「为什么这批文档被剔了」无从查起。连 error 分支也要带。
    out["质量门口径"] = {
        "本次": doc_type if exclude_rejected else "未启用（exclude_rejected=false）",
        "说明": ("批量工具默认 auto＝通用/公文标准。技术文档请显式传 "
                "doc_type=技术文档，否则会按公文口径判成拒收、不参与建画像；"
                "这与 analyze_document 指定类型后的结论可能不同"),
    }

    if dup:
        # 分组只带文件名。同名不同目录的两份，名字上分不开——这时做「一致性对比」
        # 会把同一份文件拿来跟自己对，得出「完全一致」的假结论。所以直接不做。
        out["警告"] = (f"输入里有重名文件（{'、'.join(dup[:3])}"
                       f"{' 等' if len(dup) > 3 else ''}），分组结果无法区分它们，"
                       "一致性对比已跳过；请确认是否有重复导入")

    if with_compare and not out.get("error") and not dup:
        by_name = {Path(x).name: x for x in paths}
        usable = [by_name[n] for g in out.get("分组", []) for n in g.get("文档", [])
                  if n in by_name]
        if len(usable) >= 2:
            out["一致性对比"] = compare_documents(
                usable, cluster_groups=out.get("分组"))
    return out


@mcp.tool()
@_safe
def get_profile(profile_id: str | None = None, purpose: str | None = None,
                track: bool = True) -> dict:
    """读画像。传 profile_id 读一份（DNA + 规则 + 硬指标 + meta），不传则列清单。

    读一份会记一次使用（use_count +1、更新 last_used_at），用于事后识别
    「建了却从来没人拉」的死画像。只是查看不想留痕就 track=False。

    清单里 use_count 长期为 0、或状态含「过期/废弃」的，不要拿来当参考。
    """
    if profile_id:
        return load_profile(profile_id, purpose=purpose, track=track)
    return list_profiles()


@mcp.tool()
@_safe
def edit_rules(profile_id: str, action: str, rule: dict | None = None,
               rule_id: str | None = None, patch: dict | None = None,
               reason: str = "") -> dict:
    """改单条规则，不用重写整份 DNA。

    action：
      add     rule = {"assertion","judge","evidence","confidence"}
      update  rule_id + patch（只允许改 assertion/judge/evidence/confidence）
      delete  rule_id + reason（reason 必填，删除要留痕）

    evidence 必须是 `spec.xxx`（服务端会校验该字段真实存在）或 `原文：<片段>`。
    任何改动都会把「生效」画像退回候选，需要重新 promote——规则变了就得重新确认。
    """
    if action == "add":
        if not rule:
            return {"error": "action=add 需要 rule 参数",
                    "示例": {"assertion": "单句不超过 26 字", "judge": "统计句长",
                             "evidence": "spec.voice.截尾平均句长=18.2", "confidence": "高"}}
        return add_rule(profile_id, rule)
    if action == "update":
        if not rule_id or not patch:
            return {"error": "action=update 需要 rule_id 和 patch"}
        return update_rule(profile_id, rule_id, patch)
    if action == "delete":
        if not rule_id or not reason:
            return {"error": "action=delete 需要 rule_id 和 reason（删除要留痕）"}
        return delete_rule(profile_id, rule_id, reason)
    return {"error": f"未知 action：{action}", "可用": ["add", "update", "delete"]}


@mcp.tool()
@_safe
def manage_profile(profile_id: str, action: str, confirmed_by: str = "",
                   reason: str = "", note: str = "", new_id: str = "",
                   patch: dict | None = None, conflict_index: int | None = None,
                   decision: str = "", out_path: str = "") -> dict:
    """画像状态与元信息的总入口。

    action：
      promote           候选 → 生效。**confirmed_by 必填**（谁确认的），模型不能自签
      retire            → 废弃。reason 必填。文件保留，可追溯「当时为什么这么写」
      clone             复制成 new_id 试改，使用统计清零
      update_meta       改 patch 里的元信息（文档类型/面向人群/置信度…）；不能改 status
      resolve_conflict  裁决合并冲突：conflict_index + decision
                        （keep_both / keep_a / keep_b / dismiss）
      export            导出成单个 JSON，out_path 必填

    合并多份画像不走这里，用 merge_profiles（它要多传 dna_markdown 和多个来源）。
    """
    if action == "promote":
        if not confirmed_by:
            return {"error": "action=promote 需要 confirmed_by（谁确认的），模型不能自签"}
        return promote_profile(profile_id, confirmed_by, note=note)
    if action == "retire":
        if not reason:
            return {"error": "action=retire 需要 reason"}
        return retire_profile(profile_id, reason)
    if action == "clone":
        if not new_id:
            return {"error": "action=clone 需要 new_id"}
        return clone_profile(profile_id, new_id, note=note)
    if action == "update_meta":
        if not patch:
            return {"error": "action=update_meta 需要 patch"}
        return update_profile_meta(profile_id, patch, reason=reason)
    if action == "resolve_conflict":
        if conflict_index is None or not decision:
            return {"error": "action=resolve_conflict 需要 conflict_index 和 decision",
                    "decision 可选": ["keep_both", "keep_a", "keep_b", "dismiss"]}
        return resolve_conflict(profile_id, conflict_index, decision, note=note)
    if action == "export":
        if not out_path:
            return {"error": "action=export 需要 out_path"}
        return export_profile(profile_id, out_path)
    return {"error": f"未知 action：{action}",
            "可用": ["promote", "retire", "clone", "update_meta",
                     "resolve_conflict", "export"]}


def _call_stats() -> dict:
    """从 calls.jsonl 汇总运行指标。没有日志时明确说「没有数据」，不编。"""
    if not CALL_LOG.exists():
        return {"状态": "暂无调用记录（calls.jsonl 不存在）",
                "说明": "指标从第一次工具调用开始采集"}
    rows = []
    for line in CALL_LOG.read_text(encoding="utf-8").splitlines():
        try:
            rows.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    if not rows:
        return {"状态": "调用记录为空"}

    by_tool = {}
    for r in rows:
        t = r.get("tool", "?")
        d = by_tool.setdefault(t, {"次数": 0, "ok": 0, "rejected": 0,
                                   "exception": 0, "_ms": []})
        d["次数"] += 1
        st = r.get("s", "ok")
        d[st] = d.get(st, 0) + 1
        d["_ms"].append(r.get("ms") or 0)

    out = {}
    for t, d in sorted(by_tool.items(), key=lambda x: -x[1]["次数"]):
        ms = sorted(d.pop("_ms"))
        d["平均耗时(ms)"] = round(sum(ms) / len(ms), 1) if ms else 0
        # 给极值：耗时是否被截断，看 min/max 的区间宽度就能判断。
        # 若计时只包住函数尾部，min 和 max 会同时趋近 0 且几乎相等；
        # 真实分布里「解析 docx」与「读写小 JSON」会差两三个数量级。
        d["最小(ms)"] = round(ms[0], 1) if ms else 0
        d["最大(ms)"] = round(ms[-1], 1) if ms else 0
        d["P95(ms)"] = (round(ms[min(len(ms) - 1, int(len(ms) * 0.95))], 1)
                        if ms else 0)
        # 三档分开：失败=服务缺陷（要修）、被拒=前置条件不满足（正常拦截）、
        # 参数错=调用方传错（看文档）。合并成一个数会掩盖问题。
        d["失败(bug)"] = d.get("bug", 0) + d.get("exception", 0)
        d["被拒"] = d.get("rejected", 0)
        d["参数错"] = d.get("bad_input", 0)
        out[t] = {k: v for k, v in d.items()
                  if k in ("次数", "平均耗时(ms)", "最小(ms)", "最大(ms)", "P95(ms)",
                           "失败(bug)", "被拒", "参数错")}

    status = Counter(_verdict_of_from_row(r) for r in rows)
    ok_n = status.get("ok", 0)
    bug_n = status.get("bug", 0)
    rej_n = status.get("rejected", 0)
    bad_n = status.get("bad_input", 0)
    total = len(rows)

    # 成功率不能把「正常拦截」和「调用方传错」算成失败。
    # 一份健康服务如果拒绝了两次不合规请求、被传错一次路径，
    # 按 ok/总数 只有 75%——那是在惩罚它做对了事。
    # 分母只用 ok + bug：这两者才反映「服务自己行不行」。
    served = ok_n + bug_n
    return {
        "总调用": total,
        "区间": {"最早": rows[0].get("t"), "最近": rows[-1].get("t")},
        "结果分布": dict(status),
        "成功率": round(ok_n / served, 3) if served else None,
        "成功率口径": "ok / (ok + bug)。分母不含被拒与参数错——"
                     "那两类是调用方或前置条件的问题，不是服务能力问题",
        "拦截率": round(rej_n / total, 3) if total else None,
        "参数错误率": round(bad_n / total, 3) if total else None,
        "按工具": out,
    }


def _verdict_of_from_row(r: dict) -> str:
    """归一历史记录：早期写过 exception/error 两种写法，统一到 bug。"""
    st = r.get("s", "ok")
    return "bug" if st in ("exception", "error") else st


def _library_stats() -> dict:
    """画像库现状：状态、scope、口径版本分布。用来回答「库干不干净」。"""
    if not PROFILE_ROOT.exists():
        return {"画像数": 0}
    stats = {"画像数": 0, "按状态": {}, "按scope": {}, "按口径版本": {},
             "缺适用描述": [], "从未使用": []}
    for sub in sorted(PROFILE_ROOT.iterdir()):
        if not sub.is_dir() or sub.name.startswith(("_", ".")):
            continue
        if not (sub / "DNA.md").exists():
            continue
        meta = _load_json(sub / "meta.json", {}) or {}
        spec = _load_json(sub / "spec.json", {}) or {}
        stats["画像数"] += 1
        for key, bucket in (("status", "按状态"), ("scope", "按scope")):
            v = str(meta.get(key) or ("项目" if key == "scope" else "未知"))
            stats[bucket][v] = stats[bucket].get(v, 0) + 1
        ver = spec.get("engine_version") or "未记录"
        stats["按口径版本"][ver] = stats["按口径版本"].get(ver, 0) + 1
        if not (meta.get("适用描述") or "").strip():
            stats["缺适用描述"].append(sub.name)
        if int(meta.get("use_count", 0)) == 0:
            stats["从未使用"].append(sub.name)
    return stats


def _human_size(n: int) -> str:
    """按量级选单位。取整到整数 KB 会让 375 字节的配置显示成「0 KB」，
    看起来像文件是空的——一个真实存在的小文件不该这个待遇。"""
    if n < 1024:
        return f"{n} 字节"
    if n < 1024 * 1024:
        return f"{n / 1024:.1f} KB"
    return f"{n / 1024 / 1024:.1f} MB"


def _data_files() -> dict:
    out = {}
    for name, path in (("配置", service_config.CONFIG_PATH),
                       ("审计日志", AUDIT_PATH),
                       ("调用日志", CALL_LOG),
                       ("画像库", PROFILE_ROOT)):
        if not path.exists():
            out[name] = "不存在"
            continue
        if path.is_dir():
            n = sum(1 for _ in path.rglob("*") if _.is_file())
            size = sum(f.stat().st_size for f in path.rglob("*") if f.is_file())
            out[name] = f"{n} 个文件，{_human_size(size)}"
        else:
            out[name] = _human_size(path.stat().st_size)
    return out


@mcp.tool()
@_safe
def service_control(action: str, patch: dict | None = None,
                    confirm: bool = False) -> dict:
    """服务运维总入口。

    action：
      status  当前口径版本与四层口径自述（含本版变更、是否需重算）、
              每个引擎模块是否待重载、画像统计、生效配置
      backup  整库打成一个 JSON（patch={"out_path": "..."}），换机器/清理前留底
      reload  立刻重载引擎模块。改了 style_probe / quality_gate / cluster_gate 等
              10 个模块后用；**server.py 自身改不了**，那要重启进程
      config  读当前配置
      set     改配置（只放判定阈值）。**不带 confirm 只返回预览，不落盘**；
              confirm=True 才写入。统计口径不在这里，口径变了必须升 version 并全量重算
      stats   运行指标：调用次数、成功率、各工具耗时、画像库分布。
              要判断「这个服务靠不靠得住」看这里，而不是 audit.jsonl
              （后者记的是「谁改了什么」，不是「跑得怎么样」）
    """
    if action == "status":
        return service_status()
    if action == "stats":
        return {
            "口径版本": engine_version.ENGINE_VERSION,
            "调用": _call_stats(),
            "画像库": _library_stats(),
            "数据文件": _data_files(),
            "说明": "调用指标自 calls.jsonl 采集（滚动保留最近 5000 条）。"
                    "三档要分开看：「失败(bug)」是服务缺陷，要修；"
                    "「被拒」是前置条件不满足（如 promote 缺 confirmed_by），属正常拦截；"
                    "「参数错」是调用方传错，看文档即可。三者都返回 200，含义完全不同。"
                    "因此成功率只按 ok/(ok+bug) 算，被拒与参数错另列比率，"
                    "不要把它们读成「服务不可用」。"
                    "耗时是进程内 wall time（含文件读写，不含 MCP 传输与客户端解析）；"
                    "读几个 KB 的 JSON 在页缓存命中下就是亚毫秒级，属正常。"
                    "判断计时是否被截断看「最小/最大」：解析 1MB docx 与读写小 JSON "
                    "应差两三个数量级，区间若挤在一起才是异常",
        }
    if action == "reload":
        return reload_engine(force=True)
    if action == "config":
        return get_config()
    if action == "backup":
        if not patch or not patch.get("out_path"):
            return {"error": "action=backup 需要 patch={\"out_path\": \"...\"}"}
        return backup_profiles(patch["out_path"])
    if action == "set":
        if not patch:
            return {"error": "action=set 需要 patch，例如 {\"quality\":{\"pass\":85}}"}
        return set_config(patch, confirm=confirm)
    return {"error": f"未知 action：{action}",
            "可用": ["status", "reload", "config", "set", "backup", "stats"]}


if __name__ == "__main__":
    mcp.run("stdio")
