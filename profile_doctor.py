#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
profile_doctor.py — 画像库体检

画像库最危险的失效方式不是文件乱，而是**悄悄过期而没人知道**：
文件完整、格式漂亮、引用通畅，但里面的数字是上一个口径下算的。

八项检查：
  1 完整性       DNA / rules / spec / meta 是否齐备
  2 状态匹配     meta 状态与文件是否自洽（active 却缺规则 → 异常）
  3 证据可解析   rules 里每条 spec.xxx 是否能在 spec.json 真实解析到
  4 数字一致性   DNA.md 里写的百分比与 spec.json 实测值是否矛盾
  5 口径版本     spec 记录的口径版本与当前引擎是否一致
  6 源文档       spec 是否记录了源路径与时间戳，能否判断源文档被动过
  7 规则厚度     规则数是否足以支撑复用（<5 条偏薄）
  8 重复画像     同一文档类型是否存在多个生效画像（可能拆错，该合并）

用法：
  python3 profile_doctor.py                # 体检全部
  python3 profile_doctor.py --json out.json
  python3 profile_doctor.py --fix          # 只做安全动作：把该标陈旧的标成陈旧
"""

import re
import sys
import json
import argparse
from pathlib import Path

sys.path.insert(0, ".")
import version  # noqa: E402

ROOT = Path(__file__).resolve().parent
# 路径只能在 service_config 里定义一处。曾经三处各写一份
# （server.py / profile_doctor.py / 配置），迁库时漏掉体检脚本，
# 结果它报「画像库不存在」——数据在那儿，只是它看错地方。
import service_config
PROFILE_ROOT = service_config.PROFILE_ROOT

REQUIRED_FILES = {"DNA.md": "画像正文", "rules.json": "规则清单",
                  "spec.json": "硬指标", "meta.json": "元信息"}

# DNA 里会写成的百分比指标 → spec.voice 里对应的字段
NUMERIC_METRICS = {
    "硬性要求占比": "硬性要求占比",
    "技术词占比": "技术词占比",
    "列表化段落占比": "列表化段落占比",
}


def _resolve_strict(spec, path: str):
    """返回 (是否完整解析, 段数, 值)。支持中文键与含空格的键（Heading 1）。"""
    if not isinstance(spec, dict) or not path.startswith("spec."):
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


def _evidence_path(spec, evidence: str) -> bool:
    """从长到短试前缀。至少 2 段，或单段指向非字典节点（列表/标量）。"""
    if not evidence.startswith("spec."):
        return True
    limit = min(len(evidence), 120)
    for end in range(limit, 4, -1):
        cand = evidence[:end].rstrip("：:，。、,（）() \t")
        ok, seg, val = _resolve_strict(spec, cand)
        if ok and (seg >= 2 or not isinstance(val, dict)):
            return True
    return False


def _pct_in_dna(dna: str, label: str):
    """从 DNA.md 里抠出「标签 …… 数字%」的数字"""
    m = re.search(re.escape(label) + r"[^\n]{0,40}?(\d+(?:\.\d+)?)\s*%", dna)
    return float(m.group(1)) / 100 if m else None


def inspect(pid: str):
    d = PROFILE_ROOT / pid
    res = {"profile_id": pid, "checks": [], "rebuild": False, "actions": []}

    def add(name, ok, detail, level="高"):
        res["checks"].append({"检查": name, "结果": "通过" if ok else "问题",
                              "级别": level, "说明": detail})
        return ok

    # 1 完整性
    missing = [f for f in REQUIRED_FILES if not (d / f).exists()]
    add("完整性", not missing,
        "文件齐备" if not missing else
        "缺少：" + "、".join(f"{f}（{REQUIRED_FILES[f]}）" for f in missing))

    meta = {}
    if (d / "meta.json").exists():
        try:
            meta = json.loads((d / "meta.json").read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            add("完整性", False, "meta.json 不是合法 JSON")
    rules, spec, dna = [], {}, ""
    if (d / "rules.json").exists():
        try:
            rules = json.loads((d / "rules.json").read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            add("完整性", False, "rules.json 不是合法 JSON")
    if (d / "spec.json").exists():
        try:
            spec = json.loads((d / "spec.json").read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            add("完整性", False, "spec.json 不是合法 JSON")
    if (d / "DNA.md").exists():
        dna = (d / "DNA.md").read_text(encoding="utf-8")

    status = meta.get("status", "未知")

    # 2 状态匹配
    if status.startswith("生效") or status == "active":
        ok = bool(rules) and bool(spec)
        add("状态匹配", ok, "生效画像规则与硬指标齐备" if ok else
            "标记为生效，但缺规则或硬指标，不能用")
    else:
        add("状态匹配", True, f"状态「{status}」，无需生效校验", "低")

    # 3 证据可解析
    if rules:
        bad = []
        for r in rules:
            ev = (r.get("evidence") or "").strip()
            if ev.startswith("spec.") and not _evidence_path(spec, ev):
                bad.append(f"{r.get('id', '?')}→{ev[:40]}")
        add("证据可解析", not bad,
            f"{len(rules)} 条规则全部命中" if not bad else
            f"{len(bad)} 条证据指向不存在的字段：" + "、".join(bad[:5]))
    else:
        add("证据可解析", False, "无规则可校验")

    # 4 数字一致性
    if spec.get("voice") and dna:
        conflicts = []
        for label, key in NUMERIC_METRICS.items():
            in_dna = _pct_in_dna(dna, label)
            in_spec = spec["voice"].get(key)
            if in_dna is None or in_spec is None:
                continue
            if abs(in_dna - in_spec) > 0.02:
                conflicts.append(f"「{label}」DNA 写 {in_dna:.1%}，spec 实测 {in_spec:.1%}")
        add("数字一致性", not conflicts,
            "DNA 与 spec 的数字一致" if not conflicts else "；".join(conflicts))
        if conflicts:
            res["rebuild"] = True
            res["actions"].append("按 spec.json 重写 DNA 中引用的数字，或整份重建")
    else:
        add("数字一致性", False, "缺 spec.json 或 voice 字段，无法比对 DNA 里的数字", "中")

    # 5 口径版本
    need, why = version.needs_rebuild(spec.get("engine_version"))
    add("口径版本", not need, why)
    if need:
        res["rebuild"] = True
        res["actions"].append("按当前口径重算并重建画像")

    # 6 源文档
    src_path = spec.get("source_path")
    if not src_path:
        add("源文档", False, "spec 未记录源路径与时间戳，无法判断源文档是否被改动过", "中")
        res["actions"].append("保存画像时应记录 source_path / source_mtime / source_size")
    elif not Path(src_path).exists():
        add("源文档", False, f"源文档已不存在：{src_path}")
        res["rebuild"] = True
    else:
        mtime = Path(src_path).stat().st_mtime
        saved = spec.get("source_mtime")
        if saved and mtime > saved + 1:
            add("源文档", False, "源文档在画像生成后被修改过，硬指标可能已失效")
            res["rebuild"] = True
        else:
            add("源文档", True, "源文档存在且未变更")

    # 7 规则厚度
    n = len(rules)
    add("规则厚度", n >= 5, f"{n} 条规则" + ("" if n >= 5 else "，偏薄，复用价值有限"), "中")
    if n < 5:
        res["actions"].append("补充样本或补充内容套路层，把规则补到 5 条以上")

    # 8 文档类型元信息（重复检测与按类型检索的前提）
    doc_type = (meta.get("文档类型") or meta.get("doc_type") or "").strip()
    if doc_type:
        add("元信息", True, f"文档类型：{doc_type}", "中")
    else:
        add("元信息", False,
            "meta 缺「文档类型」，重复画像检测与按类型检索会失效", "中")
        res["actions"].append("用 manage_profile(action=\"update_meta\") 补 文档类型")

    # 9 重复（同类型多生效画像在 main 里统一判）
    res["status"] = status
    res["doc_type"] = doc_type  # 缺失就留空，不要拿 pid 冒充类型
    res["rule_count"] = n
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", dest="json_out")
    ap.add_argument("--fix", action="store_true", help="只做安全动作：把该标陈旧的标成陈旧")
    args = ap.parse_args()

    if not PROFILE_ROOT.exists():
        print("画像库不存在")
        return 1

    pids = [p.name for p in sorted(PROFILE_ROOT.iterdir())
            if p.is_dir() and not p.name.startswith("_")]
    if not pids:
        print("画像库为空")
        return 0

    results = [inspect(p) for p in pids]

    # 8 重复画像：同文档类型多个非过期画像
    by_type = {}
    untagged = []
    for r in results:
        if "过期" in r["status"] or "废" in r["status"]:
            continue
        if not r.get("doc_type"):
            untagged.append(r["profile_id"])
            continue
        by_type.setdefault(r["doc_type"], []).append(r["profile_id"])
    dup = {k: v for k, v in by_type.items() if len(v) > 1}

    print("=" * 96)
    print(f"画像库体检  口径版本 {version.ENGINE_VERSION}")
    print(f"根目录 {PROFILE_ROOT}")
    print("=" * 96)

    for r in results:
        mark = "❌" if any(c["级别"] == "高" and c["结果"] == "问题" for c in r["checks"]) else "✅"
        print(f"\n{mark} {r['profile_id']}   状态 {r['status']}   规则 {r['rule_count']} 条")
        for c in r["checks"]:
            icon = "✅" if c["结果"] == "通过" else ("❌" if c["级别"] == "高" else "⚠️")
            print(f"   {icon} {c['检查']:6s} {c['说明']}")

    if dup:
        print("\n" + "=" * 96)
        print("重复画像（同文档类型有多个可用画像，可能拆错了）")
        print("=" * 96)
        for t, v in dup.items():
            print(f"  {t}：{'、'.join(v)}  → 建议先跑 cluster_documents 确认是否真该拆开")

    if untagged:
        print("\n" + "=" * 96)
        print("⚠️ 未标文档类型的画像（这些画像不参与重复检测）")
        print("=" * 96)
        print(f"  {'、'.join(untagged)}")
        print("  → 用 manage_profile(action=\"update_meta\") 补「文档类型」，否则同类画像共存也发现不了")

    need = [r["profile_id"] for r in results if r["rebuild"]]
    print("\n" + "=" * 96)
    print(f"汇总：{len(results)} 个画像，{len(need)} 个需要重算" + (f"（{'、'.join(need)}）" if need else ""))
    print("=" * 96)

    todo = [(r["profile_id"], a) for r in results for a in r["actions"]]
    if todo:
        print("\n建议动作")
        for pid, a in todo:
            print(f"  - {pid}：{a}")

    if args.fix:
        print("\n执行安全修复（只改状态标记，不动规则与数字）")
        for r in results:
            if not r["rebuild"]:
                continue
            mf = PROFILE_ROOT / r["profile_id"] / "meta.json"
            if not mf.exists():
                continue
            m = json.loads(mf.read_text(encoding="utf-8"))
            if "过期" in m.get("status", ""):
                print(f"  {r['profile_id']}：已是过期状态，跳过")
                continue
            m["status"] = "候选（数据过期，待重建）"
            m["stale_reason"] = "口径版本变更或数字不一致，体检发现"
            mf.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"  {r['profile_id']}：已标记为过期")

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as fh:
            json.dump({"engine_version": version.ENGINE_VERSION,
                       "profiles": results, "duplicates": dup},
                      fh, ensure_ascii=False, indent=2)
        print(f"\n[JSON] {args.json_out}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
