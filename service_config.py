#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
service_config.py — 服务可调参数

把原先散落在各模块里的硬编码阈值集中到一处，让 agent 能通过 set_config 调整。

注意：这里只放「判定阈值」，不放「口径」。口径改了就影响历史数据可比性，
必须走 version.py 升版本 + 全量重算，不能让 agent 随手改。
"""

import json
import os
from pathlib import Path

# 配置必须和画像库放在一起，不能一个在用户目录、一个在工作区。
# 之前画像库迁到 ~/.workbuddy/doc-style 了，配置还留在会话工作区——
# 结果是「清理旧工作区丢画像」修好了，「丢配置」还在。
# 环境变量 DOC_STYLE_HOME 与 server.py 保持一致。
DATA_HOME = Path(os.environ.get("DOC_STYLE_HOME",
                                Path.home() / ".workbuddy" / "doc-style")).expanduser()
CONFIG_PATH = DATA_HOME / "config.json"
PROFILE_ROOT = DATA_HOME / "profiles"      # 画像库，与 server.py 共用同一处定义

# 旧位置（模块同目录）的配置一次性搬过来
_LEGACY_CONFIG = Path(__file__).resolve().parent / "config.json"
if _LEGACY_CONFIG.exists() and not CONFIG_PATH.exists():
    try:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        import shutil as _shutil
        _shutil.move(str(_LEGACY_CONFIG), str(CONFIG_PATH))
    except OSError:
        pass    # 迁移失败就继续用旧位置，不能挡住启动

DEFAULTS = {
    "quality": {
        "pass": 80,             # >= 准入
        "weighted": 60,         # 60–79 降权准入
        "min_paragraphs": 20,   # 样本充分度下限（段落）
        "min_chars": 500,       # 样本充分度下限（字数）
    },
    "cluster": {
        "strong": 1,            # 差异 <= 此值自动合并
        "weak": 2,              # 差异 == 此值挂起待裁决；>= 3 拆画像
    },
    "feedback": {
        "rule_confirm_count": 3,   # 同类反馈累计几次才允许改画像
        "stale_unanswered": 3,     # 未答复累计几次标陈旧
    },
    "metrics": {
        "max_sentence": 26,        # 推荐单句上限（字）
    },
    "promotion": {
        "min_rules": 5,            # 提升为生效所需最少规则数（项目画像）
        "min_rules_shared": 8,     # scope=通用 的画像门槛更高：跨项目复用，影响面更大
        "block_on_stale": True,    # 口径过期时禁止提升
        "block_on_conflict": True, # 有未裁决冲突时禁止提升
    },
}

_cache = {"mtime": None, "cfg": None}


def _merge(base: dict, patch: dict) -> dict:
    out = dict(base)
    for k, v in patch.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _merge(out[k], v)
        else:
            out[k] = v
    return out


def load() -> dict:
    """读配置（带 mtime 缓存），文件不存在或损坏就回落到默认值。"""
    if not CONFIG_PATH.exists():
        return _merge(DEFAULTS, {})
    m = CONFIG_PATH.stat().st_mtime
    if _cache["mtime"] == m and _cache["cfg"] is not None:
        return _cache["cfg"]
    try:
        user = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        user = {}
    cfg = _merge(DEFAULTS, user if isinstance(user, dict) else {})
    _cache.update(mtime=m, cfg=cfg)
    return cfg


def save(cfg: dict) -> dict:
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    _cache.update(mtime=None, cfg=None)
    return load()


def patch(changes: dict) -> dict:
    return save(_merge(load(), changes))


def validate(changes: dict, current: dict | None = None) -> list:
    """校验一次改动，返回问题列表（空列表 = 通过）。

    只做「能明确判错」的三类检查：未知键、类型不符、跨字段矛盾。
    未知键必须拦：否则写进去的配置项永远不会被读取，调用方却以为改生效了——
    这比报错更难发现。
    """
    if not isinstance(changes, dict) or not changes:
        return ['patch 必须是非空对象，例如 {"quality": {"pass": 85}}']

    problems = []
    for sec, val in changes.items():
        if sec not in DEFAULTS:
            problems.append(f"未知配置段「{sec}」，可用：{'、'.join(DEFAULTS)}")
            continue
        if not isinstance(val, dict):
            problems.append(f"{sec} 必须是对象，例如 {{\"{next(iter(DEFAULTS[sec]))}\": 1}}")
            continue
        for k, v in val.items():
            if k not in DEFAULTS[sec]:
                problems.append(f"未知配置项「{sec}.{k}」，可用：{'、'.join(DEFAULTS[sec])}")
                continue
            default = DEFAULTS[sec][k]
            if isinstance(default, bool):
                if not isinstance(v, bool):
                    problems.append(f"{sec}.{k} 必须是 true / false，收到 {v!r}")
            elif isinstance(default, int) and (isinstance(v, bool) or not isinstance(v, int)):
                problems.append(f"{sec}.{k} 必须是整数，收到 {type(v).__name__}（{v!r}）")
    if problems:
        return problems

    merged = _merge(current if isinstance(current, dict) else load(), changes)
    q, c, f = merged.get("quality", {}), merged.get("cluster", {}), merged.get("feedback", {})
    p, m = merged.get("promotion", {}), merged.get("metrics", {})
    try:
        if not (0 <= q["weighted"] < q["pass"] <= 100):
            problems.append(f"quality 需满足 0 ≤ weighted({q['weighted']}) "
                            f"< pass({q['pass']}) ≤ 100")
        if q["min_paragraphs"] < 1:
            problems.append("quality.min_paragraphs 至少为 1")
        if q["min_chars"] < 100:
            problems.append("quality.min_chars 至少为 100（再低等于取消样本充分度门槛）")
        if not (0 <= c["strong"] < c["weak"]):
            problems.append(f"cluster 需满足 0 ≤ strong({c['strong']}) "
                            f"< weak({c['weak']})")
        if f["rule_confirm_count"] < 1 or f["stale_unanswered"] < 1:
            problems.append("feedback 的两个阈值都至少为 1")
        if p["min_rules"] < 1:
            problems.append("promotion.min_rules 至少为 1")
        if m["max_sentence"] < 5:
            problems.append("metrics.max_sentence 至少为 5（再低正常句子都会被判超长）")
    except TypeError as e:
        problems.append(f"配置内部类型不一致，无法比较：{e}")
    return problems
