# -*- coding: utf-8 -*-
"""核心：配置管理、文本归一化、数据模型。"""
from __future__ import annotations

import hashlib
import json
import os
import re
from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Optional

INDEX_DIR = ".cb-index"

DEFAULT_CONFIG: Dict[str, Any] = {
    "workspace_name": "",
    # 文件编号正则；为空表示首次运行时自动推断
    "doc_no_patterns": [],
    # 条款号正则（用于条款级定位，不作为文件主键）
    "clause_pattern": r"\d{1,2}(?:\.\d{1,2}){0,3}",
    "version_keys": ["版本号", "版本", "修订号", "版次", "Rev", "Revision"],
    "effective_date_keys": ["生效日期", "实施日期", "发布日期", "生效日"],
    "owner_keys": ["编制", "责任人", "归口部门", "编写"],
    "ignore_globs": ["~$*", "*.tmp", "*.bak", "*.old", "*.lnk", "*.ini", "Thumbs.db"],
    "ignore_dirs": [
        "旧版本", "作废", "作废文件", "历史版本", "备份", "backup",
        "_archive", ".cb-index", "System Volume Information",
    ],
    "max_files_warning": 2000,
    "history_keep": 24,
    "booktitle_refs": True,
    "filename_refs": True,
    "diff_min_section_len": 2,
}

# 编号自动推断的候选模式（按精确度从高到低排列；命中数相同时取靠前者）
CANDIDATE_DOC_NO_PATTERNS = [
    r"[A-Z]{2,4}[-_/]\d{2}[-_/]\d{2,3}",                                   # QP-03-012
    r"[A-Z]{2,4}[-_/][0-9]{1,2}(?:\.[0-9]{1,2})*(?:[-_/][0-9]{1,3})?",      # QP-7.5-01 / QP-7.5 / QM-01
    r"[A-Z]{2,4}[-_/]\d{3,4}",                                              # QP-0123 / QM-001
    r"[A-Z]{1,3}/[A-Z]{1,3}[-_/]\d{2,3}",                                   # QM/QP-01
    r"[A-Z]{2,}-\d{3,}",                                                    # ABC-001
    r"[A-Z]{2,4}\d{4,6}",                                                   # QP0001
]


# ---------------------------------------------------------------- 文本处理

_ZERO_WIDTH = re.compile(r"[\u200b-\u200f\ufeff]")
_WS = re.compile(r"\s+")


def norm_text(s: str) -> str:
    """归一化文本：去零宽字符、全角空格、合并连续空白。不做标点转换，避免误判。"""
    if not s:
        return ""
    s = _ZERO_WIDTH.sub("", s)
    s = s.replace("\u3000", " ").replace("\xa0", " ")
    return _WS.sub(" ", s).strip()


def sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()


def short(h: str, n: int = 12) -> str:
    return h[:n]


# ---------------------------------------------------------------- 数据模型

@dataclass
class Section:
    """条款/段落级单元。"""
    sid: str
    level: int
    number: str
    text: str
    hash: str

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class SheetData:
    """Excel 工作表。cells 为 addr -> 归一化文本（空值不存）。"""
    name: str
    dims: str
    cells: Dict[str, str] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {"name": self.name, "dims": self.dims, "cells": self.cells}


@dataclass
class Ref:
    """一处引用。"""
    target: str          # 引用目标：编号或文件名
    kind: str            # code | booktitle | filename
    context: str         # 上下文片段
    section_id: str      # 所在条款 sid（Excel 为 sheet 名）

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class DocModel:
    """统一文档模型。"""
    doc_key: str                 # 主键：优先编号，无编号退化为规范化相对路径
    rel_path: str
    ext: str
    size: int = 0
    mtime: float = 0.0
    parse_status: str = "failed"  # ok | partial | failed
    parse_note: str = ""
    title: str = ""
    doc_no: str = ""
    version: str = ""
    effective_date: str = ""
    last_modified_by: str = ""
    owner: str = ""
    sections: List[Section] = field(default_factory=list)
    sheets: List[SheetData] = field(default_factory=list)
    refs: List[Ref] = field(default_factory=list)
    content_hash: str = ""

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["sections"] = [s.to_dict() for s in self.sections]
        d["sheets"] = [s.to_dict() for s in self.sheets]
        d["refs"] = [r.to_dict() for r in self.refs]
        return d

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "DocModel":
        d = dict(d)
        d["sections"] = [Section(**s) for s in d.get("sections", [])]
        d["sheets"] = [SheetData(**s) for s in d.get("sheets", [])]
        d["refs"] = [Ref(**r) for r in d.get("refs", [])]
        known = {f for f in DocModel.__dataclass_fields__}  # type: ignore[attr-defined]
        return DocModel(**{k: v for k, v in d.items() if k in known})

    def all_text(self) -> str:
        parts = [self.title]
        parts += [f"{s.number} {s.text}" for s in self.sections]
        for sh in self.sheets:
            parts.append(sh.name)
            parts += [v for v in sh.cells.values() if v]
        return "\n".join(p for p in parts if p)


# ---------------------------------------------------------------- 配置

def load_config(root: str) -> Dict[str, Any]:
    """读取配置，缺失则返回默认配置（不写盘）。"""
    path = os.path.join(root, INDEX_DIR, "config.json")
    cfg = dict(DEFAULT_CONFIG)
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg


def save_config(root: str, cfg: Dict[str, Any]) -> str:
    d = os.path.join(root, INDEX_DIR)
    os.makedirs(d, exist_ok=True)
    path = os.path.join(d, "config.json")
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)
    return path


def infer_doc_no_patterns(samples: List[str]) -> List[str]:
    """从文件名与正文样本中统计推断编号正则，返回按命中数排序的候选。"""
    scored = []
    for idx, pat in enumerate(CANDIDATE_DOC_NO_PATTERNS):
        rx = re.compile(pat)
        hits = set()
        for s in samples:
            for m in rx.findall(s):
                hits.add(m if isinstance(m, str) else "".join(m))
        if len(hits) >= 3:
            # 命中数降序；命中数相同时取更精确的（候选列表中更靠前的）
            scored.append((-len(hits), idx, pat))
    scored.sort()
    return [p for _, _, p in scored[:3]]


def match_doc_no(text: str, patterns: List[str]) -> Optional[str]:
    for pat in patterns:
        try:
            m = re.search(pat, text)
        except re.error:
            continue
        if m:
            return m.group(0)
    return None


def is_ignored(rel_path: str, cfg: Dict[str, Any]) -> bool:
    """判断相对路径是否应被忽略。"""
    import fnmatch

    parts = rel_path.replace("\\", "/").split("/")
    ignore_dirs = {d.lower() for d in cfg.get("ignore_dirs", [])}
    for p in parts[:-1]:
        if p.lower() in ignore_dirs or p.startswith("~"):
            return True
    name = parts[-1]
    for g in cfg.get("ignore_globs", []):
        if fnmatch.fnmatch(name.lower(), g.lower()):
            return True
    return False


def safe_key(rel_path: str) -> str:
    """无编号时，用规范化相对路径作为主键（去掉扩展名）。"""
    p = rel_path.replace("\\", "/")
    base = os.path.splitext(p)[0]
    return "path:" + base


def strip_base(base: str) -> str:
    """文件名去掉前导序号（"0xxx"、"1.2 xxx"）与「附件N」前缀，并去掉尾部日期。

    解析正文引用时，人们通常只写文件真名（"详见文件化信息控制程序"），
    不会带前导序号，所以建索引时必须同时登记这两种写法。
    """
    b = re.sub(r"^[\d\s.\-—]+", "", base)
    b = re.sub(r"^附件[一二三四五六七八九十\d]*\s*[:：]?\s*", "", b).strip()
    b = re.sub(r"[-_ ]?(19|20)\d{2}[-_./]?\d{0,2}[-_./]?\d{0,2}\s*$", "", b).strip()
    b = re.sub(r"[-_ ]?(ok|OK|Ok|终|定稿|final)\s*$", "", b).strip()
    return b
