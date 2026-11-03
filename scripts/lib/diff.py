# -*- coding: utf-8 -*-
"""差异比对：文件级 + 条款级 + 单元格级。

比对的唯一依据是「解析后内容的哈希」，不是原始文件哈希。
docx/xlsx 打开再保存会改变二进制，但内容哈希不变 —— 这正是过滤假变更的关键。
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from difflib import SequenceMatcher
from typing import Any, Dict, List, Optional, Tuple

from .core import DocModel
from .store import get_object

MAX_DETAIL_PER_FILE = 200


@dataclass
class FileChange:
    doc_key: str
    change_type: str            # added | removed | modified | renamed
    doc_no: str = ""
    title: str = ""
    old_path: str = ""
    new_path: str = ""
    level: str = "none"         # substantive | metadata | none
    meta_changes: Dict[str, Tuple[str, str]] = field(default_factory=dict)
    details: List[Dict[str, Any]] = field(default_factory=list)
    truncated: bool = False
    old_status: str = ""
    new_status: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


META_FIELDS = [
    ("version", "版本号"),
    ("effective_date", "生效日期"),
    ("owner", "责任人/归口"),
    ("title", "标题"),
    ("last_modified_by", "最后修改人"),
]


def _diff_sections(old: DocModel, new: DocModel, limit: int) -> Tuple[List[Dict[str, Any]], bool]:
    """条款级序列比对。"""
    a = [s.hash for s in old.sections]
    b = [s.hash for s in new.sections]
    sm = SequenceMatcher(a=a, b=b, autojunk=False)
    out: List[Dict[str, Any]] = []
    truncated = False

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        if tag in ("replace", "delete", "insert"):
            o_seg = old.sections[i1:i2]
            n_seg = new.sections[j1:j2]
            n = max(len(o_seg), len(n_seg))
            for k in range(n):
                if len(out) >= limit:
                    truncated = True
                    break
                o = o_seg[k] if k < len(o_seg) else None
                nn = n_seg[k] if k < len(n_seg) else None
                if o is not None and nn is not None:
                    op = "replace"
                elif o is not None:
                    op = "delete"
                else:
                    op = "insert"
                out.append({
                    "kind": "section",
                    "op": op,
                    "number": (nn.number if nn else (o.number if o else "")),
                    "old_text": (o.text if o else "")[:400],
                    "new_text": (nn.text if nn else "")[:400],
                })
        if truncated:
            break
    return out, truncated


def _diff_sheets(old: DocModel, new: DocModel, limit: int) -> Tuple[List[Dict[str, Any]], bool]:
    """Excel：工作表级 + 单元格级比对。"""
    out: List[Dict[str, Any]] = []
    truncated = False
    o_sh = {s.name: s for s in old.sheets}
    n_sh = {s.name: s for s in new.sheets}

    for name in o_sh:
        if name not in n_sh:
            out.append({"kind": "sheet", "op": "delete", "sheet": name})
    for name in n_sh:
        if name not in o_sh:
            out.append({"kind": "sheet", "op": "insert", "sheet": name})

    for name in n_sh:
        if name not in o_sh:
            continue
        o_cells = o_sh[name].cells
        n_cells = n_sh[name].cells
        keys = set(o_cells) | set(n_cells)
        for addr in sorted(keys):
            if len(out) >= limit:
                truncated = True
                break
            ov = o_cells.get(addr)
            nv = n_cells.get(addr)
            if ov == nv:
                continue
            if ov is None:
                op = "insert"
            elif nv is None:
                op = "delete"
            else:
                op = "replace"
            out.append({
                "kind": "cell", "op": op, "sheet": name, "addr": addr,
                "old_text": (ov or "")[:300], "new_text": (nv or "")[:300],
            })
        if truncated:
            break
    return out, truncated


def _diff_meta(old_e: Dict[str, Any], new_e: Dict[str, Any]) -> Dict[str, Tuple[str, str]]:
    ch: Dict[str, Tuple[str, str]] = {}
    for f, label in META_FIELDS:
        ov = str(old_e.get(f, "") or "")
        nv = str(new_e.get(f, "") or "")
        if ov != nv:
            ch[label] = (ov, nv)
    return ch


def compare(root: str, old_m: Dict[str, Any], new_m: Dict[str, Any]) -> List[FileChange]:
    """比对两份快照清单，返回变更列表。"""
    o_files: Dict[str, Any] = old_m.get("files", {})
    n_files: Dict[str, Any] = new_m.get("files", {})
    changes: List[FileChange] = []

    for key in n_files:
        if key not in o_files:
            e = n_files[key]
            changes.append(FileChange(
                doc_key=key, change_type="added", doc_no=e.get("doc_no", ""),
                title=e.get("title", ""), new_path=e.get("rel_path", ""),
                level="substantive", new_status=e.get("parse_status", ""),
            ))

    for key in o_files:
        if key not in n_files:
            e = o_files[key]
            changes.append(FileChange(
                doc_key=key, change_type="removed", doc_no=e.get("doc_no", ""),
                title=e.get("title", ""), old_path=e.get("rel_path", ""),
                level="substantive", old_status=e.get("parse_status", ""),
            ))

    for key in n_files:
        if key not in o_files:
            continue
        oe, ne = o_files[key], n_files[key]

        # 内容完全相同：打开再保存造成的假变更，直接跳过
        if oe.get("content_hash") == ne.get("content_hash"):
            if oe.get("rel_path") != ne.get("rel_path"):
                changes.append(FileChange(
                    doc_key=key, change_type="renamed", doc_no=ne.get("doc_no", ""),
                    title=ne.get("title", ""), old_path=oe.get("rel_path", ""),
                    new_path=ne.get("rel_path", ""), level="none",
                    old_status=oe.get("parse_status", ""), new_status=ne.get("parse_status", ""),
                ))
            continue

        old_doc = get_object(root, oe.get("obj", ""))
        new_doc = get_object(root, ne.get("obj", ""))
        details: List[Dict[str, Any]] = []
        truncated = False

        if old_doc and new_doc:
            if new_doc.sections or old_doc.sections:
                d1, t1 = _diff_sections(old_doc, new_doc, MAX_DETAIL_PER_FILE)
                details += d1
                truncated = truncated or t1
            if new_doc.sheets or old_doc.sheets:
                d2, t2 = _diff_sheets(old_doc, new_doc, MAX_DETAIL_PER_FILE)
                details += d2
                truncated = truncated or t2

        meta_ch = _diff_meta(oe, ne)
        if details:
            level = "substantive"
        elif meta_ch:
            level = "metadata"
        else:
            level = "none"

        # 不可解析文件：无法比对条款内容，用 大小+修改时间 兜底告警
        if oe.get("parse_status") != "ok" and ne.get("parse_status") != "ok":
            osz, nsz = oe.get("size", 0), ne.get("size", 0)
            omt, nmt = int(oe.get("mtime", 0)), int(ne.get("mtime", 0))
            if osz != nsz or omt != nmt:
                import time as _t
                fm = lambda t: _t.strftime("%Y-%m-%d %H:%M", _t.localtime(t)) if t else "-"
                meta_ch["文件指纹（不可解析，仅按大小/时间判断）"] = (
                    f"{osz} B / {fm(omt)}", f"{nsz} B / {fm(nmt)}")
                if level == "none":
                    level = "metadata"

        changes.append(FileChange(
            doc_key=key,
            change_type="renamed" if oe.get("rel_path") != ne.get("rel_path") else "modified",
            doc_no=ne.get("doc_no", ""), title=ne.get("title", ""),
            old_path=oe.get("rel_path", ""), new_path=ne.get("rel_path", ""),
            level=level, meta_changes=meta_ch, details=details, truncated=truncated,
            old_status=oe.get("parse_status", ""), new_status=ne.get("parse_status", ""),
        ))

    order = {"added": 0, "removed": 1, "modified": 2, "renamed": 3}
    changes.sort(key=lambda c: (order.get(c.change_type, 9), c.doc_key))
    return changes
