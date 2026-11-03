# -*- coding: utf-8 -*-
"""快照存储：内容寻址对象库 + 快照清单。

不使用 Git，自研极简版本库：
- objects/ 按内容哈希存 DocModel，相同内容只存一份（天然去重）
- snapshots/ 只存「doc_key -> object_hash」的清单，体积极小
- 主键是文件编号（无编号时退化为路径），因此文件重命名/移动不会丢失历史
"""
from __future__ import annotations

import json
import os
import shutil
from typing import Any, Dict, List, Optional

from .core import DocModel, INDEX_DIR


def index_dir(root: str) -> str:
    return os.path.join(root, INDEX_DIR)


def ensure_layout(root: str) -> str:
    d = index_dir(root)
    for sub in ("objects", "snapshots", "reports", "tmp"):
        os.makedirs(os.path.join(d, sub), exist_ok=True)
    return d


def _atomic_write_json(path: str, data: Any) -> None:
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


# ---------------------------------------------------------------- 对象库

def _obj_path(root: str, h: str) -> str:
    return os.path.join(index_dir(root), "objects", h[:2], h + ".json")


def put_object(root: str, doc: DocModel) -> str:
    """写入对象，返回内容哈希（即 object hash）。"""
    h = doc.content_hash
    p = _obj_path(root, h)
    if os.path.isfile(p):
        return h
    os.makedirs(os.path.dirname(p), exist_ok=True)
    _atomic_write_json(p, doc.to_dict())
    return h


def get_object(root: str, h: str) -> Optional[DocModel]:
    p = _obj_path(root, h)
    if not os.path.isfile(p):
        return None
    try:
        with open(p, "r", encoding="utf-8") as f:
            return DocModel.from_dict(json.load(f))
    except Exception:
        return None


# ---------------------------------------------------------------- 快照

def _snap_path(root: str, ts: str) -> str:
    return os.path.join(index_dir(root), "snapshots", ts + ".json")


def latest_pointer(root: str) -> str:
    return os.path.join(index_dir(root), "snapshots", "latest.json")


def read_latest(root: str) -> Optional[Dict[str, Any]]:
    lp = latest_pointer(root)
    if not os.path.isfile(lp):
        return None
    try:
        with open(lp, "r", encoding="utf-8") as f:
            ptr = json.load(f)
    except Exception:
        return None
    ts = ptr.get("ts")
    if not ts:
        return None
    p = _snap_path(root, ts)
    if not os.path.isfile(p):
        return None
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)


def write_snapshot(root: str, manifest: Dict[str, Any]) -> str:
    """写入快照并更新 latest 指针（先写快照再移指针，保证原子性）。"""
    ensure_layout(root)
    ts = manifest["ts"]
    _atomic_write_json(_snap_path(root, ts), manifest)
    _atomic_write_json(latest_pointer(root), {"ts": ts, "label": manifest.get("label", "")})
    return ts


def list_snapshots(root: str) -> List[str]:
    d = os.path.join(index_dir(root), "snapshots")
    if not os.path.isdir(d):
        return []
    out = []
    for n in os.listdir(d):
        if n.endswith(".json") and n != "latest.json":
            out.append(n[:-5])
    return sorted(out)


def read_snapshot(root: str, ts: str) -> Optional[Dict[str, Any]]:
    p = _snap_path(root, ts)
    if not os.path.isfile(p):
        return None
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)


def prune_snapshots(root: str, keep: int) -> int:
    """保留最近 keep 份快照，返回删除数量。objects 不动（内容寻址安全）。"""
    snaps = list_snapshots(root)
    if len(snaps) <= keep:
        return 0
    latest = None
    lp = latest_pointer(root)
    if os.path.isfile(lp):
        try:
            with open(lp, "r", encoding="utf-8") as f:
                latest = json.load(f).get("ts")
        except Exception:
            latest = None
    doomed = snaps[:-keep]
    n = 0
    for ts in doomed:
        if ts == latest:
            continue
        try:
            os.remove(_snap_path(root, ts))
            n += 1
        except Exception:
            pass
    return n


# ---------------------------------------------------------------- 决策记录

def decisions_path(root: str) -> str:
    return os.path.join(index_dir(root), "decisions.json")


def load_decisions(root: str) -> Dict[str, Any]:
    p = decisions_path(root)
    if not os.path.isfile(p):
        return {"dismissed": {}}
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"dismissed": {}}


def save_decisions(root: str, data: Dict[str, Any]) -> None:
    ensure_layout(root)
    _atomic_write_json(decisions_path(root), data)


def status_path(root: str) -> str:
    return os.path.join(index_dir(root), "STATUS.md")


def reset_workspace(root: str, keep_converted: bool = True) -> None:
    """清空索引（不影响原始文件）。

    keep_converted=True 时保留 converted/ —— 那一批 .doc→.docx 副本很贵
    （上百个文件要逐个启动 Office 转换），不能因为重建基线就被连带删掉。
    """
    d = index_dir(root)
    if not os.path.isdir(d):
        return
    keep = {"converted"} if keep_converted else set()
    for name in os.listdir(d):
        if name in keep:
            continue
        p = os.path.join(d, name)
        if os.path.isdir(p):
            shutil.rmtree(p, ignore_errors=True)
        else:
            try:
                os.remove(p)
            except OSError:
                pass
