# -*- coding: utf-8 -*-
"""引用图谱构建与影响传播。

节点主键 = doc_key（编号优先），因此文件重命名不会切断关系。
影响传播 = 从变更节点沿「反向引用边」做 BFS，1 跳为直接影响，2 跳为间接影响。
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Optional, Set, Tuple

from .core import DocModel, Ref, strip_base


def _norm(s: str) -> str:
    s = (s or "").strip()
    s = s.replace("《", "").replace("》", "").replace("<", "").replace(">", "")
    s = re.sub(r"\s+", "", s)
    return s.lower()


@dataclass
class Edge:
    source: str          # 引用方 doc_key
    target: str          # 被引方 doc_key
    kind: str            # code | booktitle | filename | hierarchy
    context: str
    section_id: str
    fuzzy: bool = False  # True = 靠模糊包含匹配上的，可信度降级

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


# 看起来像「本单位体系文件」的词（用于把未命中的引用判成失效引用）
INTERNAL_HINT_RX = re.compile(
    r"(程序|手册|作业|指导|规范|制度|办法|规定|流程|标准书|计划书|报告|记录|表单|"
    r"台账|清单|一览表|控制计划|检验|SOP|WI|预案|方案|职责|管理)"
)
# 明显是外部法规/标准，未命中也不算问题
# 注意：《质量管理体系—要求》就是 ISO9001 的中文名，必须单独认出来，
# 否则每本手册都会报一堆假失效引用。
EXTERNAL_RX = re.compile(
    r"(^质量管理体系.{0,6}要求$|中华人民共和国|法律法规|ISO|IEC|GB/?T?|IATF|"
    r"VDA|AS9|JIS|DIN|ANSI|MIL|国标|国军标|条例$|[^方]法$|认证规则)"
)


def _strip_base(base: str) -> str:      # 单一口径，与 core.strip_base 一致
    return strip_base(base)


@dataclass
class Impact:
    target_key: str
    target_title: str
    target_path: str
    target_owner: str
    distance: int
    confidence: str
    reasons: List[str] = field(default_factory=list)
    evidence: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


class Graph:
    def __init__(self) -> None:
        self.nodes: Dict[str, DocModel] = {}
        self.edges: List[Edge] = []
        self._by_no: Dict[str, str] = {}
        self._by_title: Dict[str, str] = {}
        self._by_base: Dict[str, str] = {}
        self.dangling: List[Tuple[str, str, str]] = []   # (from_key, target, context)
        self.external: List[Tuple[str, str, str]] = []   # 外部法规/标准，非问题
        self.unmatched: List[Tuple[str, str, str]] = []  # 既不像内部也不像外部，待人工判
        self._seen_dangling: Set[Tuple[str, str]] = set()

    # ---------------------------------------------------------- 构建

    def add(self, doc: DocModel) -> None:
        self.nodes[doc.doc_key] = doc
        if doc.doc_no:
            self._by_no[_norm(doc.doc_no)] = doc.doc_key
        if doc.title:
            self._by_title[_norm(doc.title)] = doc.doc_key
        base = os.path.splitext(os.path.basename(doc.rel_path))[0]
        if base:
            self._by_base.setdefault(_norm(base), doc.doc_key)
            sb = _strip_base(base)
            # 正文里引用时通常不带前导序号（"详见文件化信息控制程序"），
            # 所以去序号后的写法也要登记，否则匹配不上。
            if sb and len(sb) >= 4:
                self._by_base.setdefault(_norm(sb), doc.doc_key)

    def build_edges(self, use_filename_refs: bool = True) -> None:
        for key, doc in self.nodes.items():
            for r in doc.refs:
                tgt, fuzzy = self._resolve(r.target, r.kind, key, use_filename_refs)
                if tgt is None:
                    # 解析不到：区分「本单位文件失效」与「外部法规/标准」，
                    # 不能静默丢弃——大量体系文件没有编号，书名号引用是主要建图手段。
                    self._record_unresolved(key, r.target, r.context)
                    continue
                if tgt == key:      # 自引用（台账里列出自己）不算
                    continue
                self.edges.append(Edge(source=key, target=tgt, kind=r.kind,
                                       context=r.context, section_id=r.section_id,
                                       fuzzy=fuzzy))
        self._add_hierarchy_edges()

    def _record_unresolved(self, key: str, target: str, context: str) -> None:
        t = (target or "").strip()
        if not t:
            return
        dedup = (key, _norm(t))
        if dedup in self._seen_dangling:
            return
        self._seen_dangling.add(dedup)
        if EXTERNAL_RX.search(t):
            self.external.append((key, t, context))
        elif INTERNAL_HINT_RX.search(t):
            self.dangling.append((key, t, context))
        else:
            self.unmatched.append((key, t, context))

    def _resolve(self, target: str, kind: str, from_key: str,
                 use_filename_refs: bool) -> Tuple[Optional[str], bool]:
        """返回 (目标 doc_key 或 None, 是否模糊匹配)。"""
        n = _norm(target)
        if not n:
            return None, False
        if kind == "code":
            if n in self._by_no:
                return self._by_no[n], False
            for k, v in self._by_no.items():
                if k.endswith(n) or n.endswith(k):
                    return v, True
            return None, False
        if kind == "booktitle":
            for tbl in (self._by_title, self._by_base):
                if n in tbl:
                    return tbl[n], False
            # 模糊：两边都至少 5 个字且互为子串，才算同一个文件
            for tbl in (self._by_title, self._by_base):
                for k, v in tbl.items():
                    if min(len(n), len(k)) >= 5 and (n in k or k in n):
                        return v, True
            return None, False
        if kind == "filename" and use_filename_refs:
            if n in self._by_base:
                return self._by_base[n], False
            for k, v in self._by_base.items():
                if min(len(n), len(k)) >= 5 and (n in k or k in n):
                    return v, True
            return None, False
        return None, False

    def _add_hierarchy_edges(self) -> None:
        """按编号前缀推断上下级：QP-7.5 是 QP-7.5-01 的上级。"""
        nos = [(k, d.doc_no) for k, d in self.nodes.items() if d.doc_no]
        for k1, n1 in nos:
            for k2, n2 in nos:
                if k1 == k2:
                    continue
                if n2.startswith(n1 + "-") or n2.startswith(n1 + "."):
                    self.edges.append(Edge(source=k1, target=k2, kind="hierarchy",
                                           context="由编号前缀推断的上下级关系",
                                           section_id=""))

    # ---------------------------------------------------------- 查询

    def reverse_edges(self) -> Dict[str, List[Edge]]:
        rev: Dict[str, List[Edge]] = {}
        for e in self.edges:
            rev.setdefault(e.target, []).append(e)
        return rev

    def propagate(self, changed_keys: Set[str], max_hops: int = 2) -> List[Impact]:
        """从变更集合沿反向边 BFS。"""
        rev = self.reverse_edges()
        seen: Dict[str, Impact] = {}
        frontier: Set[str] = set(changed_keys)
        visited: Set[str] = set(changed_keys)

        for hop in range(1, max_hops + 1):
            nxt: Set[str] = set()
            for cur in frontier:
                for e in rev.get(cur, []):
                    if e.source in changed_keys or e.source in visited:
                        continue
                    tgt_doc = self.nodes.get(e.source)
                    if tgt_doc is None:
                        continue
                    imp = seen.get(e.source)
                    if imp is None:
                        imp = Impact(
                            target_key=e.source,
                            target_title=tgt_doc.title or tgt_doc.rel_path,
                            target_path=tgt_doc.rel_path,
                            target_owner=tgt_doc.owner or tgt_doc.last_modified_by,
                            distance=hop,
                            confidence=_confidence(hop, e.kind, e.fuzzy),
                        )
                        seen[e.source] = imp
                    imp.distance = min(imp.distance, hop)
                    imp.confidence = _best(imp.confidence,
                                           _confidence(hop, e.kind, e.fuzzy))
                    # cur = 本次变更的文件（被引方）；e.source = 受影响文件（引用方）
                    src = self.nodes.get(cur)
                    src_name = (src.title if src else "") or (src.rel_path if src else cur)
                    ev = {
                        "changed_key": cur,
                        "changed_title": src_name,
                        "self_key": e.source,
                        "section": e.section_id,
                        "context": e.context,
                        "kind": e.kind,
                    }
                    if ev not in imp.evidence:
                        imp.evidence.append(ev)
                    reason = f"本文件引用了变更文件《{src_name}》（{_kind_label(e.kind)}）"
                    if reason not in imp.reasons:
                        imp.reasons.append(reason)
                    nxt.add(e.source)
            visited |= nxt
            frontier = nxt
            if not frontier:
                break

        out = list(seen.values())
        order = {"high": 0, "medium": 1, "low": 2}
        out.sort(key=lambda i: (i.distance, order.get(i.confidence, 9), i.target_key))
        return out

    def orphans(self) -> List[str]:
        """既未引用他人、也未被他人引用的孤立文件。"""
        used: Set[str] = set()
        for e in self.edges:
            used.add(e.source)
            used.add(e.target)
        return sorted(k for k in self.nodes if k not in used)

    def to_visual(self, changed: Set[str]) -> Dict[str, Any]:
        nodes = []
        for k, d in self.nodes.items():
            top = d.rel_path.replace("\\", "/").split("/")
            group = top[0] if len(top) > 1 else "(根目录)"
            nodes.append({
                "id": k,
                "label": d.doc_no or os.path.splitext(os.path.basename(d.rel_path))[0],
                "title": d.title or d.rel_path,
                "path": d.rel_path,
                "group": group,
                "status": d.parse_status,
                "owner": d.owner or d.last_modified_by,
                "changed": k in changed,
            })
        edges = [{
            "source": e.source, "target": e.target,
            "kind": e.kind, "context": e.context, "fuzzy": e.fuzzy,
            "hot": (e.source in changed) or (e.target in changed),
        } for e in self.edges]
        return {"nodes": nodes, "edges": edges}


def enrich_filename_refs(docs: List[DocModel], min_len: int = 6) -> int:
    """补充「文件名引用」。

    很多体系文件没有编号规则（本例即如此），正文引用时只写文件全名，
    如"详见质量管理体系策划控制程序"。此时文件名引用是建图的主要手段。
    采用一次正则扫描（长名优先），避免 O(n²) 逐名查找。
    """
    cands: Dict[str, Tuple[str, str]] = {}
    for d in docs:
        base = strip_base(os.path.splitext(os.path.basename(d.rel_path))[0])
        names = set()
        if len(base) >= min_len:
            names.add(base)
        if d.title and len(d.title.strip()) >= min_len:
            names.add(d.title.strip())
        for nm in names:
            cands.setdefault(_norm(nm), (nm, d.doc_key))

    if not cands:
        return 0

    ordered = sorted(cands.keys(), key=len, reverse=True)
    try:
        rx = re.compile("|".join(re.escape(k) for k in ordered))
    except re.error:
        return 0

    added = 0
    for d in docs:
        tn = _norm(d.all_text())
        if not tn:
            continue
        seen = set()
        for m in rx.finditer(tn):
            k = m.group(0)
            if k in seen:
                continue
            seen.add(k)
            orig, key = cands[k]
            if key == d.doc_key:
                continue
            d.refs.append(Ref(target=orig, kind="filename",
                              context=f"正文出现文件名「{orig}」", section_id=""))
            added += 1
    return added


def _confidence(hop: int, kind: str, fuzzy: bool = False) -> str:
    if fuzzy:
        return "low"                      # 模糊匹配上的关系只作提示
    if hop == 1:
        return "high" if kind in ("code", "hierarchy") else "medium"
    return "low"


def _best(a: str, b: str) -> str:
    order = {"high": 0, "medium": 1, "low": 2}
    return a if order.get(a, 9) <= order.get(b, 9) else b


def _kind_label(kind: str) -> str:
    return {
        "code": "编号引用",
        "booktitle": "书名号引用",
        "filename": "文件名引用",
        "hierarchy": "上下级关系",
    }.get(kind, kind)
