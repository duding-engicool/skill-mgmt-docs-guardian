# -*- coding: utf-8 -*-
"""主入口：预检 / 扫描 / 比对 / 提交快照 / 出报告。

用法：
  python scan.py --path "D:\\体系文件" --preflight
  python scan.py --path "D:\\体系文件"
  python scan.py --path "D:\\体系文件" --label "2026年10月换版"
  python scan.py --path "D:\\体系文件" --dismiss "no:QP-001" --reason "已人工确认无需修改"
  python scan.py --path "D:\\体系文件" --reset

铁律：目标文件夹内除 .cb-index/ 外的文件全程只读。
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime
from typing import Any, Dict, List, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from lib.core import (  # noqa: E402
    DocModel, INDEX_DIR, infer_doc_no_patterns, is_ignored,
    load_config, save_config,
)
from lib import diff as D  # noqa: E402
from lib.graph import Graph, enrich_filename_refs  # noqa: E402
from lib import parser as P  # noqa: E402
from lib import report as R  # noqa: E402
from lib import store as S  # noqa: E402

from graph_html import render_graph  # noqa: E402


def collect_files(root: str, cfg: Dict[str, Any]) -> List[Tuple[str, str]]:
    """遍历目录，返回 (绝对路径, 相对路径)，已应用忽略规则。"""
    ignore_dirs = {d.lower() for d in cfg.get("ignore_dirs", [])}
    out: List[Tuple[str, str]] = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(
            d for d in dirnames
            if d.lower() not in ignore_dirs and not d.startswith("~")
        )
        for fn in filenames:
            abs_p = os.path.join(dirpath, fn)
            rel = os.path.relpath(abs_p, root)
            if is_ignored(rel, cfg):
                continue
            out.append((abs_p, rel))
    out.sort(key=lambda t: t[1])
    return out


def entry_of(doc: DocModel, obj_hash: str) -> Dict[str, Any]:
    return {
        "obj": obj_hash,
        "content_hash": doc.content_hash,
        "rel_path": doc.rel_path,
        "ext": doc.ext,
        "title": doc.title,
        "doc_no": doc.doc_no,
        "version": doc.version,
        "effective_date": doc.effective_date,
        "last_modified_by": doc.last_modified_by,
        "owner": doc.owner,
        "parse_status": doc.parse_status,
        "parse_note": doc.parse_note,
        "size": doc.size,
        "mtime": doc.mtime,
    }


def resolve_key_collisions(docs: List[DocModel]) -> List[Tuple[str, str]]:
    """同一编号出现在多份文件时，给主键追加路径后缀，避免互相覆盖。

    注意：只要有重复，所有重复项都加后缀（而不是只改后来的），保证主键在
    多次扫描之间保持稳定，否则删除其中一份会导致其余文件的 key 漂移。
    """
    from collections import Counter

    cnt = Counter(d.doc_key for d in docs)
    dup_keys = {k for k, v in cnt.items() if v > 1}
    dups: List[Tuple[str, str]] = []
    for d in docs:
        if d.doc_key in dup_keys:
            dups.append((d.doc_no, d.rel_path))
            d.doc_key = f"{d.doc_key}|{d.rel_path}"
    return dups


def parse_all(files: List[Tuple[str, str]], cfg: Dict[str, Any],
              quiet: bool = False, conv_dir: str = "") -> List[DocModel]:
    """解析全部文件。

    conv_dir 存在时，.doc 会改用其中的 .docx 副本解析——但 doc 的身份
    （rel_path / doc_key / 台账记录）仍然是原 .doc，原始文件不受影响。
    """
    docs: List[DocModel] = []
    n = len(files)
    for i, (abs_p, rel) in enumerate(files, 1):
        if not quiet and (i % 50 == 0 or i == n):
            print(f"  解析中 {i}/{n}", flush=True)
        target = abs_p
        if conv_dir and os.path.splitext(abs_p)[1].lower() == ".doc":
            cand = os.path.join(conv_dir, os.path.splitext(rel)[0] + ".docx")
            if os.path.isfile(cand):
                target = cand
        docs.append(P.parse_file(target, rel, cfg))
    return docs


def stale_copies(files: List[Tuple[str, str]], conv_dir: str) -> List[str]:
    """找出「原 .doc 已被修改、但转换副本还是旧的」的文件。

    这是最危险的静默失败：用户明明改了文件，工具却因为读旧副本而报 0 变更。
    """
    if not conv_dir:
        return []
    out = []
    for abs_p, rel in files:
        if os.path.splitext(abs_p)[1].lower() != ".doc":
            continue
        cand = os.path.join(conv_dir, os.path.splitext(rel)[0] + ".docx")
        if not os.path.isfile(cand):
            continue
        try:
            if int(os.path.getmtime(abs_p)) > int(os.path.getmtime(cand)) + 1:
                out.append(rel)
        except OSError:
            pass
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="管理体系文件一致性追溯技能")
    ap.add_argument("--path", required=True, help="目标文件夹绝对路径")
    ap.add_argument("--preflight", action="store_true", help="只做预检，不写任何文件")
    ap.add_argument("--label", default="", help="本次快照标签")
    ap.add_argument("--hops", type=int, default=2, help="影响传播跳数，默认 2")
    ap.add_argument("--dismiss", default="", help="标记某影响项为已忽略（传 doc_key）")
    ap.add_argument("--reason", default="", help="忽略原因")
    ap.add_argument("--reset", action="store_true", help="删除索引目录重建（不影响原文件）")
    ap.add_argument("--baseline", action="store_true",
                    help="以当前状态重设基线，不与上次快照比对（不删任何文件；"
                         "适合解析能力刚提升、旧快照已不可比的情况）")
    ap.add_argument("--no-graph", action="store_true", help="不生成关系图")
    ap.add_argument("--converted", default="",
                    help="转换副本目录（默认自动探测 <目标>/.cb-index/converted）")
    args = ap.parse_args()

    root = os.path.abspath(args.path)
    if not os.path.isdir(root):
        print(f"[错误] 路径不存在或不是文件夹：{root}", file=sys.stderr)
        return 2

    if args.reset:
        S.reset_workspace(root)
        print(f"[完成] 已删除索引目录 {os.path.join(root, INDEX_DIR)}，原始文件未受影响。")
        print("      下次运行将重新建立基线。")
        return 0

    cfg = load_config(root)
    files = collect_files(root, cfg)

    # .doc 转换副本目录：存在则用它解析内容（原文件仍只读）
    conv_dir = args.converted or os.path.join(root, INDEX_DIR, "converted")
    conv_dir = conv_dir if os.path.isdir(conv_dir) else ""

    # ---- 副本新鲜度告警（必须在解析之前，否则会静默漏检）----
    stale = stale_copies(files, conv_dir)
    if not args.preflight and stale:
        print("=" * 60)
        print(f"⚠️  有 {len(stale)} 份 .doc 原文件已被修改，但转换副本还是旧的")
        print("=" * 60)
        print("继续扫描会读旧副本 → 这些文件的改动检测不到。请先刷新副本：")
        print()
        print(f'  python convert_docs.py --path "{root}"')
        print()
        for r in stale[:10]:
            print("  - " + r)
        if len(stale) > 10:
            print(f"  …另有 {len(stale)-10} 份")
        print()

    # ---- 预检 ----
    if args.preflight:
        total = len(files)
        warn = cfg.get("max_files_warning", 2000)
        exts: Dict[str, int] = {}
        for _, rel in files:
            e = os.path.splitext(rel)[1].lower() or "(无扩展名)"
            exts[e] = exts.get(e, 0) + 1
        supportable = sum(v for k, v in exts.items()
                          if k in P.HANDLERS or k in (".txt", ".md", ".csv"))
        # 已有转换副本的 .doc 实际可解析，必须计入，否则预检会虚报盲区
        n_conv = 0
        if conv_dir:
            n_conv = sum(1 for _, r in files
                         if os.path.splitext(r)[1].lower() == ".doc"
                         and os.path.isfile(os.path.join(
                             conv_dir, os.path.splitext(r)[0] + ".docx")))
        n_stale = len(stale_copies(files, conv_dir))
        print("=" * 60)
        print("路径预检")
        print("=" * 60)
        print(f"目标目录 : {root}")
        print(f"文件总数 : {total}")
        if n_conv:
            print(f"可解析   : {supportable + n_conv}（原生 {supportable} + 转换副本 {n_conv}）")
        else:
            print(f"可解析   : {supportable}（格式受支持）")
        print(f"不可解析 : {total - supportable - n_conv}")
        if n_stale:
            print(f"⚠️ 其中 {n_stale} 份 .doc 原文件已改动、副本过期，需先重转：")
            print('   python convert_docs.py --path "<本目录>"')
        print()
        print("格式分布：")
        for e, c in sorted(exts.items(), key=lambda x: -x[1])[:15]:
            note = ""
            if e in P.NEEDS_CONVERT:
                note = f"　← {P.NEEDS_CONVERT[e]}"
            print(f"  {e:<12}{c:>5}{note}")
        print()
        n_doc = sum(1 for _, r in files if os.path.splitext(r)[1].lower() == ".doc")
        if n_doc:
            print(f"老版 .doc：{n_doc} 个（纯 Python 无法直接解析）")
            if conv_dir and n_conv:
                lack = n_doc - n_conv
                print(f"  已转换副本：{n_conv} 个（convert_docs.py 产出，扫描时自动使用）")
                if lack:
                    print(f"  尚缺副本：{lack} 个，建议：python convert_docs.py --path \"<本目录>\"")
            else:
                print("  建议：python convert_docs.py --path \"<本目录>\"   （生成副本，原文件不动）")
            print()

        deps = P.dependency_report()
        missing = [c for _, _, ok, c in deps if c]
        if missing:
            print("依赖缺失（对应格式将标记为盲区）：")
            for m in missing:
                print(f"  {m}")
            print()
        if total == 0:
            print("[警告] 该目录下未发现任何文件，请确认路径是否正确。")
            return 3
        if total > warn:
            print(f"[警告] 文件数 {total} 超过阈值 {warn}，请确认这是体系文件目录而不是更大的范围。")
        print("确认无误后，去掉 --preflight 执行全量扫描。")
        return 0

    # ---- 忽略标注 ----
    if args.dismiss:
        dec = S.load_decisions(root)
        dec.setdefault("dismissed", {})[args.dismiss] = {
            "reason": args.reason,
            "at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        S.save_decisions(root, dec)
        print(f"[完成] 已标记 {args.dismiss} 为已忽略，后续扫描不再进入待办清单。")
        return 0

    if not files:
        print("[错误] 目录下未发现可纳管的文件。", file=sys.stderr)
        return 3

    print("=" * 60)
    print(f"管理体系文件一致性追溯技能 · {root}")
    print("=" * 60)

    fresh = not os.path.isdir(os.path.join(root, INDEX_DIR))
    S.ensure_layout(root)

    # ---- 第一轮解析 ----
    print(f"[1/6] 解析 {len(files)} 个文件…")
    if conv_dir:
        print(f"      使用转换副本目录：{conv_dir}")
    docs = parse_all(files, cfg, conv_dir=conv_dir)

    # ---- 编号规则自动推断（仅首次） ----
    if fresh and not cfg.get("doc_no_patterns"):
        samples = [os.path.basename(a) for a, _ in files]
        samples += [d.title for d in docs if d.title]
        for d in docs[:80]:
            if d.sections:
                samples.append(d.sections[0].text[:200])
        pats = infer_doc_no_patterns(samples)
        if pats:
            cfg["doc_no_patterns"] = pats
            save_config(root, cfg)
            print(f"      自动识别到编号规则 {pats}，重新解析以应用…")
            docs = parse_all(files, cfg, quiet=True, conv_dir=conv_dir)

    # 冲突消解必须在最终解析结果上执行（首轮用空编号规则解析时不会暴露冲突）
    dups = resolve_key_collisions(docs)
    if dups:
        print(f"      发现 {len(dups)} 份文件存在同编号冲突，已按路径区分")

    # 无编号规则时，文件名引用是建图的主要手段
    if cfg.get("filename_refs", True):
        n_ref = enrich_filename_refs(docs)
        if n_ref:
            print(f"      补充文件名引用 {n_ref} 条")

    # ---- 写入对象库 ----
    print("[2/6] 写入内容寻址对象库…")
    manifest_files: Dict[str, Any] = {}
    for d in docs:
        h = S.put_object(root, d)
        manifest_files[d.doc_key] = entry_of(d, h)
    ts = datetime.now().strftime("%Y-%m-%dT%H%M%S")
    new_m = {"ts": ts, "label": args.label, "root": root, "files": manifest_files}

    # ---- 比对 ----
    old_m = None if args.baseline else S.read_latest(root)
    changes: List[D.FileChange] = []
    if old_m is None:
        print("[3/6] 建立基线（不做变更比对）"
              + ("（--baseline：已忽略历史快照）" if args.baseline else "（首次运行）"))
    else:
        print(f"[3/6] 与上次快照 {old_m.get('ts')} 比对…")
        changes = D.compare(root, old_m, new_m)

    # ---- 图谱 ----
    print("[4/6] 构建引用图谱…")
    g = Graph()
    for d in docs:
        g.add(d)
    g.build_edges(use_filename_refs=cfg.get("filename_refs", True))

    changed_keys = {c.doc_key for c in changes
                    if c.change_type in ("added", "removed", "modified") and c.level != "none"}
    impacts: List[Any] = []
    if old_m is not None and changed_keys:
        impacts = g.propagate(changed_keys, max_hops=max(1, args.hops))

    dec = S.load_decisions(root)
    dismissed = dec.get("dismissed", {})

    # ---- 产物 ----
    print("[5/6] 生成报告与图谱…")
    day = datetime.now().strftime("%Y-%m-%d")
    inv = R.write_inventory(root, docs, g)
    graph_path = ""
    if not args.no_graph:
        graph_path = render_graph(root, g.to_visual(changed_keys), day)

    if old_m is None:
        rep = R.write_baseline_report(root, docs, g, cfg, dups)
        todo = ""
    else:
        rep = R.write_change_report(root, old_m, new_m, changes, impacts, g, docs,
                                    dismissed, dups)
        todo = R.write_todo(root, impacts, dismissed, {})

    # ---- 提交快照（放在最后，保证不会覆盖掉比对基准） ----
    print("[6/6] 提交快照…")
    S.write_snapshot(root, new_m)
    pruned = S.prune_snapshots(root, int(cfg.get("history_keep", 24)))
    n_blind = len([d for d in docs if d.parse_status == "failed"])
    R.write_status(root, old_m.get("ts", "") if old_m else "", ts, len(docs),
                   len(changes), len([i for i in impacts if i.target_key not in dismissed]),
                   n_blind, rep, graph_path or rep)

    print()
    print("=" * 60)
    print("完成")
    print("=" * 60)
    print(f"纳管文件     : {len(docs)}")
    print(f"未解析(盲区) : {n_blind}")
    print(f"引用关系     : {len(g.edges)} 条")
    if old_m is not None:
        print(f"本次变更     : {len(changes)} 处 "
              f"（新增 {len([c for c in changes if c.change_type=='added'])} / "
              f"删除 {len([c for c in changes if c.change_type=='removed'])} / "
              f"修改 {len([c for c in changes if c.change_type=='modified'])}）")
        print(f"需同步复核   : {len([i for i in impacts if i.target_key not in dismissed])} 个文件")
    if pruned:
        print(f"历史快照清理 : 删除 {pruned} 份旧快照")
    print()
    print("产物：")
    print(f"  报告   : {rep}")
    if todo:
        print(f"  待办   : {todo}")
    print(f"  台账   : {inv}")
    if graph_path:
        print(f"  关系图 : {graph_path}")
    print(f"  状态   : {os.path.join(root, INDEX_DIR, 'STATUS.md')}")
    if n_blind:
        print()
        print(f"⚠️ 有 {n_blind} 份文件未解析，影响结论可能不完整，请先阅读报告开头的盲区声明。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
