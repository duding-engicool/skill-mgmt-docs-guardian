# -*- coding: utf-8 -*-
"""报告生成：变更影响报告（Markdown）、待办清单（CSV）、体系台账（CSV）。"""
from __future__ import annotations

import csv
import os
from datetime import datetime
from typing import Any, Dict, List, Set

from .core import DocModel, INDEX_DIR
from .diff import FileChange
from .graph import Graph, Impact


def _reports_dir(root: str) -> str:
    d = os.path.join(root, INDEX_DIR, "reports")
    os.makedirs(d, exist_ok=True)
    return d


def _today() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def _esc(s: str) -> str:
    return (s or "").replace("|", "\\|").replace("\n", " ").strip()


# ---------------------------------------------------------------- 台账

def write_inventory(root: str, docs: List[DocModel], graph: Graph) -> str:
    path = os.path.join(_reports_dir(root), "体系台账.csv")
    deg: Dict[str, int] = {}
    for e in graph.edges:
        deg[e.target] = deg.get(e.target, 0) + 1

    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["编号", "标题", "相对路径", "格式", "版本", "生效日期",
                    "责任人", "最后修改人", "解析状态", "解析备注", "被引用次数"])
        for d in sorted(docs, key=lambda x: (x.doc_no or "~", x.rel_path)):
            w.writerow([
                d.doc_no, d.title, d.rel_path, d.ext, d.version, d.effective_date,
                d.owner, d.last_modified_by, d.parse_status, d.parse_note,
                deg.get(d.doc_key, 0),
            ])
    return path


# ---------------------------------------------------------------- 待办清单

def write_todo(root: str, impacts: List[Impact], dismissed: Dict[str, Any],
               changed_titles: Dict[str, str]) -> str:
    path = os.path.join(_reports_dir(root), f"{_today()}_待办清单.csv")
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["序号", "受影响文件", "相对路径", "责任人", "影响距离",
                    "置信度", "触发源", "依据", "状态"])
        n = 0
        for i in impacts:
            if i.target_key in dismissed:
                continue
            n += 1
            ev_titles = {e.get("changed_title") or e.get("from_title", "") for e in i.evidence}
            srcs = "；".join(sorted(t for t in ev_titles if t)[:3])
            reason = "；".join(i.reasons[:2])
            w.writerow([n, i.target_title, i.target_path, i.target_owner,
                        f"{i.distance} 跳", i.confidence, srcs, reason, "待处理"])
    return path


# ---------------------------------------------------------------- 主报告

def write_change_report(root: str, old_m: Dict[str, Any], new_m: Dict[str, Any],
                        changes: List[FileChange], impacts: List[Impact],
                        graph: Graph, docs: List[DocModel],
                        dismissed: Dict[str, Any],
                        duplicates: List[Any] = ()) -> str:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    day = _today()
    path = os.path.join(_reports_dir(root), f"{day}_变更影响报告.md")

    unparsed = [d for d in docs if d.parse_status != "ok"]
    blind = [d for d in docs if d.parse_status == "failed"]

    added = [c for c in changes if c.change_type == "added"]
    removed = [c for c in changes if c.change_type == "removed"]
    renamed = [c for c in changes if c.change_type == "renamed"]
    modified = [c for c in changes if c.change_type == "modified"]
    subst = [c for c in modified if c.level == "substantive"]
    meta_only = [c for c in modified if c.level == "metadata"]

    # 版本脱节：内容实质变更但版本号未动
    drift = []
    for c in subst:
        if "版本号" not in c.meta_changes:
            drift.append(c)

    L: List[str] = []
    A = L.append
    A(f"# 体系文件变更影响报告")
    A("")
    A(f"- **扫描目录**：`{root}`")
    A(f"- **上次快照**：{old_m.get('ts','-')}")
    A(f"- **本次快照**：{new_m.get('ts','-')}")
    A(f"- **生成时间**：{ts}")
    A(f"- **纳管文件**：{len(docs)} 份（可解析 {len(docs)-len(unparsed)} 份）")
    A("")

    # 盲区声明置顶
    A("## ⚠️ 分析盲区（务必先读）")
    A("")
    if not blind:
        A("本次无未解析文件，分析覆盖全部纳管文件。")
    else:
        A(f"**本次有 {len(blind)} 份文件未能解析内容。** 以下影响结论**未包含**这些文件的引用关系，")
        A("若其中包含被引用的核心文件，影响清单可能不完整。建议将其转换为 .docx 后重新扫描。")
        A("")
        A("| 相对路径 | 格式 | 原因 |")
        A("|---|---|---|")
        for d in blind[:60]:
            A(f"| {_esc(d.rel_path)} | {d.ext} | {_esc(d.parse_note)} |")
        if len(blind) > 60:
            A(f"| … | | 另有 {len(blind)-60} 份，详见体系台账.csv |")
    A("")

    # 变更概要
    A("## 一、变更概要")
    A("")
    A("| 类别 | 数量 |")
    A("|---|---|")
    A(f"| 新增文件 | {len(added)} |")
    A(f"| 删除文件 | {len(removed)} |")
    A(f"| 重命名 / 移动 | {len(renamed)} |")
    A(f"| 实质内容变更 | {len(subst)} |")
    A(f"| 仅元信息变更 | {len(meta_only)} |")
    A(f"| 需同步复核的文件 | {len([i for i in impacts if i.target_key not in dismissed])} |")
    A("")

    # 变更明细
    A("## 二、变更明细")
    A("")

    if added:
        A(f"### 2.1 新增文件（{len(added)}）")
        A("")
        A("| 编号 | 标题 | 路径 | 解析状态 |")
        A("|---|---|---|---|")
        for c in added:
            A(f"| {_esc(c.doc_no)} | {_esc(c.title)} | {_esc(c.new_path)} | {c.new_status} |")
        A("")

    if removed:
        A(f"### 2.2 删除文件（{len(removed)}）")
        A("")
        A("| 编号 | 标题 | 原路径 |")
        A("|---|---|---|")
        for c in removed:
            A(f"| {_esc(c.doc_no)} | {_esc(c.title)} | {_esc(c.old_path)} |")
        A("")
        A("> 删除文件若仍被其他文件引用，会在「四、一致性问题」中报为失效引用。")
        A("")

    if subst:
        A(f"### 2.3 实质内容变更（{len(subst)}）")
        A("")
        for c in subst:
            A(f"#### {_esc(c.title or c.doc_no or c.doc_key)}")
            A("")
            A(f"- 编号：`{c.doc_no or '（无编号）'}`　路径：`{c.new_path}`")
            if c.meta_changes:
                mc = "；".join(f"{k}：{v[0]} → {v[1]}" for k, v in c.meta_changes.items())
                A(f"- 元信息变化：{mc}")
            A("")
            secs = [d for d in c.details if d.get("kind") == "section"]
            cells = [d for d in c.details if d.get("kind") == "cell"]
            if secs:
                A("| 变更类型 | 条款 | 变更前 | 变更后 |")
                A("|---|---|---|---|")
                op_lbl = {"replace": "修改", "insert": "新增", "delete": "删除"}
                for d in secs[:40]:
                    A(f"| {op_lbl.get(d['op'], d['op'])} | {_esc(d.get('number',''))} "
                      f"| {_esc(d.get('old_text',''))[:120]} | {_esc(d.get('new_text',''))[:120]} |")
                if len(secs) > 40:
                    A(f"| … | | 另有 {len(secs)-40} 处条款变更 | |")
                A("")
            if cells:
                A("| 变更类型 | 工作表 | 单元格 | 变更前 | 变更后 |")
                A("|---|---|---|---|---|")
                op_lbl = {"replace": "修改", "insert": "新增", "delete": "删除"}
                for d in cells[:40]:
                    A(f"| {op_lbl.get(d['op'], d['op'])} | {_esc(d.get('sheet',''))} "
                      f"| {_esc(d.get('addr',''))} | {_esc(d.get('old_text',''))[:80]} "
                      f"| {_esc(d.get('new_text',''))[:80]} |")
                if len(cells) > 40:
                    A(f"| … | | | 另有 {len(cells)-40} 处单元格变更 | |")
                A("")
            if c.truncated:
                A("> 变更明细已截断，仅显示前 200 条。")
                A("")

    if meta_only:
        A(f"### 2.4 仅元信息变更（{len(meta_only)}）")
        A("")
        A("| 编号 | 标题 | 元信息变化 |")
        A("|---|---|---|")
        for c in meta_only:
            mc = "；".join(f"{k}：{v[0]} → {v[1]}" for k, v in c.meta_changes.items())
            A(f"| {_esc(c.doc_no)} | {_esc(c.title)} | {_esc(mc)} |")
        A("")

    if renamed:
        A(f"### 2.5 重命名 / 移动（{len(renamed)}）")
        A("")
        A("| 编号 | 标题 | 原路径 | 新路径 |")
        A("|---|---|---|---|")
        for c in renamed:
            A(f"| {_esc(c.doc_no)} | {_esc(c.title)} | {_esc(c.old_path)} | {_esc(c.new_path)} |")
        A("")

    if not (added or removed or subst or meta_only or renamed):
        A("本次未检测到任何变更（内容哈希与上次快照完全一致）。")
        A("")

    # 影响分析
    A("## 三、影响分析（需同步复核的文件）")
    A("")
    live = [i for i in impacts if i.target_key not in dismissed]
    if not live:
        A("未发现需要同步复核的文件。")
    else:
        for conf, label in (("high", "高置信 · 建议必改"),
                            ("medium", "中置信 · 建议复核"),
                            ("low", "低置信 · 仅需知晓")):
            grp = [i for i in live if i.confidence == conf]
            if not grp:
                continue
            A(f"### {label}（{len(grp)}）")
            A("")
            A("| 受影响文件 | 责任人 | 距离 | 触发的变更文件 | 依据 |")
            A("|---|---|---|---|---|")
            for i in grp:
                ev_titles = {e.get("changed_title") or e.get("from_title", "") for e in i.evidence}
                srcs = "；".join(sorted(t for t in ev_titles if t)[:2])
                reason = "；".join(i.reasons[:1])
                A(f"| {_esc(i.target_title)} | {_esc(i.target_owner)} | {i.distance} 跳 "
                  f"| {_esc(srcs)} | {_esc(reason)} |")
            A("")
    if dismissed:
        A(f"> 已忽略 {len(dismissed)} 项（历史标注，不进入待办清单）。")
        A("")

    # 一致性问题
    A("## 四、一致性问题（规则层判定，无需模型）")
    A("")
    if graph.dangling:
        A(f"### 4.1 失效引用（{len(graph.dangling)} 处）")
        A("")
        A("正文引用了本单位体系文件，但库里找不到对应文件：")
        A("")
        A("| 引用方 | 引用对象（找不到） | 上下文 |")
        A("|---|---|---|")
        for from_key, tgt, ctx in graph.dangling[:40]:
            d = graph.nodes.get(from_key)
            A(f"| {_esc(d.title if d else from_key)} | `{_esc(tgt)}` | {_esc(ctx)[:80]} |")
        if len(graph.dangling) > 40:
            A(f"| … | | 另有 {len(graph.dangling)-40} 处 |")
        A("")
    else:
        A("### 4.1 失效引用：**无**")
        A("")

    if drift:
        A(f"### 4.2 版本脱节（{len(drift)} 份）")
        A("")
        A("内容发生实质变更，但文档内版本号未更新——体系管理中属于典型不符合项：")
        A("")
        A("| 文件 | 编号 | 当前版本号 |")
        A("|---|---|---|")
        for c in drift:
            e = new_m["files"].get(c.doc_key, {})
            A(f"| {_esc(c.title)} | {_esc(c.doc_no)} | {_esc(e.get('version','') or '（未识别）')} |")
        A("")
    else:
        A("### 4.2 版本脱节：**无**")
        A("")

    if duplicates:
        A(f"### 4.3 同编号重复（{len(duplicates)} 份）")
        A("")
        A("同一编号出现在多份文件中，通常是新旧版本并存，需确认哪份现行有效：")
        A("")
        A("| 编号 | 相对路径 |")
        A("|---|---|")
        for no, p in duplicates[:40]:
            A(f"| {_esc(no)} | {_esc(p)} |")
        A("")

    orphans = graph.orphans()
    A(f"### 4.{4 if duplicates else 3} 孤立文件（{len(orphans)} 份）")
    A("")
    if orphans:
        A("既未引用他文件、也未被他文件引用，可能是废弃文件或引用抽取不完整：")
        A("")
        for k in orphans[:30]:
            d = graph.nodes[k]
            A(f"- `{d.rel_path}`")
        if len(orphans) > 30:
            A(f"- …另有 {len(orphans)-30} 份")
    else:
        A("无。")
    A("")

    # 待办
    A("## 五、待办清单")
    A("")
    A(f"见同目录 `{day}_待办清单.csv`，可直接分发给各文件责任人。")
    A("")
    A("## 六、产物索引")
    A("")
    A(f"- 体系台账：`reports/体系台账.csv`")
    A(f"- 关系图：`reports/{day}_关系图.html`（双击打开）")
    A(f"- 待办清单：`reports/{day}_待办清单.csv`")
    A("")
    A("---")
    A("")
    A("> 本报告由规则层自动生成，未调用语言模型。所有结论均附原文依据，请人工复核后执行修改。")

    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    return path


# ---------------------------------------------------------------- 基线报告

def write_baseline_report(root: str, docs: List[DocModel], graph: Graph,
                          cfg: Dict[str, Any],
                          duplicates: List[Any] = ()) -> str:
    day = _today()
    path = os.path.join(_reports_dir(root), f"{day}_基线建立报告.md")
    ok = [d for d in docs if d.parse_status == "ok"]
    partial = [d for d in docs if d.parse_status == "partial"]
    failed = [d for d in docs if d.parse_status == "failed"]

    L: List[str] = []
    A = L.append
    A("# 体系文件基线建立报告")
    A("")
    A(f"- **扫描目录**：`{root}`")
    A(f"- **生成时间**：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    A(f"- **纳管文件**：{len(docs)} 份")
    A("")
    A("## 一、首次扫描结果")
    A("")
    A("| 类别 | 数量 |")
    A("|---|---|")
    A(f"| 完整解析 | {len(ok)} |")
    A(f"| 部分解析（无正文） | {len(partial)} |")
    A(f"| 无法解析 | {len(failed)} |")
    A(f"| 识别出的引用关系 | {len(graph.edges)} 条 |")
    A("")
    A("> **本次为首次运行，仅建立基线，不做变更比对。** 下次扫描将自动与本次快照对比并输出变更影响报告。")
    A("")

    A("## 二、识别到的文件编号规则")
    A("")
    pats = cfg.get("doc_no_patterns") or []
    if pats:
        A("已自动推断并写入配置：")
        A("")
        for p in pats:
            A(f"- `{p}`")
        A("")
        A("如识别有误，编辑 `.cb-index/config.json` 中的 `doc_no_patterns` 后重新扫描即可。")
    else:
        A("未自动识别出文件编号规则。影响分析将只能依赖书名号/文件名引用，精度会下降。")
        A("")
        A("建议在 `.cb-index/config.json` 中手工填写 `doc_no_patterns`（正则）。")
    A("")

    # ---- 体系健康体检：首次运行也要给出可执行结论，不能只报统计 ----
    orphans = graph.orphans()
    A("## 三、体系健康体检（规则层判定，无需模型）")
    A("")

    A(f"### 3.1 失效引用（{len(graph.dangling)} 处）")
    A("")
    if graph.dangling:
        A("正文引用了**本单位体系文件**，但库里找不到对应文件。"
          "常见原因：文件已改名/删除、名称写法与文件名不一致、文件缺失。")
        A("")
        A("| 引用方 | 引用对象（找不到） | 上下文 |")
        A("|---|---|---|")
        for from_key, tgt, ctx in graph.dangling[:40]:
            d = graph.nodes.get(from_key)
            A(f"| {_esc(d.title if d else from_key)} | `{_esc(tgt)}` | {_esc(ctx)[:80]} |")
        if len(graph.dangling) > 40:
            A(f"| … | | 另有 {len(graph.dangling)-40} 处 |")
    else:
        A("无。")
    A("")

    if graph.unmatched:
        A(f"### 3.1b 待确认引用（{len(graph.unmatched)} 处）")
        A("")
        A("既不像本单位文件、也不像外部法规，需人工判断：")
        A("")
        A("| 引用方 | 引用对象 | 上下文 |")
        A("|---|---|---|")
        for from_key, tgt, ctx in graph.unmatched[:25]:
            d = graph.nodes.get(from_key)
            A(f"| {_esc(d.title if d else from_key)} | `{_esc(tgt)}` | {_esc(ctx)[:60]} |")
        if len(graph.unmatched) > 25:
            A(f"| … | | 另有 {len(graph.unmatched)-25} 处 |")
        A("")

    if graph.external:
        A(f"### 3.1c 外部法规/标准引用（{len(graph.external)} 处，仅供参考，非问题）")
        A("")
        names = sorted({t for _, t, _ in graph.external})
        A("、".join(f"《{_esc(x)}》" for x in names[:60]))
        if len(names) > 60:
            A(f"\n\n…另有 {len(names)-60} 种")
        A("")

    A(f"### 3.2 孤立文件（{len(orphans)} 份）")
    A("")
    if orphans:
        A("既未引用他文件、也未被他文件引用，可能是废弃文件或引用抽取不完整：")
        A("")
        for k in orphans[:30]:
            d = graph.nodes[k]
            A(f"- `{d.rel_path}`")
        if len(orphans) > 30:
            A(f"- …另有 {len(orphans)-30} 份")
    else:
        A("无。")
    A("")

    A(f"### 3.3 同编号重复（{len(duplicates)} 份）")
    A("")
    if duplicates:
        A("同一编号出现在多份文件中，通常是新旧版本并存，需确认哪份现行有效：")
        A("")
        A("| 编号 | 相对路径 |")
        A("|---|---|")
        for no, p in duplicates[:40]:
            A(f"| {_esc(no)} | {_esc(p)} |")
    else:
        A("无。")
    A("")

    if failed:
        A(f"## 四、⚠️ 无法解析的文件（{len(failed)}）")
        A("")
        A("| 相对路径 | 格式 | 原因 |")
        A("|---|---|---|")
        for d in failed[:60]:
            A(f"| {_esc(d.rel_path)} | {d.ext} | {_esc(d.parse_note)} |")
        if len(failed) > 60:
            A(f"| … | | 另有 {len(failed)-60} 份，详见体系台账.csv |")
        A("")

    A("## 产物索引")
    A("")
    A("- 体系台账：`reports/体系台账.csv`")
    A(f"- 关系图：`reports/{day}_关系图.html`（双击打开）")
    A("")
    A("---")
    A("")
    A("> 索引目录 `.cb-index/` 可安全删除，删除后仅失去历史对比能力，不影响原始文件。")

    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    return path


# ---------------------------------------------------------------- STATUS

def write_status(root: str, old_ts: str, new_ts: str, n_docs: int,
                 n_changes: int, n_impacts: int, n_blind: int,
                 report_path: str, graph_path: str) -> str:
    p = os.path.join(root, INDEX_DIR, "STATUS.md")
    rel_r = os.path.relpath(report_path, root).replace("\\", "/")
    rel_g = os.path.relpath(graph_path, root).replace("\\", "/")
    txt = f"""# 工作区状态（供 Agent 读取的交接笔记）

- **目标目录**：`{root}`
- **上次快照**：{old_ts or '（无，本次为首次）'}
- **本次快照**：{new_ts}
- **纳管文件**：{n_docs} 份
- **本次变更**：{n_changes} 处
- **需同步复核**：{n_impacts} 个文件
- **未解析文件（盲区）**：{n_blind} 份

## 最近产物

- 报告：`{rel_r}`
- 关系图：`{rel_g}`
- 台账：`{INDEX_DIR}/reports/体系台账.csv`

## Agent 行动提示

1. 每次运行先读本文件，再读 `config.json`。
2. 向用户汇报时，**必须先声明盲区**（未解析文件数量）。
3. 原目录只读，禁止改写体系文件；修改建议只出现在报告里。
4. 用户标注"不用改"的影响项，用 `scan.py --dismiss <doc_key>` 记录。
"""
    with open(p, "w", encoding="utf-8") as f:
        f.write(txt)
    return p
