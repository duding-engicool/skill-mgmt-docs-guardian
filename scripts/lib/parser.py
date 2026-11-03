# -*- coding: utf-8 -*-
"""多格式解析与优雅降级。

原则：
1. 任何格式读不动都必须返回 DocModel（parse_status=failed），绝不抛异常中断整体扫描。
2. 依赖缺失时给出明确的安装提示，而不是崩溃。
3. 抽取的是"解析后内容"，后续所有比对基于内容哈希，绝不基于原始文件哈希。
"""
from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Optional, Tuple

from .core import (
    DocModel, Ref, Section, SheetData,
    match_doc_no, norm_text, safe_key, sha256, strip_base,
)

MAX_SHEET_CELLS = 5000
MAX_PDF_CHARS = 400_000

_BOOKTITLE = re.compile(r"[《<]([^》>]{2,60})[》>]")
_REQUIREMENTS: Dict[str, Optional[str]] = {}

# 封面上的模板字样，不能当文件标题（否则满屏都是"程  序  文  件"）
_TITLE_GENERIC = {
    "程序文件", "质量手册", "管理手册", "作业指导书", "作业标准书", "记录表单",
    "表单", "记录", "文件", "管理制度", "支持性文件", "一级文件", "二级文件",
    "三级文件", "四级文件", "受控文件", "章节", "目录",
}
_TITLE_LIKE = re.compile(
    r"(程序|控制|管理|办法|规定|规范|制度|手册|计划|报告|清单|一览表|"
    r"台账|指导书|标准书|流程|预案|方案|职责|表$)"
)
# 表头常见词，被误当成责任人/版本/日期时直接丢掉
_META_JUNK = {
    "审核", "批准", "编制", "校对", "会签", "发放", "受控状态", "版本", "版次",
    "日期", "页次", "修改前内容", "修改后内容", "部门", "责任人", "生效日期",
    "序号", "备注", "内容", "修订内容", "修改内容",
}


def _norm_cmp(s: str) -> str:
    return re.sub(r"\s+", "", s or "").lower()


def _pick_title(cp_title: str, chunks: List[Tuple[int, str]], path: str) -> str:
    """挑最像「文件真名」的标题。

    规则（实测有效）：文件名去掉序号后往往就是文件真名，标题候选里
    与之相同的那个最可信；其次是像文件名的候选；模板字样（"程 序 文 件"）
    打分最低。
    """
    guess = _norm_cmp(strip_base(os.path.splitext(os.path.basename(path))[0]))
    cands: List[str] = []
    if cp_title and cp_title.strip():
        cands.append(cp_title.strip())
    cands.extend(t for _, t in chunks[:40])

    best, best_score = "", -99
    for c in cands:
        t = norm_text(c)
        if not t:
            continue
        n = _norm_cmp(t)
        if len(n) > 60:
            continue
        score = 0
        if guess and len(guess) >= 3:
            if n == guess:
                score += 6
            elif len(guess) >= 5 and (guess in n or n in guess):
                score += 3
        if _TITLE_LIKE.search(n):
            score += 3
        if 4 <= len(n) <= 30:
            score += 1
        if n in _TITLE_GENERIC:
            score -= 3
        if len(n) > 40:
            score -= 2
        if score > best_score:
            best, best_score = t[:60], score
    if best_score <= 1 and chunks:      # 没找到更可信的，退回原行为
        return norm_text(chunks[0][1])[:60]
    return best


def _clean_meta(v: str, kind: str) -> str:
    """洗掉从表头/栏位名误抓的元数据。"""
    v = (v or "").strip()
    if not v or len(v) > 24 or v in _META_JUNK:
        return ""
    if kind == "version" and not re.search(r"[A-Za-z0-9]", v):
        return ""                        # "修改前内容" 这类纯中文不是版本号
    if kind == "date" and not re.search(r"\d", v):
        return ""
    if kind == "owner" and not re.search(r"[\u4e00-\u9fa5A-Za-z]{2,}", v):
        return ""                        # "："、"。" 这类标点残留
    return v


def _need(module: str, pip: str) -> bool:
    """检查依赖是否可用，结果缓存（True/False，不用 None 以免被误判）。"""
    if module in _REQUIREMENTS:
        return _REQUIREMENTS[module] is True
    try:
        __import__(module)
        _REQUIREMENTS[module] = True
        return True
    except Exception:
        _REQUIREMENTS[module] = False
        return False


def dependency_report() -> List[Tuple[str, str, bool, str]]:
    """返回 (能力, 模块, 是否可用, 安装命令) 列表。"""
    items = [
        ("Word .docx", "docx", "python-docx"),
        ("Excel .xlsx/.xlsm", "openpyxl", "openpyxl"),
        ("Excel 老版 .xls", "xlrd", "xlrd"),
        ("PDF 文本层", "pdfplumber", "pdfplumber"),
        ("PPT .pptx", "pptx", "python-pptx"),
    ]
    out = []
    for label, mod, pip in items:
        ok = _need(mod, pip)
        out.append((label, mod, ok, "" if ok else f"pip install {pip}"))
    return out


# ---------------------------------------------------------------- 元数据抽取

def _extract_meta_lines(text: str, keys: List[str]) -> str:
    """在文本中查找 `键：值` / `键 值` 形式的元数据。"""
    for key in keys:
        m = re.search(
            re.escape(key) + r"\s*[:：]?\s*([^\s，,；;。\n]{1,30})", text
        )
        if m:
            return m.group(1).strip()
    return ""


def _next_col(col: str) -> str:
    """A -> B, Z -> AA, AZ -> BA"""
    s = list(col)
    i = len(s) - 1
    while i >= 0:
        if s[i] != "Z":
            s[i] = chr(ord(s[i]) + 1)
            return "".join(s)
        s[i] = "A"
        i -= 1
    return "A" + "".join(s)


def _col_letter(idx: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA"""
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _tables_to_sheets(tables) -> List[SheetData]:
    """把 Word 表格转成 SheetData，复用表格型元数据抽取（编制/审核、版本号等）。"""
    out: List[SheetData] = []
    for ti, tb in enumerate(tables):
        cells: Dict[str, str] = {}
        try:
            for ri, row in enumerate(tb.rows):
                for ci, c in enumerate(row.cells):
                    v = norm_text(c.text)
                    if v:
                        cells[_col_letter(ci) + str(ri + 1)] = v
        except Exception:
            pass
        if cells:
            out.append(SheetData(name=f"表{ti + 1}", dims="", cells=cells))
    return out


def _col_row(addr: str) -> Tuple[str, str]:
    return ("".join(c for c in addr if c.isalpha()),
            "".join(c for c in addr if c.isdigit()))


def _extract_meta_cell(sheets: List[SheetData], keys: List[str],
                       own_no: str = "") -> str:
    """表格布局的元数据抽取。

    1) 台账模式：若知道本文件编号，先定位该编号所在行，再按表头列取值
       （避免把台账第一行的责任人误当成自己的）。
    2) 通用模式：关键词单元格 + 同行右侧一格。
    """
    n_own = norm_text(own_no).lower()
    if n_own:
        for sh in sheets:
            header: Dict[str, Tuple[str, int]] = {}
            for addr, v in sh.cells.items():
                if len(v) > 20:
                    continue
                col, row = _col_row(addr)
                if not row:
                    continue
                for key in keys:
                    if v == key or v.startswith(key):
                        header.setdefault(key, (col, int(row)))
            if not header:
                continue
            for addr, v in sh.cells.items():
                if norm_text(v).lower() != n_own:
                    continue
                _, row = _col_row(addr)
                if not row:
                    continue
                for key in keys:
                    if key not in header:
                        continue
                    col, hrow = header[key]
                    if int(row) <= hrow:
                        continue
                    val = sh.cells.get(col + row, "")
                    if val and len(val) <= 30:
                        return val

    for sh in sheets:
        for addr, v in sh.cells.items():
            if len(v) > 20:
                continue
            col, row = _col_row(addr)
            if not row:
                continue
            for key in keys:
                if v != key and not v.startswith(key):
                    continue
                val = sh.cells.get(_next_col(col) + row, "")
                if val and len(val) <= 30 and val != v:
                    return val
    return ""


def _heading_level(style_name: Optional[str]) -> int:
    if not style_name:
        return 0
    s = style_name.lower()
    m = re.search(r"heading\s*(\d)", s) or re.search(r"标题\s*(\d)", s)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return 1
    if "heading" in s or "标题" in s or "title" in s:
        return 1
    return 0


def _split_number(text: str, clause_pat: str) -> Tuple[str, str]:
    """把段首的条款号与正文分开。返回 (number, rest)。"""
    try:
        # 用命名组，避免 clause_pat 内部的捕获组干扰索引
        # 分隔符允许为空（"4.1文件化信息包括" 中间没有空格），
        # 但空分隔时正文不得以数字开头，否则 "2017年…" 会被当成条款号。
        m = re.match(r"^\s*(?P<num>" + clause_pat + r")(?P<sep>[\s、.:：,，]*)"
                     r"(?P<body>\S.*)$", text)
    except re.error:
        return "", text
    if m:
        num = m.group("num").strip()
        body = m.group("body").strip()
        if not m.group("sep") and body[:1].isdigit():
            return "", text
        return num, body
    return "", text


def _build_sections(chunks: List[Tuple[int, str]], clause_pat: str) -> List[Section]:
    """chunks: (heading_level, text)

    条款号会向下继承：形如「5.2 隔离」的标题确定条款号后，其下的正文段落
    即使自身不以数字开头，也会带上 5.2，便于定位"第几条改了"。
    """
    out: List[Section] = []
    current = ""
    for idx, (level, raw) in enumerate(chunks):
        t = norm_text(raw)
        if len(t) < 1:
            continue
        number, body = _split_number(t, clause_pat)
        if number:
            current = number
        else:
            number = current
        sid = f"s{idx:05d}"
        out.append(Section(sid=sid, level=level, number=number,
                           text=body or t, hash=sha256(t)))
    return out


def _extract_refs(sections: List[Section], cfg: Dict[str, Any],
                  source_hint: str = "") -> List[Ref]:
    """抽取引用：编号引用 + 书名号引用。文件名引用在图谱阶段补充。"""
    refs: List[Ref] = []
    patterns = cfg.get("doc_no_patterns", []) or []
    seen = set()
    for s in sections:
        text = s.text
        if patterns:
            for pat in patterns:
                try:
                    rx = re.compile(pat)
                except re.error:
                    continue
                for m in rx.finditer(text):
                    tgt = m.group(0)
                    if tgt == source_hint:
                        continue
                    k = ("code", tgt, s.sid)
                    if k in seen:
                        continue
                    seen.add(k)
                    ctx = text[max(0, m.start() - 20):m.end() + 20]
                    refs.append(Ref(target=tgt, kind="code",
                                    context=norm_text(ctx), section_id=s.sid))
        if cfg.get("booktitle_refs", True):
            for m in _BOOKTITLE.finditer(text):
                tgt = m.group(1).strip()
                k = ("booktitle", tgt, s.sid)
                if k in seen:
                    continue
                seen.add(k)
                ctx = text[max(0, m.start() - 15):m.end() + 15]
                refs.append(Ref(target=tgt, kind="booktitle",
                                context=norm_text(ctx), section_id=s.sid))
    return refs


# ---------------------------------------------------------------- 各格式解析

def _parse_docx(path: str, cfg: Dict[str, Any]) -> DocModel:
    import docx  # noqa: F401
    from docx import Document

    doc = Document(path)
    chunks: List[Tuple[int, str]] = []
    header_texts: List[str] = []

    for sec in doc.sections:
        try:
            for p in sec.header.paragraphs:
                if p.text.strip():
                    header_texts.append(p.text.strip())
        except Exception:
            pass

    body = doc.element.body
    # 按文档流顺序遍历段落与表格
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            p = Paragraph(child, doc)
            t = p.text.strip()
            if t:
                chunks.append((_heading_level(p.style.name if p.style else None), t))
        elif tag == "tbl":
            tb = Table(child, doc)
            for row in tb.rows:
                cells = [norm_text(c.text) for c in row.cells]
                line = " | ".join(c for c in cells if c)
                if line.strip(" |"):
                    chunks.append((0, line))

    table_sheets = _tables_to_sheets(doc.tables)

    cp = doc.core_properties
    title = _pick_title((cp.title or ""), chunks, path)

    meta_src = "\n".join(header_texts) + "\n" + "\n".join(t for _, t in chunks[:40])
    vk = cfg.get("version_keys", [])
    ek = cfg.get("effective_date_keys", [])
    ok_ = cfg.get("owner_keys", [])
    own_no = match_doc_no(os.path.basename(path), cfg.get("doc_no_patterns", []) or []) or ""
    version = (_extract_meta_cell(table_sheets, vk, own_no)
               or _extract_meta_lines(meta_src, vk))
    eff = (_extract_meta_cell(table_sheets, ek, own_no)
           or _extract_meta_lines(meta_src, ek))
    owner = (_extract_meta_cell(table_sheets, ok_, own_no)
             or _extract_meta_lines(meta_src, ok_))

    version = _clean_meta(version, "version")
    eff = _clean_meta(eff, "date")
    owner = _clean_meta(owner, "owner")

    sections = _build_sections(chunks, cfg.get("clause_pattern", r"\d{1,2}(?:\.\d{1,2}){0,3}"))
    return DocModel(
        doc_key="", rel_path="", ext="docx",
        parse_status="ok", title=title,
        version=version, effective_date=eff, owner=owner,
        last_modified_by=(cp.last_modified_by or "").strip(),
        sections=sections,
    )


def _parse_xlsx(path: str, cfg: Dict[str, Any]) -> DocModel:
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheets: List[SheetData] = []

    for ws in wb.worksheets:
        cells: Dict[str, str] = {}
        n = 0
        try:
            for row in ws.iter_rows():
                for c in row:
                    if c.value is None:
                        continue
                    v = norm_text(str(c.value))
                    if not v:
                        continue
                    cells[c.coordinate] = v
                    n += 1
                    if n >= MAX_SHEET_CELLS:
                        break
                if n >= MAX_SHEET_CELLS:
                    break
        except Exception:
            pass
        try:
            dims = f"{ws.max_row} 行 × {ws.max_column} 列"
        except Exception:
            dims = ""
        sheets.append(SheetData(name=ws.title, dims=dims, cells=cells))
    try:
        wb.close()
    except Exception:
        pass

    return _build_sheet_doc(path, cfg, sheets, "xlsx")


def _build_sheet_doc(path: str, cfg: Dict[str, Any], sheets: List[SheetData],
                     ext: str) -> DocModel:
    """xlsx / xls 共用的收尾：元数据与引用抽取。"""
    refs: List[Ref] = []
    meta_parts: List[str] = []
    for sh in sheets:
        joined = " ".join(sh.cells.values())
        if len(meta_parts) < 3:
            meta_parts.append(joined[:1200])
        for pat in cfg.get("doc_no_patterns", []) or []:
            try:
                rx = re.compile(pat)
            except re.error:
                continue
            for m in rx.finditer(joined):
                refs.append(Ref(target=m.group(0), kind="code",
                                context=norm_text(joined[max(0, m.start() - 15):m.end() + 15]),
                                section_id=sh.name))
        if cfg.get("booktitle_refs", True):
            for m in _BOOKTITLE.finditer(joined):
                refs.append(Ref(target=m.group(1).strip(), kind="booktitle",
                                context=norm_text(joined[max(0, m.start() - 10):m.end() + 10]),
                                section_id=sh.name))

    meta_src = "\n".join(meta_parts)
    vk = cfg.get("version_keys", [])
    ek = cfg.get("effective_date_keys", [])
    ok_ = cfg.get("owner_keys", [])
    own_no = match_doc_no(os.path.basename(path), cfg.get("doc_no_patterns", []) or []) or ""
    return DocModel(
        doc_key="", rel_path="", ext=ext, parse_status="ok",
        title=os.path.splitext(os.path.basename(path))[0],
        version=_extract_meta_cell(sheets, vk, own_no) or _extract_meta_lines(meta_src, vk),
        effective_date=_extract_meta_cell(sheets, ek, own_no) or _extract_meta_lines(meta_src, ek),
        owner=_extract_meta_cell(sheets, ok_, own_no) or _extract_meta_lines(meta_src, ok_),
        sheets=sheets, refs=refs,
    )


def _parse_xls(path: str, cfg: Dict[str, Any]) -> DocModel:
    """老版 Excel .xls（xlrd 2.x 只支持 .xls，正好互补 openpyxl）。"""
    import xlrd

    book = xlrd.open_workbook(path)
    sheets: List[SheetData] = []
    for sh in book.sheets():
        cells: Dict[str, str] = {}
        n = 0
        try:
            for r in range(sh.nrows):
                for c in range(sh.ncols):
                    v = sh.cell_value(r, c)
                    if v is None or v == "":
                        continue
                    s = norm_text(str(v))
                    if not s:
                        continue
                    cells[_col_letter(c) + str(r + 1)] = s
                    n += 1
                    if n >= MAX_SHEET_CELLS:
                        break
                if n >= MAX_SHEET_CELLS:
                    break
        except Exception:
            pass
        sheets.append(SheetData(name=sh.name, dims=f"{sh.nrows} 行 × {sh.ncols} 列",
                                cells=cells))
    return _build_sheet_doc(path, cfg, sheets, "xls")


def _parse_pdf(path: str, cfg: Dict[str, Any]) -> DocModel:
    import pdfplumber

    texts: List[str] = []
    total = 0
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            total += len(t)
            texts.append(t)
            if total > MAX_PDF_CHARS:
                break
    raw = "\n".join(texts)
    if len(norm_text(raw)) < 20:
        return DocModel(doc_key="", rel_path="", ext="pdf",
                        parse_status="failed",
                        parse_note="PDF 无文本层，疑似扫描件（需 OCR，当前环境未启用）")
    chunks = [(0, ln) for ln in raw.splitlines()]
    sections = _build_sections(chunks, cfg.get("clause_pattern", r"\d{1,2}(?:\.\d{1,2}){0,3}"))
    meta_src = raw[:3000]
    return DocModel(
        doc_key="", rel_path="", ext="pdf", parse_status="ok",
        title=os.path.splitext(os.path.basename(path))[0],
        version=_extract_meta_lines(meta_src, cfg.get("version_keys", [])),
        effective_date=_extract_meta_lines(meta_src, cfg.get("effective_date_keys", [])),
        sections=sections,
    )


def _parse_pptx(path: str, cfg: Dict[str, Any]) -> DocModel:
    from pptx import Presentation

    prs = Presentation(path)
    chunks: List[Tuple[int, str]] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                t = "".join(r.text for r in p.runs).strip()
                if t:
                    chunks.append((0, t))
    sections = _build_sections(chunks, cfg.get("clause_pattern", r"\d{1,2}(?:\.\d{1,2}){0,3}"))
    return DocModel(doc_key="", rel_path="", ext="pptx", parse_status="ok",
                    title=os.path.splitext(os.path.basename(path))[0],
                    sections=sections)


def _parse_text(path: str, cfg: Dict[str, Any], ext: str) -> DocModel:
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        raw = f.read()
    chunks: List[Tuple[int, str]] = []
    for ln in raw.splitlines():
        m = re.match(r"^(#{1,6})\s*(.*)$", ln)
        if m:
            chunks.append((len(m.group(1)), m.group(2)))
        else:
            chunks.append((0, ln))
    sections = _build_sections(chunks, cfg.get("clause_pattern", r"\d{1,2}(?:\.\d{1,2}){0,3}"))
    meta_src = raw[:3000]
    return DocModel(doc_key="", rel_path="", ext=ext, parse_status="ok",
                    title=os.path.splitext(os.path.basename(path))[0],
                    version=_extract_meta_lines(meta_src, cfg.get("version_keys", [])),
                    effective_date=_extract_meta_lines(meta_src, cfg.get("effective_date_keys", [])),
                    owner=_extract_meta_lines(meta_src, cfg.get("owner_keys", [])),
                    sections=sections)


# ---------------------------------------------------------------- 入口

HANDLERS = {
    ".docx": ("docx", "python-docx", _parse_docx),
    ".xlsx": ("openpyxl", "openpyxl", _parse_xlsx),
    ".xlsm": ("openpyxl", "openpyxl", _parse_xlsx),
    ".xls": ("xlrd", "xlrd", _parse_xls),
    ".pdf": ("pdfplumber", "pdfplumber", _parse_pdf),
    ".pptx": ("pptx", "python-pptx", _parse_pptx),
    ".txt": (None, None, None),
    ".md": (None, None, None),
    ".csv": (None, None, None),
}

NEEDS_CONVERT = {
    ".doc": "老版 Word 格式，建议另存为 .docx",
    ".wps": "WPS 专有格式，建议另存为 .docx",
    ".et": "WPS 表格专有格式，建议另存为 .xlsx",
    ".dps": "WPS 演示专有格式，建议另存为 .pptx",
    ".jpg": "图片，需 OCR，建议转为可编辑文本",
    ".png": "图片，需 OCR，建议转为可编辑文本",
    ".tif": "图片，需 OCR，建议转为可编辑文本",
}


def parse_file(path: str, rel_path: str, cfg: Dict[str, Any]) -> DocModel:
    """解析单个文件，永不抛异常。"""
    ext = os.path.splitext(path)[1].lower()
    try:
        st = os.stat(path)
        size, mtime = st.st_size, st.st_mtime
    except Exception:
        size, mtime = 0, 0.0

    m: Optional[DocModel] = None
    err = ""
    try:
        if ext in (".txt", ".md", ".csv"):
            m = _parse_text(path, cfg, ext.lstrip("."))
        elif ext in HANDLERS:
            mod, pip, fn = HANDLERS[ext]
            if mod and not _need(mod, pip or ""):
                err = f"缺少依赖 {pip}，请执行 pip install {pip}"
            else:
                m = fn(path, cfg)
        else:
            err = NEEDS_CONVERT.get(ext, f"暂不支持的格式 {ext or '(无扩展名)'}")
    except Exception as e:
        m = None
        err = f"解析异常：{type(e).__name__}: {e}"[:200]

    if m is None:
        m = DocModel(doc_key="", rel_path=rel_path, ext=ext.lstrip("."),
                     size=size, mtime=mtime, parse_status="failed", parse_note=err)

    # ---- 统一后处理：解析成功与失败都要走，否则失败文件主键会全为空 ----
    m.rel_path = rel_path
    m.size, m.mtime = size, mtime
    m.ext = ext.lstrip(".")

    # 主键：编号优先，无编号退化为路径
    fname = os.path.basename(path)
    patterns = cfg.get("doc_no_patterns", []) or []
    no = match_doc_no(fname, patterns) or match_doc_no(m.title, patterns)
    if not no and m.sections:
        no = match_doc_no(m.sections[0].text[:80], patterns) if m.sections[0].text else None
    m.doc_no = no or ""
    m.doc_key = ("no:" + no) if no else safe_key(rel_path)

    if not m.refs and m.sections:
        m.refs = _extract_refs(m.sections, cfg, source_hint=no or "")

    # 无正文内容时（不可解析文件），用 大小+修改时间 兜底做指纹，
    # 至少能发现"这个文件被动过"，但无法定位改了什么
    text = m.all_text()
    if text.strip():
        m.content_hash = sha256(text)
    else:
        # 注意：content_hash 会用作对象库文件名，不能含 Windows 非法字符（: 等）
        m.content_hash = "blob_" + sha256(f"{size}|{int(mtime)}|{rel_path}")
        if m.parse_status == "ok":
            m.parse_status = "partial"
            m.parse_note = (m.parse_note + "；未提取到正文内容").strip("；")
    return m
