# -*- coding: utf-8 -*-
"""批量把 .doc 转成 .docx 副本，供分析使用。

铁律：原始文件只读。转换结果写到独立目录（默认 .cb-index/converted/），
保持原目录的层级结构；scan.py 会自动优先用副本解析，但文件身份（编号/路径/台账）
仍然记在原 .doc 名下。

用法：
  python convert_docs.py --path "D:\\体系文件"
  python convert_docs.py --path "D:\\体系文件" --out "D:\\转换副本"
  python convert_docs.py --path "D:\\体系文件" --dry-run

依赖：pip install pywin32（Windows + 已安装 Microsoft Office 或 WPS）
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from typing import List, Tuple

from lib.core import INDEX_DIR, is_ignored, load_config

WD_FORMAT_DOCX = 16


def collect_doc_files(root: str, cfg) -> List[str]:
    out = []
    ignore_dirs = {d.lower() for d in cfg.get("ignore_dirs", [])}
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames
                             if d.lower() not in ignore_dirs and not d.startswith("~"))
        for fn in filenames:
            if os.path.splitext(fn)[1].lower() != ".doc":
                continue
            rel = os.path.relpath(os.path.join(dirpath, fn), root)
            if is_ignored(rel, cfg):
                continue
            out.append(rel)
    return sorted(out)


def open_app():
    """优先 Word，其次 WPS。

    注意（实测坑）：
      - 本机常常并没有真·Word，Word.Application 也是 WPS 在模拟，
        WPS 的行为是「关掉最后一个文档就整个进程退出」，后续调用全部
        报 RPC 服务器不可用。解决办法：全程常驻一个空白文档。
      - SaveAs2 是 Word 2010+ 的方法，老版本没有，需要回退到 SaveAs。
    """
    import win32com.client as win32

    last_err = None
    for prog in ("Word.Application", "KWps.Application"):
        try:
            app = win32.Dispatch(prog)
            app.Visible = False
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            return app, prog
        except Exception as e:
            last_err = e
    raise RuntimeError(f"未能启动 Word 或 WPS（{last_err}）。请确认已安装桌面版 Office/WPS。")


def make_keeper(app):
    """常驻一个空白文档，防止 WPS 关掉最后一个文档后自杀。失败返回 None。"""
    try:
        return app.Documents.Add()
    except Exception:
        return None


def app_alive(app) -> bool:
    try:
        app.Name  # noqa: B018
        return True
    except Exception:
        return False


def save_as(doc, dst: str) -> None:
    """SaveAs2 在 Word 2007 及更早版本不存在，回退 SaveAs。"""
    try:
        doc.SaveAs2(os.path.abspath(dst), FileFormat=WD_FORMAT_DOCX)
    except AttributeError:
        doc.SaveAs(os.path.abspath(dst), FileFormat=WD_FORMAT_DOCX)


def main() -> int:
    ap = argparse.ArgumentParser(description="批量 .doc → .docx 转换（副本，不改原文件）")
    ap.add_argument("--path", required=True, help="目标文件夹绝对路径")
    ap.add_argument("--out", default="", help=f"输出目录，默认 <目标>/{INDEX_DIR}/converted")
    ap.add_argument("--dry-run", action="store_true", help="只列出将要转换的文件，不实际操作")
    ap.add_argument("--force", action="store_true", help="已存在副本也重新转换")
    ap.add_argument("--no-refresh", action="store_true",
                    help="不检查副本新鲜度（默认：原文件比副本新时自动重转）")
    args = ap.parse_args()

    root = os.path.abspath(args.path)
    if not os.path.isdir(root):
        print(f"[错误] 路径不存在：{root}", file=sys.stderr)
        return 2

    out_dir = os.path.abspath(args.out) if args.out else os.path.join(root, INDEX_DIR, "converted")
    cfg = load_config(root)
    rels = collect_doc_files(root, cfg)

    print("=" * 60)
    print(".doc → .docx 转换（副本）")
    print("=" * 60)
    print(f"源目录  : {root}")
    print(f"输出目录: {out_dir}")
    print(f"待转换  : {len(rels)} 个 .doc")
    print("说明    : 原始文件不会被修改或移动")
    print()

    if args.dry_run:
        for r in rels[:30]:
            print("  " + r)
        if len(rels) > 30:
            print(f"  …另有 {len(rels)-30} 个")
        return 0

    if not rels:
        print("没有需要转换的 .doc 文件。")
        return 0

    try:
        import win32com.client  # noqa: F401
    except Exception:
        print("[错误] 缺少 pywin32，请先执行：pip install pywin32", file=sys.stderr)
        return 3

    os.makedirs(out_dir, exist_ok=True)
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    app, progname = open_app()
    keeper = make_keeper(app)
    print(f"已启动 {progname}（后台运行，请勿手动关闭）")
    print()

    ok, skipped, failed, refreshed = 0, 0, 0, 0
    failures: List[Tuple[str, str]] = []
    t0 = time.time()

    try:
        for i, rel in enumerate(rels, 1):
            src = os.path.join(root, rel)
            dst = os.path.join(out_dir, os.path.splitext(rel)[0] + ".docx")
            os.makedirs(os.path.dirname(dst), exist_ok=True)

            # 副本新鲜度：原文件被改过就必须重转，否则 scan.py 读的是旧副本，
            # 会「改了文件却报 0 变更」——这是最危险的静默失败。
            stale = False
            if os.path.isfile(dst):
                try:
                    stale = int(os.path.getmtime(src)) > int(os.path.getmtime(dst)) + 1
                except OSError:
                    stale = False

            if os.path.isfile(dst) and not args.force and not (stale and not args.no_refresh):
                skipped += 1
                continue

            if i % 20 == 0 or i == len(rels):
                print(f"  转换中 {i}/{len(rels)}（成功 {ok} / 失败 {failed}）", flush=True)

            if not app_alive(app):          # 进程被意外杀掉时自动重启，不中断整批
                try:
                    app, progname = open_app()
                    keeper = make_keeper(app)
                    print(f"  （{progname} 意外退出，已自动重启）", flush=True)
                except Exception as e:
                    failed += 1
                    failures.append((rel, f"重启失败: {e}"[:160]))
                    continue

            doc = None
            try:
                doc = app.Documents.Open(os.path.abspath(src), ReadOnly=True)
                save_as(doc, dst)
                if stale:
                    refreshed += 1
                    print(f"  ↻ 原文件已更新，重转：{rel}", flush=True)
                else:
                    ok += 1
            except Exception as e:
                failed += 1
                failures.append((rel, f"{type(e).__name__}: {e}"[:160]))
            finally:
                if doc is not None:
                    try:
                        doc.Close(0)
                    except Exception:
                        pass
    finally:
        if keeper is not None:
            try:
                keeper.Close(0)
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass

    dt = time.time() - t0
    print()
    print("=" * 60)
    print("转换完成")
    print("=" * 60)
    print(f"新转   : {ok}")
    print(f"刷新   : {refreshed}（原文件被改过，旧副本已作废重转）")
    print(f"跳过   : {skipped}（副本仍是最新的）")
    print(f"失败   : {failed}")
    print(f"耗时   : {dt/60:.1f} 分钟")
    print(f"输出   : {out_dir}")
    if failures:
        print()
        print("失败清单（这些文件仍会作为盲区出现在报告中）：")
        for r, e in failures[:20]:
            print(f"  - {r}  →  {e}")
        if len(failures) > 20:
            print(f"  …另有 {len(failures)-20} 个")
    print()
    print("下一步：重新运行 scan.py，它会自动优先使用副本内容解析（台账仍记在原 .doc 名下）。")
    return 0


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    sys.exit(main())
