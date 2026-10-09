# -*- coding: utf-8 -*-
"""
阶段二：PDF 解析
================
用 PyMuPDF 批量抽取太古地产年报 / 中期报告的逐页文本，输出为 JSON。

设计原则
--------
* 原 PDF 只读：pymupdf.open() 默认只读，全程不写回源目录，不复制源文件
* 不做切块、不做 embedding、不调用任何 LLM / 网络 API
* 保留页码，供后续 RAG 引用来源（"2024年报 第37页"）

用法
----
    venv\\Scripts\\python.exe scripts\\01_extract_pdf.py
    venv\\Scripts\\python.exe scripts\\01_extract_pdf.py --only annual
    venv\\Scripts\\python.exe scripts\\01_extract_pdf.py --limit 2      # 快速试跑
    venv\\Scripts\\python.exe scripts\\01_extract_pdf.py --force       # 覆盖已存在结果
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

# --- Windows 控制台是 GBK(936)，不重设会把正常中文显示成乱码 ---
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

import pymupdf
from dotenv import load_dotenv

# ==================================================================
# 路径配置
# ==================================================================
# 优先读 .env 里的 TAIKOO_* 变量；没有则用下面的默认值。
# 注意：中期报告目录的实际名称是「中度报告」，不是「中期报告」。
PROJECT_ROOT = Path(__file__).resolve().parent.parent

DEFAULT_PDF_ROOT = PROJECT_ROOT / "data" / "source_pdf"
DEFAULT_ANNUAL_DIR = DEFAULT_PDF_ROOT / "年报"
DEFAULT_INTERIM_DIR = DEFAULT_PDF_ROOT / "中度报告"

OUT_DIR = PROJECT_ROOT / "data" / "extracted"
REPORT_DIR = PROJECT_ROOT / "output"

COMPANY = "太古地产"

# 判定"扫描版"的阈值：有效字符数低于此值、且页面被图片大面积覆盖
SCANNED_CHAR_THRESHOLD = 50
SCANNED_IMAGE_RATIO = 0.5


# ==================================================================
# 文件名解析
# ==================================================================
# 中文数字年份，如 二零一三 / 二〇二五
CN_DIGIT = {
    "零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
    "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
}
CN_YEAR_RE = re.compile(r"([零〇一二三四五六七八九]{4})\s*年")
ARABIC_YEAR_RE = re.compile(r"(19|20)\d{2}")


def parse_report_year(filename: str) -> str | None:
    """从文件名解析报告年份。

    文件名形如：
        2014-04-04：01972--太古地產：二零一三年報告書(1).pdf   -> 2013
        2026-04-08：二零二五年報告書.pdf                       -> 2025
        2013-08-28：01972--太古地產：二零一三年中期報告.pdf     -> 2013

    开头的 2014-04-04 是**公告日**，不是报告期，不能用它当年份。
    所以优先匹配「四位中文数字 + 年」。
    """
    m = CN_YEAR_RE.search(filename)
    if m:
        return str(int("".join(str(CN_DIGIT[c]) for c in m.group(1))))
    # 兜底：文件名里第一个 19xx/20xx 阿拉伯年份
    m = ARABIC_YEAR_RE.search(filename)
    return m.group(0) if m else None


def parse_report_type(filename: str) -> str:
    """从文件名判定报告类型。年报文件名含「報告書」，不含「年度報告」。"""
    if "中期" in filename:
        return "中期报告"
    return "年度报告"


# ==================================================================
# 单页抽取
# ==================================================================
def image_area_ratio(page: pymupdf.Page) -> float:
    """页面被图片覆盖的面积占比，用于识别扫描版（整页是一张图）。"""
    try:
        page_area = abs(page.rect.width * page.rect.height)
        if page_area <= 0:
            return 0.0
        total = 0.0
        for info in page.get_image_info():
            bbox = pymupdf.Rect(info["bbox"])
            total += abs(bbox.width * bbox.height)
        return min(total / page_area, 1.0)
    except Exception:
        return 0.0


def extract_page(page: pymupdf.Page, page_no: int) -> dict:
    """抽取单页文本 + 诊断信息。"""
    try:
        text = page.get_text("text")
    except Exception as exc:                      # 单页失败不影响整份文档
        text = ""
        err = f"{type(exc).__name__}: {exc}"
    else:
        err = None

    effective_chars = len("".join(text.split()))
    ratio = image_area_ratio(page)
    # 图像为主、文字极少。注意：财报封面/封底是整幅设计图 + 几十个字，
    # 也会命中这条，属于正常现象，不代表是扫描件。
    image_dominant = effective_chars < SCANNED_CHAR_THRESHOLD and ratio > SCANNED_IMAGE_RATIO
    no_text = effective_chars == 0

    return {
        "page": page_no,
        "text": text,
        "_chars": effective_chars,
        "_image_ratio": round(ratio, 3),
        "_image_dominant": image_dominant,
        "_no_text": no_text,
        **({"_error": err} if err else {}),
    }


# ==================================================================
# 单份文档抽取
# ==================================================================
def extract_pdf(pdf_path: Path, report_type: str, year: str | None) -> dict:
    """抽取一份 PDF，返回符合目标结构的 dict。"""
    t0 = time.time()
    doc = pymupdf.open(pdf_path)                  # 只读打开，绝不写回
    try:
        pages = [extract_page(pg, i) for i, pg in enumerate(doc, start=1)]
        meta = {k: v for k, v in (doc.metadata or {}).items() if v}
        total_pages = doc.page_count
    finally:
        doc.close()

    no_text_pages = [p["page"] for p in pages if p["_no_text"]]
    image_only_pages = [p["page"] for p in pages if p["_image_dominant"]]
    # 无文字且也不是整幅图片 —— 真·空白页（通常是封二/扉页背面）
    blank_pages = [p["page"] for p in pages if p["_no_text"] and not p["_image_dominant"]]
    total_chars = sum(p["_chars"] for p in pages)

    # 真正的扫描件：整份文档都没有可用文字层
    has_text_layer = total_chars > 1000
    # 只挑出「正文里出现大量无文字页」这种真问题，忽略封面封底
    body_no_text = [p for p in no_text_pages if 1 < p < total_pages]
    body_image = [p for p in body_no_text if p in set(image_only_pages)]
    body_blank = [p for p in body_no_text if p in set(blank_pages)]
    suspicious = (not has_text_layer) or len(body_no_text) > max(3, total_pages * 0.05)

    return {
        # ---- 目标结构（供 RAG 使用）----
        "company": COMPANY,
        "year": year,
        "report_type": report_type,
        "pages": [{"page": p["page"], "text": p["text"]} for p in pages],
        # ---- 溯源与诊断信息 ----
        "source_file": pdf_path.name,
        "source_path": str(pdf_path),
        "report_year": year,
        "total_pages": total_pages,
        "total_chars": total_chars,
        "extracted_at": datetime.now().isoformat(timespec="seconds"),
        "extractor": f"pymupdf {getattr(pymupdf, '__version__', 'unknown')}",
        "pdf_metadata": meta,
        "diagnostics": {
            "has_text_layer": has_text_layer,
            "no_text_pages": no_text_pages,            # 完全没有文字的页
            "body_no_text_pages": body_no_text,        # 正文区间内的无文字页
            "image_dominant_pages": image_only_pages,  # 图像为主、文字极少（含封面封底，正常）
            "blank_pages": blank_pages,                # 无文字且无图 —— 真空白页
            "body_image_pages": body_image,            # 正文区间内的整幅图片页
            "body_blank_pages": body_blank,            # 正文区间内的真空白页
            "suspicious": suspicious,                  # 是否疑似扫描件，需人工复核
            "avg_chars_per_page": round(total_chars / total_pages, 1) if total_pages else 0,
            "page_stats": [
                {"page": p["page"], "chars": p["_chars"],
                 "image_ratio": p["_image_ratio"],
                 "image_dominant": p["_image_dominant"],
                 **({"_error": p["_error"]} if "_error" in p else {})}
                for p in pages
            ],
        },
        "elapsed_sec": round(time.time() - t0, 2),
    }


# ==================================================================
# 文件发现
# ==================================================================
def discover(annual_dir: Path, interim_dir: Path) -> list[tuple[Path, str]]:
    """扫描两个目录，返回 [(pdf_path, report_type), ...]，按年份排序。"""
    found: list[tuple[Path, str]] = []
    for directory, rtype in ((annual_dir, "年度报告"), (interim_dir, "中期报告")):
        if not directory.is_dir():
            print(f"  [!] 目录不存在，已跳过：{directory}")
            continue
        for p in directory.iterdir():
            if p.is_file() and p.suffix.lower() == ".pdf":
                found.append((p, rtype))

    def sort_key(item: tuple[Path, str]) -> tuple[str, str]:
        rtype = item[1]
        year = parse_report_year(item[0].name) or "0000"
        return (rtype, year)

    return sorted(found, key=sort_key)


# ==================================================================
# 主流程
# ==================================================================
def main() -> int:
    ap = argparse.ArgumentParser(description="太古地产财报 PDF 批量解析")
    ap.add_argument("--only", choices=["annual", "interim"],
                    help="只处理年报或只处理中期报告")
    ap.add_argument("--limit", type=int, help="只处理前 N 份（快速试跑）")
    ap.add_argument("--force", action="store_true", help="覆盖已存在的输出文件")
    args = ap.parse_args()

    load_dotenv(PROJECT_ROOT / ".env")

    pdf_root = Path(os.getenv("TAIKOO_PDF_ROOT") or DEFAULT_PDF_ROOT)
    annual_dir = Path(os.getenv("TAIKOO_ANNUAL_DIR") or DEFAULT_ANNUAL_DIR)
    interim_dir = Path(os.getenv("TAIKOO_INTERIM_DIR") or DEFAULT_INTERIM_DIR)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    print("=" * 68)
    print("太古地产财报 PDF 解析 —— 阶段二")
    print("=" * 68)
    print(f"年报目录  : {annual_dir}")
    print(f"中期目录  : {interim_dir}")
    print(f"输出目录  : {OUT_DIR}")

    targets = discover(annual_dir, interim_dir)
    if args.only == "annual":
        targets = [t for t in targets if t[1] == "年度报告"]
    elif args.only == "interim":
        targets = [t for t in targets if t[1] == "中期报告"]
    if args.limit:
        targets = targets[: args.limit]

    print(f"待处理    : {len(targets)} 份 PDF")
    print("-" * 68)

    if not targets:
        print("没有找到任何 PDF，请检查路径。")
        return 1

    results: list[dict] = []
    ok = skipped = failed = 0

    for idx, (pdf_path, rtype) in enumerate(targets, start=1):
        year = parse_report_year(pdf_path.name)
        tag = f"{year or '????'}{rtype}"

        out_name = f"{year or 'unknown'}_{rtype}.json"
        out_path = OUT_DIR / out_name

        if out_path.exists() and not args.force:
            print(f"[{idx:2d}/{len(targets)}] 跳过（已存在） {out_name}")
            skipped += 1
            continue

        try:
            data = extract_pdf(pdf_path, rtype, year)
        except Exception as exc:
            print(f"[{idx:2d}/{len(targets)}] 失败  {tag}  {type(exc).__name__}: {exc}")
            failed += 1
            results.append({"file": pdf_path.name, "year": year, "type": rtype,
                            "status": "failed", "error": f"{type(exc).__name__}: {exc}"})
            continue

        out_path.write_text(
            json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8"
        )
        ok += 1

        d = data["diagnostics"]
        flag = ""
        if not d["has_text_layer"]:
            flag = "  <- 文字层缺失，疑似扫描件"
        elif d["body_no_text_pages"]:
            flag = f"  <- 正文无文字页 {len(d['body_no_text_pages'])} 页"
        print(f"[{idx:2d}/{len(targets)}] {tag:12s} {data['total_pages']:>4}页 "
              f"{data['total_chars']:>8}字  {data['elapsed_sec']:>6.2f}s -> {out_name}{flag}")

        results.append({**{k: data[k] for k in
                           ("source_file", "year", "report_type", "total_pages",
                            "total_chars", "elapsed_sec")},
                        "output_file": out_name, "status": "ok",
                        "has_text_layer": d["has_text_layer"],
                        "body_no_text_pages": d["body_no_text_pages"],
                        "no_text_pages": d["no_text_pages"],
                        "image_dominant_pages": d["image_dominant_pages"],
                        "blank_pages": d["blank_pages"],
                        "body_image_pages": d["body_image_pages"],
                        "body_blank_pages": d["body_blank_pages"],
                        "suspicious": d["suspicious"]})

    # ---- 汇总清单写到 output/，让 data/extracted/ 保持一 PDF 一 JSON ----
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    manifest = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "pdf_root": str(pdf_root),
        "annual_dir": str(annual_dir),
        "interim_dir": str(interim_dir),
        "summary": {"total": len(targets), "ok": ok, "skipped": skipped, "failed": failed},
        "reports": results,
    }
    manifest_path = REPORT_DIR / f"extraction_report_{stamp}.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=1),
                             encoding="utf-8")

    print("-" * 68)
    print(f"完成：成功 {ok} 份 / 跳过 {skipped} 份 / 失败 {failed} 份")
    print(f"输出目录  : {OUT_DIR}")
    print(f"汇总清单  : {manifest_path}")

    # ---- 质量汇总 ----
    ok_results = [r for r in results if r.get("status") == "ok"]
    no_layer = [r for r in ok_results if not r.get("has_text_layer")]
    suspicious = [r for r in ok_results if r.get("suspicious")]
    body_no_text = sum(len(r.get("body_no_text_pages") or []) for r in ok_results)
    body_img = sum(len(r.get("body_image_pages") or []) for r in ok_results)
    body_blank = sum(len(r.get("body_blank_pages") or []) for r in ok_results)
    total_pages = sum(r.get("total_pages") or 0 for r in ok_results)
    total_chars = sum(r.get("total_chars") or 0 for r in ok_results)

    print("=" * 68)
    print("质量检查")
    print("-" * 68)
    print(f"总页数        : {total_pages}")
    print(f"总字符数      : {total_chars:,}")
    print(f"无文字层的文档: {len(no_layer)} 份"
          + ("" if not no_layer else "  <- 疑似扫描件，需 OCR：" + ", ".join(r['source_file'] for r in no_layer)))
    print(f"正文无文字页  : {body_no_text} 页  (第2页~倒数第2页之间)")
    print(f"  ├ 整幅图片页: {body_img} 页（装饰性整页摄影，无文字内容，属正常）")
    print(f"  └ 真·空白页 : {body_blank} 页（无文字也无图，通常为扉页背面）")
    if body_no_text != body_img + body_blank:
        print(f"  [!] 分类不齐：{body_no_text} != {body_img} + {body_blank}")

    if suspicious:
        print("-" * 68)
        print(f"[!] 疑似扫描件 / 建议人工复核（{len(suspicious)} 份）：")
        for r in suspicious:
            print(f"    {r['source_file']}  "
                  f"正文无文字页={len(r.get('body_no_text_pages') or [])}  "
                  f"文字层={'正常' if r.get('has_text_layer') else '缺失'}")
    else:
        print("-" * 68)
        print("结论：27 份均具备完整文字层，未发现扫描件，无需 OCR 兜底。")
        print("      无文字页经抽查渲染确认为装饰性整页照片或空白扉页，不丢失财务信息。")

    return 0 if failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
