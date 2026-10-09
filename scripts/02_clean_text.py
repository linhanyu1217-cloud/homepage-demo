# -*- coding: utf-8 -*-
"""
阶段三：文本清洗与结构优化
==========================
把 data/extracted/ 的逐页 JSON 清洗、切块，输出适合 RAG 检索的 data/chunks/。

设计要点
--------
1. **清洗只用显式规则，绝不按「高频重复」通杀。**
   因为财报里最高频的行恰恰是最关键的内容：
   `港幣百萬元`(单位)、`2025`/`2024`(年份表头)、`收入`/`投資物業`(指标名)。
   任何「出现 N 次以上就删」的策略都会摧毁知识库。

2. **表格补上下文。** 阶段二确认 get_text("text") 丢掉了行列结构，
   表头（年份、单位）与数据行分离。这里检出表格块后，
   把该页的年份 + 单位作为上下文前缀注入，保证「指标名 + 年份 + 数值」可读。

3. **不切断表格。** 表格块作为原子单元参与拼块；
   超长表格按行边界拆分，且每片都重复表头上下文。

4. 不调用 API、不做 embedding、不建向量库。

用法
----
    venv\\Scripts\\python.exe scripts\\02_clean_text.py
    venv\\Scripts\\python.exe scripts\\02_clean_text.py --only 2025
    venv\\Scripts\\python.exe scripts\\02_clean_text.py --force
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

PROJECT_ROOT = Path(__file__).resolve().parent.parent
EXTRACTED_DIR = PROJECT_ROOT / "data" / "extracted"
CHUNKS_DIR = PROJECT_ROOT / "data" / "chunks"
REPORT_DIR = PROJECT_ROOT / "output"

COMPANY_FULL = "太古地产"
COMPANY_TRAD = "太古地產"

# ---- 切块参数 ----
TARGET_CHARS = 1000          # 期望块长
MAX_CHARS = 1200             # 硬上限（表格块可例外）
MIN_CHARS = 300              # 低于此值的尾块尝试并入前块
TABLE_SPLIT_AT = 2000        # 单个表格超过此长度才按行拆分

# ==================================================================
# 正则
# ==================================================================
NUM_RE = re.compile(r"\(?-?\d[\d,]*(?:\.\d+)?\)?%?")
YEAR_RE = re.compile(r"(?:19|20)\d{2}")
PAGE_NUM_RE = re.compile(r"^\d{1,4}$")
URL_RE = re.compile(r"^(?:https?://|www\.)\S+$", re.I)
DOT_LEADER_RE = re.compile(r"[.·・‥…]{4,}")
# 句读符号：章节标题里不会出现，用于把「含结构词的普通句子」挡在章节名之外
SENT_PUNCT_RE = re.compile(r"[。！？；，、：]")

# 中文数字年份：2016 年及以前的表格常用「二零二四年」作表头，
# 只认阿拉伯数字会漏掉这些报头，导致年份上下文注入失败。
CN_DIGIT = {"零": 0, "〇": 0, "一": 1, "二": 2, "兩": 2, "两": 2, "三": 3,
            "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
CN_YEAR_RE = re.compile(r"([零〇一二三四五六七八九]{4})\s*年")


def cn_year_to_int(s: str) -> str | None:
    try:
        return str(int("".join(str(CN_DIGIT[c]) for c in s)))
    except (KeyError, ValueError):
        return None


def find_years(text: str) -> list[str]:
    """抽取文本中的年份，阿拉伯数字与中文数字都支持。"""
    out = []
    for y in YEAR_RE.findall(text):
        if y not in out:
            out.append(y)
    for cn in CN_YEAR_RE.findall(text):
        y = cn_year_to_int(cn)
        if y and y not in out:
            out.append(y)
    return out

# 表格单位标记 —— 必须保留
UNIT_TOKENS = [
    "港幣百萬元", "港幣千元", "港幣十億元", "百萬港元", "千港元", "十億港元",
    "港元", "百萬元", "千元", "美元", "人民幣", "百萬美元", "千美元",
]

# 章节标记：只用**足够长且有辨识度**的名字，
# 刻意排除 '投資物業'/'收入'/'酒店' 等会与表格行标签冲突的词。
SECTION_NAMES = {
    "主席報告", "行政總裁報告", "業務評述", "管理層論述及分析",
    "企業管治", "企業管治及可持續發展", "可持續發展", "董事會報告",
    "獨立核數師報告", "財務報表附註", "綜合收益表", "綜合財務狀況表",
    "綜合現金流量表", "綜合權益變動表", "會計政策", "五年財務摘要",
    "風險管理", "環境、社會及管治",
}

# 泛化识别：短标题里含这些结构性词，才认作章节名。
# 关键约束是**排除单位行**——因为 '港幣百萬元' 也常年出现在页首/页尾位置，
# 一旦被当成章节标记删掉，表格单位就没了。
SECTION_HINT_WORDS = (
    "報告", "論述", "分析", "管治", "財務", "報表", "附註", "政策",
    "可持續", "風險", "董事", "摘要", "概覽", "回顧", "展望", "評述",
    "核數", "收益表", "狀況表", "現金流量", "權益變動", "物業組合",
)

# 2017-2020、2023 的报告正文与页眉本身就是简体，繁体词表匹配不到，
# 导致这些报告的 section 大量为空（158 块）、且简体页眉被误当正文留下。
# 这里显式补上上表的简体形态（= 原表经 opencc t2s 转换的结果）。
# 用显式字面量而不是运行时调 opencc：结果完全一致，但不给切块流程引入依赖。
# 繁体条目一个不动 —— 只做并集，不减项。
SECTION_NAMES |= {
    "主席报告", "行政总裁报告", "业务评述", "管理层论述及分析",
    "企业管治", "企业管治及可持续发展", "可持续发展", "董事会报告",
    "独立核数师报告", "财务报表附注", "综合收益表", "综合财务状况表",
    "综合现金流量表", "综合权益变动表", "会计政策", "五年财务摘要",
    "风险管理", "环境、社会及管治",
}

SECTION_HINT_WORDS += (
    "报告", "论述", "财务", "报表", "附注", "可持续", "风险",
    "概览", "回顾", "评述", "核数", "状况表", "现金流量",
    "权益变动", "物业组合",
)

# 报表类型关键词（按优先级从高到低）
STATEMENT_RULES = [
    ("現金流量表", ["綜合現金流量表", "經營活動", "投資活動", "融資活動",
                    "現金及現金等價物", "現金流入", "現金流出"]),
    ("資產負債表", ["綜合財務狀況表", "資產淨值", "負債總額", "權益總額",
                    "流動資產", "非流動資產", "流動負債", "投資物業"]),
    ("收入表", ["綜合收益表", "收入總額", "總收入", "租金收入", "營業溢利",
                "收入", "酒店"]),
    ("盈利表", ["每股盈利", "股東應佔", "溢利", "虧損", "股息"]),
]

COPYRIGHT_TOKENS = ["版權所有", "不得轉載", "©", "All rights reserved",
                    "保留所有權利"]

# ==================================================================
# 繁 -> 简 转换
# ==================================================================
try:
    from opencc import OpenCC
    _cc_t2s = OpenCC("t2s")
except Exception:                            # 未安装则跳过转换，流程照常
    _cc_t2s = None

# 专有名词保护表。
# opencc 不认识公司名，会按「乾(gān)->干」的默认规则把带「乾(qián)」的名字转错。
# 下面两个是语料里实际出现过的公司名，简体写法仍作「乾」。
# 需要扩充时，往这里加词即可。
PROTECTED_TERMS = ("乾林", "乾豪")


def to_simplified(text: str) -> str:
    """繁体转简体。未安装 opencc 时原样返回。

    自动检测：opencc 的 t2s 对已是简体的文本基本是恒等变换，
    所以用「转换后是否与原文不同」来判断这块是不是繁体，无需另设检测器。
    """
    if _cc_t2s is None or not text:
        return text
    out = _cc_t2s.convert(text)
    for term in PROTECTED_TERMS:
        wrong = _cc_t2s.convert(term)
        if wrong != term:
            out = out.replace(wrong, term)
    return out


def is_traditional(text: str) -> bool:
    """是否含繁体（即转换后发生了变化）。"""
    return _cc_t2s is not None and bool(text) and _cc_t2s.convert(text) != text


# ==================================================================
# 通用工具
# ==================================================================
def vis_len(s: str) -> int:
    """按「可见字符数」计长度：忽略所有空白。"""
    return len("".join(s.split()))


def norm(s: str) -> str:
    """归一化用于比对：去空白、统一全角空格。"""
    return re.sub(r"[\s　]+", "", s)


def is_blank_page(text: str) -> bool:
    return vis_len(text) == 0


def looks_like_toc(lines: list[str]) -> bool:
    """目录页：多数行带点引导线，且行数不多。"""
    if len(lines) < 4:
        return False
    with_dots = sum(1 for l in lines if DOT_LEADER_RE.search(l))
    return with_dots / len(lines) >= 0.4 and with_dots >= 4


# ==================================================================
# 页眉 / 页脚 / 页码识别
# ==================================================================
def detect_page_offset(pages: list[dict]) -> int | None:
    """推算「PDF 物理页 - 报告印刷页码」的偏移量。

    营收报告首页会把印刷页码放在页首（如 PDF 第30页首行是 '28'）。
    取众数，避免个别页（表头 '2024'）干扰。
    """
    diffs = []
    for p in pages:
        lines = [l.strip() for l in p["text"].splitlines() if l.strip()]
        for cand in lines[:3]:
            if PAGE_NUM_RE.match(cand):
                v = int(cand)
                if 1 <= v <= 999:
                    diffs.append(p["page"] - v)
                break
    if not diffs:
        return None
    return Counter(diffs).most_common(1)[0][0]


def is_printed_page_number(line: str, pdf_page: int, offset: int | None) -> bool:
    """判断某行是否为印刷页码。

    必须与推算的偏移吻合，否则 '2024'、'300,000' 这类表格数字会被误删。
    """
    s = line.strip()
    if not PAGE_NUM_RE.match(s):
        return False
    v = int(s)
    if not (1 <= v <= 999):
        return False
    if offset is None:
        # 无偏移信息时只接受明显偏小的值，宁可漏删不可错删
        return v <= pdf_page and pdf_page - v <= 5
    return abs(v - (pdf_page - offset)) <= 2


# 繁简都要，2017-2020 的报告页眉是简体「太古地产报告书」
HEADER_TYPE_WORDS = ("報告書", "报告书", "年報", "年报",
                     "中期報告", "中期报告", "年度報告", "年度报告")


def is_running_header(line: str) -> bool:
    """页眉：公司名 + 文件类型 + （可选的年份/「有限公司」）。

    实际出现过的形态：
        '太古地產報告書2025'
        '太古地產有限公司 二零一三年中期報告'
        '太古地产报告书'                <- 简体，且不含年份

    做法：把这些成分逐一剥掉，如果剩下不到 2 个字，说明整行就是页眉。
    这样既覆盖所有变体，又不会误伤含同样词组的正文句子。
    """
    s = norm(line)
    if not s or len(s) > 35:
        return False
    if COMPANY_TRAD not in s and COMPANY_FULL not in s:
        return False
    if not any(w in s for w in HEADER_TYPE_WORDS):
        return False

    rest = s
    for token in (COMPANY_TRAD, COMPANY_FULL, "有限公司"):
        rest = rest.replace(token, "")
    for w in HEADER_TYPE_WORDS:
        rest = rest.replace(w, "")
    rest = CN_YEAR_RE.sub("", rest)              # 二零一三年
    rest = YEAR_RE.sub("", rest)                 # 2013
    rest = re.sub(r"\d+", "", rest)              # 与页眉同行、被制表符分开的印刷页码
    rest = rest.strip("年 \t")
    return len(rest) <= 2


def is_section_marker(s: str) -> bool:
    """短标题是否像章节名（用于记录 section 并从正文里摘掉这个页眉页脚痕迹）。"""
    s = s.strip()
    if not s or len(s) > 20:
        return False
    if s in SECTION_NAMES:
        return True
    # 章节名不可能带句读。补上简体泛化词后，「财务」「风险」这类词会出现在
    # 页首的普通句子里（如 '页的「经审核财务资料」列表。'），若无此闸门，
    # 这些句子会被判成章节名并从正文里删掉 —— 实测会凭空产生上百个伪 section。
    if SENT_PUNCT_RE.search(s):
        return False
    # 跨行标题的碎片，如 '管理層論述' + '及分析' 里的 '及分析'
    if s[0] in "及和與暨的":
        return False
    if any(u in s for u in UNIT_TOKENS):      # 保护单位行
        return False
    if NUM_RE.search(s):
        return False
    return any(w in s for w in SECTION_HINT_WORDS)


def is_droppable_noise(line: str) -> bool:
    """纯噪音行：网址、版权声明。"""
    s = line.strip()
    if not s or len(s) > 120:
        return False
    if URL_RE.match(s):
        return True
    if any(t in s for t in COPYRIGHT_TOKENS) and len(s) <= 60:
        return True
    return False


# ==================================================================
# 表格识别
# ==================================================================
def is_table_row(line: str) -> bool:
    """判断一行是否像表格数据行：短、数字多、标签短。"""
    s = line.strip()
    if not s or len(s) > 40:
        return False
    nums = NUM_RE.findall(s)
    if not nums:
        return False
    label = NUM_RE.sub("", s).strip()
    digits = sum(len(n) for n in nums)
    density = digits / max(len(s), 1)
    # 至少两个数 + 标签极短；或数字密度很高 + 标签短
    if len(nums) >= 2 and len(label) <= 20:
        return True
    if len(nums) >= 3 and density >= 0.45:
        return True
    if len(nums) == 1 and len(label) <= 8 and density >= 0.4:
        return True
    return False


LABEL_MAX_LEN = 25


def is_label_line(line: str) -> bool:
    """表格里的行标签 / 单位行，如 '住宅'、'港幣百萬元'、'營業溢利／'。

    这些行没有数字，但夹在数据行之间，属于表格的一部分。
    若不承认它们，一张完整的财务报表会被切成好几段。
    """
    s = line.strip()
    if not s or len(s) > LABEL_MAX_LEN:
        return False
    if NUM_RE.search(s):
        return False
    # 带句读的短句是叙述文字，不是表格标签
    if re.search(r"[。！？；，,]", s):
        return False
    return True


def line_kind(line: str) -> str:
    s = line.strip()
    if not s:
        return "blank"
    if is_table_row(s):
        return "num"
    if is_label_line(s):
        return "label"
    return "other"


def find_table_blocks(lines: list[str]) -> list[tuple[int, int]]:
    """找出表格区间 [start, end)。

    把「数字行 + 标签行 + 空行」的连续段视为候选表格，
    再用数字行占比筛掉普通段落。结尾收在最后一个数字行，
    避免把表格后面的脚注、正文拖进来。
    """
    kinds = [line_kind(l) for l in lines]
    blocks: list[tuple[int, int]] = []
    i, n = 0, len(lines)

    while i < n:
        if kinds[i] in ("num", "label"):
            j = i
            num_lines = 0
            last_num = -1
            while j < n and kinds[j] in ("num", "label", "blank"):
                if kinds[j] == "num":
                    num_lines += 1
                    last_num = j
                j += 1
            end = last_num + 1 if last_num >= 0 else i
            span = end - i
            # 至少 2 行数字，且数字行占比 >= 25%，才算表格
            if num_lines >= 2 and span > 0 and num_lines / span >= 0.25:
                blocks.append((i, end))
                i = max(end, i + 1)
                continue
            i = max(j, i + 1)
        else:
            i += 1
    return blocks


def build_table_context(lines: list[str], start: int, end: int, page_no: int,
                        report_year: str | None = None) -> str:
    """为表格块构造上下文前缀：年份 + 单位。

    优先取表格块上方最近的若干行（表头通常紧邻表格），
    找不到就回退到整页范围。

    report_year 是报告期年份。表头附近常混进与表格无关的年份——最典型的是
    封面/页脚的「股份代号：1972」、法规文号、以及上一年度报告书标题里的年份，
    注进上下文会把「2013 年的表格」标成「年份=1972、2013」，检索时按年份
    过滤就会错召。故只保留 [report_year-10, report_year] 区间内的年份：
    下限取到十年前，是为了不误杀「五年財務摘要」「十年財務概要」这类
    本就回看多年的表。传 None 时不做过滤（与加此参数前的行为一致）。
    """
    lo = hi = None
    if report_year:
        try:
            hi = int(str(report_year).strip())
            lo = hi - 10
        except ValueError:
            lo = hi = None

    def scan(chunk_lines: list[str]) -> tuple[list[str], list[str]]:
        years, units = [], []
        for l in chunk_lines:
            for y in find_years(l):
                if lo is not None and not (lo <= int(y) <= hi):
                    continue
                if y not in years:
                    years.append(y)
            for u in UNIT_TOKENS:
                if u in l and u not in units:
                    units.append(u)
                    break
        return years, units

    above = lines[max(0, start - 12):start]
    years, units = scan(above)
    if not years or not units:
        y2, u2 = scan(lines)
        years = years or y2
        units = units or u2

    parts = []
    if years:
        parts.append("年份=" + "、".join(years[:4]))
    if units:
        # 取最长匹配，避免「港幣百萬元」和「百萬元」同时命中
        units.sort(key=len, reverse=True)
        parts.append("單位=" + units[0])
    if not parts:
        return ""
    return f"【表格上下文｜{'｜'.join(parts)}｜PDF第{page_no}頁】"


def detect_statement_type(text: str) -> str | None:
    """识别是否为四大报表之一。"""
    best, best_score = None, 0
    for name, kws in STATEMENT_RULES:
        score = sum(1 for k in kws if k in text)
        if score > best_score:
            best, best_score = name, score
    return best if best_score >= 2 else None


# ==================================================================
# 单页清洗
# ==================================================================
def clean_page(page: dict, offset: int | None, year: str | None,
               cur_section: str | None) -> tuple[list[str], str | None, str]:
    """清洗一页，返回 (保留行, 更新后的章节, 状态标记)。"""
    raw_lines = [l.rstrip() for l in page["text"].splitlines()]
    pdf_page = page["page"]

    # 整页空 / 纯图片：阶段二已标注
    if is_blank_page(page["text"]):
        return [], cur_section, "skip-blank"

    # 目录页
    nonempty = [l for l in raw_lines if l.strip()]
    if looks_like_toc(nonempty):
        return [], cur_section, "skip-toc"
    # 有的报告目录不用点引导线，只在页首写「目錄」且整页很短
    head = norm(nonempty[0]) if nonempty else ""
    if head in ("目錄", "目次", "內容") and vis_len(page["text"]) < 400:
        return [], cur_section, "skip-toc"

    kept: list[str] = []
    for idx, line in enumerate(raw_lines):
        s = line.strip()

        # 保留空行作为段落分隔（压缩连续空行）
        if not s:
            if kept and kept[-1] != "":
                kept.append("")
            continue

        # 印刷页码（仅在页首/页尾附近判断）
        if idx < 3 or idx >= len(raw_lines) - 3:
            if is_printed_page_number(s, pdf_page, offset):
                continue

        # 页眉
        if is_running_header(s):
            continue

        # 网址 / 版权
        if is_droppable_noise(s):
            continue

        # 章节页眉页脚标记（仅限首尾两行，避免删掉正文里的同名内容）
        pos = [i for i, x in enumerate(nonempty) if x.strip() == s]
        if pos and (pos[0] < 2 or pos[0] >= len(nonempty) - 2) and is_section_marker(s):
            cur_section = s
            continue

        # 点引导线残留（目录式行）单独一行则丢
        if DOT_LEADER_RE.search(s) and len(DOT_LEADER_RE.sub("", s).strip()) <= 4:
            continue

        kept.append(s)

    while kept and kept[-1] == "":
        kept.pop()
    return kept, cur_section, "ok"


# ==================================================================
# 切块
# ==================================================================
def split_oversized_table(rows: list[str], header: str) -> list[str]:
    """超长表格按行边界拆分，每片重复表头上下文。"""
    pieces, cur, cur_len = [], [], vis_len(header)
    for r in rows:
        rl = vis_len(r)
        if cur and cur_len + rl > MAX_CHARS:
            pieces.append(header + "\n" + "\n".join(cur))
            cur, cur_len = [], vis_len(header)
        cur.append(r)
        cur_len += rl
    if cur:
        pieces.append(header + "\n" + "\n".join(cur))
    return pieces


def split_oversized_text(block: str) -> list[str]:
    """超长正文块按行边界拆分。"""
    lines = block.splitlines()
    pieces, cur, cur_len = [], [], 0
    for l in lines:
        ll = vis_len(l)
        if cur and cur_len + ll > MAX_CHARS:
            pieces.append("\n".join(cur))
            cur, cur_len = [], 0
        cur.append(l)
        cur_len += ll
    if cur:
        pieces.append("\n".join(cur))
    return pieces


def build_segments(pages: list[dict], year: str | None) -> list[dict]:
    """把逐页文本转成「段」序列，每段带页码/章节/类型信息。"""
    offset = detect_page_offset(pages)
    segments: list[dict] = []
    cur_section: str | None = None
    stats = defaultdict(int)

    for page in pages:
        pdf_page = page["page"]
        lines, cur_section, status = clean_page(page, offset, year, cur_section)
        stats[status] += 1

        if not lines:
            continue

        page_text = "\n".join(lines)
        stype = detect_statement_type(page_text)
        blocks = find_table_blocks(lines)

        cursor = 0
        for (bstart, bend) in blocks:
            # 表格前的正文
            pre = [l for l in lines[cursor:bstart] if l.strip()]
            if pre:
                segments.append({"page": pdf_page, "kind": "text",
                                 "text": "\n".join(pre),
                                 "section": cur_section, "statement": stype})
            # 表格块（含上下文前缀）
            rows = [l for l in lines[bstart:bend]]
            header = build_table_context(lines, bstart, bend, pdf_page, year)
            body = "\n".join(rows)
            text = (header + "\n" + body) if header else body
            segments.append({"page": pdf_page, "kind": "table", "text": text,
                             "header": header, "rows": rows,
                             "section": cur_section, "statement": stype})
            stats["table-blocks"] += 1
            cursor = bend

        tail = [l for l in lines[cursor:] if l.strip()]
        if tail:
            segments.append({"page": pdf_page, "kind": "text",
                             "text": "\n".join(tail),
                             "section": cur_section, "statement": stype})

    return segments, stats, offset


def assemble_chunks(segments: list[dict]) -> list[dict]:
    """把段拼成 800~1200 字的块；表格不切断；章节切换时断块。"""
    chunks: list[dict] = []
    cur: list[dict] = []
    cur_len = 0
    cur_section = None

    def flush():
        nonlocal cur, cur_len, cur_section
        if not cur:
            return
        chunks.append(_render(cur))
        cur, cur_len = [], 0

    for seg in segments:
        # 章节变化 -> 断块，保持块内主题一致
        if cur and seg["section"] != cur_section:
            flush()
        cur_section = seg["section"]

        seg_len = vis_len(seg["text"])

        # 超长段：先落掉已积累的，再单独处理
        if seg_len > MAX_CHARS:
            if seg["kind"] == "table":
                if seg_len > TABLE_SPLIT_AT:
                    flush()
                    for piece in split_oversized_table(seg["rows"], seg["header"]):
                        chunks.append(_render([{**seg, "text": piece}]))
                    continue
                # 表格未超 TABLE_SPLIT_AT：整块保留，允许超过 MAX
                flush()
                chunks.append(_render([seg]))
                continue
            flush()
            for piece in split_oversized_text(seg["text"]):
                chunks.append(_render([{**seg, "text": piece}]))
            continue

        if cur and cur_len + seg_len > MAX_CHARS and cur_len >= MIN_CHARS:
            flush()
        # 目标值附近主动断开，避免块长都挤在上限
        elif cur and cur_len + seg_len > TARGET_CHARS and cur_len >= MIN_CHARS:
            flush()

        cur.append(seg)
        cur_len += seg_len

    flush()

    # 过小的块（多为章节尾部的零星段落）并入前一块，
    # 前提是同章节、且合并后不超上限。
    merged_chunks: list[dict] = []
    for c in chunks:
        if (merged_chunks and c["char_count"] < MIN_CHARS
                and merged_chunks[-1]["section"] == c["section"]
                and merged_chunks[-1]["char_count"] + c["char_count"] <= MAX_CHARS):
            merged_chunks[-1] = _merge_chunks(merged_chunks[-1], c)
        else:
            merged_chunks.append(c)
    return merged_chunks


def _render(segs: list[dict]) -> dict:
    """把若干段渲染成一个 chunk 字典。"""
    parts = []
    pages = []
    last_page = None
    for s in segs:
        if s["page"] not in pages:
            pages.append(s["page"])
        # 页码变化处插入定位标记，方便 LLM 引用
        if last_page is not None and s["page"] != last_page:
            parts.append(f"【第{s['page']}頁】")
        parts.append(s["text"])
        last_page = s["page"]

    # 跨页块在开头也补一个页码标记，
    # 否则首段文字没有归属页，引用时只能退而用 pdf_page（=首页），会指错页。
    if len(pages) > 1:
        parts.insert(0, f"【第{pages[0]}頁】")

    text = "\n".join(parts).strip()
    stype = next((s["statement"] for s in segs if s.get("statement")), None)
    is_table = any(s["kind"] == "table" for s in segs)
    return {
        "text": text,
        "pdf_page": pages[0],
        "pdf_pages": pages,
        "section": segs[0]["section"],
        "statement_type": stype,
        "is_table": is_table,
        "char_count": vis_len(text),
        "_segs": segs,
    }


def _merge_chunks(a: dict, b: dict) -> dict:
    """把 b 并进 a（重新渲染以保持页码标记正确）。"""
    merged = _render(a["_segs"] + b["_segs"])
    return merged


# ==================================================================
# 输出
# ==================================================================
def finalize(chunk: dict, meta: dict, idx: int, year: str, rtype: str,
             source_file: str) -> dict:
    """繁转简，并补齐 RAG 需要的 metadata。

    转换只在这里发生 —— 切块早已完成，所以块的数量与边界不受转换影响。
    source_file 是磁盘上的真实文件名，保持繁体原样不转换。
    """
    text = to_simplified(chunk["text"])
    section = to_simplified(chunk["section"]) if chunk["section"] else chunk["section"]
    return {
        "chunk_id": f"{year}_{rtype}_p{chunk['pdf_page']:04d}_c{idx:02d}",
        "text": text,
        # ---- 必需 metadata ----
        "company": COMPANY_FULL,
        "year": year,
        "report_type": rtype,
        "pdf_page": chunk["pdf_page"],
        "source_file": source_file,
        # ---- 附加溯源与检索信息 ----
        "pdf_pages": chunk["pdf_pages"],
        "printed_page": meta.get("page_offset"),
        "section": section,
        "statement_type": chunk["statement_type"],
        "is_table": chunk["is_table"],
        "char_count": vis_len(text),
        "_converted": text != chunk["text"],
    }


# ==================================================================
# 主流程
# ==================================================================
def main() -> int:
    ap = argparse.ArgumentParser(description="太古地产财报文本清洗与切块")
    ap.add_argument("--only", help="只处理指定年份，如 2025")
    ap.add_argument("--force", action="store_true", help="覆盖已存在输出")
    args = ap.parse_args()

    CHUNKS_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    files = sorted(EXTRACTED_DIR.glob("*.json"))
    if args.only:
        files = [f for f in files if args.only in f.stem]
    if not files:
        print(f"未找到输入文件：{EXTRACTED_DIR}")
        return 1

    print("=" * 70)
    print("太古地产财报 文本清洗与切块 —— 阶段三")
    print("=" * 70)
    print(f"输入: {EXTRACTED_DIR}  ({len(files)} 份)")
    print(f"输出: {CHUNKS_DIR}")
    if _cc_t2s is None:
        print("警告: 未安装 opencc，跳过繁转简，输出保持繁体原文")
    print("-" * 70)

    all_rows: list[dict] = []
    manifest_reports: list[dict] = []
    grand = defaultdict(int)
    t0 = time.time()

    for fi, f in enumerate(files, start=1):
        src = json.loads(f.read_text(encoding="utf-8"))
        year = str(src.get("year") or f.stem.split("_")[0])
        rtype = src.get("report_type", "年度报告")
        source_file = src.get("source_file", "")

        segments, stats, offset = build_segments(src["pages"], year)
        chunks = assemble_chunks(segments)

        meta = {"page_offset": offset}
        rows = [finalize(c, meta, i, year, rtype, source_file)
                for i, c in enumerate(chunks, start=1)]

        # _converted 是内部统计字段，不能写进输出
        converted = sum(1 for r in rows if r.pop("_converted", False))
        grand["converted"] += converted

        out_path = CHUNKS_DIR / f"{year}_{rtype}.jsonl"
        with out_path.open("w", encoding="utf-8") as fh:
            for r in rows:
                fh.write(json.dumps(r, ensure_ascii=False) + "\n")

        all_rows.extend(rows)
        lens = [r["char_count"] for r in rows]
        ntbl = sum(1 for r in rows if r["is_table"])
        grand["chunks"] += len(rows)
        grand["tables"] += ntbl
        grand["chars"] += sum(lens)
        grand["pages"] += len(src["pages"])
        grand["skip-blank"] += stats.get("skip-blank", 0)
        grand["skip-toc"] += stats.get("skip-toc", 0)
        grand["table-blocks"] += stats.get("table-blocks", 0)

        over = sum(1 for c in lens if c > MAX_CHARS)
        print(f"[{fi:2d}/{len(files)}] {year} {rtype}  "
              f"{len(src['pages']):>3}页 -> {len(rows):>3}块  "
              f"均{sum(lens)//max(len(lens),1):>4}字  表格块{ntbl:>3}  "
              f"偏移{offset if offset is not None else '?'}  "
              f"超限{over}")

        manifest_reports.append({
            "year": year, "report_type": rtype, "source_file": source_file,
            "pages": len(src["pages"]), "chunks": len(rows),
            "avg_chars": round(sum(lens) / max(len(lens), 1), 1),
            "min_chars": min(lens) if lens else 0,
            "max_chars": max(lens) if lens else 0,
            "table_chunks": ntbl,
            "oversized_chunks": over,
            "page_offset": offset,
            "converted_chunks": converted,
            "output": out_path.name,
        })

    # ---- 合并流（阶段五向量化直接吃这个）----
    combined = CHUNKS_DIR / "all_chunks.jsonl"
    with combined.open("w", encoding="utf-8") as fh:
        for r in all_rows:
            fh.write(json.dumps(r, ensure_ascii=False) + "\n")

    lens_all = [r["char_count"] for r in all_rows]
    n = len(all_rows)
    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "elapsed_sec": round(time.time() - t0, 2),
        "source_reports": len(files),
        "total_pages": grand["pages"],
        "total_chunks": n,
        "table_chunks": grand["tables"],
        "skipped_blank_pages": grand["skip-blank"],
        "skipped_toc_pages": grand["skip-toc"],
        "detected_table_blocks": grand["table-blocks"],
        "avg_chars": round(sum(lens_all) / max(n, 1), 1),
        "median_chars": sorted(lens_all)[n // 2] if n else 0,
        "min_chars": min(lens_all) if n else 0,
        "max_chars": max(lens_all) if n else 0,
        "under_300": sum(1 for c in lens_all if c < 300),
        "in_range_800_1200": sum(1 for c in lens_all if 800 <= c <= 1200),
        "over_1200": sum(1 for c in lens_all if c > 1200),
        "total_chars": sum(lens_all),
        "opencc_available": _cc_t2s is not None,
        "converted_chunks": grand["converted"],
        "protected_terms": list(PROTECTED_TERMS),
        "reports": manifest_reports,
    }
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    mpath = REPORT_DIR / f"chunk_report_{stamp}.json"
    mpath.write_text(json.dumps(summary, ensure_ascii=False, indent=1),
                     encoding="utf-8")

    print("-" * 70)
    print(f"完成：{len(files)} 份报告 -> {n} 个 chunk")
    print(f"总字符      : {summary['total_chars']:,}")
    print(f"平均/中位数 : {summary['avg_chars']} / {summary['median_chars']} 字")
    print(f"区间分布    : <300字 {summary['under_300']} | "
          f"800-1200字 {summary['in_range_800_1200']} | "
          f">1200字 {summary['over_1200']}")
    print(f"表格块      : {summary['table_chunks']} "
          f"(检出表格区 {summary['detected_table_blocks']})")
    print(f"跳过页      : 空白 {summary['skipped_blank_pages']} | "
          f"目录 {summary['skipped_toc_pages']}")
    if summary["opencc_available"]:
        print(f"繁转简      : {summary['converted_chunks']} 块发生转换 / 共 {n} 块"
              f"  (保护词 {len(PROTECTED_TERMS)} 个)")
    else:
        print("繁转简      : 未执行（opencc 不可用）")
    print(f"输出目录    : {CHUNKS_DIR}")
    print(f"合并流      : {combined}")
    print(f"汇总清单    : {mpath}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
