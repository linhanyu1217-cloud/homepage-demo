# -*- coding: utf-8 -*-
"""Word 周报 + PPT 页渲染（由原 生成Word.py 与 更新PPT.py 整合而来）。

对外接口
--------
    generate_word(out_dir, force=False, only=None)  -> 生成 本周公司公告周报.docx / 本周重要资讯.docx
    generate_ppt(week=None, report_dir=None, start=None, end=None,
                 entries_file=None, must_include=None,
                 max_entry_lines=3)                  -> 生成 房地产周报_*.pptx
    verify_word(out_dir)                            -> Word 产出校验

路径约定（全部相对项目根目录，靠 __file__ 推导，不写死绝对路径）
    01_模板\公司公告周报模板.docx     TEMPLATE_WORD
    01_模板\房地产周报PPT模板.pptx    TEMPLATE_PPT
    05_输出\<YYYY-Wxx>\              OUT_ROOT
    04_工作区\                        WORK_ROOT

安全边界
    · 模板原件只读：先复制再改副本，结束校验模板 sha256 未变
    · 只允许写入 04_工作区\ 与 05_输出\（见 WRITABLE_ROOTS）
    · 拒绝写入历史实习目录（HISTORY_DIR，由 agent_core 按个人配置注入；发行版为 None）
"""
import argparse
import copy
import datetime as dt
import hashlib
import os
import random
import re
import shutil
import sys
import zipfile

# 旧 更新PPT.py 里是 `from copy import deepcopy`；两个脚本合并后
# copy 与 deepcopy 都要在，所以这行必须保留。
from copy import deepcopy

try:
    from docx import Document
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import qn, nsdecls
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run
    from lxml import etree
except ImportError:
    sys.stderr.write(
        "缺少 python-docx / lxml，请先安装：python -m pip install python-docx lxml\n")
    sys.exit(2)

try:
    from pptx import Presentation
    from pptx.oxml.ns import qn as _qn_pptx
    from pptx.text.text import _Paragraph
except ImportError:
    sys.stderr.write("缺少 python-pptx，请先安装：python -m pip install python-pptx\n")
    sys.exit(2)

try:
    from . import zh_convert as _TC          # 同包内的零依赖繁简转换
except (ImportError, ValueError):
    try:
        import zh_convert as _TC
    except ImportError:                      # 模块缺失时降级为「不转换」，不打断生成
        _TC = None

# --------------------------------------------------------------- 路径（相对化）
_HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(_HERE)                       # 项目根目录

TEMPLATE_DIR = os.path.join(ROOT, "01_模板")
TEMPLATE_WORD = os.path.join(TEMPLATE_DIR, "公司公告周报模板.docx")
TEMPLATE_PPT = os.path.join(TEMPLATE_DIR, "房地产周报PPT模板.pptx")

OUT_ROOT = os.path.join(ROOT, "05_输出")
WORK_ROOT = os.path.join(ROOT, "04_工作区")

# 兼容旧正文里的命名：旧 生成Word.py 用 BASE 拼 05_输出 路径，
# 这里让它等于 ROOT，语义不变（旧脚本里 BASE 就是项目根）。
BASE = ROOT
# 允许写入的根目录。旧版是 (04_工作区, 05_输出)，新版对应 (工作区, 输出)。
WRITABLE_ROOTS = (WORK_ROOT, OUT_ROOT)

# 历史实习资料目录（个人版专有）。由 agent_core 在启动时按 配置\个人配置.json 注入；
# 发行版没有该文件，这里保持 None，安全护栏自动退化为「不检查该项」。
HISTORY_DIR = None


def log(msg=""):
    """GBK 控制台下也不会崩的输出。"""
    try:
        print(msg)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "ascii"
        print(msg.encode(enc, "replace").decode(enc, "replace"))


# ==================================================================
#  第一部分：Word（原 生成Word.py）
# ==================================================================
# (源 md, 输出 docx) —— Word 只有一份完整底稿，没有正式/备选之分
JOBS = [
    ("本周公司公告汇总.md", "本周公司公告周报.docx"),
    ("本周重要资讯.md", "本周重要资讯.docx"),
]

W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"

# "一行空行"的段前间距：260 twips = 13pt，正好是模板日标题自带的分节日距，
# 也约等于 11pt 正文的一行高（11pt × 1.158 = 12.74pt）。
GAP_TWIPS = 260

# w:pPr 里排在 w:contextualSpacing **之后**的元素——插入新元素时要用到
_PPR_AFTER_CONTEXTUAL = (
    "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection",
    "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId",
    "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)

# w:pPr 里排在 w:pageBreakBefore **之后**的元素（ECMA-376 规定的 pPr 子元素顺序）
_PPR_AFTER_PAGEBREAK = (
    "w:framePr", "w:widowControl", "w:numPr", "w:suppressLineNumbers",
    "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
    "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
    "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
    "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
    "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
    "w:rPr", "w:sectPr", "w:pPrChange",
)

# 日期块标题：先认模板的日标题写法「广发证券房地产每日新闻（2026.09.04）——李怡慧」，
# 再认「标题就是日期本身」的写法（`## 2026.09.04`）。
# 刻意**不**认「含日期但另有主题」的小标题（如 `## 2026年半年度业绩`），
# 否则普通章节会被误判成日期块而多插一个分页。
RE_DATE = re.compile(r"\d{4}\s*[.\-/年]\s*\d{1,2}\s*[.\-/月]\s*\d{1,2}\s*日?")
RE_DATE_ONLY = re.compile(
    r"^[（(]?\s*\d{4}\s*[.\-/年]\s*\d{1,2}\s*[.\-/月]\s*\d{1,2}\s*日?\s*[）)]?"
    r"\s*(?:[—\-–]{1,2}\s*\S+\s*)?$")


def is_date_head(text):
    """该标题是否开启一个新的日期块。"""
    t = (text or "").strip()
    if "每日新闻" in t and RE_DATE.search(t):
        return True
    return bool(RE_DATE_ONLY.match(t))


class WordToolError(Exception):
    """工具层拒绝执行（安全约束或调用错误）。"""





# ----------------------------------------------------------- 路径安全
def _abs(p):
    return os.path.abspath(p)


def _under(path, root):
    p = _abs(path).rstrip("\\/").lower()
    r = _abs(root).rstrip("\\/").lower()
    return p == r or p.startswith(r + os.sep)


def assert_writable_target(out):
    """输出必须落在工作区内，且不得是模板原件或历史实习目录。"""
    if _abs(out).lower() == _abs(TEMPLATE_WORD).lower():
        raise WordToolError("拒绝覆盖模板原件：%s" % out)
    if _under(out, HISTORY_DIR):
        raise WordToolError("拒绝写入历史实习目录（只读）：%s" % out)
    if not any(_under(out, r) for r in WRITABLE_ROOTS):
        raise WordToolError(
            "输出路径必须位于 04_工作区 或 05_输出 之内：%s\n  允许的根目录：\n    - %s"
            % (out, "\n    - ".join(WRITABLE_ROOTS)))
    if not out.lower().endswith(".docx"):
        raise WordToolError("输出文件必须是 .docx：%s" % out)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# ----------------------------------------------------------- markdown 解析
RE_TABLE = re.compile(r"^\s*\|.*\|\s*$")
RE_ENTRY = re.compile(r"^(\s*)-\s+(.*)$")
RE_RULE = re.compile(r"^-{3,}$")

# 「模板体」行：不带任何 markdown 标记。本周公司公告正式版.md 通篇是这种格式，
# 它本来就是照 Word 模板的样子写出来的，所以必须单独认出来，否则会被当成
# 普通段落——日期分组会丢掉 16pt 加粗、条目会丢掉 List Paragraph 和编号。
RE_PLAIN_HEAD = re.compile(
    r"^[^【].{0,20}（\s*\d{4}\s*[.\-/年]\s*\d{1,2}\s*[.\-/月]\s*\d{1,2}\s*日?\s*）")
RE_PLAIN_ENTRY = re.compile(r"^【[^】]{1,24}】.+")


def parse_md(text):
    """把 md 拆成块序列：[(kind, text)]，kind ∈
    head1/head2/head3 / news / entry / sub / quote / para / rule / table
    """
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i].rstrip()
        s = line.strip()
        if not s:
            i += 1
            continue

        if RE_TABLE.match(line):
            rows = []
            while i < n and RE_TABLE.match(lines[i].rstrip()):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                # 跳过 |---|---| 分隔行
                if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
                    rows.append(cells)
                i += 1
            blocks.append(("table", rows))
            continue

        if s.startswith("#"):
            lvl = len(s) - len(s.lstrip("#"))
            blocks.append(("head%d" % min(lvl, 3), s.lstrip("#").strip()))
            i += 1
            continue

        if s == "【公司新闻】":
            blocks.append(("news", s))
            i += 1
            continue

        if RE_RULE.match(s):
            blocks.append(("rule", ""))
            i += 1
            continue

        if s.startswith(">"):
            # 连续的 > 行在 markdown 里是**同一个**引用块，合并成一段（不插分隔符，
            # 原文件是软换行，断点落在「；」「、」上）
            parts = []
            while i < n and lines[i].strip().startswith(">"):
                parts.append(lines[i].strip().lstrip(">").strip())
                i += 1
            blocks.append(("quote", "".join(parts)))
            continue

        m = RE_ENTRY.match(line)
        if m:
            indent = len(m.group(1).expandtabs(2))
            blocks.append(("entry" if indent < 2 else "sub", m.group(2).strip()))
            i += 1
            continue

        # 缩进且无标记的行 = markdown 软换行的续行，稍后并回上一块
        if line[:1] in (" ", "\t"):
            blocks.append(("cont", s))
            i += 1
            continue

        if RE_PLAIN_HEAD.match(s):
            blocks.append(("head2", s))
            i += 1
            continue

        if RE_PLAIN_ENTRY.match(s):
            blocks.append(("entry", s))
            i += 1
            continue

        blocks.append(("para", s))
        i += 1

    # 把缩进续行并回上一块（软换行不该另起段落，更不该多占一个编号）
    merged = []
    for k, v in blocks:
        if k == "cont" and merged and merged[-1][0] != "table":
            pk, pv = merged[-1]
            merged[-1] = (pk, pv + v)
            continue
        merged.append((k, v))
    return _simplify_blocks(merged)


def _simplify_blocks(blocks):
    """最终中文输出统一转简体（港股公告原文常为繁体）。

    为什么放在 parse_md 的出口、而不是写盘处：
        本脚本的 verify() 会把 docx 里的文本与 md 原文**逐条比对**，
        而两侧都走 parse_md 解析。转换放在解析出口，docx 与 md 天然自洽；
        若只在写入侧转，verify() 会整片报 FAIL。

    只做字形转换：数字 / 金额 / 百分比 / 日期 / 股票代码 / 英文缩写一律不动。
    分层说明：`00_配置\\行业公司池规则.md` 要求公司池保留**公告原文繁体全称**，
    因此此处只作用于面向读者的最终中文输出，不回溯改写公司池与来源清单。
    """
    if _TC is None or not _TC.available():
        return blocks
    out = []
    for k, v in blocks:
        if k == "table":
            out.append((k, [[_TC.to_simplified(c) for c in row] for row in v]))
        elif isinstance(v, str) and v:
            out.append((k, _TC.to_simplified(v)))
        else:
            out.append((k, v))
    return out


def parse_inline(s):
    """把 `**x**` 拆成 [(文字, 是否加粗)]；非标记字符一个不丢。

    加粗值为 None 表示"沿用模具自带的格式"（正式版条目因此与模板同款）。
    """
    out = []
    for part in re.split(r"(\*\*.+?\*\*)", s):
        if not part:
            continue
        if len(part) > 4 and part.startswith("**") and part.endswith("**"):
            out.append((part[2:-2], True))
        else:
            out.append((part, None))
    return out


def plain_text(s):
    """去掉行内 markdown 标记后的纯文字（用于校验）。"""
    return "".join(t for t, _ in parse_inline(s))


# ----------------------------------------------------------- 模具与写盘
def _strip_para_ids(el):
    """deepcopy 出来的段落会带重复的 w14:paraId / w14:textId，去掉让 Word 重新分配。"""
    for e in el.iter():
        for a in list(e.attrib):
            if a.startswith(W14) and a[len(W14):] in ("paraId", "textId"):
                del e.attrib[a]


def set_para_text(p_el, segments):
    """把段落内容整体替换为 segments，格式沿用该段第一个 run 的 rPr。"""
    proto = p_el.find(qn("w:r"))
    for r in p_el.findall(qn("w:r")):
        p_el.remove(r)
    if proto is None:          # 空段落：没有可继承的 run，保持为空
        return
    for text, bold in segments:
        r = copy.deepcopy(proto)
        for ch in list(r):
            if ch.tag != qn("w:rPr"):
                r.remove(ch)
        run = Run(r, None)
        run.text = text
        if bold is not None:
            run.bold = bold
        p_el.append(r)


class Renderer(object):
    """用模板自带的段落当模具，往 body 里追加内容。"""

    def __init__(self, doc):
        self.doc = doc
        self.body = doc.element.body
        self.sectPr = self.body.find(qn("w:sectPr"))
        if self.sectPr is None:
            raise WordToolError("模板缺少 w:sectPr，无法安全清空正文")

        paras = doc.paragraphs
        if len(paras) < 5:
            raise WordToolError("模板段落数异常（%d），与预期结构不符" % len(paras))
        self.d_head = paras[0]._p       # 16pt 加粗日标题
        self.d_news = paras[2]._p       # 【公司新闻】
        self.d_entry = paras[4]._p      # List Paragraph 条目（带 numPr）

        # 全文第一段不施加上间距，避免正文顶部凭空多一行
        self._first = True
        self._last = None               # 最近追加的段落（给表格补上方间距用）

        # 结构自检：模具段落必须认得出来，否则宁可报错也不写坏
        if "每日新闻" not in paras[0].text:
            raise WordToolError("模板第 1 段不是日标题：%r" % paras[0].text[:40])
        if paras[2].text.strip() != "【公司新闻】":
            raise WordToolError("模板第 3 段不是【公司新闻】：%r" % paras[2].text[:40])
        if paras[4].style.name != "List Paragraph":
            raise WordToolError("模板第 5 段不是 List Paragraph：%r" % paras[4].style.name)

        self._nb = None
        self._next_abs = self._next_num = None
        self.date_blocks = 0        # 本次渲染遇到的日期块数（render 结束时回填）

    # -- 编号：每个日期分组一个新列表实例，编号从 1 重启 ---------------
    def _numbering(self):
        if self._nb is not None:
            return self._nb
        try:
            part = self.doc.part.numbering_part
        except (KeyError, NotImplementedError):
            part = None
        if part is None or part.element is None:
            raise WordToolError("模板没有 numbering.xml，无法还原编号格式")
        nb = part.element
        self._nb = nb
        self._next_abs = 1 + max(
            int(x) for x in re.findall(r'<w:abstractNum w:abstractNumId="(\d+)"',
                                       nb.xml.encode("utf-8").decode("utf-8")))
        self._next_num = 1 + max(
            int(x) for x in re.findall(r'<w:num w:numId="(\d+)"',
                                       nb.xml.encode("utf-8").decode("utf-8")))
        return nb

    def new_num_id(self):
        """注册一份新的 abstractNum + numId，返回新 numId（编号重启）。"""
        nb = self._numbering()
        src = nb.find(qn("w:abstractNum"))
        if src is None:
            raise WordToolError("numbering.xml 里没有 abstractNum 可克隆")

        aid = self._next_abs
        self._next_abs += 1

        new_abs = copy.deepcopy(src)
        new_abs.set(qn("w:abstractNumId"), str(aid))
        # nsid / tmpl 是"列表模板"指纹，必须换掉，否则 Word 会把两份定义
        # 认成同一个列表而继续编号
        for tag in ("w:nsid", "w:tmpl"):
            el = new_abs.find(qn(tag))
            if el is not None:
                el.set(qn("w:val"), "%08X" % random.randint(0, 0xFFFFFFFF))

        first_num = nb.find(qn("w:num"))
        if first_num is not None:
            first_num.addprevious(new_abs)
        else:
            nb.append(new_abs)

        nid = self._next_num
        self._next_num += 1
        new_num = OxmlElement("w:num")
        new_num.set(qn("w:numId"), str(nid))
        aid_el = OxmlElement("w:abstractNumId")
        aid_el.set(qn("w:val"), str(aid))
        new_num.append(aid_el)
        nb.append(new_num)
        return nid

    # -- 间距：所有层级统一"一行空行"，用段前间距实现，不插空段落 --------
    def _apply_gap(self, el):
        pPr = el.get_or_add_pPr()

        # List Paragraph 样式带 <w:contextualSpacing/>，会吞掉相邻同款段落之间的
        # 间距（公告会全部粘在一起）。必须在段落级显式覆写为 0。
        cs = pPr.find(qn("w:contextualSpacing"))
        if cs is None:
            cs = OxmlElement("w:contextualSpacing")
            pPr.insert_element_before(cs, *_PPR_AFTER_CONTEXTUAL)
        cs.set(qn("w:val"), "0")

        # get_or_add_spacing 会插到 w:ind / w:contextualSpacing 之前，顺序合法；
        # 已存在的 spacing（日标题的 before=260 after=260 line=416）只改 before/after，
        # line / lineRule 原样保留 → 模板行距不变。
        sp = pPr.get_or_add_spacing()
        sp.set(qn("w:before"), "0" if self._first else str(GAP_TWIPS))
        sp.set(qn("w:after"), "0")
        self._first = False
        self._last = el

    def gap_after_last(self, twips=GAP_TWIPS):
        """给上一段补**段后**间距——表格不参与段前间距，只能从上方那段补一行。"""
        if self._last is None:
            return
        sp = self._last.get_or_add_pPr().get_or_add_spacing()
        sp.set(qn("w:after"), str(twips))

    # -- 分页：日期块各自另起一页 -------------------------------------
    def set_page_break_before(self, el):
        """让该段**从新的一页顶部开始**（真正的 Word 分页，不是空行凑的）。

        用 `<w:pageBreakBefore/>`（段落属性「段前分页」）而不是手工分页符
        `<w:br w:type="page"/>`，原因有二：
          1. Word 在"本段已经位于页首"时会自动忽略它，**不会凭空多出一张空白页**；
             手工分页符落在页首时是会真的翻出一页空白的。
          2. 它只是段落属性，不新增 run、不新增段落、不产生任何可见字符，
             段落的文字 / 字体 / 编号 / 间距全部不受影响。
        两者都是 Word 原生分页，都不涉及空行——满足"必须使用真正的分页符"。
        """
        pPr = el.get_or_add_pPr()
        pbb = pPr.find(qn("w:pageBreakBefore"))
        if pbb is None:
            pbb = OxmlElement("w:pageBreakBefore")
            pPr.insert_element_before(pbb, *_PPR_AFTER_PAGEBREAK)
        # 与 Word 自己的写法一致：勾选时写裸元素，取消勾选才写 w:val="0"
        if pbb.get(qn("w:val")) is not None:
            del pbb.attrib[qn("w:val")]
        return pbb

    # -- 追加段落 -----------------------------------------------------
    def _append(self, donor, segments):
        el = copy.deepcopy(donor)
        _strip_para_ids(el)
        set_para_text(el, segments)
        self._apply_gap(el)
        self.sectPr.addprevious(el)
        return el

    def head(self, text):
        return self._append(self.d_head, parse_inline(text))

    def date_head(self, text, new_page=True):
        """日期块标题：需要时让它另起一页（第一个日期块 new_page=False）。"""
        el = self.head(text)
        if new_page:
            self.set_page_break_before(el)
        return el

    def news(self, text):
        return self._append(self.d_news, parse_inline(text))

    def entry(self, text, num_id):
        el = self._append(self.d_entry, parse_inline(text))
        pPr = el.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.get_or_add_numId().val = num_id
        numPr.get_or_add_ilvl().val = 0
        return el

    def body_para(self, text, indent=None):
        el = self._append(self.d_news, parse_inline(text))
        if indent:
            pPr = el.get_or_add_pPr()
            pPr.get_or_add_ind().set(qn("w:left"), str(indent))
        return el

    # -- 表格 ---------------------------------------------------------
    def table(self, rows):
        if not rows:
            return None
        ncols = max(len(r) for r in rows)
        t = self.doc.add_table(rows=len(rows), cols=ncols)

        tblPr = t._tbl.tblPr
        borders = parse_xml(
            '<w:tblBorders %s>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '</w:tblBorders>' % nsdecls("w"))
        tblPr.insert_element_before(
            borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
            "w:tblCaption", "w:tblDescription", "w:tblPrChange")

        for ri, row in enumerate(rows):
            for ci in range(ncols):
                cell = t.cell(ri, ci)
                txt = row[ci] if ci < len(row) else ""
                p = cell.paragraphs[0]
                run = p.add_run(txt)
                run.font.name = "Calibri"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
                if ri == 0:
                    run.bold = True
        self.sectPr.addprevious(t._tbl)
        self._last = None          # 表格后面那段自己的段前间距负责留白
        return t


# ----------------------------------------------------------- 渲染主流程
def render(doc, blocks):
    r = Renderer(doc)

    # 清空正文（保留 sectPr）——模板里的旧公告内容在这一步被清掉
    for child in list(r.body):
        if child is not r.sectPr:
            r.body.remove(child)

    # 空行一律由 Renderer._apply_gap 的段前间距产生，这里不再插入任何空段落——
    # 所有层级（日标题 / 【公司新闻】 / 条目之间 / 日块之间）因此都是同一行距。
    i = 0
    n = len(blocks)
    date_blocks = 0          # 已开始的日期块数
    while i < n:
        kind, val = blocks[i]

        if kind in ("head1", "head2", "head3", "news"):
            if kind != "news" and is_date_head(val):
                date_blocks += 1
                # 第 1 个日期块不加分页（第一页本来就在页首）；
                # 之后每进入一个新日期，都让它从新的一页顶部开始，
                # 因此新日期标题永远不会落在上一页末尾。
                r.date_head(val, new_page=(date_blocks > 1))
            elif kind == "news":
                r.news(val)
            else:
                r.head(val)
            i += 1
            continue

        if kind == "entry":
            group = []
            while i < n and blocks[i][0] in ("entry", "sub"):
                group.append(blocks[i])
                i += 1
            num_id = r.new_num_id()          # 本组编号从 1 重启
            for k, t in group:
                if k == "entry":
                    r.entry(t, num_id)
                else:
                    r.body_para(t, indent=1080)   # 子项：缩进、不占编号
            continue

        if kind == "table":
            r.gap_after_last()               # 表格上方补一行（见 gap_after_last）
            r.table(val)
            i += 1
            continue

        if kind == "rule":
            i += 1
            continue

        if kind == "quote":
            r.body_para(val, indent=720)
            i += 1
            continue

        r.body_para(val)
        i += 1

    r.date_blocks = date_blocks
    return r


def generate(out_dir, force=False, template=TEMPLATE_WORD, jobs=None):
    """生成 docx（公告周报 + 重要资讯）。返回 [(输出路径, 段落数)]。"""
    out_dir = _abs(out_dir)
    if not os.path.isfile(template):
        raise WordToolError("模板不存在：%s" % template)
    tpl_sha = sha256(template)

    jobs = jobs or JOBS
    missing = [src for src, _ in jobs
               if not os.path.isfile(os.path.join(out_dir, src))]
    if missing:
        raise WordToolError("源 md 不存在：\n    - %s" % "\n    - ".join(
            os.path.join(out_dir, m) for m in missing))

    results = []
    for src_name, out_name in jobs:
        src = os.path.join(out_dir, src_name)
        out = os.path.join(out_dir, out_name)
        assert_writable_target(out)
        if os.path.exists(out) and not force:
            raise WordToolError(
                "目标文件已存在，拒绝覆盖：%s\n  如需覆盖请显式加 --force" % out)

        # 1) 复制模板 → 副本；模板原件此后不再被打开
        shutil.copyfile(template, out)
        # 2) 打开副本修改
        doc = Document(out)
        blocks = parse_md(open(src, encoding="utf-8").read())
        r = render(doc, blocks)
        doc.save(out)

        # 3) 模板必须原封不动
        if sha256(template) != tpl_sha:
            raise WordToolError("模板被改动了！%s" % template)

        results.append((out, len(doc.paragraphs), src))
        counts = {}
        for k, _ in blocks:
            counts[k] = counts.get(k, 0) + 1
        log("  [OK] %-24s <- %s" % (out_name, src_name))
        log("       分块：%s" % "  ".join(
            "%s=%d" % (k, v) for k, v in sorted(counts.items())))
        if r.date_blocks:
            log("       分页：%d 个日期块 → %d 个分页符（首个日期块不加）"
                % (r.date_blocks, r.date_blocks - 1))
        else:
            log("       分页：无日期块，未插入分页符")
    return results


# ----------------------------------------------------------- 校验
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _canon(el):
    """元素的规范化字节：忽略属性顺序、命名空间前缀写法与空白。

    python-docx 存盘时会重排整份 XML 的序列化，所以**不能**拿字节比对——
    字节不同不代表内容被改了。规范化之后比对才是"语义是否相同"。
    """
    try:
        return etree.tostring(el, method="c14n2", with_comments=False)
    except (ValueError, TypeError):
        return etree.tostring(el, method="c14n", with_comments=False)


def _template_fidelity(template, out, problems):
    """核对模板原有的**字体 / 字号 / 编号 / 页边距 / 页眉页脚 / 段落样式**未被改动。

    直接比对整包字节是不行的（见 `_canon`），因此改为逐项核对真正决定版式的四处：
      `styles.xml`      → 段落样式 / 字体 / 字号
      `header*` `footer*` → 页眉页脚（这两类部件我们从不触碰，要求逐字节相同）
      `w:sectPr`        → 纸张与页边距
      `abstractNum`     → 编号格式（模板原有的定义必须一个不少、内容不变）
    """
    zt, zo = zipfile.ZipFile(template), zipfile.ZipFile(out)
    tn, on = set(zt.namelist()), set(zo.namelist())

    for name in sorted(tn):
        base = name.rsplit("/", 1)[-1]
        hf = base.startswith("header") or base.startswith("footer")
        if name != "word/styles.xml" and not hf:
            continue
        if name not in on:
            problems.append("模板部件丢失：%s" % name)
            continue
        a, b = zt.read(name), zo.read(name)
        if hf:
            if a != b:
                problems.append("页眉/页脚被改动：%s" % name)
        elif _canon(etree.fromstring(a)) != _canon(etree.fromstring(b)):
            problems.append("样式表被改动（字体/字号/段落样式可能已变）：%s" % name)

    def _sectpr(z):
        return etree.fromstring(z.read("word/document.xml")).find(
            ".//" + W_NS + "sectPr")

    sa, sb = _sectpr(zt), _sectpr(zo)
    if sa is None or sb is None:
        problems.append("document.xml 里找不到 w:sectPr，无法核对页面设置")
    elif _canon(sa) != _canon(sb):
        problems.append("页面设置（纸张 / 页边距）被改动")

    def _abstract_nums(z):
        return {e.get(W_NS + "abstractNumId"): _canon(e)
                for e in etree.fromstring(z.read("word/numbering.xml")).iter(
                    W_NS + "abstractNum")}

    na, nb = _abstract_nums(zt), _abstract_nums(zo)
    for k, v in na.items():
        if k not in nb:
            problems.append("模板的编号定义 abstractNumId=%s 丢失" % k)
        elif nb[k] != v:
            problems.append("模板的编号定义 abstractNumId=%s 被改动" % k)


def verify(out_dir, jobs=None, template=TEMPLATE_WORD):
    """读回 docx：zip 完整性 / 内容与 md 逐条一致 / 排版硬规则 / 模板保真。"""
    jobs = jobs or JOBS
    ok = True
    for src_name, out_name in jobs:
        src = os.path.join(out_dir, src_name)
        out = os.path.join(out_dir, out_name)
        problems = []

        # a) zip 完整
        if not zipfile.is_zipfile(out):
            problems.append("不是合法的 zip/docx")
        else:
            try:
                z = zipfile.ZipFile(out)
                bad = z.testzip()
                if bad:
                    problems.append("zip 损坏：%s" % bad)
                for name in ("word/document.xml", "word/styles.xml",
                             "word/numbering.xml"):
                    if name not in z.namelist():
                        problems.append("缺少 %s" % name)
            except Exception as exc:                       # noqa: BLE001
                problems.append("zip 读取失败：%s" % exc)

        # b) 文件能被 Word 解析器打开、且内容与 md 逐条对齐
        try:
            d = Document(out)
        except Exception as exc:                           # noqa: BLE001
            problems.append("python-docx 打不开：%s" % exc)
            d = None

        if d is not None:
            # 按 body 真实子元素顺序取内容（表格必须落在它该在的位置，
            # 不能把表格单元格一股脑追加到末尾——那样测不出顺序错乱）
            got = []
            for ch in d.element.body:
                if ch.tag == qn("w:p"):
                    t = Paragraph(ch, d).text
                    if t.strip():
                        got.append(t)
                elif ch.tag == qn("w:tbl"):
                    for row in Table(ch, d).rows:
                        got.append(" | ".join(c.text for c in row.cells))
            if len(d.tables) == 0 and any(k == "table" for k, _ in parse_md(
                    open(src, encoding="utf-8").read())):
                problems.append("md 有表格，docx 里没有")

            src_blocks = parse_md(open(src, encoding="utf-8").read())
            want = []
            for k, v in src_blocks:
                if k == "table":
                    want += [" | ".join(plain_text(c) for c in row)
                             for row in v]
                elif plain_text(v).strip():
                    want.append(plain_text(v))

            if got != want:
                problems.append(
                    "内容与 md 不是逐条逐序一致（docx %d 条 / md %d 条）"
                    % (len(got), len(want)))
                for i, (a, b) in enumerate(zip(got, want)):
                    if a != b:
                        problems.append("  首个差异 @%d docx=%r md=%r"
                                        % (i, a[:40], b[:40]))
                        break

            # c) 公告条目必须真的是 List Paragraph 且带编号
            #    （防"模板体"行被当成普通段落——那会丢掉编号和加粗）
            entry_texts = set(plain_text(v) for k, v in src_blocks if k == "entry")
            # 只有日期/章节分组标题才是 16pt 加粗；【公司新闻】在模板里是普通段落
            flat = [plain_text(v) for k, v in src_blocks
                    if k in ("head1", "head2", "head3")]
            for p in d.paragraphs:
                if p.text in entry_texts:
                    pPr = p._element.pPr
                    has_num = pPr is not None and pPr.find(qn("w:numPr")) is not None
                    if p.style.name != "List Paragraph" or not has_num:
                        problems.append(
                            "条目缺 List Paragraph/编号：%r（样式=%s，编号=%s）"
                            % (p.text[:26], p.style.name, has_num))
                    # List Paragraph 自带的 contextualSpacing 若没被覆写，
                    # 相邻条目之间不会留白 → 公告会粘成一片
                    cs = pPr.find(qn("w:contextualSpacing")) if pPr is not None else None
                    if cs is None or cs.get(qn("w:val")) not in ("0", "false"):
                        problems.append(
                            "条目未覆写 contextualSpacing（会与下一条粘住）：%r"
                            % p.text[:26])
                if p.text in flat and p.text.strip():
                    r0 = p.runs[0] if p.runs else None
                    if not (r0 and r0.font.size and r0.font.bold):
                        problems.append(
                            "分组标题不是 16pt 加粗：%r" % p.text[:26])
            # d) 排版硬规则：全文不插空段落，所有层级统一"一行空行"
            body_paras = [Paragraph(ch, d)
                          for ch in d.element.body if ch.tag == qn("w:p")]
            blanks = [i for i, p in enumerate(body_paras) if not p.text.strip()]
            if blanks:
                problems.append(
                    "出现 %d 个空段落（空行必须由段前间距产生，不能手插空行）：位置 %s"
                    % (len(blanks), blanks[:6]))
            for i, p in enumerate(body_paras):
                pPr = p._element.pPr
                sp = pPr.find(qn("w:spacing")) if pPr is not None else None
                if sp is None:
                    problems.append("第 %d 段没有显式 spacing：%r" % (i, p.text[:24]))
                    continue
                before = int(sp.get(qn("w:before")) or 0)
                after = int(sp.get(qn("w:after")) or 0)
                want = 0 if i == 0 else GAP_TWIPS
                if before != want or after != 0:
                    problems.append(
                        "第 %d 段间距不是统一一行空行：before=%d after=%d（应为 %d/0）：%r"
                        % (i, before, after, want, p.text[:24]))

            # e) 分页硬规则：每个日期块另起一页（首个日期块除外，第一页不加）
            date_heads = [plain_text(v) for k, v in src_blocks
                          if k in ("head1", "head2", "head3") and is_date_head(v)]
            want_pb = max(0, len(date_heads) - 1)
            got_pb = []
            for j, p in enumerate(body_paras):
                pPr = p._element.pPr
                pbb = pPr.find(qn("w:pageBreakBefore")) if pPr is not None else None
                if pbb is not None and pbb.get(qn("w:val")) not in ("0", "false", "off"):
                    got_pb.append(j)
            if len(got_pb) != want_pb:
                problems.append(
                    "分页符数量不对：docx %d 个 / md %d 个日期块 → 应为 %d 个"
                    % (len(got_pb), len(date_heads), want_pb))
            for j in got_pb:
                if not is_date_head(body_paras[j].text):
                    problems.append(
                        "分页符落在了非日期块标题上（第 %d 段：%r）"
                        % (j, body_paras[j].text[:30]))
            if date_heads:
                first_head = date_heads[0]
                for j in got_pb:
                    if body_paras[j].text.strip() == first_head.strip():
                        problems.append(
                            "首个日期块不应加段前分页（第一页本来就在页首）：%r"
                            % first_head[:30])
            # 分页必须落在日期块标题上、且不得用空行凑
            if got_pb and blanks:
                problems.append("同时存在分页符与空段落，分页可能由空行凑出")

            if not entry_texts:
                problems.append("md 里一条公告条目都没识别出来（分块规则可能失效）")

            # f) 模板保真：字体/字号/编号/页边距/页眉页脚/段落样式必须原样保留
            if template and os.path.isfile(template):
                try:
                    _template_fidelity(template, out, problems)
                except Exception as exc:                   # noqa: BLE001
                    problems.append("与模板比对失败：%s" % exc)

        log("  [%s] %-24s 段落 %s，表格 %s"
            % ("OK" if not problems else "FAIL", out_name,
               len(d.paragraphs) if d else "-",
               len(d.tables) if d else "-"))
        for p in problems:
            ok = False
            log("         - %s" % p)
    return ok


def pick_jobs(only):
    """把 --only 解析成 JOBS 的子集。

    支持传输出名 / 源 md 名 / 二者中的关键字（如「汇总」「重要资讯」）——
    用户说"只重新生成某一份"时不必记全名。
    """
    if not only:
        return list(JOBS)
    q = only.strip()
    hits = [j for j in JOBS if q in (j[0], j[1]) or q in j[0] or q in j[1]]
    if len(hits) != 1:
        raise WordToolError(
            "--only %r 匹配到 %d 项，请写得更明确。可选：\n    - %s"
            % (only, len(hits), "\n    - ".join(o for _, o in JOBS)))
    return hits


def current_week_dir():
    root = os.path.join(BASE, "05_输出")
    subs = [d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))]
    if not subs:
        raise WordToolError("05_输出 下没有周次目录")
    return os.path.join(root, sorted(subs)[-1])



def main_word():
    ap = argparse.ArgumentParser(
        description="把周报 md 交付物按模板渲染成 docx（模板原件只读）")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--week", help="周次，例如 2026-W37（等价于 --outdir 05_输出\\2026-W37）")
    g.add_argument("--outdir", help="直接指定输出目录")
    ap.add_argument("--force", action="store_true", help="允许覆盖已存在的 docx")
    ap.add_argument("--verify-only", action="store_true", help="只校验，不生成")
    ap.add_argument("--only", help="只处理其中一项，例如 --only 本周公司公告周报.docx "
                                   "或 --only 汇总；不填则两项都处理")
    ap.add_argument("--template", default=TEMPLATE_WORD)
    args = ap.parse_args()

    if args.outdir:
        out_dir = _abs(args.outdir)
    elif args.week:
        out_dir = os.path.join(BASE, "05_输出", args.week)
    else:
        out_dir = current_week_dir()

    log("模板  : %s" % _abs(args.template))
    log("目录  : %s" % out_dir)
    log("")

    try:
        jobs = pick_jobs(args.only)
        if args.only:
            log("范围  : 只处理 %s" % ", ".join(o for _, o in jobs))
            log("")
        if not args.verify_only:
            log("生成：")
            generate(out_dir, force=args.force, template=args.template, jobs=jobs)
            log("")
        log("校验：")
        ok = verify(out_dir, jobs=jobs, template=args.template)
    except WordToolError as exc:
        log("[拒绝] %s" % exc)
        return 3

    log("")
    log("模板 sha256 未变：%s" % sha256(args.template))
    return 0 if ok else 1




# ==================================================================
#  第二部分：PPT（原 更新PPT.py）
# ==================================================================
CONTENT_SHAPE_INDEX = 3      # 正文形状（实测）
LINE_SPACING_FACTOR = 1.3    # 保守行高系数（PowerPoint 中文正文约 1.2–1.5）

TIER_FORMAL = "正式入选"
TIER_FILLED = "正式入选·自动补位"   # 版面自适应从备选池补进来的
TIER_BACKUP = "备选"

TITLE_END = "。"                                    # 小标题必须以它收尾
RE_BRACKET = re.compile(r"[【】（）()\[\]〔〕{}〈〉《》「」『』]")

# 单条长度硬约束（任务卡 §PPT 正文长度控制）
MAX_ENTRY_LINES = 3          # 每条（小标题 + 事实 + 点评）最多占几行
MAX_TITLE_LINE_RATIO = 0.5   # 小标题最多占半行（"标题本身也要简短"）

# 压缩时的分句点：句号 / 分号（含全角半角）。事实常写成「A；B；C」的长句，
# 按分句砍才能保证每次只丢一个完整语义单元，不会砍出半句话。
RE_CLAUSE = re.compile(r"(?<=[。；;])")

# 点评禁用语（任务卡 §PPT 点评规则）——出现即校验不通过
BANNED_WORDS = ("明显利好", "显著利好", "强烈看好", "建议买入", "建议增持", "买入评级",
                "增持评级", "推荐评级", "目标价", "上调评级", "必涨", "翻倍")

# 版面自适应阈值
SAFE_MAX_RATIO = 0.92   # 补位后不得超过预算的 92%（留余量吸收折行估算误差）
BLANK_RATIO = 0.78      # 低于此比例 = 下方仍有明显大块空白，可尝试补 1 条

RE_TIER = re.compile(r"^\[([^\]]+)\]\s*(.*)$")
RE_FIELD = re.compile(r"^(事实|点评|理由|公司)\s*[：:]\s*(.*)$")
# markdown 结构行：分隔线 / ATX 标题 / 表格 / 引用 / 列表项
RE_STRUCT = re.compile(r"^(#{1,6}\s|[-*_]{3,}\s*$|\||>|[-*+]\s|\d+[.、)]\s)")
# 同上，但按「每行行首」检查（用于字段内容，避免把字段里偶然出现的 `|` 误判成表格）
RE_STRUCT_MULTILINE = re.compile(
    r"^(#{1,6}\s|[-*_]{3,}\s*$|\||>|[-*+]\s|\d+[.、)]\s)", re.M)




# ----------------------------------------------------------------- 周次计算
def resolve_window(week=None, end=None):
    """返回 (week_label, start_date, end_date)。窗口 = 周六 → 周五。"""
    if end:
        e = dt.date.fromisoformat(end)
    else:
        today = dt.date.today()
        back = (today.weekday() + 1) % 7          # 到最近周五的天数
        e = today if back == 0 else today - dt.timedelta(days=back)
    s = e - dt.timedelta(days=6)
    if not week:
        y, w, _ = e.isocalendar()
        week = "%d-W%02d" % (y, w)
    return week, s, e


# ----------------------------------------------------------------- 容错读取
def read_text(path):
    if not os.path.isfile(path):
        return None
    for enc in ("utf-8-sig", "utf-8", "gbk"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, LookupError):
            continue
    return None


def split_sentences(text):
    """按中文句末标点切句，保留标点。"""
    parts = re.split(r"(?<=[。！？])", text)
    return [p.strip() for p in parts if p.strip()]


# ----------------------------------------------------------------- 公司名识别
def load_company_pool(report_dir):
    """本周公司简称集合，取自 本周公司公告汇总.md 的 `【公司简称】`。

    不依赖外部公司池文件，天然与本周实际处理的公司同步。
    """
    t = read_text(os.path.join(report_dir, "本周公司公告汇总.md")) or ""
    names = {m.group(1).strip() for m in re.finditer(r"^【(.+?)】", t, flags=re.M)}
    names.discard("公司新闻")          # 版式固定行，不是公司
    return names


def lead_company(lead, pool):
    """从无括号小标题里识别公司名：取能作为前缀的最长公司简称。

    新体例的标题形如「保利发展销售同比承压，……。」——公司名在最前面且没有分隔符，
    因此用本周公司池做最长前缀匹配；匹配不到返回 None（跳过同公司合并检查）。
    """
    s = (lead or "").strip()
    hits = [c for c in pool if c and s.startswith(c)]
    return max(hits, key=len) if hits else None


# ----------------------------------------------------------------- 内容来源
def normalize_lead(lead):
    """把作者写的导语规整成合规小标题：去括号残留 → 确保以「。」结尾。

    返回 (title, rest)：rest 是从标题里切出来、应并入正文的剩余部分。
    """
    s = (lead or "").strip()
    # 兼容作者误写的老体例 【公司：结论】 / （公司：结论）→ 直接报错，不静默改写
    if RE_BRACKET.search(s):
        raise ValueError(
            "PPT候选内容.md 的小标题含括号，违反【标题硬格式】：%r\n"
            "请改成「公司名 + 结论，……。」（无括号、以句号结尾）" % s[:80])
    if not s:
        return "", ""
    if s.endswith(TITLE_END):
        return s, ""
    # 没以句号结尾：以第一个句号为界，前面当标题，后面并回正文
    i = s.find(TITLE_END)
    if i >= 0:
        return s[:i + 1], s[i + 1:].strip()
    return s + TITLE_END, ""


def parse_candidate_md(text, pool=None):
    """解析 PPT候选内容.md（条目块 + [正式入选]/[备选] 标记）。

    每条返回 {"tier", "lead", "facts", "comment", "reason", "company", "data", "_lineno"}：
        lead    —— 无括号、以「。」结尾的小标题（幻灯片上渲染为红色加粗）
        data    —— facts + comment 拼接（幻灯片上渲染为黑色正文，点评接在最后）
    解析不出 tier 的条目一律不采纳（避免把说明性文字当正文塞进幻灯片）。
    """
    entries, cur = [], None
    for lineno, raw in enumerate(text.splitlines()):
        line = raw.rstrip()
        s = line.strip()
        if not s:
            continue

        m = RE_TIER.match(s)
        if m:
            tag = m.group(1).strip()
            if tag.startswith(TIER_FORMAL):
                tier = TIER_FILLED if "自动补位" in tag else TIER_FORMAL
            elif tag.startswith(TIER_BACKUP):
                tier = TIER_BACKUP
            else:
                tier = None          # 不是条目（如 [1] 之类的说明行）
            if tier is None:
                continue
            if cur:
                entries.append(cur)
            cur = {"tier": tier, "lead": m.group(2).strip(),
                   "facts": "", "comment": "", "reason": "", "company": "",
                   "_lineno": lineno}
            continue

        if cur is None:
            continue

        fm = RE_FIELD.match(s)
        if fm:
            key = {"事实": "facts", "点评": "comment",
                   "理由": "reason", "公司": "company"}[fm.group(1)]
            if key == "company":
                cur["company"] = fm.group(2).strip()
            else:
                cur[key] = (cur[key] + fm.group(2).strip())
            cur["_field"] = key
            continue

        # markdown 结构行（分隔线 / 标题 / 表格 / 引用 / 列表）不是字段内容，
        # 必须在这里截断当前字段——否则 `---`、`## 备选` 会被当成续行
        # 并进上一条的正文，悄悄跑到幻灯片上。
        if RE_STRUCT.match(s):
            cur["_field"] = None
            continue

        # 字段的续行（md 里写成长句换行时）——并回当前字段
        if cur.get("_field"):
            cur[cur["_field"]] += s

    if cur:
        entries.append(cur)

    out = []
    for e in entries:
        e.pop("_field", None)
        for k in ("facts", "comment", "reason"):
            # 按行首锚定检查，不能对整个字段用 search —— 那样字段里偶尔出现一个
            # `|` 就会被误判成表格行而拒绝生成。
            if RE_STRUCT_MULTILINE.search(e[k]):
                raise ValueError(
                    "PPT候选内容.md 的 %s 字段混入了 markdown 结构行，"
                    "拒绝生成（会污染幻灯片）：%r" % (k, e[k][:70]))
        title, rest = normalize_lead(e["lead"])
        e["lead"] = title
        if rest:
            e["facts"] = rest + e["facts"]
        parts = [x for x in (e["facts"], e["comment"]) if x]
        e["data"] = "".join(parts)
        e["src"] = "PPT候选内容.md"
        if not e.get("company"):
            e["company"] = lead_company(title, pool or set()) or ""
        if e["lead"]:
            out.append(e)
    return out


def parse_summary_md(text, pool=None):
    """兜底：从 本周公司公告汇总.md 机械提取（导语需人工润色）。

    模板行格式：【公司简称】公告标题：正文  →  小标题写成「公司简称 + 公告标题。」
    """
    entries = []
    for raw in text.splitlines():
        m = re.match(r"^【(.+?)】(.+?)[：:](.*)$", raw.strip())
        if not m:
            continue
        comp, title, body = m.group(1), m.group(2), m.group(3).strip()
        title = re.sub(r"的?公告$", "", title).strip()
        title = re.sub(r"^关于", "", title).strip()      # 去公文式前缀
        lead = (comp + title).strip()
        lead = lead if lead.endswith(TITLE_END) else lead + TITLE_END
        entries.append({"tier": TIER_BACKUP, "lead": lead, "facts": body,
                        "comment": "", "reason": "机械转换，未标注入选层级",
                        "company": comp, "data": body, "_lineno": None,
                        "src": "本周公司公告汇总.md"})
    return entries


def _simplify_entries(entries):
    """把面向读者的最终中文输出统一转简体（港股公告原文常为繁体）。

    只做**字形转换**：数字 / 金额 / 百分比 / 日期 / 股票代码 / 英文缩写一律不动。
    注意分层：`00_配置\\行业公司池规则.md` 要求公司池保留**公告原文繁体全称**，
    因此繁简转换只作用于最终中文输出，不回溯改写公司池与来源清单。
    """
    if _TC is None or not _TC.available():
        return entries
    for e in entries:
        for k in ("lead", "facts", "comment", "reason"):
            if e.get(k):
                e[k] = _TC.to_simplified(e[k])
        _rebuild_data(e)
    return entries


def load_entries(report_dir, explicit=None):
    """按优先级取候选内容，返回 (entries, 来源说明)。

    返回前统一过一次繁→简（见 _simplify_entries）。
    """
    pool = load_company_pool(report_dir)
    if explicit:
        t = read_text(explicit)
        if t:
            return _simplify_entries(parse_candidate_md(t, pool)), \
                os.path.basename(explicit)

    for fname, parser, note in (
        ("PPT候选内容.md", parse_candidate_md, ""),
        ("本周公司公告汇总.md", parse_summary_md, "（机械转换，需人工润色）"),
        ("本周重要资讯.md", parse_candidate_md, ""),
    ):
        t = read_text(os.path.join(report_dir, fname))
        if not t:
            continue
        e = parser(t, pool)
        if e:
            return _simplify_entries(e), fname + note

    return [], "（无可用内容来源）"


# ----------------------------------------------------------------- 排版度量
def eff_units(s):
    """半角单位数：全角=2，半角=1（用于估算折行）。"""
    return sum(2 if ord(c) > 0x2E80 else 1 for c in s)


def measure_template():
    """读模板样板 XML + 标定行容量。所有数字都来自模板自身，不写死。"""
    prs = Presentation(TEMPLATE_PPT)
    slide = prs.slides[0]
    shape = slide.shapes[CONTENT_SHAPE_INDEX]
    tf = shape.text_frame

    content_paras, blank_para = [], None
    for p in tf.paragraphs:
        if p.runs:
            content_paras.append(p)
        elif blank_para is None:
            blank_para = p

    if not content_paras:
        raise RuntimeError("模板正文形状未找到有文字的段落，形状索引可能已变化")
    if blank_para is None:
        raise RuntimeError("模板正文未找到空段样板")

    # 字号：取首个有显式字号的 run
    pt = None
    for r in content_paras[0].runs:
        if r.font.size:
            pt = r.font.size.pt
            break
    if not pt:
        raise RuntimeError("模板正文未取到字号，无法标定行容量")

    emu_half = int(pt * 12700) // 2                     # 半角字符宽
    text_w = shape.width - tf.margin_left - tf.margin_right
    text_h = shape.height - tf.margin_top - tf.margin_bottom
    units_per_line = int(text_w / emu_half)             # 向下取整，偏保守
    line_h = int(pt * LINE_SPACING_FACTOR * 12700)

    # 模板自身用量 —— 这是"已被证明放得下"的行数，作为安全上限
    geom_lines = int(text_h / line_h)
    tpl_lines = 0
    for p in tf.paragraphs:
        u = eff_units(p.text)
        tpl_lines += (-(-u // units_per_line)) if u else 1     # 空段占 1 行
    budget_lines = min(geom_lines, tpl_lines)

    return {
        "prs": prs,
        "tf": tf,
        "body": tf._txBody,
        "donor_content": deepcopy(content_paras[0]._p),
        "donor_blank": deepcopy(blank_para._p),
        "font_pt": pt,
        "units_per_line": units_per_line,
        "geom_lines": geom_lines,
        "tpl_lines": tpl_lines,
        "budget_lines": budget_lines,
        "baseline_count": len(content_paras),
        "auto_size": str(tf.auto_size),
        "max_entry_lines": MAX_ENTRY_LINES,
    }


def entry_lines(entry, units_per_line):
    """一条占用几行（只算文本行，不含条间空段）。"""
    u = eff_units(entry["lead"] + " " + entry["data"])
    return max(1, -(-u // units_per_line))


def total_lines(entries, m):
    """含条间空段的整页行占用。"""
    text_lines = sum(entry_lines(e, m["units_per_line"]) for e in entries)
    return text_lines + max(0, len(entries) - 1), text_lines


def split_clauses(text):
    """按句号/分号切分句，保留标点。用于"从末尾砍次要信息"。"""
    return [p for p in RE_CLAUSE.split(text or "") if p.strip()]


def _rebuild_data(e):
    """facts + comment → data（幻灯片上渲染的黑色正文）。每次改动后必须重算。"""
    e["data"] = "".join(x for x in (e.get("facts"), e.get("comment")) if x)


def compress_entry(entry, m, cap=None):
    """把**单条**压到 cap 行以内，只减字、绝不缩小字号。

    压缩顺序（任务卡 §PPT 正文长度控制）：
        0. 点评超过 1 行 → 只留首句（点评本身要求"极短"）
        1. 砍事实的**末尾**分句 —— 这些汇总都是"最重要的数字在前"，
           因此末位分句就是最次要的信息，符合"优先删掉次要信息"
        2. 仍超 → 整条舍弃点评，只留事实
    返回被砍掉的内容 [(部位, 原文), ...]，供写进 PPT生成说明.md 留痕。
    """
    cap = cap or m["max_entry_lines"]
    upl = m["units_per_line"]
    dropped = []

    # 0) 点评要"极短"：超过 1 行就只留首句
    c = entry.get("comment") or ""
    if c and eff_units(c) > upl:
        cl = split_clauses(c)
        if len(cl) > 1:
            dropped.append(("点评（超 1 行，只留首句）", "".join(cl[1:])))
            entry["comment"] = cl[0]
            _rebuild_data(entry)

    # 1) 从末尾砍事实的次要分句
    while entry_lines(entry, upl) > cap:
        cl = split_clauses(entry.get("facts"))
        if len(cl) <= 1:
            break
        dropped.append(("事实（末尾次要分句）", cl[-1]))
        entry["facts"] = "".join(cl[:-1])
        _rebuild_data(entry)

    # 2) 仍超 → 舍弃点评
    if entry_lines(entry, upl) > cap and (entry.get("comment") or ""):
        dropped.append(("点评（整条）", entry["comment"]))
        entry["comment"] = ""
        _rebuild_data(entry)

    return dropped


def compress_entries(entries, m, cap=None):
    """逐条压缩，返回 (records, still_over)。

    records   —— [(条目序号, 小标题, [(部位, 原文), ...]), ...]，只含真正被压过的
    still_over—— 压完仍超 cap 的条目序号（说明小标题本身就太长，需人工改写）
    """
    cap = cap or m["max_entry_lines"]
    records, still_over = [], []
    for i, e in enumerate(entries, 1):
        if entry_lines(e, m["units_per_line"]) <= cap:
            continue
        before = entry_lines(e, m["units_per_line"])
        dropped = compress_entry(e, m, cap)
        after = entry_lines(e, m["units_per_line"])
        records.append((i, e["lead"], before, after, dropped))
        if after > cap:
            still_over.append(i)
    return records, still_over


def title_too_long(entry, m):
    """小标题是否超过半行（"标题本身也要简短"）。"""
    limit = m["units_per_line"] * MAX_TITLE_LINE_RATIO
    return eff_units(entry["lead"]) > limit


def select(entries, m):
    """按行预算挑条数：放不下就减条数（从低优先级往后砍），不缩字号。

    条数**不是固定值**（4/5/6 条都可），完全由内容长度与行预算决定。
    """
    chosen, text_lines = [], 0
    for e in entries:
        nl = entry_lines(e, m["units_per_line"])
        # 加进来之后的空段数 = len(chosen)
        if chosen and (text_lines + nl + len(chosen) > m["budget_lines"]):
            break
        if not chosen and nl > m["budget_lines"]:
            chosen.append(e)        # 首条就超容量：保留但下面会告警
            text_lines += nl
            break
        chosen.append(e)
        text_lines += nl
    blanks = max(0, len(chosen) - 1)
    return chosen, text_lines, blanks


def fill_from_backup(chosen, pool, m):
    """版面自适应：下半页若明显留白，按重要性从备选池补条目。

    规则（任务卡 §PPT 版面自适应）:
      1. 只在下半页明显留白时补 —— 页面已较满就不再增加
      2. 从备选池按**重要性**（文件内顺序）挑，不跳级
      3. 补进去后若溢出 / 页面过满，立即**撤回**该条
      4. 全程不缩字号、不压行距

    pool 是可变列表，被选中的条目会被移出。
    返回 (filled, stop_reason)：stop_reason 说明为什么停止补位，写进运行日志留痕。
    """
    filled = []
    stop = "备选池已空"
    while pool:
        cur_total, _ = total_lines(chosen, m)
        cur_ratio = cur_total / m["budget_lines"]
        if cur_ratio >= BLANK_RATIO:
            stop = "页面已占 %.0f%%，较满，不再增加" % (cur_ratio * 100)
            break
        cand = pool[0]
        nl = entry_lines(cand, m["units_per_line"])
        new_total = cur_total + nl + 1          # +1 = 与上一条之间的空段
        new_ratio = new_total / m["budget_lines"]
        if new_ratio > SAFE_MAX_RATIO:
            stop = ("备选池首条（%s）补入后达 %.0f%%，超出 %.0f%% 安全线，撤回该条"
                    % (cand["lead"][:16], new_ratio * 100, SAFE_MAX_RATIO * 100))
            break
        chosen.append(pool.pop(0))
        filled.append(cand)
    return filled, stop


# ----------------------------------------------------------------- 写入副本
def write_entries(m, entries):
    """把 entries 写入副本正文形状，保留原格式。

    模板 run[0] = 红色加粗小标题，run[1] = 黑色正文；
    这里只改文字、删掉多余 run，因此配色/字号/行距/项目符号全部沿用模板。
    小标题（含结尾句号）与正文同处一个项目符号段落，句号后直接接正文。
    """
    tf, body = m["tf"], m["body"]
    donor_c, donor_b = m["donor_content"], m["donor_blank"]

    for p in body.findall(qn("a:p")):
        body.remove(p)

    for i, e in enumerate(entries):
        p_el = deepcopy(donor_c)
        body.append(p_el)
        runs = _Paragraph(p_el, tf).runs
        if not runs:
            continue
        runs[0].text = e["lead"]                       # 小标题（红 + 粗，沿用模板）
        if len(runs) >= 2:
            runs[1].text = e["data"] or ""             # 事实 + 点评，紧接句号
            for r in runs[2:]:
                r._r.getparent().remove(r._r)
        if i < len(entries) - 1:              # 条间空段，最后一条后不加
            body.append(deepcopy(donor_b))


# ----------------------------------------------------------------- 层级回写
def update_tier_markers(path, entries):
    """把自动补位的条目在 PPT候选内容.md 里的标记改成「[正式入选·自动补位]」。

    **只改标记那一行**，其余正文/理由/说明原样保留，避免脚本把作者写的
    候选内容覆盖掉。幂等：补位条目改过一次后再次运行不会重复改动。
    """
    if not path or not os.path.isfile(path):
        return 0
    lines = read_text(path)
    if lines is None:
        return 0
    lines = lines.splitlines(keepends=True)
    changed = 0
    for e in entries:
        ln = e.get("_lineno")
        if ln is None or ln >= len(lines):
            continue
        old = lines[ln]
        # 只替换行首的 `[标记]`，绝不能整行替换 —— RE_TIER 会把小标题一起匹配进去，
        # 用 sub 整行替换会把小标题文本删掉。
        new = re.sub(r"^\[[^\]]+\]", "[%s]" % e["tier"], old, count=1)
        if new != old:
            lines[ln] = new
            changed += 1
    if changed:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8", newline="") as f:
            f.write("".join(lines))
        os.replace(tmp, path)
    return changed


# ----------------------------------------------------------------- 质量检查
def check(chosen, m, dropped, filled, must_include=None):
    issues = []

    empty = [i + 1 for i, e in enumerate(chosen)
             if not (e["lead"] or "").strip() and not (e["data"] or "").strip()]
    if empty:
        issues.append("存在空条目：第 %s 条" % ",".join(map(str, empty)))

    # 点评必须基于数据 → 没有事实只有点评的条目视为残缺
    nofact = [i + 1 for i, e in enumerate(chosen)
              if not (e.get("comment") or "").strip() and not (e.get("facts") or "").strip()]
    if nofact:
        issues.append("条目既无事实也无点评：第 %s 条" % ",".join(map(str, nofact)))

    # 标题硬格式：无括号 + 以「。」结尾
    brk = ["第%d条 %r" % (i + 1, e["lead"][:24])
           for i, e in enumerate(chosen) if RE_BRACKET.search(e["lead"])]
    if brk:
        issues.append("小标题含括号（违反标题硬格式）：%s" % "；".join(brk))
    noend = ["第%d条 %r" % (i + 1, e["lead"][-14:])
             for i, e in enumerate(chosen) if not e["lead"].endswith(TITLE_END)]
    if noend:
        issues.append("小标题未以「。」结尾：%s" % "；".join(noend))

    # 点评禁用营销化措辞
    banned = []
    for i, e in enumerate(chosen):
        hit = [w for w in BANNED_WORDS if w in (e.get("comment") or "")]
        if hit:
            banned.append("第%d条 %s" % (i + 1, "/".join(hit)))
    if banned:
        issues.append("点评含禁用营销化措辞：%s" % "；".join(banned))

    seen, dup = {}, []
    for i, e in enumerate(chosen):
        key = re.sub(r"\s+", "", (e["lead"] or "") + (e["data"] or ""))
        if not key:
            continue
        if key in seen:
            dup.append((seen[key], i + 1))
        else:
            seen[key] = i + 1
    if dup:
        issues.append("存在重复内容：%s" % "；".join("第%d条=第%d条" % d for d in dup))

    # 同公司重复出现（应在写作时合并成一条）
    seen_c, rep_c = {}, []
    for i, e in enumerate(chosen):
        c = (e.get("company") or lead_company(e["lead"], set()) or "").strip()
        if not c:
            continue
        if c in seen_c:
            rep_c.append("%s（第%d、%d条）" % (c, seen_c[c], i + 1))
        else:
            seen_c[c] = i + 1
    if rep_c:
        issues.append("同公司未合并：%s" % "；".join(rep_c))

    lines, text_lines = total_lines(chosen, m)
    blanks = max(0, len(chosen) - 1)
    if lines > m["budget_lines"]:
        issues.append("预估溢出：%d 行 / 可用 %d 行（模板自用 %d 行）"
                      % (lines, m["budget_lines"], m["tpl_lines"]))
    if dropped and not chosen:
        issues.append("所有条目均超容量，未写入任何内容")

    # 每条长度（任务卡 §PPT 正文长度控制）—— PPT 不是 Word 公告汇总
    per_entry = [entry_lines(e, m["units_per_line"]) for e in chosen]
    over = [i + 1 for i, n in enumerate(per_entry) if n > m["max_entry_lines"]]
    if over:
        issues.append("单条超过 %d 行，应压缩或合并（不得缩小字号）：第 %s 条"
                      % (m["max_entry_lines"], "、".join(map(str, over))))

    long_title = ["第%d条「%s」" % (i + 1, e["lead"][:20])
                  for i, e in enumerate(chosen) if title_too_long(e, m)]
    if long_title:
        issues.append("小标题超过半行，需更简短：%s" % "；".join(long_title))

    # 指定必含公司（如本周重点公司）是否真的进了 PPT
    missing = []
    if must_include:
        blob = "".join((e.get("lead") or "") + (e.get("data") or "") for e in chosen)
        missing = [c for c in must_include if c and c not in blob]
        if missing:
            issues.append("指定必含公司未出现在 PPT：%s" % "、".join(missing))

    ratio = (lines / m["budget_lines"]) if m["budget_lines"] else 1.0
    return {
        "text_lines": text_lines,
        "blank_lines": blanks,
        "total_lines": lines,
        "budget_lines": m["budget_lines"],
        "geom_lines": m["geom_lines"],
        "max_count": m["baseline_count"],
        "ratio": ratio,
        "overflow": lines > m["budget_lines"],
        "dropped": dropped,
        "filled": filled,
        "issues": issues,
        "entry_lines": per_entry,
        "must_include_missing": missing,
    }


# ----------------------------------------------------------------- 主流程

def main_ppt():
    ap = argparse.ArgumentParser(description="生成本周 PPT 页副本")
    ap.add_argument("--week", help="周次标签，如 2026-W37")
    ap.add_argument("--dir", dest="report_dir", help="输出目录")
    ap.add_argument("--start", help="窗口起始日 YYYY-MM-DD")
    ap.add_argument("--end", help="窗口结束日 YYYY-MM-DD")
    ap.add_argument("--entries-file", help="显式指定候选内容文件")
    ap.add_argument("--max-entries", type=int, default=0,
                    help="最多放几条（0=不限，由行预算决定）")
    ap.add_argument("--no-fill", action="store_true",
                    help="关闭版面自适应补位（只用正式入选条目）")
    ap.add_argument("--max-entry-lines", type=int, default=MAX_ENTRY_LINES,
                    help="每条最多占几行（默认 %d；以模板版式标定，不缩字号）"
                         % MAX_ENTRY_LINES)
    ap.add_argument("--must-include", default="",
                    help="逗号分隔的必含公司名；缺席即记为校验问题（用于本周重点公司）")
    args = ap.parse_args()

    week = args.week or os.environ.get("REPORT_WEEK")
    start = args.start or os.environ.get("REPORT_START")
    end = args.end or os.environ.get("REPORT_END")
    rdir = args.report_dir or os.environ.get("REPORT_DIR")

    week, s, e = resolve_window(week, end)
    if start:
        s = dt.date.fromisoformat(start)
    if not rdir:
        rdir = os.path.join(OUT_ROOT, week)
    os.makedirs(rdir, exist_ok=True)

    log("=" * 62)
    log("地产周报 PPT 页生成")
    log("=" * 62)
    log("周次      : %s" % week)
    log("窗口      : %s ~ %s" % (s.isoformat(), e.isoformat()))
    log("输出目录  : %s" % rdir)

    if not os.path.isfile(TEMPLATE_PPT):
        log("[错误] 模板不存在：%s" % TEMPLATE_PPT)
        return 3
    with open(TEMPLATE_PPT, "rb") as f:
        tpl_sha_before = hashlib.sha256(f.read()).hexdigest()

    m = measure_template()
    m["max_entry_lines"] = args.max_entry_lines
    if args.max_entries > 0:
        m["baseline_count"] = args.max_entries
    log("模板正文  : %d 条 / %gpt / 每行 %d 半角单位"
        % (m["baseline_count"], m["font_pt"], m["units_per_line"]))
    log("行容量    : 几何 %d 行 / 模板自用 %d 行 → 预算 %d 行"
        % (m["geom_lines"], m["tpl_lines"], m["budget_lines"]))
    log("单条上限  : %d 行（= %d 半角单位 ≈ %d 个汉字，含标题）"
        % (m["max_entry_lines"], m["units_per_line"] * m["max_entry_lines"],
           m["units_per_line"] * m["max_entry_lines"] // 2))

    entries, src = load_entries(rdir, args.entries_file)
    log("内容来源  : %s" % src)
    if not entries:
        log("[跳过] 无可用候选内容，未生成 PPT。")
        log("        请准备 PPT候选内容.md（格式见脚本头注释）")
        _write_note(rdir, week, s, e, src, [], [], None, tpl_sha_before, None, m)
        return 0

    formal = [x for x in entries if x["tier"] in (TIER_FORMAL, TIER_FILLED)]
    backup = [x for x in entries if x["tier"] == TIER_BACKUP]
    log("二次筛选  : 正式入选 %d 条 / 备选 %d 条" % (len(formal), len(backup)))
    if not formal:
        log("[提醒] PPT候选内容.md 里没有标 [正式入选] 的条目，本次不写入任何内容。")

    # 先压缩、再按行预算定条数：宁可多保留几家公司，也不让某一条变长。
    records, still_over = compress_entries(formal, m)
    if records:
        log("单条压缩  : %d 条超过 %d 行，已按「先砍末尾次要分句、再缩短点评」压到上限："
            % (len(records), m["max_entry_lines"]))
        for i, lead, before, after, dropped in records:
            log("            #%d %s  %d 行 → %d 行" % (i, lead[:22], before, after))
            for where, text in dropped:
                log("                 -%s：%s" % (where, text[:52]))
    if still_over:
        log("  [警告] 第 %s 条压到上限后仍超行 —— 小标题本身过长，需人工改写标题"
            % "、".join(map(str, still_over)))

    chosen, _, _ = select(formal, m)
    dropped = len(formal) - len(chosen)
    if dropped > 0:
        log("容量筛选  : 正式入选保留 %d 条，其余 %d 条因页面放不下未写入"
            % (len(chosen), dropped))
        log("            按规则应先合并同公司内容，再减少低优先级条目；")
        log("            这些条目已在 PPT候选内容.md 里，可直接移到 [备选] 作为人工替换项。")

    # 版面自适应：只在正式入选全部放得下、且下半页明显留白时才补位
    pool = list(backup)
    if args.no_fill:
        filled, stop = [], "已用 --no-fill 关闭"
    elif dropped:
        filled, stop = [], "正式入选已超容量，不补位"
    else:
        filled, stop = fill_from_backup(chosen, pool, m)
    if filled:
        log("版面自适应: 按重要性从备选池补入 %d 条：" % len(filled))
        for x in filled:
            log("            + %s" % x["lead"])
        log("            停止补位：%s" % stop)
        # 补位条目同样受单条行数上限约束
        rec2, so2 = compress_entries(filled, m)
        for i, lead, before, after, _d in rec2:
            log("            [压缩] %s  %d 行 → %d 行" % (lead[:22], before, after))
        if so2:
            log("            [警告] 补位条目仍超行：%s" % "、".join(map(str, so2)))
    else:
        log("版面自适应: 未补入条目 —— %s" % stop)

    # 把补位结果回写到 PPT候选内容.md 的标记行（只动标记，不动正文）
    if filled and (args.entries_file is None):
        for x in filled:
            x["tier"] = TIER_FILLED
        n = update_tier_markers(os.path.join(rdir, "PPT候选内容.md"), filled)
        if n:
            log("标记回写  : PPT候选内容.md 中 %d 条改为 [%s]" % (n, TIER_FILLED))

    m2 = measure_template()          # 取一份干净的再写
    write_entries(m2, chosen)

    out_name = "房地产周报_%s-%s.pptx" % (
        s.strftime("%Y.%m.%d"), e.strftime("%Y.%m.%d"))
    out_path = os.path.join(rdir, out_name)
    m2["prs"].save(out_path)
    log("已生成    : %s" % out_name)

    must = [c.strip() for c in (args.must_include or "").split(",") if c.strip()]
    chk = check(chosen, m, dropped, filled, must_include=must)
    log("行占用    : %d 行（正文 %d + 空段 %d）/ 预算 %d 行 → %.0f%%   共 %d 条"
        % (chk["total_lines"], chk["text_lines"], chk["blank_lines"],
           chk["budget_lines"], chk["ratio"] * 100, len(chosen)))
    log("每条行数  : %s"
        % " / ".join("#%d=%d行" % (i + 1, n) for i, n in enumerate(chk["entry_lines"])))
    if chk["issues"]:
        for it in chk["issues"]:
            log("  [校验] %s" % it)
    else:
        log("  [校验] 标题无括号且以句号结尾 / 无空条目 / 无禁用措辞 / 同公司已合并 / "
            "未溢出 / 单条≤%d 行 / 标题≤半行" % m["max_entry_lines"])

    with open(TEMPLATE_PPT, "rb") as f:
        intact = (tpl_sha_before == hashlib.sha256(f.read()).hexdigest())
    log("模板完整性: %s (%s)" % ("未改动 [OK]" if intact else "已被改动 [FAIL]",
                                 tpl_sha_before[:16]))

    # 回读验证：页数、条数、小标题是否红色加粗、正文是否黑色不加粗
    rb = Presentation(out_path)
    rb_tf = rb.slides[0].shapes[CONTENT_SHAPE_INDEX].text_frame
    rp = [p for p in rb_tf.paragraphs if p.runs]
    ok_bold, ok_red, ok_body, ok_brk, ok_end = True, True, True, True, True
    for p in rp:
        r0 = p.runs[0]
        if not r0.font.bold:
            ok_bold = False
        try:
            c = r0.font.color.rgb
        except (AttributeError, TypeError):
            c = None
        if c is None or str(c).upper() != "C00000":
            ok_red = False
        if len(p.runs) >= 2 and p.runs[1].font.bold:
            ok_body = False
        if RE_BRACKET.search(r0.text):
            ok_brk = False
        if not r0.text.endswith(TITLE_END):
            ok_end = False
    log("回读验证  : %d 页 / 正文 %d 条 / 小标题红色加粗 %s / 正文不加粗 %s / 无括号 %s / 以句号结尾 %s"
        % (len(rb.slides), len(rp),
           "OK" if (ok_bold and ok_red) else "FAIL", "OK" if ok_body else "FAIL",
           "OK" if ok_brk else "FAIL", "OK" if ok_end else "FAIL"))
    if len(rb.slides) != 1:
        log("  [校验] 页数不为 1，模板结构可能已变")

    _write_note(rdir, week, s, e, src, chosen, backup, chk, tpl_sha_before, out_name, m)
    log("说明文件  : PPT生成说明.md")
    log("=" * 62)
    return 0



def _write_note(rdir, week, s, e, src, entries, backup, chk, tpl_sha, out_name, m):
    lines = [
        "# PPT 生成说明",
        "",
        "| 项 | 值 |",
        "|---|---|",
        "| 周次 | %s |" % week,
        "| 窗口 | %s ~ %s |" % (s.isoformat(), e.isoformat()),
        "| 内容来源 | %s |" % src,
        "| 生成文件 | %s |" % (out_name or "（未生成）"),
        "| 模板 sha256 | %s |" % tpl_sha[:16],
        "| 模板字号 | %g pt |" % m["font_pt"],
        "| 行容量预算 | %d 行（几何 %d / 模板自用 %d）|"
        % (m["budget_lines"], m["geom_lines"], m["tpl_lines"]),
        "",
    ]
    if not entries:
        lines += [
            "**本次未生成 PPT**：无可用候选内容。",
            "",
            "请准备 `PPT候选内容.md`，条目块格式：",
            "",
            "```markdown",
            "[正式入选] 保利发展销售同比承压，可转债与拿地同步推进。",
            "事实：客观数据与进展，数字优先。",
            "点评：1-2 句，只从上面的数据推导。",
            "```",
            "",
            "小标题不得含括号，且必须以「。」结尾；句号后直接接正文，同处一个项目符号段落。",
            "脚本会按优先级依次尝试：`PPT候选内容.md` → "
            "`本周公司公告汇总.md`（机械转换）→ `本周重要资讯.md`。",
            "",
        ]
    else:
        lines += ["## 本次纳入 PPT 的条目（%d 条）" % len(entries), ""]
        elines = (chk or {}).get("entry_lines") or []
        for i, x in enumerate(entries, 1):
            tag = "（版面自适应补入）" if x["tier"] == TIER_FILLED else ""
            n = (elines[i - 1] if i - 1 < len(elines)
                 else entry_lines(x, m["units_per_line"]))
            lines.append("%d. **[%d 行]** **%s**%s %s"
                         % (i, n, x["lead"], tag, x["data"]))
        if backup:
            lines += ["", "## 备选条目（%d 条，本期未进 PPT）" % len(backup), ""]
            for x in backup:
                lines.append("- **%s** %s" % (x["lead"], x.get("reason") or ""))
        lines += ["", "## 校验", ""]
        if chk:
            el = chk.get("entry_lines") or []
            lines += [
                "- 行占用：%d / %d 行（正文 %d + 空段 %d）→ %.0f%%"
                % (chk["total_lines"], chk["budget_lines"],
                   chk["text_lines"], chk["blank_lines"], chk["ratio"] * 100),
                "- 每条行数：%s（上限 %d 行；3 行 = %d 汉字含标题）"
                % ("、".join("第%d条 %d 行" % (i + 1, n) for i, n in enumerate(el)),
                   m["max_entry_lines"],
                   m["units_per_line"] * m["max_entry_lines"] // 2),
                "- 页底留白：占用 %.0f%%，%s"
                % (chk["ratio"] * 100,
                   "下沿较满，不应再加条目" if chk["ratio"] >= BLANK_RATIO
                   else "下沿留白明显，可再补 1 条"),
                "- 溢出：%s" % ("是" if chk["overflow"] else "否"),
                "- 因容量未写入的正式入选条数：%d（仍在 PPT候选内容.md，可移作备选）"
                % chk["dropped"],
                "- 版面自适应补入：%d 条%s"
                % (len(chk["filled"] or []),
                   ("（" + "、".join(x["lead"] for x in chk["filled"]) + "）")
                   if chk["filled"] else ""),
            ]
            if chk.get("must_include_missing") is not None:
                mi = chk.get("must_include_missing")
                lines.append("- 指定必含公司：%s"
                             % ("全部在版" if not mi else "⚠️ 缺席 " + "、".join(mi)))
            if chk["issues"]:
                for it in chk["issues"]:
                    lines.append("- ⚠️ %s" % it)
            else:
                lines.append("- 标题无括号且以句号结尾 / 无空条目 / 无禁用措辞 / "
                             "同公司已合并 / 未溢出 / 单条≤%d 行 / 标题≤半行"
                             % m["max_entry_lines"])
    lines += [
        "",
        "> 本文件由 `internal\\agent_core.py ppt` 自动生成。",
        "> 仅改动副本，`01_模板` 中的模板未被改动。",
        "> 未缩小任何字号；未调整文本框位置、字体、配色或页面尺寸。",
        "> 条数不固定，由行预算与版面留白共同决定。",
        "",
    ]
    with open(os.path.join(rdir, "PPT生成说明.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))




# ==================================================================
#  统一入口
# ==================================================================
def generate_word(out_dir, force=False, only=None):
    """生成 Word 交付物。返回 0 表示成功，非 0 为退出码。

    注意：generate() 返回的是「生成了哪些文件」的明细列表，不是退出码。
    这里必须把它折算成 int —— cmd_word 是按 `rc == 0` 判断成败的，
    直接把列表当 rc 用会导致**每次都报 FAIL**（哪怕文件已经正确生成）。
    明细列表如需保留，请直接调用 generate()。
    """
    r = generate(out_dir, force=force, jobs=pick_jobs(only))
    return 0 if r else 1


def verify_word(out_dir):
    """校验 Word 产出。返回 verify() 的结果。"""
    return verify(out_dir)


def generate_ppt(week=None, report_dir=None, start=None, end=None,
                 entries_file=None, must_include=None, max_entry_lines=MAX_ENTRY_LINES,
                 no_fill=False, extra_args=None):
    """生成 PPT 页。返回 0 表示成功，非 0 为退出码。"""
    argv = []
    if week:
        argv += ['--week', week]
    if report_dir:
        argv += ['--dir', report_dir]
    if start:
        argv += ['--start', start]
    if end:
        argv += ['--end', end]
    if entries_file:
        argv += ['--entries-file', entries_file]
    if must_include:
        argv += ['--must-include', must_include]
    if max_entry_lines:
        argv += ['--max-entry-lines', str(max_entry_lines)]
    if no_fill:
        argv += ['--no-fill']
    if extra_args:
        argv += list(extra_args)
    old = sys.argv
    try:
        sys.argv = ['word_ppt'] + argv
        return main_ppt()
    finally:
        sys.argv = old


def main():
    """命令行入口：word / ppt / verify。"""
    ap = argparse.ArgumentParser(description='Word 周报 + PPT 页渲染')
    sub = ap.add_subparsers(dest='cmd')
    sub.add_parser('word', help='生成 Word')
    sub.add_parser('ppt', help='生成 PPT')
    sub.add_parser('verify', help='校验 Word')
    if len(sys.argv) > 1 and sys.argv[1] in ('word', 'ppt', 'verify'):
        return {'word': main_word, 'ppt': main_ppt, 'verify': main_word}[sys.argv[1]]()
    # 兼容老用法：不带子命令时按原脚本参数解析
    return main_word()


if __name__ == "__main__":
    sys.exit(main())
