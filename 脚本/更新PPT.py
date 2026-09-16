# -*- coding: utf-8 -*-
"""基于固定模板生成「本周 PPT 页」副本。

用法:
    python 脚本\\更新PPT.py --week 2026-W37
    python 脚本\\更新PPT.py --dir 05_输出\\2026-W37 --start 2026-09-05 --end 2026-09-11
    python 脚本\\更新PPT.py --week 2026-W37 --entries-file 某文件.md

不传参数时，读取 run_weekly.ps1 注入的环境变量
REPORT_WEEK / REPORT_DIR / REPORT_START / REPORT_END。

模板:
    01_模板\\房地产周报PPT模板.pptx   —— 只有 1 页，只读，永不修改

输出:
    05_输出\\<YYYY-Wxx>\\房地产周报_YYYY.MM.DD-YYYY.MM.DD.pptx
    05_输出\\<YYYY-Wxx>\\PPT生成说明.md          （校验结果留痕）
    05_输出\\<YYYY-Wxx>\\PPT候选内容.md          （**只改层级标记行**，正文原样保留）

内容来源（按优先级，PPT 是**二次筛选**）:
    1. PPT候选内容.md        —— 已按 PPT 体例写好，且已标好 [正式入选] / [备选]
    2. 本周公司公告汇总.md    —— 机械转换兜底（导语=公司+标题），需人工润色
    3. 本周重要资讯.md

    Word（本周公司公告周报.docx / 本周公司公告汇总.md）不分「正式版 / 备选版」；
    分层**只发生在 PPT 层**：`PPT候选内容.md` 里每条都用 `[正式入选]` / `[备选]` 标记。

PPT候选内容.md 条目块格式（块间空行分隔）:
    [正式入选] 保利发展销售同比承压，可转债与拿地同步推进。
    事实：客观数据与进展，数字优先。
    点评：1-2 句，只从上面的数据推导。

    [备选] 九龙仓置业中期派息同比增长42%，出售资产降杠杆。
    事实：……
    点评：……
    理由：为什么本期暂不放入（页面空间不足 / 同类内容较多 / 重要性略低 / 与正式内容重复）。

【小标题硬格式】（任务卡 §PPT 标题格式）
    1. 红底加粗小标题 + 黑色正文，同处**一个项目符号段落**
    2. 小标题**不得出现任何括号**（【】（）()[]〔〕《》「」…）
    3. 小标题**必须以中文句号「。」结束**
    4. 句号后**直接接正文**，不另起一行、不再加空格
    5. 正文黑色、正常字重（点评接在事实之后，同在正文 run 内）
    违反 2/3 直接 raise，拒绝生成 —— 宁可报错也不把不合规的标题写上幻灯片。

【版面自适应】（任务卡 §PPT 版面自适应）
    正式入选条目全部入版后，若下半页**仍有明显大块空白**，且备选池里还有内容，
    则按重要性（文件内顺序）补 1 条；补完若溢出 / 过满则**撤回**。
    全程不缩字号、不压行距；条数不固定（4/5/6 条都可）。

设计原则:
    - 只改副本，不改模板（结束时校验模板 sha256 未变）
    - 保留模板原字体 / 字号 / 文本框位置 / 配色 / 页面尺寸
    - **小标题沿用模板 run[0] 自带的"红色加粗"**（模板本身就是 C00000 加粗），
      正文沿用 run[1] 的黑色，因此点评自然接在正文最后
    - 行容量由模板几何+模板自身用量标定，不自造数字
"""
import argparse
import datetime as dt
import hashlib
import os
import re
import sys
from copy import deepcopy

try:
    from pptx import Presentation
    from pptx.oxml.ns import qn
    from pptx.text.text import _Paragraph
except ImportError:
    sys.stderr.write("缺少 python-pptx，请先安装：python -m pip install python-pptx\n")
    sys.exit(2)

BASE = r"C:\Users\15258\Desktop\地产周报Agent"
HISTORY_DIR = r"C:\Users\15258\Desktop\实习"
TEMPLATE = os.path.join(BASE, "01_模板", "房地产周报PPT模板.pptx")
OUT_ROOT = os.path.join(BASE, "05_输出")

CONTENT_SHAPE_INDEX = 3      # 正文形状（实测）
LINE_SPACING_FACTOR = 1.3    # 保守行高系数（PowerPoint 中文正文约 1.2–1.5）

TIER_FORMAL = "正式入选"
TIER_FILLED = "正式入选·自动补位"   # 版面自适应从备选池补进来的
TIER_BACKUP = "备选"

TITLE_END = "。"                                    # 小标题必须以它收尾
RE_BRACKET = re.compile(r"[【】（）()\[\]〔〕{}〈〉《》「」『』]")

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


def log(msg=""):
    """GBK 控制台下也不会崩的输出。"""
    try:
        print(msg)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "ascii"
        print(msg.encode(enc, "replace").decode(enc, "replace"))


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


def load_entries(report_dir, explicit=None):
    """按优先级取候选内容，返回 (entries, 来源说明)。"""
    pool = load_company_pool(report_dir)
    if explicit:
        t = read_text(explicit)
        if t:
            return parse_candidate_md(t, pool), os.path.basename(explicit)

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
            return e, fname + note

    return [], "（无可用内容来源）"


# ----------------------------------------------------------------- 排版度量
def eff_units(s):
    """半角单位数：全角=2，半角=1（用于估算折行）。"""
    return sum(2 if ord(c) > 0x2E80 else 1 for c in s)


def measure_template():
    """读模板样板 XML + 标定行容量。所有数字都来自模板自身，不写死。"""
    prs = Presentation(TEMPLATE)
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
    }


def entry_lines(entry, units_per_line):
    """一条占用几行（只算文本行，不含条间空段）。"""
    u = eff_units(entry["lead"] + " " + entry["data"])
    return max(1, -(-u // units_per_line))


def total_lines(entries, m):
    """含条间空段的整页行占用。"""
    text_lines = sum(entry_lines(e, m["units_per_line"]) for e in entries)
    return text_lines + max(0, len(entries) - 1), text_lines


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
def check(chosen, m, dropped, filled):
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
    }


# ----------------------------------------------------------------- 主流程
def main():
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

    if not os.path.isfile(TEMPLATE):
        log("[错误] 模板不存在：%s" % TEMPLATE)
        return 3
    with open(TEMPLATE, "rb") as f:
        tpl_sha_before = hashlib.sha256(f.read()).hexdigest()

    m = measure_template()
    if args.max_entries > 0:
        m["baseline_count"] = args.max_entries
    log("模板正文  : %d 条 / %gpt / 每行 %d 半角单位"
        % (m["baseline_count"], m["font_pt"], m["units_per_line"]))
    log("行容量    : 几何 %d 行 / 模板自用 %d 行 → 预算 %d 行"
        % (m["geom_lines"], m["tpl_lines"], m["budget_lines"]))

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

    chk = check(chosen, m, dropped, filled)
    log("行占用    : %d 行（正文 %d + 空段 %d）/ 预算 %d 行 → %.0f%%   共 %d 条"
        % (chk["total_lines"], chk["text_lines"], chk["blank_lines"],
           chk["budget_lines"], chk["ratio"] * 100, len(chosen)))
    if chk["issues"]:
        for it in chk["issues"]:
            log("  [校验] %s" % it)
    else:
        log("  [校验] 标题无括号且以句号结尾 / 无空条目 / 无禁用措辞 / 同公司已合并 / 未溢出")

    with open(TEMPLATE, "rb") as f:
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
        for i, x in enumerate(entries, 1):
            tag = "（版面自适应补入）" if x["tier"] == TIER_FILLED else ""
            lines.append("%d. **%s**%s %s" % (i, x["lead"], tag, x["data"]))
        if backup:
            lines += ["", "## 备选条目（%d 条，本期未进 PPT）" % len(backup), ""]
            for x in backup:
                lines.append("- **%s** %s" % (x["lead"], x.get("reason") or ""))
        lines += ["", "## 校验", ""]
        if chk:
            lines += [
                "- 行占用：%d / %d 行（正文 %d + 空段 %d）→ %.0f%%"
                % (chk["total_lines"], chk["budget_lines"],
                   chk["text_lines"], chk["blank_lines"], chk["ratio"] * 100),
                "- 溢出：%s" % ("是" if chk["overflow"] else "否"),
                "- 因容量未写入的正式入选条数：%d（仍在 PPT候选内容.md，可移作备选）"
                % chk["dropped"],
                "- 版面自适应补入：%d 条%s"
                % (len(chk["filled"] or []),
                   ("（" + "、".join(x["lead"] for x in chk["filled"]) + "）")
                   if chk["filled"] else ""),
            ]
            if chk["issues"]:
                for it in chk["issues"]:
                    lines.append("- ⚠️ %s" % it)
            else:
                lines.append("- 标题无括号且以句号结尾 / 无空条目 / 无禁用措辞 / "
                             "同公司已合并 / 未溢出")
    lines += [
        "",
        "> 本文件由 `脚本\\更新PPT.py` 自动生成。",
        "> 仅改动副本，`01_模板` 中的模板未被改动。",
        "> 未缩小任何字号；未调整文本框位置、字体、配色或页面尺寸。",
        "> 条数不固定，由行预算与版面留白共同决定。",
        "",
    ]
    with open(os.path.join(rdir, "PPT生成说明.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


if __name__ == "__main__":
    sys.exit(main())
