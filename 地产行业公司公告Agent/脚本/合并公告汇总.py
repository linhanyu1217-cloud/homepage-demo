# -*- coding: utf-8 -*-
"""【一次性迁移工具 · W37 用】把旧的「正式版 + 备选」合并成新的《本周公司公告汇总.md》。

背景
----
2026-09-16 起输出体例变更：

    旧：本周公司公告正式版.md  +  本周公司公告备选.md     （md 层分「正式 / 备选」）
    新：本周公司公告汇总.md                              （md 层**不分类**，完整底稿）
        分类只发生在 PPT 层：PPT候选内容.md 标【正式入选】/【备选】

W37 是用旧体例跑出来的，本脚本把已有的两份 md 合并成新体例，**不重新解析原始公告**。

合并规则
--------
1. 逐日合并：先放原「正式版」条目，再接原「备选」里同日、且未重复的条目；
2. 去重键 = 公司简称 + 公告标题（标题逐字比较，不去前缀）——
   同一公司的**不同**公告不会被误合并；
3. 原「备选」中标记「（已上提正式版）」的条目已在正式版内，按去重键自动跳过；
4. 备选条目下的 `筛选原因 / 备注 / 待人工确认` 子项**不进汇总**——
   它们属于筛选留痕，落在 `本周公司公告来源清单.csv` 的「筛选原因」列与 run_log.md；
5. 文字**逐字搬运**，不改写、不补全、不调整任何数字。

命令行
------
    python 脚本\\合并公告汇总.py --dir 05_输出\\2026-W37
    python 脚本\\合并公告汇总.py --dir 05_输出\\2026-W37 --force
"""
import argparse
import os
import re
import sys

BASE = r"C:\Users\15258\Desktop\地产周报Agent"
HISTORY_DIR = r"C:\Users\15258\Desktop\实习"
WRITABLE_ROOTS = (
    os.path.join(BASE, "04_工作区"),
    os.path.join(BASE, "05_输出"),
)

RE_FORMAL_DAY = re.compile(r"^(.+?)（(\d{4})\.(\d{2})\.(\d{2})）(.*)$")
RE_FORMAL_ENTRY = re.compile(r"^【(.+?)】(.+?)[：:](.*)$")
RE_BX_DAY = re.compile(r"^##\s+(\d{4})-(\d{2})-(\d{2})\s*$")
RE_BX_ENTRY = re.compile(r"^-\s+\*\*(.+?)\*\*\s*[：:]\s*(.*)$")
RE_BX_HEAD = re.compile(r"^【(.+?)】(.+)$")


class MergeError(Exception):
    """工具层拒绝执行。"""


def log(msg=""):
    try:
        print(msg)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "ascii"
        print(msg.encode(enc, "replace").decode(enc, "replace"))


def _abs(p):
    return os.path.abspath(p)


def _under(path, root):
    p = _abs(path).rstrip("\\/").lower()
    r = _abs(root).rstrip("\\/").lower()
    return p == r or p.startswith(r + os.sep)


def assert_writable(out):
    if _under(out, HISTORY_DIR):
        raise MergeError("拒绝写入历史实习目录（只读）：%s" % out)
    if not any(_under(out, r) for r in WRITABLE_ROOTS):
        raise MergeError("输出必须位于 04_工作区 或 05_输出 之内：%s" % out)


# ------------------------------------------------------------- 解析旧体例
def parse_formal(text):
    """本周公司公告正式版.md → OrderedDict{YYYY-MM-DD: [(公司, 标题, 正文, 原始行)]}"""
    days, cur = {}, None
    for raw in text.split("\n"):
        line = raw.rstrip()
        if not line:
            continue
        if line.strip() == "【公司新闻】":
            continue
        m = RE_FORMAL_DAY.match(line)
        if m and "每日新闻" in line:
            cur = "%s-%s-%s" % (m.group(2), m.group(3), m.group(4))
            days.setdefault(cur, [])
            days[cur].append(("__HEAD__", "", "", line))
            continue
        m = RE_FORMAL_ENTRY.match(line)
        if m and cur:
            days[cur].append((m.group(1), m.group(2).strip(), m.group(3).strip(), line))
    return days


def parse_backup(text):
    """本周公司公告备选.md → OrderedDict{YYYY-MM-DD: [(公司, 标题, 正文, 原始行)]}"""
    days, cur = {}, None
    for raw in text.split("\n"):
        line = raw.rstrip()
        m = RE_BX_DAY.match(line)
        if m:
            cur = "%s-%s-%s" % (m.group(1), m.group(2), m.group(3))
            days.setdefault(cur, [])
            continue
        if line.startswith("## "):
            cur = None
            continue
        m = RE_BX_ENTRY.match(line)
        if not m or not cur:
            continue
        head = m.group(1).replace("（已上提正式版）", "").strip()
        hm = RE_BX_HEAD.match(head)
        if not hm:
            continue
        days[cur].append((hm.group(1), hm.group(2).strip(), m.group(2).strip(), line))
    return days


def key_of(company, title):
    """去重键：公司简称 + 公告标题（逐字，不去前缀，避免误合并同公司不同公告）。"""
    return re.sub(r"\s+", "", "%s|%s" % (company, title))


# ------------------------------------------------------------- 合并
def merge(formal, backup):
    days = {}
    for d in set(formal) | set(backup):
        rows = list(formal.get(d, []))
        seen = set(k for k in (key_of(c, t) for c, t, _b, _l in rows if c != "__HEAD__"))
        added, skipped = [], []
        for company, title, body, _line in backup.get(d, []):
            k = key_of(company, title)
            if k in seen:
                skipped.append((company, title))
                continue
            seen.add(k)
            rows.append((company, title, body, ""))
            added.append((company, title))
        days[d] = {"rows": rows, "added": added, "skipped": skipped}
    return days


def render(days, header_fmt):
    """输出汇总 md：沿用正式版的日标题体例，Word 渲染脚本可直接吃。"""
    out = []
    for d in sorted(days):
        head = None
        for c, t, _b, ln in days[d]["rows"]:
            if c == "__HEAD__":
                head = ln
                break
        if head is None:
            y, m, dd = d.split("-")
            head = header_fmt % (y, m, dd)
        out.append(head)
        out.append("【公司新闻】")
        for c, t, b, ln in days[d]["rows"]:
            if c == "__HEAD__":
                continue
            if ln and not ln.startswith("-"):
                out.append(ln)                      # 正式版原行，逐字搬运
            else:
                out.append("【%s】%s：%s" % (c, t, b))
        out.append("")
    while out and not out[-1].strip():
        out.pop()
    return "\n".join(out) + "\n"


def main():
    ap = argparse.ArgumentParser(description="把正式版+备选合并为本周公司公告汇总.md（一次性迁移）")
    ap.add_argument("--dir", required=True, help="输出目录，如 05_输出\\2026-W37")
    ap.add_argument("--formal", default="本周公司公告正式版.md")
    ap.add_argument("--backup", default="本周公司公告备选.md")
    ap.add_argument("--out", default="本周公司公告汇总.md")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    rdir = _abs(args.dir)
    src_f = os.path.join(rdir, args.formal)
    src_b = os.path.join(rdir, args.backup)
    out = os.path.join(rdir, args.out)

    for p in (src_f, src_b):
        if not os.path.isfile(p):
            log("[错误] 缺少输入：%s" % p)
            return 3
    assert_writable(out)
    if os.path.exists(out) and not args.force:
        log("[拒绝] 目标已存在：%s\n  如需覆盖请显式加 --force" % out)
        return 3

    formal = parse_formal(open(src_f, encoding="utf-8").read())
    backup = parse_backup(open(src_b, encoding="utf-8").read())

    # 日标题样板取自正式版本身，不另造格式
    header_fmt = "广发证券房地产每日新闻（%s.%s.%s）——李怡慧"
    for d, rows in formal.items():
        for c, _t, _b, ln in rows:
            if c == "__HEAD__":
                header_fmt = ln.replace(d.replace("-", "."), "%s.%s.%s")
                break
        break

    days = merge(formal, backup)
    text = render(days, header_fmt)

    n_formal = sum(1 for d in days for c, _t, _b, _l in days[d]["rows"] if c != "__HEAD__")
    n_added = sum(len(days[d]["added"]) for d in days)
    n_skip = sum(len(days[d]["skipped"]) for d in days)

    log("合并结果：")
    for d in sorted(days):
        n = sum(1 for c, _t, _b, _l in days[d]["rows"] if c != "__HEAD__")
        log("  %s  共 %2d 条（其中来自备选 %d 条）" % (d, n, len(days[d]["added"])))
    log("  合计 %d 条 = 正式版 %d + 备选新增 %d；按去重键跳过 %d 条"
        % (n_formal, n_formal - n_added, n_added, n_skip))
    for c, t in [x for d in sorted(days) for x in days[d]["skipped"]]:
        log("    · 跳过（已在正式版内）：【%s】%s" % (c, t))

    with open(out, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    log("已写出：%s（%d 字节）" % (out, os.path.getsize(out)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
