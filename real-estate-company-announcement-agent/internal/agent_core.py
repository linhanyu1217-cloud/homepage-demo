# -*- coding: utf-8 -*-
r"""地产公司公告 Agent —— 核心调度层（唯一对外 Python 接口）。

主入口是项目根目录的 `地产公司公告Agent.ps1`；本模块只负责 Python 侧的实际
干活，并对外暴露 9 个公开函数（规范 §八）：

    read_documents()                    读取 docx/pptx/xlsx/pdf/txt 纯文本
    merge_announcements()               正式版 + 备选 → 本周公司公告汇总.md
    convert_traditional_to_simplified() 繁→简（只换字形）
    generate_word()                     生成 本周公司公告周报.docx / 本周重要资讯.docx
    generate_ppt()                      生成 房地产周报_<起>-<止>.pptx
    build_company_pool()                盘点公告 → 行业公司池.csv / 历史覆盖统计.csv
    validate_outputs()                  校验本周产出是否齐全
    extract_hk_companies()              港股公告 → 港股公司提取.csv / 港股待确认.csv
    check_dependencies()                只检查依赖，**绝不自动安装**

分层
    agent_core.py   调度 + 配置 + 路径解析（本文件）
    word_ppt.py     Word / PPT 渲染
    data_tools.py   读取、合并、盘点、公司池、Excel
    zh_convert.py   零依赖繁→简

进度输出约定
    每完成一步打印一行：  [STEP] [i/n] 说明
    主机（状态窗口）靠这个格式算百分比。这是 **Agent 流程步骤**的完成度，
    不是模型内部 token 完成度。

安全边界
    · 只写 04_工作区\ 与 05_输出\（+ 显式允许的 03_个人资料\）
    · 绝不写入历史实习目录（只读）
    · 不打印、不记录、不写死任何 API Key / Token；.env 只由 PowerShell 侧读取
"""
import argparse
import datetime as dt
import json
import os
import re
import shutil
import sys

# --------------------------------------------------------------- 路径（全相对）
_HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(_HERE)                      # 项目根目录

# 让 internal\ 里的模块能互相 import（脚本方式运行时 __package__ 为空，
# word_ppt / data_tools 里的相对导入会回退成绝对导入，靠的就是这一行）
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import data_tools as D                             # noqa: E402
import word_ppt as W                               # noqa: E402

try:
    import zh_convert as Z
except ImportError:                                # 极端情况下不拖垮主流程
    Z = None

CONFIG_DIR = os.path.join(ROOT, "00_配置")
AGENT_CONFIG = os.path.join(CONFIG_DIR, "Agent配置.json")
PERSONAL_CONFIG = os.path.join(CONFIG_DIR, "个人配置.json")
RULES_MD = os.path.join(CONFIG_DIR, "规则.md")

TEMPLATE_DIR = os.path.join(ROOT, "01_模板")
INPUT_ROOT = os.path.join(ROOT, "02_待处理")
OUT_ROOT = os.path.join(ROOT, "05_输出")
WORK_ROOT = os.path.join(ROOT, "04_工作区")
LOG_ROOT = os.path.join(ROOT, "日志")
RELEASE_DIR = os.path.join(ROOT, "发行版")
PERSONAL_DIR = os.path.join(ROOT, "03_个人资料")

# 待处理目录下的输入分区（规范 §四）
INPUT_PARTS = ("A股公告", "港股公告", "研报", "新闻资讯")

WEEK_DIR_RE = None      # 延迟编译，见 _week_re()

# 每周应当产出的文件（规范 §三十二 验收项）。{start}/{end} 为日期区间。
EXPECTED_OUTPUTS = (
    "本周公司公告汇总.md",
    "本周公司公告周报.docx",
    "本周重要资讯.docx",
    "房地产周报_{start}-{end}.pptx",
    "PPT候选内容.md",
    "PPT生成说明.md",
    "周报更新说明.md",
    "run_log.md",
)

DEFAULT_AGENT_CONFIG = {
    "input": {"dir": "02_待处理", "parts": list(INPUT_PARTS)},
    "output": {"dir": "05_输出"},
    "workspace": {"dir": "04_工作区"},
    "logs": {"dir": "日志"},
    "templates": {"dir": "01_模板"},
    "watcher": {"stable_minutes": 5, "max_auto_retries": 2},
    "ppt": {"approx_lines_per_item": 3},
    "language": {"hk_output": "simplified_chinese"},
}


# ==================================================================
#  1. 日志与进度
# ==================================================================
def log(msg=""):
    """GBK 控制台下也不会崩的输出。"""
    try:
        print(msg)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "ascii"
        print(msg.encode(enc, "replace").decode(enc, "replace"))


def step(i, n, desc):
    """主机状态窗口按 [STEP] [i/n] 解析进度。"""
    log("[STEP] [%d/%d] %s" % (i, n, desc))


def result(key, value):
    """给 PowerShell 解析的机器可读行。"""
    log("RESULT|%s|%s" % (key, value))


# ==================================================================
#  2. 配置
# ==================================================================
def _merge_dict(base, over):
    out = dict(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _merge_dict(out[k], v)
        else:
            out[k] = v
    return out


def load_agent_config():
    """通用配置。文件不存在/损坏时退回内置默认值，绝不因此中断。"""
    cfg = json.loads(json.dumps(DEFAULT_AGENT_CONFIG))
    if os.path.isfile(AGENT_CONFIG):
        try:
            with open(AGENT_CONFIG, encoding="utf-8") as f:
                cfg = _merge_dict(cfg, json.load(f))
        except Exception as e:                     # noqa: BLE001
            log("[警告] Agent配置.json 读取失败，改用内置默认值：%s" % e)
    return cfg


def load_personal_config():
    """个人配置。发行版**没有**这个文件，返回空 dict，一切增强功能自动关闭。"""
    if not os.path.isfile(PERSONAL_CONFIG):
        return {}
    try:
        with open(PERSONAL_CONFIG, encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception as e:                         # noqa: BLE001
        log("[警告] 个人配置.json 读取失败，按「无个人配置」运行：%s" % e)
        return {}


def apply_personal():
    """把个人配置注入内部模块。核心逻辑不依赖任何一项。

    返回生效的个人配置（可能为空 dict）。
    """
    p = load_personal_config()

    hist = p.get("historical_reference_path") or ""
    if p.get("use_personal_history") and hist and os.path.isdir(hist):
        D.HISTORY_DIR = hist
        W.HISTORY_DIR = hist
        log("[个人版] 历史参考资料目录已启用：%s" % hist)
    else:
        D.HISTORY_DIR = None
        W.HISTORY_DIR = None
        if hist and not os.path.isdir(hist):
            log("[个人版] 历史参考资料目录不存在，本次不启用：%s" % hist)

    if p.get("use_priority_company_pool"):
        log("[个人版] 重点关注公司池：已启用")
    if p.get("use_historical_coverage_stats"):
        log("[个人版] 历史覆盖统计：已启用")
    if not p:
        log("[发行版] 未发现 个人配置.json，按标准模式运行（不影响核心流程）")
    return p


def personal_enabled(key):
    """某项个人增强是否启用（缺失即 False）。"""
    return bool(load_personal_config().get(key))


# ==================================================================
#  3. 周次 / 路径
# ==================================================================
def _week_re():
    global WEEK_DIR_RE
    if WEEK_DIR_RE is None:
        import re
        WEEK_DIR_RE = re.compile(r"^W(\d{1,2})公司公告$")
    return WEEK_DIR_RE


def parse_week(tag):
    """'2026-W38' → (2026, 38)。非法则抛 ValueError。"""
    import re
    m = re.match(r"^(\d{4})-W(\d{1,2})$", (tag or "").strip())
    if not m:
        raise ValueError("周次格式应为 YYYY-Wxx，收到：%r" % tag)
    y, w = int(m.group(1)), int(m.group(2))
    if not 1 <= w <= 53:
        raise ValueError("周次超出范围：%r" % tag)
    return y, w


def week_tag(year, week):
    return "%04d-W%02d" % (year, week)


def week_range(tag):
    """ISO 周次 → (周一, 周日) 的 date。"""
    y, w = parse_week(tag)
    mon = dt.date.fromisocalendar(y, w, 1)
    return mon, mon + dt.timedelta(days=6)


def week_dir(tag):
    return os.path.join(OUT_ROOT, tag)


def input_week_dirs(part, number):
    """某个输入分区下 W<number>公司公告 的目录（不存在返回 None）。"""
    p = os.path.join(INPUT_ROOT, part, "W%d公司公告" % number)
    return p if os.path.isdir(p) else None


def dir_stat(path):
    """目录下有效文件数与总字节。跳过 Office 临时文件。"""
    n = size = 0
    for r, _d, fs in os.walk(path):
        for fn in fs:
            if fn.startswith("~$"):
                continue
            try:
                size += os.path.getsize(os.path.join(r, fn))
                n += 1
            except OSError:
                pass
    return n, size


def get_input_weeks():
    """待处理里出现过的周次号 → {周号: {分区: 目录}}。"""
    found = {}
    for part in INPUT_PARTS:
        base = os.path.join(INPUT_ROOT, part)
        if not os.path.isdir(base):
            continue
        for name in sorted(os.listdir(base)):
            m = _week_re().match(name)
            if not m:
                continue
            num = int(m.group(1))
            found.setdefault(num, {})[part] = os.path.join(base, name)
    return found


def latest_common_week():
    """最新「共同周次」：A股公告与港股公告都有的最大周号。

    两者没有交集时退回「所有分区里最大的周号」。都没有则返回 None。
    """
    found = get_input_weeks()
    if not found:
        return None
    both = [n for n, parts in found.items()
            if "A股公告" in parts and "港股公告" in parts]
    return max(both) if both else max(found)


def resolve_week_number(number):
    """周号 → 周次标签。年份取「含该周周一的那一年」的 ISO 年。"""
    today = dt.date.today()
    for y in (today.isocalendar()[0], today.isocalendar()[0] - 1,
              today.isocalendar()[0] + 1):
        try:
            mon = dt.date.fromisocalendar(y, number, 1)
        except ValueError:
            continue
        # 只要该周一落在今天前后 8 周内，就认为年份对得上
        if abs((mon - today).days) <= 56:
            return week_tag(y, number)
    y = today.isocalendar()[0]
    return week_tag(y, number)


def input_weeks_ready(number):
    """该周号在 A股公告 / 港股公告 下是否都已有目录。"""
    found = get_input_weeks().get(number, {})
    return ("A股公告" in found, "港股公告" in found)


# ==================================================================
#  4. 公开函数（规范 §八）
# ==================================================================
def read_documents(path, max_chars=4000):
    """读取单个文档的纯文本。支持 .docx/.pptx/.xlsx/.pdf/.txt/.md/.csv。"""
    if not os.path.isfile(path):
        raise FileNotFoundError("文件不存在：%s" % path)
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return D.read_docx(path, max_chars)
    if ext == ".pptx":
        return D.read_pptx(path, max_chars)
    if ext == ".xlsx":
        return D.read_xlsx(path, max_chars)
    if ext == ".pdf":
        text, _pages = D.pdf_text(path, max_chars)
        return text
    if ext in (".txt", ".md", ".csv", ".json", ".py"):
        with open(path, encoding="utf-8", errors="replace") as f:
            return f.read(max_chars)
    raise ValueError("暂不支持的文件类型：%s" % ext)


def merge_announcements(report_dir, force=False,
                        formal="本周公司公告正式版.md",
                        backup="本周公司公告备选.md",
                        out_name="本周公司公告汇总.md"):
    """正式版 + 备选 → 本周公司公告汇总.md。返回统计 dict。"""
    rdir = os.path.abspath(report_dir)
    src_f = os.path.join(rdir, formal)
    src_b = os.path.join(rdir, backup)
    out = os.path.join(rdir, out_name)

    for p in (src_f, src_b):
        if not os.path.isfile(p):
            raise FileNotFoundError("缺少输入：%s" % p)
    D.assert_writable(out)
    if os.path.exists(out) and not force:
        raise FileExistsError("目标已存在（加 force 覆盖）：%s" % out)

    with open(src_f, encoding="utf-8") as f:
        formal_days = D.parse_formal(f.read())
    with open(src_b, encoding="utf-8") as f:
        backup_days = D.parse_backup(f.read())

    # 日标题样板取自正式版本身，不另造格式
    header_fmt = "广发证券房地产每日新闻（%s.%s.%s）——李怡慧"
    for d, rows in formal_days.items():
        for c, _t, _b, ln in rows:
            if c == "__HEAD__":
                header_fmt = ln.replace(d.replace("-", "."), "%s.%s.%s")
                break
        break

    days = D.merge(formal_days, backup_days)
    text = D.render(days, header_fmt)

    n_formal = sum(1 for d in days for c, _t, _b, _l in days[d]["rows"]
                   if c != "__HEAD__")
    n_added = sum(len(days[d]["added"]) for d in days)
    n_skip = sum(len(days[d]["skipped"]) for d in days)

    with open(out, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)

    return {"out": out, "days": sorted(days), "formal": n_formal,
            "added": n_added, "skipped": n_skip,
            "bytes": os.path.getsize(out)}


def convert_traditional_to_simplified(text, force=True):
    """繁→简。只换字形，不动数字/金额/百分比/日期/代码/英文。

    返回 (简体文本, 是否发生变化)。
    """
    if Z is None or not text:
        return text, False
    new = Z.to_simplified(text, force=force)
    return new, (new != text)


def generate_word(report_dir, force=False, only=None):
    """生成 Word 交付物（本周公司公告周报.docx / 本周重要资讯.docx）。

    走的是已验证的 word_ppt.generate_word，不另起炉灶。
    返回 0 表示成功。
    """
    return W.generate_word(os.path.abspath(report_dir), force=force, only=only)


def generate_ppt(report_dir=None, week=None, start=None, end=None,
                 entries_file=None, must_include=None, max_entry_lines=None,
                 no_fill=False):
    """生成 房地产周报 PPT。max_entry_lines 默认取 Agent配置.json 的 ppt 项。"""
    if max_entry_lines is None:
        max_entry_lines = int(
            load_agent_config().get("ppt", {}).get("approx_lines_per_item", 3))
    return W.generate_ppt(week=week, report_dir=report_dir, start=start, end=end,
                          entries_file=entries_file, must_include=must_include,
                          max_entry_lines=max_entry_lines, no_fill=no_fill)


def extract_hk_companies(src=None, out_dir=None):
    """从港股公告提取公司名与股份代号。

    src 默认取个人版历史参考资料目录；没有则退化为 02_待处理\\港股公告。
    产出 港股公司提取.csv / 港股待确认.csv（保留公告原文繁体，供对照）。
    返回统计 dict。
    """
    import csv
    src = src or D.HISTORY_DIR or os.path.join(INPUT_ROOT, "港股公告")
    out_dir = out_dir or D.PAND
    if not os.path.isdir(src):
        raise FileNotFoundError("港股公告目录不存在：%s" % src)
    os.makedirs(out_dir, exist_ok=True)

    recs = []
    for root, _dirs, files in os.walk(src):
        rel_root = os.path.relpath(root, src)
        if "香港" not in rel_root and rel_root != ".":
            continue
        week = rel_root.split(os.sep)[0] if rel_root != "." else ""
        for fn in sorted(files):
            if not fn.lower().endswith(".pdf") or fn.startswith("~$"):
                continue
            stem = os.path.splitext(fn)[0]
            date = stem.split("：")[0] if "：" in stem else ""
            title = stem.split("：", 1)[1] if "：" in stem else stem
            try:
                code, name = D.extract(os.path.join(root, fn))
            except Exception as e:                 # noqa: BLE001
                code, name = "", ""
                title = "%s [ERROR:%s]" % (title, e)
            recs.append({"周": week, "日期": date, "股份代号": code,
                         "公司名称": name,
                         "置信度": "高" if name else "待确认",
                         "标题": title, "文件": fn})

    fields = ["周", "日期", "股份代号", "公司名称", "置信度", "标题", "文件"]
    dst = os.path.join(out_dir, "港股公司提取.csv")
    with open(dst, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in sorted(recs, key=lambda x: (x["周"], x["日期"], x["公司名称"])):
            w.writerow(r)

    low = [r for r in recs if r["置信度"] != "高"]
    dst_low = os.path.join(out_dir, "港股待确认.csv")
    with open(dst_low, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in low:
            w.writerow(r)

    return {"total": len(recs),
            "high": sum(1 for r in recs if r["置信度"] == "高"),
            "low": len(low), "companies": len({r["公司名称"] for r in recs if r["公司名称"]}),
            "out": dst}


def build_company_pool(src=None, out_dir=None, scan_out=None):
    """盘点公告 → 行业公司池.csv / 历史覆盖统计.csv。

    需要先有盘点中间产物（A股公司统计.csv / 港股公司提取.csv）。
    缺失时**不报错**，只提示——发行版没有历史资料，这一步是可选取增强。
    """
    import csv
    out_dir = out_dir or (PERSONAL_DIR if personal_enabled("use_priority_company_pool")
                          or os.path.isdir(PERSONAL_DIR) else D.WORK_ROOT)
    os.makedirs(out_dir, exist_ok=True)

    missing = [p for p in (os.path.join(D.PAND, "A股公司统计.csv"),
                           os.path.join(D.PAND, "港股公司提取.csv"))
               if not os.path.isfile(p)]
    if missing:
        for p in missing:
            log("[提示] 缺少盘点中间产物，跳过公司池：%s" % p)
        log("       先运行 `盘点公告`（需要历史参考资料目录）再构建公司池。")
        return {"ok": False, "missing": missing}

    a = D.load_ashare()
    hk, unresolved = D.load_hk()

    watch_by_key = {v[1]: k for k, v in D.WATCH_MATCH.items()}
    matched = set()
    pool, hist = [], []

    for code, e in a.items():
        wname = watch_by_key.get(code, "")
        is_watch = "是" if wname else "否"
        if wname:
            matched.add(wname)
        pool.append({
            "公司名称": e["名称"], "证券代码": code, "市场": "A股",
            "行业分类体系": D.TODO, "一级行业": D.TODO,
            "二级行业": D.TODO, "三级行业": D.TODO,
            "是否历史覆盖": "是", "历史出现次数": e["次数"],
            "是否重点关注": is_watch, "主要公告类型": e["类型"],
            "跟踪优先级": "高" if is_watch == "是" else "常规",
            "备注": ("重点名单匹配：%s（A股公告文件名证券简称完全一致）；" % wname
                     if wname else "") + "行业分类待补充（本地无申万分类数据源）",
        })
        hist.append({"公司名称": e["名称"], "市场": "A股", "证券代码": code,
                     "历史出现次数": e["次数"], "主要公告类型": e["类型"],
                     "来源样例": e["样例"], "备注": ""})

    import re
    for name, e in sorted(hk.items(), key=lambda kv: (-kv[1]["count"], kv[0])):
        codes = "/".join(sorted(e["codes"])) if e["codes"] else ""
        wname = D.watch_hk(name)
        is_watch = "是" if wname else "否"
        if wname:
            matched.add(wname)
        types = " / ".join("%s(%d)" % (k, v) for k, v in
                           sorted(e["types"].items(), key=lambda kv: -kv[1]))
        note = "".join(sorted(e.get("notes", [])))
        note += ("重点名单匹配：%s（港股公告全称含关键字）；" % wname) if wname else ""
        if not codes:
            note += "股份代号待补充；"
        note += "行业分类待补充（本地无恒生分类数据源）"
        pool.append({
            "公司名称": name, "证券代码": codes, "市场": "港股",
            "行业分类体系": D.TODO, "一级行业": D.TODO,
            "二级行业": D.TODO, "三级行业": D.TODO,
            "是否历史覆盖": "是", "历史出现次数": e["count"],
            "是否重点关注": is_watch, "主要公告类型": types,
            "跟踪优先级": "高" if is_watch == "是" else "常规",
            "备注": note,
        })
        hist.append({"公司名称": name, "市场": "港股", "证券代码": codes,
                     "历史出现次数": e["count"], "主要公告类型": types,
                     "来源样例": " ; ".join(e["samples"]), "备注": note})

    for wname, grp, wnote in D.WATCHLIST:
        base = re.sub(r"（.*?）", "", wname)
        if base in matched:
            continue
        note = "重点关注名单公司，本地历史公告中未见对应记录；"
        note += D.WATCH_RELATED.get(wname, "")
        note += " 代码/市场/行业分类均待确认" if wname not in D.WATCH_RELATED else ""
        pool.append({
            "公司名称": wname, "证券代码": D.TODO, "市场": D.TODO,
            "行业分类体系": D.TODO, "一级行业": D.TODO,
            "二级行业": D.TODO, "三级行业": D.TODO,
            "是否历史覆盖": "否", "历史出现次数": 0,
            "是否重点关注": "是", "主要公告类型": "",
            "跟踪优先级": "中" if grp == "建筑系" else "高",
            "备注": note + (" 名单属性：%s" % grp) + (("；" + wnote) if wnote else ""),
        })

    fields_pool = ["公司名称", "证券代码", "市场", "行业分类体系", "一级行业",
                   "二级行业", "三级行业", "是否历史覆盖", "历史出现次数",
                   "是否重点关注", "主要公告类型", "跟踪优先级", "备注"]
    fields_hist = ["公司名称", "市场", "证券代码", "历史出现次数", "主要公告类型",
                   "来源样例", "备注"]
    order = {"A股": 0, "港股": 1, D.TODO: 2}
    pool.sort(key=lambda r: (order.get(r["市场"], 3), -r["历史出现次数"], r["公司名称"]))
    hist.sort(key=lambda r: (order.get(r["市场"], 3), -r["历史出现次数"], r["公司名称"]))

    dst_pool = os.path.join(out_dir, "行业公司池.csv")
    dst_hist = os.path.join(out_dir, "历史覆盖统计.csv")
    with open(dst_pool, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields_pool)
        w.writeheader()
        w.writerows(pool)
    with open(dst_hist, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields_hist)
        w.writeheader()
        w.writerows(hist)

    return {"ok": True, "pool": len(pool), "hist": len(hist),
            "ashare": sum(1 for r in pool if r["市场"] == "A股"),
            "hk": sum(1 for r in pool if r["市场"] == "港股"),
            "watch": sum(1 for r in pool if r["是否重点关注"] == "是"),
            "matched_watch": len(matched),
            "unresolved_hk": len(unresolved),
            "out_dir": out_dir}


def validate_outputs(week):
    """校验本周产出是否齐全。返回 (ok, 说明列表)。"""
    rdir = week_dir(week)
    mon, sun = week_range(week)
    subs = {"start": mon.isoformat(), "end": sun.isoformat()}

    problems, present = [], []
    if not os.path.isdir(rdir):
        return False, ["输出目录不存在：%s" % rdir]

    for tmpl in EXPECTED_OUTPUTS:
        fn = tmpl.format(**subs)
        p = os.path.join(rdir, fn)
        if not os.path.isfile(p):
            problems.append("缺少：%s" % fn)
        elif os.path.getsize(p) == 0:
            problems.append("空文件：%s" % fn)
        else:
            present.append(fn)

    # 繁→简体检：交付的 docx/pptx 必须是简体。用幂等性判断，
    # 不用字符集粗筛（简繁同形字会误报）。
    if Z is not None and Z.available():
        for fn in ("本周公司公告周报.docx",):
            p = os.path.join(rdir, fn)
            if not os.path.isfile(p):
                continue
            try:
                txt = D.docx_text(p)
                again, _ = convert_traditional_to_simplified(txt)
                if again != txt:
                    problems.append("仍含繁体字形：%s" % fn)
            except Exception as e:                 # noqa: BLE001
                problems.append("简体体检失败 %s：%s" % (fn, e))

    return (not problems), (problems or ["全部 %d 项产出齐全" % len(EXPECTED_OUTPUTS)])


def check_dependencies():
    """只检查依赖，**绝不自动安装**。返回 (ok, 缺失列表)。"""
    missing = []
    for display, mod, _purpose in D.DEPS:
        try:
            __import__(mod)
        except ImportError:
            missing.append(display)
    return (not missing), missing


# ==================================================================
#  5. 目录初始化
# ==================================================================
def ensure_dirs():
    """建立运行所需目录。只创建，不删除任何东西。"""
    made = []
    for p in (CONFIG_DIR, INPUT_ROOT, OUT_ROOT, WORK_ROOT, LOG_ROOT,
              RELEASE_DIR, PERSONAL_DIR, TEMPLATE_DIR,
              os.path.join(WORK_ROOT, "locks"), D.PAND):
        if not os.path.isdir(p):
            os.makedirs(p, exist_ok=True)
            made.append(p)
    for part in INPUT_PARTS:
        p = os.path.join(INPUT_ROOT, part)
        if not os.path.isdir(p):
            os.makedirs(p, exist_ok=True)
            made.append(p)
    return made


def env_state():
    """只报告 .env 是否存在与缺哪些键，**绝不读取或打印任何值**。"""
    p = os.path.join(ROOT, ".env")
    if not os.path.isfile(p):
        return {"exists": False, "keys": []}
    keys = []
    try:
        with open(p, encoding="utf-8-sig") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                keys.append((k.strip(), bool(v.strip())))
    except Exception:                              # noqa: BLE001
        return {"exists": True, "keys": []}
    return {"exists": True, "keys": keys}


# ==================================================================
#  6. 命令行
# ==================================================================
def cmd_init(args):
    log("== 初始化 ==")
    for p in ensure_dirs():
        log("  已创建目录：%s" % p)
    log("  目录检查完毕")

    ok, missing = check_dependencies()
    log("")
    log("== 依赖检查 ==")
    if ok:
        log("  OK  依赖齐全")
    else:
        log("  缺少：%s" % "、".join(missing))
        log("  请自行执行：pip install -r requirements.txt")
        log("  （本 Agent 不会自动安装任何软件包）")

    st = env_state()
    log("")
    log("== API 配置 ==")
    if not st["exists"]:
        log("  未发现 .env")
        log("  请复制 .env.example 为 .env，并填写自己的 API Key。")
    else:
        empty = [k for k, filled in st["keys"] if not filled]
        log("  已发现 .env，共 %d 个配置项" % len(st["keys"]))
        if empty:
            log("  以下项尚未填写：%s" % "、".join(empty))
        log("  （内容不打印，只报告是否填写）")

    log("")
    log("完成。可执行：地产公司公告Agent.ps1 status")
    return 0 if ok else 1


def cmd_weeks(args):
    found = get_input_weeks()
    if not found:
        log("待处理目录下未发现任何 W数字公司公告 目录")
        return 1
    log("待处理里出现的周次：")
    for n in sorted(found):
        parts = sorted(found[n])
        log("  W%-3d %s" % (n, "、".join(parts)))
    n = latest_common_week()
    if n is not None:
        log("")
        log("最新共同周次：%s（W%d）" % (resolve_week_number(n), n))
        result("latest_week", resolve_week_number(n))
    return 0


def cmd_validate(args):
    ok, notes = validate_outputs(args.week)
    log("== 校验 %s ==" % args.week)
    for x in notes:
        log("  %s" % x)
    result("validate", "OK" if ok else "FAIL")
    return 0 if ok else 3


def cmd_pool(args):
    log("== 构建公司池 ==")
    r = build_company_pool()
    if not r.get("ok"):
        log("  未完成（缺中间产物）")
        return 3
    log("  公司池 %d 家（A股 %d / 港股 %d）" % (r["pool"], r["ashare"], r["hk"]))
    log("  重点关注 %d 家，历史覆盖 %d 家" % (r["watch"], r["hist"]))
    log("  已写出到：%s" % r["out_dir"])
    result("pool_total", r["pool"])
    return 0


def cmd_deps(args):
    ok, missing = check_dependencies()
    for display, mod, purpose in D.DEPS:
        mark = "OK  " if mod not in missing and _importable(mod) else "MISS"
        log("  %s %-12s %s" % (mark, display, purpose))
    if not ok:
        log("")
        log("请自行执行：pip install -r requirements.txt")
    return 0 if ok else 1


def _importable(mod):
    try:
        __import__(mod)
        return True
    except ImportError:
        return False


def cmd_read(args):
    log(read_documents(args.path, args.max_chars))
    return 0


def cmd_merge(args):
    try:
        r = merge_announcements(args.dir, force=args.force)
    except (FileNotFoundError, FileExistsError, D.MergeError) as e:
        log("[错误] %s" % e)
        return 3
    log("合并结果：正式版 %d 条 + 备选新增 %d 条（去重跳过 %d 条）"
        % (r["formal"], r["added"], r["skipped"]))
    log("已写出：%s" % r["out"])
    return 0


def cmd_simplify(args):
    # --check：只报告哪些文件需要转换，一个字都不改
    if getattr(args, "check", False):
        target = args.path
        if not target:
            log("[错误] --check 需要 --path 指定文件或目录")
            return 3
        exts = (".md", ".csv", ".txt")
        files = []
        if os.path.isdir(target):
            for r, _d, fs in os.walk(target):
                files += [os.path.join(r, f) for f in fs if f.lower().endswith(exts)]
        else:
            files = [target]

        need = []
        for p in sorted(files):
            try:
                with open(p, encoding="utf-8", errors="replace") as f:
                    t = f.read()
            except OSError as e:
                log("  ! 读取失败 %s：%s" % (p, e))
                continue
            _new, changed = convert_traditional_to_simplified(t)
            if changed:
                need.append(p)
                log("  需转换：%s" % p)
        log("")
        log("检查 %d 个文件，其中 %d 个含未转换的繁体字形。" % (len(files), len(need)))
        if need:
            log("如需就地转换：python internal\\agent_core.py simplify --write --path <该目录>")
        result("simplify_check", len(need))
        return 0

    if args.path:
        with open(args.path, encoding="utf-8", errors="replace") as f:
            text = f.read()
    else:
        text = sys.stdin.read()
    new, changed = convert_traditional_to_simplified(text)
    if args.write and args.path and changed:
        with open(args.path, "w", encoding="utf-8", newline="") as f:
            f.write(new)
        log("已就地转换：%s" % args.path)
    else:
        log(new)
    return 0


def cmd_word(args):
    d = os.path.abspath(args.dir) if args.dir else current_week_dir_guess()
    if not d:
        log("[错误] 未指定 --dir，且无法从 05_输出\\ 推断本周目录")
        return 3
    log("== 生成 Word（%s）==" % d)
    rc = generate_word(d, force=args.force, only=args.only)
    result("word", "OK" if rc == 0 else "FAIL")
    return 0 if rc == 0 else 3


def cmd_ppt(args):
    d = os.path.abspath(args.dir) if args.dir else current_week_dir_guess()
    if not d:
        log("[错误] 未指定 --dir，且无法从 05_输出\\ 推断本周目录")
        return 3
    week = args.week or os.path.basename(d.rstrip("\\/"))
    start, end = args.start, args.end
    try:
        s, e = week_range(week)
    except ValueError:
        s = e = None
    if not start:
        start = (s or dt.date.today()).isoformat()
    if not end:
        end = (e or dt.date.today()).isoformat()

    log("== 生成 PPT（%s  %s → %s）==" % (week, start, end))
    rc = generate_ppt(report_dir=d, week=week, start=start, end=end,
                      entries_file=args.entries_file,
                      must_include=args.must_include,
                      max_entry_lines=args.max_entry_lines,
                      no_fill=args.no_fill)
    result("ppt", "OK" if rc == 0 else "FAIL")
    return 0 if rc == 0 else 3


def cmd_hk(args):
    r = extract_hk_companies(src=args.src, out_dir=args.out)
    if not r.get("ok"):
        return 3
    log("港股公司 %d 家" % r.get("count", 0))
    return 0


def current_week_dir_guess():
    """05_输出\\ 下最新（按目录名）的 YYYY-Wxx 目录。找不到返回 None。"""
    if not os.path.isdir(OUT_ROOT):
        return None
    cands = sorted(d for d in os.listdir(OUT_ROOT)
                   if os.path.isdir(os.path.join(OUT_ROOT, d))
                   and re.match(r"^\d{4}-W\d{2}$", d))
    return os.path.join(OUT_ROOT, cands[-1]) if cands else None


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="agent_core", description="地产公司公告 Agent 核心调度")
    sub = ap.add_subparsers(dest="cmd")

    sub.add_parser("init", help="建立目录、检查依赖与 .env")

    p = sub.add_parser("weeks", help="列出待处理里的周次")
    p.set_defaults(func=cmd_weeks)

    p = sub.add_parser("validate", help="校验某周产出")
    p.add_argument("--week", required=True)
    p.set_defaults(func=cmd_validate)

    p = sub.add_parser("pool", help="构建行业公司池")
    p.set_defaults(func=cmd_pool)

    p = sub.add_parser("deps", help="检查依赖（不安装）")
    p.set_defaults(func=cmd_deps)

    p = sub.add_parser("read", help="读取文档纯文本")
    p.add_argument("path")
    p.add_argument("--max-chars", type=int, default=4000)
    p.set_defaults(func=cmd_read)

    p = sub.add_parser("merge", help="正式版+备选 → 汇总")
    p.add_argument("--dir", required=True)
    p.add_argument("--force", action="store_true")
    p.set_defaults(func=cmd_merge)

    p = sub.add_parser("simplify", help="繁→简（只换字形）")
    p.add_argument("--path")
    p.add_argument("--write", action="store_true")
    p.add_argument("--check", action="store_true",
                   help="只报告哪些文件需要转换，不修改任何文件")
    p.set_defaults(func=cmd_simplify)

    p = sub.add_parser("word", help="生成 Word（本周公司公告周报.docx / 本周重要资讯.docx）")
    p.add_argument("--dir", help="输出目录（默认取 05_输出\\ 下最新周次）")
    p.add_argument("--only", help="只生成指定作业，如 docx / news")
    p.add_argument("--force", action="store_true")
    p.set_defaults(func=cmd_word)

    p = sub.add_parser("ppt", help="生成 房地产周报 PPT")
    p.add_argument("--dir", help="输出目录（默认取 05_输出\\ 下最新周次）")
    p.add_argument("--week", help="周次标签，如 2026-W38")
    p.add_argument("--start", help="窗口起始日 YYYY-MM-DD")
    p.add_argument("--end", help="窗口结束日 YYYY-MM-DD")
    p.add_argument("--entries-file", help="显式指定候选内容文件")
    p.add_argument("--must-include", default="", help="必须入选的公司，逗号分隔")
    p.add_argument("--max-entry-lines", type=int, default=None,
                   help="每条目行数上限（默认读 Agent配置.json 的 ppt.approx_lines_per_item）")
    p.add_argument("--no-fill", action="store_true", help="不从备选自动补位")
    p.set_defaults(func=cmd_ppt)

    p = sub.add_parser("hk", help="从港股公告提取公司名与代码")
    p.add_argument("--src", help="公告来源目录（默认 02_待处理\\港股公告）")
    p.add_argument("--out", help="产物目录（默认 04_工作区\\盘点）")
    p.set_defaults(func=cmd_hk)

    args = ap.parse_args(argv)

    apply_personal()

    if args.cmd is None:
        ap.print_help()
        return 0
    if args.cmd == "init":
        return cmd_init(args)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
