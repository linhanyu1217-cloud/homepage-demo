# -*- coding: utf-8 -*-
"""繁体 → 简体字形转换（零依赖，用 Windows 自带 LCMapStringW）。

为什么不用 opencc / zhconv
    本项目要求"不擅自安装大型依赖"。Windows 内核自带 `LCMapStringW`
    的 `LCMAP_SIMPLIFIED_CHINESE` 标志，就是一套系统级的繁→简**字形映射表**，
    不联网、不装包、毫秒级。因此本模块不需要任何第三方库。

只做字形转换
    LCMapStringW 是**逐字符映射**，不是语言模型，因此天然满足：
      · 不改数字 / 金额 / 百分比 / 日期 / 股票代码 / 英文缩写
      · 不改公司名称的官方识别（只换字形，不换字）
      · 不对公告事实做任何扩写、缩写或改写
    例：「股東應佔溢利」→「股东应占溢利」；「本集團」→「本集团」；
        「截至二零二六年六月三十日止」→ 不变（这些字形本来就相同）。

    注意方向性：繁→简是"多对一"，安全；简→繁是"一对多"（如 干/乾/幹），
    有歧义，本模块**只做繁→简**。

用法
    命令行（统一走主入口的子命令，不要直接调本文件）：
        python internal\\agent_core.py simplify --check --path 输出\\<YYYY-Wxx>
        python internal\\agent_core.py simplify --write --path 输出\\<YYYY-Wxx>
    代码里：
        from . import zh_convert as _TC
        s = _TC.to_simplified("本集團盈利")

退出码：0 = 正常；2 = 运行环境不支持（非 Windows / API 失败）
"""
import argparse
import ctypes
import os
import sys

# LCMapStringW 的映射标志
LCMAP_SIMPLIFIED_CHINESE = 0x02000000
# 简体中文（中国）locale
LCID_ZH_CN = 0x0804

# 默认转换这些扩展名（文本类产物）
TEXT_EXTS = (".md", ".txt", ".csv", ".json")

# **对照层**：按 配置\规则.md §6.6 分层口径，这些文件保留公告原文繁体，
# 便于与原始公告逐字对齐、避免主体合并出错。默认**不转**。
#
# 注意：只对**最终中文输出**做转换是本规则的要点；把来源清单一起转掉会让
# 「这条公告来自哪份原始文件」失去逐字可比性，属于把对照层当输出层用。
DEFAULT_EXCLUDE = (
    "本周公司公告来源清单.csv",     # 公告原文对照，保留繁体
    "行业公司池.csv",               # 公司池，保留公告原文繁体全称
    "跨市场重复清单.csv",           # 文件名/哈希对照，不该动
    "历史覆盖统计.csv",
)

# 繁体常用字探测集：只用于"快速判断是否需要转换"，
# 命中任一即认为可能含繁体，走真正的转换；都不含则直接跳过（省一次 API 调用）。
# 这是**保守**的启发式：宁可多转换一次，也不漏掉繁体。
TRAD_HINT = set(
    "麼麽個們來們時實對開關會與為無過還這進說東車馬門問間聞見觀覺語認識記設計較轉輕載輪"
    "萬億產業專屬團圓國圖書寫讀專業員嗎號幣價價值資產負債額淨營運銷售購處置權區別則剛創"
    "劃劑勞動務勝勢勵勸匯區協單賣買賣點熱愛網絡統絲線練組終結經給絡絕統續維綱綜綠維網緊"
    "級紀約紅純紙級納紐線練組終結經給絡絕統續維綱綜綠維網緊級紀約紅純紙級納紐線練組終結"
    "華發綠城遠洋億達中華汽車雅居樂貝殼碧桂園融創服務永升宋都新城建設發展金融街"
    "據樓面積萬億淨潤虧損營收費稅務貸債券擔保購回註銷總額權益轉換價協議銷售額"
    # 下面两个字 LCMapStringW 不映射、必须靠 _SUPPLEMENT 补齐，
    # 若不同时加进粗筛，只含这两个字的文本会被门控直接跳过。
    "後於"
)

# LCMapStringW 的映射表**漏掉**的常用繁体字（实测确认，非推测）。
#
# 排查方法：对 100+ 个港股公告高频繁体字逐个调用 force=True 转换，凡输出与
# 输入相同者再逐个判断是「真漏」还是「简繁同形」。其余 58 个（事/任/份/例/
# 保/借/停/先/分/利/券/史/合/告/售/基/大/完/定/年/律/恢/成/所/投/抵/押/拾/
# 明/是/普/核/案/次/款/止/毫/流/港/澳/特/理/由/盈/票/租/置/股/至/落/董/行/
# 角/通/量/金）确认是简繁同形，本就该保持不变，不在下表内。
#
#   後 → 后   港股公告高频（其後 / 之後 / 午後），繁→简一对一无歧义
#   於 → 于   港股公告高频（於二零二六年… / 對於 / 由於），与 OpenCC t2s 口径一致
#
# 注意 於→于 的已知边界：作为**姓氏**的「於」在简体里写作「於」。此处按
# OpenCC t2s 的通行口径一律转「于」，与「只转换字形」的定位一致。
_SUPPLEMENT = str.maketrans({"後": "后", "於": "于"})


class _Converter(object):
    """封装 LCMapStringW 的繁→简转换器（惰性初始化）。"""

    def __init__(self):
        self.ok = False
        self._api = None
        self._err = None
        if not sys.platform.startswith("win"):
            self._err = "非 Windows 平台，LCMapStringW 不可用"
            return
        try:
            self._api = ctypes.windll.kernel32.LCMapStringW
            self._api.argtypes = [
                ctypes.c_uint32,               # LCID
                ctypes.c_uint32,               # dwMapFlags
                ctypes.c_wchar_p,              # lpSrcStr
                ctypes.c_int,                  # cchSrc
                ctypes.c_wchar_p,              # lpDestStr
                ctypes.c_int,                  # cchDest
            ]
            self._api.restype = ctypes.c_int
            # 探针：确认系统真的有这张映射表
            probe = self._raw("股東應佔溢利")
            if probe == "股东应占溢利":
                self.ok = True
            else:
                self._err = "系统映射表行为异常，探针返回 %r" % probe
        except Exception as e:                 # noqa: BLE001 - 兜底，任何异常都降级
            self._err = "初始化失败：%s" % e

    def _raw(self, text):
        """真正的 API 调用。text 不得为空。"""
        n = self._api(LCID_ZH_CN, LCMAP_SIMPLIFIED_CHINESE, text, -1, None, 0)
        if n <= 0:
            raise RuntimeError("LCMapStringW 取长度失败（返回 %d）" % n)
        buf = ctypes.create_unicode_buffer(n)
        m = self._api(LCID_ZH_CN, LCMAP_SIMPLIFIED_CHINESE, text, -1, buf, n)
        if m <= 0:
            raise RuntimeError("LCMapStringW 转换失败（返回 %d）" % m)
        return buf.value

    def convert(self, text):
        if not text:
            return text
        return self._raw(text).translate(_SUPPLEMENT)


_CONV = None


def _conv():
    global _CONV
    if _CONV is None:
        _CONV = _Converter()
    return _CONV


def available():
    """运行环境是否支持繁→简转换。"""
    return _conv().ok


def unavailable_reason():
    return _conv()._err


def has_traditional(text):
    """粗筛：文本里是否可能出现繁体字（命中即转，避免漏转）。"""
    return any(ch in TRAD_HINT for ch in text)


def to_simplified(text, force=False):
    """繁体 → 简体。环境不支持时**原样返回**，绝不抛异常打断主流程。

    force=True 时跳过 has_traditional 粗筛，强制走一次转换。
    """
    if not text:
        return text
    c = _conv()
    if not c.ok:
        return text
    if not force and not has_traditional(text):
        return text
    try:
        return c.convert(text)
    except Exception:                          # noqa: BLE001 - 转换失败不能拖垮生成流程
        return text


def convert_file(path, write=False, force=False):
    """转换单个文本文件。返回 (原文本, 新文本, 是否发生变化)。

    write=False 时只读不写（用于 --check）。
    """
    for enc in ("utf-8-sig", "utf-8"):
        try:
            with open(path, "r", encoding=enc) as f:
                old = f.read()
            break
        except UnicodeDecodeError:
            continue
    else:
        return None, None, False

    new = to_simplified(old, force=force)
    changed = (new != old)
    if changed and write:
        # 保持原编码风格：原来是 utf-8-sig 就写回 utf-8-sig
        enc = "utf-8-sig" if old != old.lstrip("﻿") else "utf-8"
        with open(path, "w", encoding=enc, newline="") as f:
            f.write(new)
    return old, new, changed


def is_excluded(path, excludes):
    """是否属于「对照层」而应跳过（按文件名匹配）。"""
    base = os.path.basename(path)
    return any(x and x in base for x in (excludes or ()))


def iter_text_files(target, excludes=None):
    """target 是文件就返回它自己（显式指定的文件不套用豁免）；
    是目录就递归返回文本类文件，**跳过对照层**。
    """
    if os.path.isfile(target):
        return [target]
    out = []
    for root, _dirs, files in os.walk(target):
        for fn in sorted(files):
            if fn.startswith("~$"):            # Office 临时锁文件，跳过
                continue
            p = os.path.join(root, fn)
            if fn.lower().endswith(TEXT_EXTS) and not is_excluded(p, excludes):
                out.append(p)
    return out


def diff_sample(old, new, limit=6):
    """列出前若干处逐字差异，便于人工核对"只换字形"。"""
    pairs, i = [], 0
    while i < len(old) and i < len(new) and len(pairs) < limit:
        if old[i] != new[i]:
            pairs.append((old[i], new[i]))
        i += 1
    return pairs


def log(msg=""):
    """GBK 控制台下也不会崩的输出。"""
    try:
        print(msg)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "ascii"
        print(msg.encode(enc, "replace").decode(enc, "replace"))


def main():
    ap = argparse.ArgumentParser(description="繁体 → 简体字形转换（零依赖）")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--check", metavar="路径", help="只报告，不写盘")
    g.add_argument("--write", metavar="路径", help="就地转换（会改写文件）")
    g.add_argument("--text", metavar="文字", help="直接转换一段文字并打印")
    ap.add_argument("--force", action="store_true",
                    help="跳过繁体粗筛，强制转换")
    ap.add_argument("--exclude", default="",
                    help="额外豁免的文件名关键字（逗号分隔），在默认对照层之上叠加")
    ap.add_argument("--no-default-exclude", action="store_true",
                    help="**不推荐**：连对照层（来源清单/公司池 CSV）一起转")
    args = ap.parse_args()

    if not available():
        sys.stderr.write("环境不支持繁→简转换：%s\n" % unavailable_reason())
        return 2

    if args.text is not None:
        log(to_simplified(args.text, force=True))
        return 0

    target = args.write or args.check
    write = bool(args.write)

    excludes = () if args.no_default_exclude else DEFAULT_EXCLUDE
    if args.exclude:
        excludes = tuple(excludes) + tuple(
            x.strip() for x in args.exclude.split(",") if x.strip())

    files = iter_text_files(target, excludes)
    if not files:
        sys.stderr.write("未找到可转换的文本文件：%s\n" % target)
        return 1

    # 明确报出被豁免的对照层文件，避免"以为转了其实没转"
    skipped = []
    if os.path.isdir(target):
        for root, _d, fs in os.walk(target):
            for fn in fs:
                p = os.path.join(root, fn)
                if (fn.lower().endswith(TEXT_EXTS) and not fn.startswith("~$")
                        and is_excluded(p, excludes)):
                    skipped.append(p)

    n_changed = 0
    for p in files:
        old, new, changed = convert_file(p, write=write, force=args.force)
        if old is None:
            log("SKIP|%s|非 UTF-8 文本" % p)
            continue
        if not changed:
            log("OK  |%s|无需转换" % p)
            continue
        n_changed += 1
        log("%s|%s|%d 处字形变化"
            % ("DONE" if write else "TODO", p,
               sum(1 for a, b in zip(old, new) if a != b)))
        for a, b in diff_sample(old, new):
            log("       %s → %s" % (a, b))

    if skipped:
        log("")
        log("已按分层规则豁免（对照层，保留公告原文繁体）：")
        for p in skipped:
            log("  -- %s" % p)

    log("")
    log("合计：%d 个文件%s" % (n_changed, "已转换" if write else "待转换"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
