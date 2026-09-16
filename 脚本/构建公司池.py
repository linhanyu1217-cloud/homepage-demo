# -*- coding: utf-8 -*-
"""构建「行业公司池」与「历史覆盖统计」两张表（只读盘点结果，不联网）。

输入（本工作区，均为前序脚本产出）：
  04_工作区/盘点/A股公司统计.csv      A股：代码/简称/公告数/主要类型/样例
  04_工作区/盘点/港股公司提取.csv     港股：逐条公告（含置信度）
输出：
  04_工作区/行业公司池.csv
  04_工作区/历史覆盖统计.csv

原则：
  - 历史公告只用来打「历史覆盖标签」，不用来定义行业全集
  - 本地没有申万/恒生行业分类数据 → 一级/二级/三级行业一律写「待确认」，不猜
  - 重点房企只在「能与本地历史公告对上」时写入代码，否则标「待确认」
"""
import os
import csv
import re
import sys
from collections import defaultdict

BASE = r"C:\Users\15258\Desktop\地产周报Agent"
PAND = os.path.join(BASE, "04_工作区", "盘点")
OUT = os.path.join(BASE, "04_工作区")

TODO = "待确认"

# ---------------------------------------------------------------- 公告类型
TYPE_RULES = [
    ("REITs", ["reits", "reit", "不動產投資信託", "不动产投资信托", "基础设施基金"]),
    ("盈利预警", ["盈利警告", "盈利预警", "盈利預警", "业绩预亏", "预亏", "减亏"]),
    ("业绩公告", ["業績", "业绩", "中期報告", "中期报告", "年度報告", "年报", "年報",
                  "中报", "半年报", "半年度报告", "半年度報告", "半年報",
                  "季度报告", "季報", "財務報表", "财务报表",
                  "末期股息", "全年業績", "中期股息"]),
    ("销售数据", ["销售", "銷售", "签约", "簽約", "合同销售", "營運數據", "运营数据",
                  "經營數據", "经营数据", "销售情况", "销售简报", "认购", "物業銷售",
                  "经营情况简报", "經營情況簡報", "经营简报", "主要业务经营情况",
                  "主要業務經營情況", "业务经营情况"]),
    ("拿地", ["拿地", "土地", "購得", "购得", "获得房地产项目", "新增土地", "新增項目",
              "竞得", "競得", "招拍挂", "供地"]),
    ("融资", ["融资", "融資", "债券", "債券", "票据", "票據", "募集", "发行", "發行",
              "可转换", "可轉換", "可換股", "贷款", "貸款", "担保", "擔保", "增发",
              "配股", "中期票据", "公司债", "超短期", "资产支持", "有擔保",
              "拆借", "财务资助", "財務資助", "票面利率", "回售"]),
    ("债务", ["债务", "債務", "重组", "重組", "违约", "違約", "展期", "清盘", "清盤",
              "破产", "破產", "冻结", "凍結", "輪候", "轮候", "质押", "質押", "接管"]),
    ("回购", ["回购", "回購", "增持", "减持", "減持", "股份購回"]),
    ("分红", ["分红", "股息", "派息", "权益分派", "權益分派", "红利", "紅利", "派發"]),
    ("股权变动", ["股权", "股權", "股东", "股東", "收购", "收購", "要约", "要約",
                  "出售", "转让", "轉讓", "增资", "增資", "控制权", "股權高度集中",
                  "划转", "劃轉", "限制性股票", "股权激励", "股權激勵", "无偿划转"]),
    ("管理层变动", ["董事", "高管", "高级管理人员", "高級管理人員", "总裁", "總裁",
                    "辞职", "辭職", "辞任", "辭任",
                    "聘任", "委任", "离任", "離任", "行政總裁", "首席"]),
    ("重大资产交易", ["重大资产", "重大資產", "非常重大", "主要交易", "須予披露",
                      "须予披露", "剥离", "剝離", "资产处置", "資產處置"]),
    ("重大经营事项", ["诉讼", "訴訟", "仲裁", "处罚", "處罰", "监管", "監管", "复牌",
                      "復牌", "停牌", "暂停买卖", "暫停買賣", "問詢", "问询", "说明会",
                      "异常波动", "異常波動", "风险提示", "風險提示", "计提资产减值",
                      "計提資產減值", "会计政策变更", "會計政策變更"]),
]


def classify(title):
    t = (title or "").lower()
    for label, kws in TYPE_RULES:
        for kw in kws:
            if kw.lower() in t:
                return label
    return "其他"


def norm(s):
    return re.sub(r"\s+", "", s or "")


# ------------------------------------------------- 重点房企名单（来源：使用者提供的销售数据跟踪表）
# 字段：名称, 属性分组, 表格备注
WATCHLIST = [
    ("保利发展", "央企", ""),
    ("中海集团", "央企", ""),
    ("华润置地", "央企", ""),
    ("招商蛇口", "央企", ""),
    ("建发股份", "地方国企", ""),
    ("建发国际集团", "地方国企", ""),
    ("联发集团", "地方国企", ""),
    ("绿城中国", "混合所有制", ""),
    ("绿城中国（操盘）", "混合所有制", "同一主体不同口径"),
    ("绿城中国（代建）", "混合所有制", "同一主体不同口径"),
    ("中国金茂", "央企", ""),
    ("越秀地产", "地方国企", ""),
    ("滨江集团", "民企", ""),
    ("万科A", "混合所有制", ""),
    ("保利置业集团", "央企", ""),
    ("华发股份", "地方国企", ""),
    ("龙湖集团", "民企", ""),
    ("大华集团", "民企", ""),
    ("金地集团", "混合所有制", ""),
    ("美的置业", "民企", ""),
    ("信达地产", "央企", ""),
    ("新城控股", "民企", ""),
    ("卓越集团", "民企", ""),
    ("伟星房产", "民企", ""),
    ("首开股份", "地方国企", ""),
    ("仁恒置地", "民企", ""),
    ("大悦城", "央企", ""),
    ("城建发展", "地方国企", ""),
    ("中国铁建", "建筑系", ""),
    ("中建东孚八局", "建筑系", ""),
    ("中建壹品三局", "建筑系", ""),
    ("中国中铁", "建筑系", ""),
    ("中交房地产", "建筑系", ""),
    ("中建智地一局", "建筑系", ""),
    ("电建地产", "建筑系", ""),
    ("碧桂园", "违约房企", ""),
    ("绿地控股", "违约房企", ""),
    ("世茂集团", "违约房企", ""),
    ("融创中国", "违约房企", ""),
    ("远洋集团", "违约房企", ""),
    ("中国恒大", "违约房企", ""),
    ("旭辉控股集团", "违约房企", ""),
    ("宝龙地产", "违约房企", ""),
]

# ------------------------------------------------- 重点房企 ↔ 本地历史公告 匹配表
# 只登记「已在本工作区提取结果中实际出现」的公司；未登记者一律待确认。
# A股：公告文件名中的证券简称完全相同（代码唯一，可作键）
WATCH_MATCH = {
    "保利发展": ("A股", "600048"),
    "万科A": ("A股", "000002"),
    "招商蛇口": ("A股", "001979"),
    "华发股份": ("A股", "600325"),
    "滨江集团": ("A股", "002244"),
    "金地集团": ("A股", "600383"),
    "信达地产": ("A股", "600657"),
    "新城控股": ("A股", "601155"),
    "首开股份": ("A股", "600376"),
    "大悦城": ("A股", "000031"),
    "城建发展": ("A股", "600266"),
    "绿地控股": ("A股", "600606"),
}

# 港股：用公司全称关键字匹配。不用代码作键 —— 部分港股未提取到代码，
# 若以空代码作键会与所有「无代码」公司撞车（曾导致误标）。
WATCH_HK_KEYWORDS = {
    "保利置業集團": "保利置业集团",
    "華潤置地": "华润置地",
    "中國金茂": "中国金茂",
    "越秀地產": "越秀地产",
    "龍湖集團": "龙湖集团",
    "美的置業": "美的置业",
    "碧桂園控股": "碧桂园",          # 不含「碧桂園服務控股」（物业平台）
    "旭輝控股": "旭辉控股集团",
    "寶龍地產": "宝龙地产",
    "綠城中國": "绿城中国",
}


def watch_hk(name):
    """按关键字判断港股公司是否属于重点名单，返回重点名单中的名称或 ''。"""
    for kw, wname in WATCH_HK_KEYWORDS.items():
        if kw in name:
            return wname
    return ""

# 重点房企中，本地只找到「同集团物业/服务上市平台」而非母公司本身，需人工确认
WATCH_RELATED = {
    "融创中国": "本地仅见「融創服務控股有限公司」(01516)，为物业平台，非母公司",
    "远洋集团": "本地仅见「遠洋服務控股有限公司」(06677)，为物业平台，非母公司",
    "中国恒大": "本地仅见「恒大物業集團有限公司」(06666)，为物业平台，非母公司",
    "中海集团": "本地见「中國海外宏洋集團有限公司」并附代号00688，"
                "公司名与代号是否对应需人工复核",
}


def load_ashare():
    p = os.path.join(PAND, "A股公司统计.csv")
    out = {}
    for r in csv.DictReader(open(p, encoding="utf-8-sig")):
        out[r["代码"]] = {"名称": r["简称"], "次数": int(r["公告数"]),
                          "类型": r["主要类型"], "样例": r["样例"]}
    return out


def load_hk():
    """返回 公司名 -> {codes:set, count:int, types:Counter, samples:list}"""
    p = os.path.join(PAND, "港股公司提取.csv")
    recs = list(csv.DictReader(open(p, encoding="utf-8-sig")))

    # 代号 -> 已可靠识别出的公司名（同源互证，用于补齐缺名记录）
    code2name = defaultdict(set)
    for r in recs:
        if r["公司名称"] and r["股份代号"]:
            code2name[r["股份代号"]].add(r["公司名称"])
    code2name = {k: next(iter(v)) for k, v in code2name.items() if len(v) == 1}

    agg = defaultdict(lambda: {"codes": set(), "count": 0,
                               "types": defaultdict(int), "samples": []})
    unresolved = []
    for r in recs:
        name = r["公司名称"]
        via = ""
        if not name and r["股份代号"] in code2name:
            name = code2name[r["股份代号"]]
            via = "（公司名由同源公告的股份代号关联，待人工复核）"
        if not name:
            unresolved.append(r)
            continue
        e = agg[name]
        if r["股份代号"]:
            e["codes"].add(r["股份代号"])
        e["count"] += 1
        e["types"][classify(r["标题"])] += 1
        if len(e["samples"]) < 3:
            e["samples"].append("%s/%s/%s" % (r["周"], r["日期"], r["标题"][:50]))
        if via:
            e.setdefault("notes", set()).add(via)
    return agg, unresolved


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    a = load_ashare()
    hk, unresolved_hk = load_hk()

    # 重点房企 → 已匹配到的 (市场,代码) 集合，便于回填标签
    watch_by_key = {v[1]: k for k, v in WATCH_MATCH.items()}   # 代码 -> 重点名单名称
    matched_watch_names = set()

    pool = []          # 行业公司池
    hist = []          # 历史覆盖统计

    # ---------------- A股 ----------------
    for code, e in a.items():
        wname = watch_by_key.get(code, "")
        is_watch = "是" if wname else "否"
        if wname:
            matched_watch_names.add(wname)
        pool.append({
            "公司名称": e["名称"], "证券代码": code, "市场": "A股",
            "行业分类体系": TODO, "一级行业": TODO, "二级行业": TODO, "三级行业": TODO,
            "是否历史覆盖": "是", "历史出现次数": e["次数"],
            "是否重点关注": is_watch, "主要公告类型": e["类型"],
            "跟踪优先级": "高" if is_watch == "是" else "常规",
            "备注": ("重点名单匹配：%s（A股公告文件名证券简称完全一致）；" % wname
                     if wname else "") + "行业分类待补充（本地无申万分类数据源）",
        })
        hist.append({
            "公司名称": e["名称"], "市场": "A股", "证券代码": code,
            "历史出现次数": e["次数"], "主要公告类型": e["类型"],
            "来源样例": e["样例"], "备注": "",
        })

    # ---------------- 港股 ----------------
    for name, e in sorted(hk.items(), key=lambda kv: (-kv[1]["count"], kv[0])):
        codes = "/".join(sorted(e["codes"])) if e["codes"] else ""
        wname = watch_hk(name)
        is_watch = "是" if wname else "否"
        if wname:
            matched_watch_names.add(wname)
        types = " / ".join("%s(%d)" % (k, v) for k, v in
                           sorted(e["types"].items(), key=lambda kv: -kv[1]))
        note = "".join(sorted(e.get("notes", [])))
        note += ("重点名单匹配：%s（港股公告全称含关键字）；" % wname) if wname else ""
        if not codes:
            note += "股份代号待补充；"
        note += "行业分类待补充（本地无恒生分类数据源）"
        pool.append({
            "公司名称": name, "证券代码": codes, "市场": "港股",
            "行业分类体系": TODO, "一级行业": TODO, "二级行业": TODO, "三级行业": TODO,
            "是否历史覆盖": "是", "历史出现次数": e["count"],
            "是否重点关注": is_watch, "主要公告类型": types,
            "跟踪优先级": "高" if is_watch == "是" else "常规",
            "备注": note,
        })
        hist.append({
            "公司名称": name, "市场": "港股", "证券代码": codes,
            "历史出现次数": e["count"], "主要公告类型": types,
            "来源样例": " ; ".join(e["samples"]), "备注": note,
        })

    # ---------------- 重点房企中「无历史覆盖」的，也进池 ----------------
    # matched_watch_names：两种匹配方式各自命中的重点名单名称（避免重复建行）
    for wname, grp, wnote in WATCHLIST:
        base = re.sub(r"（.*?）", "", wname)   # 去掉「（操盘）」「（代建）」等口径后缀
        if base in matched_watch_names:
            continue
        note = "重点关注名单公司，本地历史公告中未见对应记录；"
        note += WATCH_RELATED.get(wname, "")
        note += " 代码/市场/行业分类均待确认" if wname not in WATCH_RELATED else ""
        pool.append({
            "公司名称": wname, "证券代码": TODO, "市场": TODO,
            "行业分类体系": TODO, "一级行业": TODO, "二级行业": TODO, "三级行业": TODO,
            "是否历史覆盖": "否", "历史出现次数": 0,
            "是否重点关注": "是", "主要公告类型": "",
            "跟踪优先级": "中" if grp == "建筑系" else "高",
            "备注": note + (" 名单属性：%s" % grp) + (("；" + wnote) if wnote else ""),
        })

    FIELDS_POOL = ["公司名称", "证券代码", "市场", "行业分类体系", "一级行业",
                   "二级行业", "三级行业", "是否历史覆盖", "历史出现次数",
                   "是否重点关注", "主要公告类型", "跟踪优先级", "备注"]
    FIELDS_HIST = ["公司名称", "市场", "证券代码", "历史出现次数", "主要公告类型",
                   "来源样例", "备注"]

    order = {"A股": 0, "港股": 1, TODO: 2}
    pool.sort(key=lambda r: (order.get(r["市场"], 3), -r["历史出现次数"], r["公司名称"]))
    hist.sort(key=lambda r: (order.get(r["市场"], 3), -r["历史出现次数"], r["公司名称"]))

    with open(os.path.join(OUT, "行业公司池.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS_POOL)
        w.writeheader()
        w.writerows(pool)

    with open(os.path.join(OUT, "历史覆盖统计.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS_HIST)
        w.writeheader()
        w.writerows(hist)

    # ---------------- 控制台汇总 ----------------
    print("行业公司池总行数: %d" % len(pool))
    print("  A股: %d 家" % sum(1 for r in pool if r["市场"] == "A股"))
    print("  港股: %d 家" % sum(1 for r in pool if r["市场"] == "港股"))
    print("  市场待确认: %d 家" % sum(1 for r in pool if r["市场"] == TODO))
    print("历史覆盖公司数: %d（A股 %d + 港股 %d）" % (
        len(hist), sum(1 for r in hist if r["市场"] == "A股"),
        sum(1 for r in hist if r["市场"] == "港股")))
    print("重点关注公司数: %d" % sum(1 for r in pool if r["是否重点关注"] == "是"))
    print("重点房企已匹配到历史覆盖: %d / %d（去重口径）" % (
        len(matched_watch_names),
        len({re.sub(r"（.*?）", "", w) for w, _, _ in WATCHLIST})))
    print("港股未能识别公司名的公告条数: %d" % len(unresolved_hk))
    print("\n--- 历史覆盖 TOP20 ---")
    for r in hist[:20]:
        print("  %-6s %-8s %-30s %2d次" % (r["市场"], r["证券代码"],
                                           r["公司名称"][:28], r["历史出现次数"]))
    print("\n--- 重点房企未匹配到历史覆盖 ---")
    matched = set(WATCH_MATCH)
    for wname, grp, _ in WATCHLIST:
        if wname not in matched:
            print("  %-14s [%s] %s" % (wname, grp,
                                       WATCH_RELATED.get(wname, "本地无记录")))
    # 类型分布体检
    tc = defaultdict(int)
    for r in hist:
        for part in r["主要公告类型"].split(" / "):
            m = re.match(r"(.+?)\((\d+)\)", part.strip())
            if m:
                tc[m.group(1)] += int(m.group(2))
            elif part.strip():
                tc[part.strip()] += 1
    print("\n--- 公告类型分布（A股）---")
    for k, v in sorted(tc.items(), key=lambda kv: -kv[1]):
        print("  %-10s %d" % (k, v))


if __name__ == "__main__":
    main()
