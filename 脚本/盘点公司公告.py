# -*- coding: utf-8 -*-
"""历史公司公告盘点（只读）。

扫描 C:\\Users\\15258\\Desktop\\实习\\6.周六周报\\公司公告，提取：
  1. A股公告：文件名形如 `日期：代码--简称--标题.pdf` → 代码 + 简称 + 类型
  2. 聚合稿：`广发证券房地产公司公告（区间）.docx` 中的 【公司名】 → 提及次数
  3. 港股公告：文件名无公司名，单独列出待处理

输出到 04_工作区/盘点/。绝不写入「实习」目录。
"""
import os
import re
import csv
import json
import sys
from collections import defaultdict

SRC = r"C:\Users\15258\Desktop\实习\6.周六周报\公司公告"
OUT = r"C:\Users\15258\Desktop\地产周报Agent\04_工作区\盘点"

# 公告类型关键词（按优先级从高到低匹配）
TYPE_RULES = [
    ("REITs", ["reits", "reit", "不動產投資信託", "不动产投资信托", "基础设施基金"]),
    ("盈利预警", ["盈利警告", "盈利预警", "盈利預警", "业绩预亏", "预亏", "减亏"]),
    ("业绩公告", ["業績", "业绩", "中期報告", "中期报告", "年度報告", "年报", "年報",
                  "中报", "半年报", "半年度报告", "半年度報告", "半年報",
                  "季度报告", "季報", "財務報表", "财务报表", "末期股息", "中期股息"]),
    ("销售数据", ["销售", "銷售", "签约", "簽約", "合同销售", "營運數據", "运营数据",
                  "經營數據", "经营数据", "销售情况", "销售简报", "认购",
                  "经营情况简报", "經營情況簡報", "经营简报", "主要业务经营情况",
                  "主要業務經營情況", "业务经营情况", "物業銷售"]),
    ("拿地", ["拿地", "土地", "購得", "购得", "获得房地产项目", "新增土地", "新增項目",
              "竞得", "競得", "招拍挂", "供地"]),
    ("融资", ["融资", "融資", "债券", "債券", "票据", "票據", "募集", "发行", "發行",
              "可转换", "可轉換", "贷款", "貸款", "担保", "擔保", "增发", "配股",
              "中期票据", "公司债", "超短期", "资产支持", "拆借", "财务资助",
              "財務資助", "票面利率", "回售"]),
    ("债务", ["债务", "債務", "重组", "重組", "违约", "違約", "展期", "清盘", "清盤",
              "破产", "破產", "冻结", "凍結", "輪候", "轮候", "质押", "質押", "接管"]),
    ("回购", ["回购", "回購", "增持", "减持", "減持", "股份購回"]),
    ("分红", ["分红", "股息", "派息", "权益分派", "權益分派", "红利", "紅利", "派發"]),
    ("股权变动", ["股权", "股權", "股东", "股東", "收购", "收購", "要约", "要約",
                  "出售", "转让", "轉讓", "增资", "增資", "控制权", "划转", "劃轉",
                  "限制性股票", "股权激励", "股權激勵", "无偿划转"]),
    ("管理层变动", ["董事", "高管", "高级管理人员", "高級管理人員", "总裁", "總裁",
                    "辞职", "辭職", "辞任", "辭任",
                    "聘任", "委任", "离任", "離任", "行政總裁", "首席"]),
    ("重大资产交易", ["重大资产", "重大資產", "非常重大", "主要交易", "剥离", "剝離",
                      "资产处置", "資產處置"]),
    ("重大经营事项", ["诉讼", "訴訟", "仲裁", "处罚", "處罰", "监管", "監管", "复牌",
                      "復牌", "停牌", "暂停买卖", "暫停買賣", "問詢", "问询", "说明会",
                      "异常波动", "異常波動", "风险提示", "風險提示", "计提资产减值",
                      "計提資產減值", "会计政策变更", "會計政策變更"]),
]


def classify(title):
    t = title.lower()
    for label, kws in TYPE_RULES:
        for kw in kws:
            if kw.lower() in t:
                return label
    return "其他"


# A股文件名：2026-09-04：600208--衢州发展--衢州发展：总裁...pdf
RE_ASHARE = re.compile(r"^(\d{4}-\d{2}-\d{2})[：:]\s*(\d{6})--(.+?)--(.*)$")


def main():
    os.makedirs(OUT, exist_ok=True)

    a_share = defaultdict(lambda: {"code": "", "name": "", "count": 0,
                                   "types": defaultdict(int), "samples": []})
    # A股简称归一：文件名中的简称可能含空格（如 "渝 开 发"）
    def norm(s):
        return re.sub(r"\s+", "", s.strip())

    hk_files = []
    other_files = []
    rows = []

    for root, dirs, files in os.walk(SRC):
        for fn in files:
            path = os.path.join(root, fn)
            rel = os.path.relpath(path, SRC)
            week = rel.split(os.sep)[0]
            stem, ext = os.path.splitext(fn)
            ext = ext.lower()

            if fn.startswith("~$"):
                continue

            m = RE_ASHARE.match(stem)
            if m and ext in (".pdf", ".docx"):
                date, code, name, title = m.groups()
                name = norm(name)
                e = a_share[code]
                e["code"] = code
                e["name"] = name
                e["count"] += 1
                t = classify(title)
                e["types"][t] += 1
                if len(e["samples"]) < 3:
                    e["samples"].append(f"{week}/{date}/{stem[:60]}")
                rows.append({"周": week, "市场": "A股", "日期": date, "代码": code,
                             "公司": name, "类型": t, "文件名": fn})
                continue

            if "香港" in rel and ext == ".pdf":
                hk_files.append({"周": week, "文件": fn, "路径": rel,
                                 "类型": classify(stem),
                                 "日期": stem.split("：")[0] if "：" in stem else ""})
                continue

            if ext in (".pdf", ".docx") and "公司公告" in fn:
                other_files.append(rel)
            rows.append({"周": week, "市场": "其他", "日期": "", "代码": "",
                         "公司": "", "类型": classify(stem), "文件名": fn})

    # 写 A股明细
    with open(os.path.join(OUT, "A股公告明细.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["周", "市场", "日期", "代码", "公司",
                                          "类型", "文件名"])
        w.writeheader()
        for r in sorted(rows, key=lambda x: (x["周"], x["日期"], x["公司"])):
            w.writerow(r)

    # 写 A股公司统计
    with open(os.path.join(OUT, "A股公司统计.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["代码", "简称", "公告数", "主要类型", "样例"])
        for code, e in sorted(a_share.items(), key=lambda kv: -kv[1]["count"]):
            types = sorted(e["types"].items(), key=lambda kv: -kv[1])
            w.writerow([code, e["name"], e["count"],
                        " / ".join(f"{k}({v})" for k, v in types),
                        " ; ".join(e["samples"])])

    # 写港股清单（无公司名，待处理）
    with open(os.path.join(OUT, "港股公告清单.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["周", "日期", "类型", "文件", "路径"])
        w.writeheader()
        for r in sorted(hk_files, key=lambda x: (x["周"], x["日期"])):
            w.writerow(r)

    print("A股公司数: %d" % len(a_share))
    print("A股公告数: %d" % sum(e["count"] for e in a_share.values()))
    print("港股公告数: %d" % len(hk_files))
    print("其他文件: %d" % len(other_files))
    print("\n--- A股公司按公告数排序（前40）---")
    for code, e in sorted(a_share.items(), key=lambda kv: -kv[1]["count"])[:40]:
        types = sorted(e["types"].items(), key=lambda kv: -kv[1])
        print("%-8s %-12s %3d  %s" % (code, e["name"], e["count"],
                                      "/".join(k for k, _ in types)))


if __name__ == "__main__":
    main()
