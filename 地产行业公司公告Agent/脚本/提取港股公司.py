# -*- coding: utf-8 -*-
"""从港股公告 PDF 中提取公司名称与股份代号（只读）。

策略（从严，宁可标记待确认也不猜测）：
  1. 取前 2 页文本，逐行处理（港股公告的公司名通常"独立成行"）
  2. 抓 `股份代號：XXXX`
  3. 公司名 = 第一行"整行就是中文公司全称"且不含样板/第三方机构字样的行
  4. 置信度：高 = 有代码 + 有名称；中 = 只有其一；待确认 = 都没有

输出：04_工作区/盘点/港股公司提取.csv、港股待确认.csv
"""
import os
import re
import csv
from collections import defaultdict

SRC = r"C:\Users\15258\Desktop\实习\6.周六周报\公司公告"
OUT = r"C:\Users\15258\Desktop\地产周报Agent\04_工作区\盘点"

CJK = "\u4e00-\u9fff"
FW = "\uff08\uff09\uff3b\uff3d"
DOT = "\u00b7\uff0e\uff65"

# 必须整段剔除的样板/第三方机构片段
EXCLUDE_SUBSTR = [
    "香港交易及結算所有限公司", "香港聯合交易所有限公司",
    "香港交易及结算所有限公司", "香港联合交易所有限公司",
    "香港中央結算有限公司", "香港中央结算有限公司",
    "香港中央證券登記有限公司", "香港中央证券登记有限公司",
    "卓佳證券登記有限公司", "卓佳证券登记有限公司",
    "註冊成立", "注册成立",
    "會計師事務所", "会计师事务所", "律師事務所", "律师事务所",
]

# 名称中不得出现的字样（避免把报告标题/第三方机构当成公司名）
NAME_FORBIDDEN = [
    "本公司", "本集團", "本集团", "附屬", "附属",
    "公告", "報告", "报告", "目錄", "目录", "主席", "年報", "年报", "中期",
    "季度", "摘要", "聲明", "声明", "職能", "职能", "名單", "名单",
    "會計師", "会计师", "律師", "律师", "證券登記", "证券登记", "代理人",
    "環境衛生", "环境卫生", "電影", "电影", "製片", "制片", "銀行", "银行",
]

RE_CODE = re.compile(r"股份代號\s*[：:﹕]\s*(\d{3,5})(?!\d)")
RE_CODE_LOOSE = re.compile(r"股份代號[^0-9]{0,6}(\d{3,5})(?!\d)")
RE_CODE_FIELD = re.compile(r"證券?代號[^0-9]{0,12}(\d{3,5})(?!\d)")
RE_CODE_EN = re.compile(r"Stock\s*Code[^0-9]{0,6}(\d{3,5})(?!\d)", re.I)
# 整行就是中文公司全称
RE_NAME_LINE = re.compile(r"^[%s%s%s]{4,30}(?:有限公司|股份有限公司)$" % (CJK, FW, DOT))


def clean(s):
    return re.sub(r"[\s\u3000]+", "", s or "")


def bad_name(name):
    if len(name) < 4 or len(name) > 30:
        return True
    for bad in EXCLUDE_SUBSTR + NAME_FORBIDDEN:
        if bad in name:
            return True
    return False


def extract(path):
    import pymupdf
    doc = pymupdf.open(path)
    try:
        lines = []
        for i in range(min(2, doc.page_count)):
            lines.extend(doc[i].get_text().splitlines())
    finally:
        doc.close()

    raw = "\n".join(lines)

    # --- 股份代号：在原始文本上匹配，避免数字与后文粘连 ---
    code = ""
    m = (RE_CODE.search(raw) or RE_CODE_LOOSE.search(raw)
         or RE_CODE_FIELD.search(raw) or RE_CODE_EN.search(raw))
    if m:
        code = m.group(1).zfill(5) if len(m.group(1)) < 5 else m.group(1)

    # --- 公司名规则 A：月報表结构化字段「公司名稱:」 ---
    name = ""
    for i, line in enumerate(lines):
        c = clean(line)
        if c in ("公司名稱:", "公司名稱：", "公司名称:", "公司名称："):
            for j in range(i + 1, min(i + 4, len(lines))):
                cand = clean(lines[j])
                if not cand:
                    continue
                cand = re.sub(r"[（(].*$", "", cand).strip()
                if RE_NAME_LINE.match(cand) and not bad_name(cand):
                    name = cand
                break
            if name:
                break

    # --- 公司名规则 B：整行即公司全称 ---
    if name:
        return code, name
    for line in lines:
        c = clean(line)
        if not c or not RE_NAME_LINE.match(c):
            continue
        if bad_name(c):
            continue
        name = c
        break
    return code, name


def main():
    os.makedirs(OUT, exist_ok=True)
    recs = []
    for root, dirs, files in os.walk(SRC):
        rel_root = os.path.relpath(root, SRC)
        if "香港" not in rel_root:
            continue
        week = rel_root.split(os.sep)[0]
        for fn in sorted(files):
            if not fn.lower().endswith(".pdf"):
                continue
            stem = os.path.splitext(fn)[0]
            date = stem.split("：")[0] if "：" in stem else ""
            title = stem.split("：", 1)[1] if "：" in stem else stem
            try:
                code, name = extract(os.path.join(root, fn))
            except Exception as e:
                code, name = "", ""
                title = "%s [ERROR:%s]" % (title, e)
            # 置信度以"公司名是否可靠识别"为准；缺代码只标注待补充
            conf = "高" if name else "待确认"
            note = "" if code else "股份代号待补充"
            recs.append({"周": week, "日期": date, "股份代号": code,
                         "公司名称": name, "置信度": conf, "标题": title,
                         "文件": fn})

    with open(os.path.join(OUT, "港股公司提取.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["周", "日期", "股份代号", "公司名称",
                                          "置信度", "标题", "文件"])
        w.writeheader()
        for r in sorted(recs, key=lambda x: (x["周"], x["日期"], x["公司名称"])):
            w.writerow(r)

    low = [r for r in recs if r["置信度"] != "高"]
    with open(os.path.join(OUT, "港股待确认.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["周", "日期", "股份代号", "公司名称",
                                          "置信度", "标题", "文件"])
        w.writeheader()
        for r in low:
            w.writerow(r)

    agg = defaultdict(lambda: {"code": set(), "count": 0, "weeks": set()})
    for r in recs:
        if r["置信度"] != "高":
            continue
        e = agg[r["公司名称"]]
        e["code"].add(r["股份代号"])
        e["count"] += 1
        e["weeks"].add(r["周"])

    print("港股公告总数: %d" % len(recs))
    print("高置信度: %d 条，覆盖 %d 家公司"
          % (sum(1 for r in recs if r["置信度"] == "高"), len(agg)))
    print("待确认: %d 条" % len(low))
    print("\n--- 高置信度公司（按公告数排序）---")
    for name, e in sorted(agg.items(), key=lambda kv: -kv[1]["count"]):
        print("%-8s %-26s %2d次  %s" % ("/".join(sorted(e["code"])),
                                        name, e["count"],
                                        ",".join(sorted(e["weeks"]))))
    if low:
        print("\n--- 待确认（全部 %d 条）---" % len(low))
        for r in low:
            print("  [%s] %s | 代号=%-6s 名=%-22s | %s" %
                  (r["置信度"], r["周"], r["股份代号"] or "-",
                   (r["公司名称"] or "-")[:22], r["文件"][:52]))


if __name__ == "__main__":
    main()
