# -*- coding: utf-8 -*-
"""
阶段五：QA 数据生成（「董秘问答」训练集）
=========================================
从 `data/chunks/` 的财报原文片段生成「投资者提问 → 董秘作答」样本，
每条样本强制携带可溯源的 source（报告文件 + PDF 页码 + 原文片段）。

**只用 `.env` 里配置的那一个 LLM，不做 embedding、不碰 vector_db、不重解析 PDF。**

六类题目（训练目标是「财报披露口径」，不是「计算器」）
----------------------------------------------------
三大财务报表及财务数据（目标占 55–65%）
- `income`     损益表：收入、销售成本、毛利、营业溢利、财务支出净额、除税前溢利、
               所得税、股东应占溢利、每股盈利、投资物业公平值变动、其他全面收益
- `balance`    资产负债表：资产总值、流动资产/非流动资产、投资物业、现金及现金等价物、
               负债总额、借款、权益总额、非控股权益、应收账款等
- `cashflow`   现金流量表：营运业务/投资业务/融资业务的现金流、来自营运的现金、
               资本开支、增购投资物业、已付利息、已付税项等

财务指标及经营财务数据（目标占 20–25%）
- `metric`     已披露的财务指标：每股盈利、每股股息、债务净额、资本净负债比率、
               权益总额、利息倍数、派息比率、基本溢利等（按「指标 × 报告」去重）

其他
- `compare`    同比/期间比较：两个报告期的 source（目标占 10–15%）
- `subjective` 主观表述/管理层观点（目标 ≤10%）

**硬性口径**
1. 指标名称必须与原文表格中的名称一致，不得用同义词替换或自行归类。
2. 只允许使用原文明确披露的数字与指标；**严禁自行计算**任何增长率、毛利率、ROE、
   杠杆比率、占比或差额（本次训练学的是「披露口径」，不是算术）。
3. 每个数字都必须能在 source_text 里找到依据，否则不生成（`{"skip": true}`）。
4. `verified` / `error_type` / `review_note` 一律为**空**：空白＝尚未人工抽检，
   与「抽检后判定错误」是两件事，不要预先填 false。

用法
----
    venv\\Scripts\\python.exe scripts\\05_generate_qa.py                    # 生成 + 出报告
    venv\\Scripts\\python.exe scripts\\05_generate_qa.py --dry-run          # 只建任务池，不调 API
    venv\\Scripts\\python.exe scripts\\05_generate_qa.py --only income      # 只生成某一类
    venv\\Scripts\\python.exe scripts\\05_generate_qa.py --only metric --limit 70
    venv\\Scripts\\python.exe scripts\\05_generate_qa.py --limit 20         # 每类最多 20 条
    venv\\Scripts\\python.exe scripts\\05_generate_qa.py --workers 8
    venv\\Scripts\\python.exe scripts\\05_generate_qa.py --prepare-only     # 只用已有 qa_raw.jsonl 出报告
    venv\\Scripts\\python.exe scripts\\05_generate_qa.py --fresh            # 清空重来

可重复运行：已生成过的 job 不会重复调用 API（按 job_key 跳过），失败/跳过的会重试。
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import random
import re
import sys
import time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from difflib import SequenceMatcher
from pathlib import Path

import requests

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover
    load_dotenv = None

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# ---------------------------------------------------------------- 路径与常量

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CHUNKS_FILE = PROJECT_ROOT / "data" / "chunks" / "all_chunks.jsonl"
EXTRACTED_DIR = PROJECT_ROOT / "data" / "extracted"
QA_DIR = PROJECT_ROOT / "data" / "qa"
QA_RAW = QA_DIR / "qa_raw.jsonl"
QA_LOG = QA_DIR / "qa_generation_log.jsonl"
QA_CSV = QA_DIR / "qa_review.csv"
OUT_DIR = PROJECT_ROOT / "output"
REPORT_MD = OUT_DIR / "qa_generation_report.md"
QUALITY_MD = OUT_DIR / "qa_quality_check.md"
SAMPLE_MD = OUT_DIR / "qa_review_sample.md"

# 目标条数（约 300 条）：三大报表 58% / 指标 22% / 同比 13.3% / 主观 6.7%
TARGETS = {"income": 78, "balance": 62, "cashflow": 34,
           "metric": 66, "compare": 40, "subjective": 20}
# 任务池比目标大，用来吸收 skip / 失败
POOLS = {"income": 130, "balance": 95, "cashflow": 52,
         "metric": 96, "compare": 120, "subjective": 40}

CATEGORY_DESC = {
    "income": "损益表及收入数据",
    "balance": "资产负债表数据",
    "cashflow": "现金流量表数据",
    "metric": "财务指标",
    "compare": "同比/期间比较",
    "subjective": "主观表述/管理层观点",
}
STATEMENT_CATS = ("income", "balance", "cashflow")

DEFAULT_WORKERS = 6
MAX_TOKENS = 3000
MAX_TOKENS_RETRY = 5000
RETRIES = 3
TIMEOUT = 300
SAMPLE_SIZE = 30
DEFAULT_SEED = 20260930

# 报表科目词表（按真实披露用词，见 2023 年报表：现金流量表用的是「营运业务/投资业务/融资业务」）
INC_KW = ["收入", "销售成本", "毛利", "营业溢利", "除税前溢利", "所得税", "股东应占溢利",
          "每股盈利", "呈报溢利", "基本溢利", "投资物业公平值", "其他全面收益",
          "营业总额", "财务支出净额", "租金收入"]
BAL_KW = ["资产总值", "流动资产", "非流动资产", "投资物业", "物业、厂房及设备",
          "现金及现金等价物", "负债总额", "流动负债", "借款", "权益总额", "非控股权益",
          "公司股东应占权益", "债务净额", "银行存款", "应收款项", "无形资产"]
CF_KW = ["营运业务", "来自营运的现金", "投资业务", "融资业务", "购买物业、厂房及设备",
         "增购投资物业", "出售投资物业所得款项", "已付利息", "已收利息", "已付税项",
         "资本开支", "偿还借款", "现金及现金等价物"]

# 章节名模式
INC_SEC = re.compile(r"综合损益表|损益表|收益表|盈利表|其他全面收益表|权益变动表")
BAL_SEC = re.compile(r"综合财务状况表|财务状况表|资产负债表|权益变动表")
CF_SEC = re.compile(r"综合现金流量表|现金流量表|现金流量摘要")
NOTE_SEC = re.compile(r"财务报表附注|简明中期财务报表附注|账目附注|财务报表附註|附註")
# 财务摘要类表格：披露财务指标最集中的地方
SUMMARY_SEC = re.compile(r"财务撮要|财务摘要|年度摘要|财务概览|主要财务数据|概览|十年财务概要")
NARR_SEC = re.compile(r"主席报告|行政总裁报告|管理层论述|业务评述|财务评述|展望|市场展望|融资")

# 正文标题行（比 section 元数据可靠：语料里真实报表块的 section 常是
# 「财务日志及投资者资讯」「中期报告」这类页眉名，而不是报表名）
INC_TITLE = re.compile(r"综合损益表|损益表|收益表|综合其他全面收益表")
BAL_TITLE = re.compile(r"综合财务状况表|财务状况表|资产负债表|综合权益变动表|权益变动表")
CF_TITLE = re.compile(r"综合现金流量表|现金流量表")

# 直接排除的章节（会计政策、目录、公司治理等，不含可问答的财报数据）
# 注意：不排除「财务日志及投资者资讯」「中期报告」——真实报表块就藏在这两个 section 下，
# 目录页由 is_toc_page() 剔除。
SEC_EXCLUDE = re.compile(r"会计政策|目录|审阅报告|企业管治|董事局报告|"
                         r"可持续发展|公司资料|股东资料|词汇|释义|封底|"
                         r"内部核数师|外聘核数师|核数师报告|员审阅表现")

# 财务指标（必须与原文表格中的名称一致；只取原文明确披露的，不做换算）
# 顺序 = 建任务池时的轮询顺序：靠前的指标优先出题，保证每个指标都有覆盖。
METRIC_INDICATORS = ["每股盈利", "每股基本盈利", "每股股息", "债务净额", "资本净负债比率",
                     "利息倍数", "派息比率", "权益总额", "基本溢利", "股东应占溢利",
                     "投资物业公平值变动", "来自营运的现金", "租金收入总额",
                     "每股基本经常性溢利", "融资前的现金流出净额"]
# 摘要表优先级：财务撮要（两期对照）> 摘要/概览 > 十年财务概要（多年并列，最易串年份）
SUMMARY_PRIORITY = [
    re.compile(r"财务撮要"),
    re.compile(r"财务摘要|年度摘要|主要财务数据"),
    re.compile(r"财务概览|概览"),
    re.compile(r"十年财务概要|财务概要"),
]

# 判断是不是「目录页」（标题行 + 纯页码行交替）——目录页会把大量指标名凑在一起，必须剔除
TOC_PAGE_NUM = re.compile(r"^\d{1,3}$")


def is_toc_page(text: str, min_pairs: int = 4) -> bool:
    lines = [x.strip() for x in (text or "").split("\n") if x.strip()]
    pairs = 0
    for i in range(len(lines) - 1):
        if (TOC_PAGE_NUM.match(lines[i + 1]) and 2 <= len(lines[i]) <= 30
                and not re.search(r"\d", lines[i])):
            pairs += 1
    return pairs >= min_pairs


# 业务板块关键词（用于让对比题落在具体业务上）
SEGMENTS = ["办公楼", "零售", "酒店", "住宅", "服务式住宅", "中国内地", "香港", "美国",
            "迈阿密", "新加坡", "东南亚", "物业投资", "物业买卖", "租金收入", "资本开支",
            "净负债", "负债", "股息", "派息", "出租率", "物业组合", "投资物业"]

SYSTEM_PROMPT = """你是「太古地产投资者关系（董秘）问答」数据集的生成器。你的唯一依据是用户给出的 source_text（公司公开财报原文片段）。
本次训练的目标是让模型学会**财报的披露口径**：某个指标在报表里叫什么、属于哪张表、公司在哪个期间披露了哪个数字。它不是算术训练集。

硬性规则：
1. 只能使用 source_text 中明确写出的信息。不得使用你自己的知识，不得推测、不得补全、不得引申。
2. answer 中出现的每一个数字、金额、百分比、年份，都必须能在 source_text 中找到对应原文。数字请照抄 source_text 的写法与单位。
3. 严禁混用：不得把不同年度、不同报告、不同 source 的信息拼成一句。若给出两个 source，它们各代表一个报告期，只能分别陈述各自披露的事实，再作对照。
4. 若 source_text 不足以支撑一个有意义、可被回答的问题，输出 {"skip": true, "reason": "简短原因"}，不要硬编。
5. 只输出一个 JSON 对象。不要 markdown 代码块，不要任何解释文字。
6. 若 source_text 是多个年度/期间并列的表格（例如「十年财务概要」这种一行标签后面跟着一串数字的表），只有当你能确定每个数字与哪个年度对应时才作答；题目已指定年度的，就只回答该年度那一列，并在 answer 中写明年度。只要对应关系不确定，就 skip。宁可少一条，也不要张冠李戴。
7. **严禁自行计算**。不得计算、推算、折算或"据此得出"任何原文没有直接披露的数值，包括但不限于：同比增长率/变动幅度、毛利率、净利率、各项回报率、杠杆比率、占比、比重、周转率、总和、差额、平均值。
   - 只有在 source_text（如财务撮要表的「变幅」列）**原文已经披露**该百分比时，才可以照抄引用，并说明这是报告披露的数据。
   - 严禁出现「由此计算得出」「据此可知」「换算得」「据此推算」「可算出」这类表述。
   - 反例（绝对禁止）：原文给收入 7,297 与 6,910，自己算出「同比增长约 5.6%」。
   - 正例：原文的变幅列写了「+6%」，可以写「报告披露该项同比增长 6%」。
8. **指标名称必须与原文一致**。question 与 answer 里提到的指标名，必须照抄 source_text 表格中该行的名称（如原文写「债务净额」就不能改写成「净负债」或「有息负债」；原文写「来自营运的现金」就不能改写成「经营活动现金流」）。不得把没有明确对应关系的科目自行归类或归并。
9. 题干不要要求计算：不要问「增长了百分之多少」「毛利率是多少」这类需要自己算的问题，除非 source_text 原文确实披露了该比率。
10. **一条问答只问一件事**：不要在一句里连问多个科目或多个子项（如「A是多少？其中B和C分别是多少？」）。数字化训练样本要「一问一答一对数字」。
11. **不得臆称数据出处**：提到来源时只能用报告元数据里给出的报告名与章节名。来源章节是「十年财务概要」「财务撮要」「业务评述」时，不要说成「损益表披露」「资产负债表显示」。

问题（question）要求：模拟投资者/分析师的提问口吻，问句必须能由 source_text 回答，不要问 source_text 里没有的期间或指标。

回答（answer）要求：以公司对外公开披露的口径作答，专业、简洁、客观，**一般控制在 500 字以内**，只写回答该问题所必需的内容，不要把整张表或整段附注照抄一遍。可以自然地使用「根据公司XXXX年年度报告」「公司XX年中期报告显示」这类引出语，但不要每条都用同一个句式，也不要写成固定的公文套话。不得出现「我认为」「预计」「有望」等主观或预测性判断——除非 source_text 原文本身就有该表述，此时要说明是报告原文的表述。"""


# ---------------------------------------------------------------- 基础设施

def log(msg: str = "") -> None:
    print(msg, flush=True)


def load_llm_config() -> dict:
    """从 .env 读 LLM 配置。全程不打印 token。"""
    env_path = PROJECT_ROOT / ".env"
    if load_dotenv is not None:
        load_dotenv(env_path, override=False)
    base = os.environ.get("ANTHROPIC_BASE_URL", "").strip().rstrip("/")
    token = os.environ.get("ANTHROPIC_AUTH_TOKEN", "").strip()
    model = os.environ.get("ANTHROPIC_MODEL", "").strip()
    missing = [k for k, v in (("ANTHROPIC_BASE_URL", base),
                              ("ANTHROPIC_AUTH_TOKEN", token),
                              ("ANTHROPIC_MODEL", model)) if not v]
    if missing:
        raise SystemExit(f"[!] .env 缺少配置项：{missing}")
    return {"base": base, "token": token, "model": model,
            "temperature": 0.6, "temperature_ok": None}


def api_call(cfg: dict, system: str, user: str, max_tokens: int) -> tuple[str, dict]:
    """调一次 Messages 接口，返回 (文本, usage)。失败抛异常。"""
    payload = {
        "model": cfg["model"],
        "max_tokens": max_tokens,
        "system": system,
        "messages": [{"role": "user", "content": user}],
    }
    if cfg.get("temperature_ok") is not False:
        payload["temperature"] = cfg["temperature"]

    headers = {
        "x-api-key": cfg["token"],
        "authorization": f"Bearer {cfg['token']}",
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    resp = requests.post(cfg["base"] + "/v1/messages", headers=headers,
                         json=payload, timeout=TIMEOUT)

    # 端点不接受 temperature（思考型模型常见）→ 关掉重试一次
    if resp.status_code == 400 and "temperature" in resp.text and cfg.get("temperature_ok") is not False:
        cfg["temperature_ok"] = False
        payload.pop("temperature", None)
        resp = requests.post(cfg["base"] + "/v1/messages", headers=headers,
                             json=payload, timeout=TIMEOUT)

    if resp.status_code != 200:
        raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:300]}")

    data = resp.json()
    text = "".join(b.get("text", "") for b in data.get("content", [])
                   if b.get("type") == "text").strip()
    usage = data.get("usage", {})
    usage["_stop_reason"] = data.get("stop_reason")
    return text, usage


def probe_endpoint(cfg: dict) -> None:
    """探一次端点：确认可用、确认 temperature 是否被接受。"""
    text, usage = api_call(cfg, "只回复两个字，不要解释。", "回复：就绪", 300)
    log(f"  LLM 就绪：model={cfg['model']} | 返回「{text[:20]}」"
        f" | temperature {'接受' if cfg.get('temperature_ok') is not False else '被拒，已关闭'}")


def extract_json(text: str) -> dict | None:
    """从模型输出里抠出 JSON 对象，容忍 ```json 包裹与前后废话。"""
    if not text:
        return None
    t = re.sub(r"```(?:json)?", "", text).strip()
    i, j = t.find("{"), t.rfind("}")
    if i < 0 or j < 0:
        return None
    try:
        obj = json.loads(t[i:j + 1])
    except Exception:
        return None
    return obj if isinstance(obj, dict) else None


# ---------------------------------------------------------------- 数字核对

CN_DIGIT = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
            "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
CN_UNIT = {"十": 10, "百": 100, "千": 1000}
CN_BIG = {"万": 10 ** 4, "萬": 10 ** 4, "亿": 10 ** 8, "億": 10 ** 8}

AR_NUM_RE = re.compile(r"\d+(?:,\d{3})*(?:\.\d+)?")   # 只把半角逗号当千分位
CN_NUM_RE = re.compile(r"[零〇一二两三四五六七八九十百千萬万億亿]{2,}")
# 财报里的中文小数写法：「百分之三点二」→ 3.2
CN_DEC_RE = re.compile(r"[零〇一二两三四五六七八九十百千萬万億亿]+[点點][零〇一二三四五六七八九]+")


def cn2num(s: str) -> float | None:
    """中文数字转数值：四十九亿 → 4.9e9；二零二六 → 2026；三点二 → 3.2。解析不了返回 None。"""
    if not s:
        return None
    if "点" in s or "點" in s:                   # 小数写法
        sep = "点" if "点" in s else "點"
        int_part, _, frac_part = s.partition(sep)
        iv = cn2num(int_part) if int_part else 0.0
        if iv is None:
            return None
        digits = "".join(str(CN_DIGIT[c]) for c in frac_part if c in CN_DIGIT)
        return None if not digits else iv + float("0." + digits)
    if all(c in CN_DIGIT for c in s):          # 纯数位串（年份写法）
        if len(s) >= 4:
            return float("".join(str(CN_DIGIT[c]) for c in s))
        return None
    total = section = num = 0.0
    seen = False
    for ch in s:
        if ch in CN_DIGIT:
            num = CN_DIGIT[ch]
            seen = True
        elif ch in CN_UNIT:
            unit = CN_UNIT[ch]
            if num == 0:
                num = 1                          # 「十」= 10
            section += num * unit
            num = 0
        elif ch in CN_BIG:
            section = (section + num) * CN_BIG[ch]
            total += section
            section = num = 0
        else:
            return None
    if not seen:
        return None
    return total + section + num


def numbers_in(text: str) -> tuple[set[float], set[str]]:
    """抽出文本里的所有数值。返回 (数值集合, 原始 token 集合)。"""
    vals: set[float] = set()
    toks: set[str] = set()
    for tok in AR_NUM_RE.findall(text or ""):
        toks.add(tok)
        try:
            vals.add(round(float(tok.replace(",", "").replace("，", "")), 6))
        except ValueError:
            pass
    for tok in CN_NUM_RE.findall(text or "") + CN_DEC_RE.findall(text or ""):
        toks.add(tok)
        v = cn2num(tok)
        if v is not None:
            vals.add(round(v, 6))
    return vals, toks


def check_numbers(answer: str, source_text: str, allowed_text: str = "") -> list[str]:
    """answer 里哪些数字在来源文本里找不到依据。

    判定通过的条件（满足其一）：token 原文出现在来源文本里；或换算成数值后能对上。
    `allowed_text` 用来把随 prompt 一起给的权威元数据（年度、PDF 页码）也算作依据——
    回答里写「根据公司2014年年度报告」不算编造。

    注意：chunk 原文带硬换行，中文数字常被切断（「二十\\n四个」），所以比对时
    额外在「去掉所有空白」的版本上再查一遍，否则会把正确引用误判成编造。
    """
    if not answer:
        return []
    raw = (source_text or "") + "\n" + (allowed_text or "")
    # 去空白 + 去掉块内的【第X页】【表格上下文｜…】标记：跨页拼接会把一个数字劈成两半
    flat = re.sub(r"\s+", "", re.sub(r"【[^】]*】", "", raw))
    src_vals = numbers_in(raw)[0] | numbers_in(flat)[0]
    # 单个中文数字（「六月三十日」里的「六」）也要算作依据，否则 6月30日 会被误判
    for ch in set(raw) & set(CN_DIGIT):
        if ch != "零":
            src_vals.add(float(CN_DIGIT[ch]))
    # 报表单位是「港币百万元」，原文写 7,028 而答案写成「七十亿二千八百万」是同一笔钱，
    # 所以再放一套「原文数值 × 10^6」的对照值。
    src_vals |= {round(v * 1e6, 6) for v in list(src_vals)}
    # 日期里的中文月/日写法（「十二月三十一日」）指的是报告期，不是披露数字：只要该中文数字
    # 紧邻「月/日」且数值在 1–31 之间，就当作日期成分放过（报告期来自元数据，不在片段里）。
    date_toks: set[str] = set()
    for m in re.finditer(r"([零〇一二两三四五六七八九十]{1,3})(?=[月日])", answer or ""):
        v = cn2num(m.group(1))
        if v is not None and 1 <= v <= 31:
            date_toks.add(m.group(1))

    _, ans_toks = numbers_in(answer)
    unsupported = []
    for tok in sorted(ans_toks):
        if tok in raw or tok in flat or tok in date_toks:
            continue
        if re.fullmatch(r"\d+", tok):
            v = float(tok)
        elif re.fullmatch(r"[\d,]+", tok):
            v = round(float(tok.replace(",", "")), 6)
        else:
            v = cn2num(tok)
            if v is None:
                continue      # 解析不出数值（如「同一零售」里被切出的「一零」）→ 不当成数字
        if any(abs(v - sv) <= max(1e-6, abs(v) * 1e-9) for sv in src_vals):
            continue
        unsupported.append(tok)
    return unsupported


# ---------------------------------------------------------------- 自行计算检测

# 本次训练要的是「披露口径」，不是计算器。下面三个特征用来提示人工重点复核：
# 答案出现推导性表述 / 问题索要原文没披露的比率 / 答案出现原文没有的百分比。
DERIVE_RE = re.compile(r"计算|据此|可算出|推算|折算|换算得|得出|换言之|折合")
# 「按营业额计算的租金」「按公司股东应占溢利计算的派息比率」是原表行标签措辞，不是自行推导
LABEL_CALC_RE = re.compile(r"按[^，。；、]{0,14}计算")
RATIO_RE = re.compile(r"增长率|增幅|跌幅|毛利率|净利率|回报率|报酬率|杠杆|ROE|ROA|"
                      r"周转率|占比|比重|平均值|倍数")
PCT_RE = re.compile(r"\d+(?:\.\d+)?\s*(?:%|％|个百分点)")


def selfcalc_suspect(rec: dict) -> list[str]:
    """返回「疑似自行计算」的标签列表（启发式，仅供人工复核参考）。"""
    ans = rec.get("answer", "") or ""
    ques = rec.get("question", "") or ""
    src = " ".join(s.get("text", "") for s in (rec.get("sources") or []))
    hits: list[str] = []
    m = DERIVE_RE.search(LABEL_CALC_RE.sub("", ans))
    if m:
        hits.append(f"答案含推导性表述「{m.group(0)}」")
    m = RATIO_RE.search(ques)
    if m and m.group(0) not in src:
        hits.append(f"问题索要「{m.group(0)}」但原文未披露该指标名")
    if PCT_RE.search(ans) and not PCT_RE.search(src):
        hits.append("答案含百分比/百分点但原文未见相应披露")
    return hits


# ---------------------------------------------------------------- 数据载入

def load_chunks() -> list[dict]:
    rows = []
    with CHUNKS_FILE.open(encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if line:
                rows.append(json.loads(line))
    return rows


def load_report_index() -> dict[str, dict]:
    """source_file → {source_path, total_pages, year, report_type}，用于溯源核对。"""
    index: dict[str, dict] = {}
    for path in sorted(EXTRACTED_DIR.glob("*.json")):
        try:
            d = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        sf = d.get("source_file")
        if sf:
            index[sf] = {
                "source_path": d.get("source_path", ""),
                "total_pages": d.get("total_pages"),
                "year": str(d.get("year", "")),
                "report_type": d.get("report_type", ""),
                "extracted_file": path.name,
            }
    return index


def normalize_q(q: str) -> str:
    return re.sub(r"[^\w一-鿿]", "", q or "").lower()


def segments_in(text: str) -> set[str]:
    return {s for s in SEGMENTS if s in (text or "")}


# ---------------------------------------------------------------- 任务池

def _hits(text: str, kws: list[str]) -> int:
    return sum(1 for k in kws if k in text)


def stmt_classify(chunk: dict) -> tuple[str, int] | None:
    """把 chunk 归入 income / balance / cashflow，返回 (类别, 档位)。档位越小越优先。

    档位 1 = 报表正文；2 = 报表附注（细分科目数据）；3 = 摘要表或评述文字里的财务数据。
    """
    if chunk.get("char_count", 0) < 250:
        return None
    sec = chunk.get("section") or ""
    if SEC_EXCLUDE.search(sec):
        return None
    text = chunk["text"]
    if is_toc_page(text):                     # 目录页会把大量科目名凑在一起
        return None
    n_inc, n_bal, n_cf = _hits(text, INC_KW), _hits(text, BAL_KW), _hits(text, CF_KW)

    head = re.sub(r"【[^】]*】", " ", text)[:130]
    if CF_TITLE.search(head):                 # 正文标题行最可靠，优先
        return "cashflow", 1
    if INC_TITLE.search(head):
        return "income", 1 if "损益表" in head or "收益表" in head else 2
    if BAL_TITLE.search(head):
        return "balance", 1 if "财务状况表" in head or "资产负债表" in head else 2
    # 财务摘要/财务概览/十年财务概要这类汇总表归入 metric，不要当成某张报表的正文，
    # 否则模型会把「十年财务概要」的数字说成「损益表披露」。
    if SUMMARY_SEC.search(sec):
        return None
    if CF_SEC.search(sec):
        return "cashflow", 1
    if "现金流量摘要" in sec:
        return "cashflow", 2
    if BAL_SEC.search(sec):
        return "balance", 2
    if INC_SEC.search(sec):
        return "income", 2
    if n_cf >= 3:
        return "cashflow", 3
    if NOTE_SEC.search(sec):                  # 附注按科目词命中数归到最贴近的那张表
        cat, n = max((("balance", n_bal), ("income", n_inc)), key=lambda x: x[1])
        return (cat, 2) if n >= 2 else None
    if n_bal >= 4:
        return "balance", 3
    if n_inc >= 3:
        return "income", 3
    return None


def build_statement_jobs(chunks: list[dict], rng: random.Random,
                         need_by_cat: dict) -> list[dict]:
    """三大报表类：一个 chunk → 一条 QA。"""
    buckets: dict[str, list[tuple[int, dict]]] = defaultdict(list)
    for c in chunks:
        got = stmt_classify(c)
        if got:
            buckets[got[0]].append((got[1], c))
    for items in buckets.values():
        # 先按档位，再按长度，加一点随机抖动避免每次都挑同一条
        items.sort(key=lambda t: (t[0], -t[1]["char_count"], rng.random()))

    jobs: list[dict] = []
    for cat in STATEMENT_CATS:
        need = need_by_cat.get(cat, 0)
        if not need or not buckets.get(cat):
            continue
        groups: dict[tuple, list[dict]] = defaultdict(list)
        for _, c in buckets[cat]:
            groups[(c["year"], c["report_type"])].append(c)
        order = sorted(groups.keys(), key=lambda k: (k[0], k[1]))
        picked, idx = [], 0
        while len(picked) < need and any(idx < len(groups[k]) for k in order):
            for key in order:
                if idx < len(groups[key]) and len(picked) < need:
                    picked.append(groups[key][idx])
            idx += 1
        for c in picked:
            jobs.append({
                "job_key": f"{cat}::{c['chunk_id']}",
                "qa_type": cat,
                "chunks": [c],
                "focus": sorted(segments_in(c["text"]))[:3],
            })
    return jobs


def build_metric_jobs(chunks: list[dict], rng: random.Random, need: int) -> list[dict]:
    """财务指标类：按「已披露指标 × 报告期」建任务，天然不会重复出题。

    只取摘要表里明确出现的指标名，名称照抄原文；同一 (指标, 年份, 报告类型)
    只出一次，优先用「财务撮要」这种两期对照表，避免用多年并列的概要表。
    """
    slots: dict[tuple, tuple[int, dict]] = {}
    for c in chunks:
        sec = c.get("section") or ""
        if SEC_EXCLUDE.search(sec) or not SUMMARY_SEC.search(sec):
            continue
        if c.get("char_count", 0) < 250 or is_toc_page(c["text"]):
            continue
        prio = next((i for i, p in enumerate(SUMMARY_PRIORITY) if p.search(sec)),
                    len(SUMMARY_PRIORITY))
        for ind in METRIC_INDICATORS:
            if ind not in c["text"]:
                continue
            key = (ind, c["year"], c["report_type"])
            cur = slots.get(key)
            if cur is None or prio < cur[0] or (prio == cur[0]
                                                and c["char_count"] > cur[1]["char_count"]):
                slots[key] = (prio, c)

    # 按「指标」分组轮询（不是按报告期轮询）：否则字典序靠前的少数指标会把配额吃光，
    # 「每股盈利」「每股股息」「资本净负债比率」这类靠后的指标一条都出不来。
    by_ind: dict[str, list[tuple[int, str, str, dict]]] = defaultdict(list)
    for (ind, year, rtype), (prio, c) in slots.items():
        by_ind[ind].append((prio, year, rtype, c))
    order = [i for i in METRIC_INDICATORS if by_ind[i]]       # 指标优先级顺序
    for ind in order:
        # 每个指标只取前几期：若按年份从新到旧排，13 个指标各取 5~6 条会全落在最近三四年。
        # 改为年份「两头交替」（最新、最旧、次新、次旧…），指标题即可同时覆盖早期与近年。
        yrs = sorted({t[1] for t in by_ind[ind]}, reverse=True)
        seq, lo, hi = [], 0, len(yrs) - 1
        while lo <= hi:
            seq.append(yrs[lo]); lo += 1
            if lo <= hi:
                seq.append(yrs[hi]); hi -= 1
        rank = {y: k for k, y in enumerate(seq)}
        by_ind[ind].sort(key=lambda t: (rank[t[1]], t[2], t[0]))
    jobs, idx = [], 0
    while len(jobs) < need and any(idx < len(by_ind[i]) for i in order):
        for ind in order:
            if idx < len(by_ind[ind]) and len(jobs) < need:
                _, _, _, c = by_ind[ind][idx]
                jobs.append({
                    "job_key": f"metric::{ind}::{c['year']}::{c['report_type']}::{c['chunk_id']}",
                    "qa_type": "metric",
                    "chunks": [c],
                    "indicator": ind,
                    "focus": [ind],
                })
        idx += 1
    return jobs


YEAR_TOKEN_RE = re.compile(r"(?:20\d{2}|[零〇一二三四五六七八九]{4})")


def distinct_years(text: str) -> int:
    return len({m for m in YEAR_TOKEN_RE.findall(text or "")})


# 同比题配对族：财务撮要类是两期对照表（还带「变幅」列），质量最好；
# 叙述性章节作为补充。多年并列的表（十年财务概要等）由 distinct_years 排除。
PAIR_FAMILIES = [
    ("财务摘要表", re.compile(r"财务撮要|财务摘要|年度摘要|主要财务数据|财务概览")),
    ("业务评述", re.compile(r"业务评述")),
    ("财务评述", re.compile(r"财务评述")),
    ("主席报告", re.compile(r"主席报告")),
    ("行政总裁报告", re.compile(r"行政总裁报告")),
    ("管理层论述", re.compile(r"管理层论述")),
    ("展望", re.compile(r"展望")),
    ("融资", re.compile(r"融资")),
]


def build_compare_jobs(chunks: list[dict], rng: random.Random, need: int) -> list[dict]:
    """同报告类型、相邻年度的两条 chunk 配对 → 同比题。"""
    have = {(c["year"], c["report_type"]) for c in chunks}
    buckets: dict[tuple, list[dict]] = defaultdict(list)
    for c in chunks:
        if c["char_count"] < 250:
            continue
        sec = c.get("section") or ""
        if SEC_EXCLUDE.search(sec) or is_toc_page(c["text"]):
            continue
        if distinct_years(c["text"]) >= 3:      # 多年份并列的表 → 容易串年份，排除
            continue
        for label, pat in PAIR_FAMILIES:
            if pat.search(sec):
                buckets[(c["report_type"], c["year"], label)].append(c)
                break

    slots = []
    for (rtype, year, label) in buckets:
        prev = str(int(year) - 1)
        key_prev = (rtype, prev, label)
        if (prev, rtype) not in have or key_prev not in buckets:
            continue
        now, before = buckets[(rtype, year, label)], buckets[key_prev]
        scored = []
        for a in now:
            sa = segments_in(a["text"])
            for b in before:
                sb = segments_in(b["text"])
                overlap = len(sa & sb)
                if overlap == 0:
                    continue
                same_sec = a["section"] == b["section"]
                scored.append((overlap + (1 if same_sec else 0),
                               -abs(a["char_count"] - b["char_count"]), a, b, sorted(sa & sb)))
        scored.sort(key=lambda t: (-t[0], -t[1], rng.random()))
        for score, _, a, b, shared in scored[:3]:   # 每个槽位最多取 3 对
            slots.append({
                "job_key": f"compare::{rtype}::{year}::{label}::{a['chunk_id']}::{b['chunk_id']}",
                "qa_type": "compare",
                "chunks": [a, b],                    # a=本期, b=上年同期
                "focus": shared,
            })

    # 轮询排布：先铺开年份和章节族，避免连着出同一年的题
    slots.sort(key=lambda j: (j["chunks"][0]["year"], j["chunks"][0]["report_type"],
                              j["chunks"][1]["chunk_id"]))
    by_family: dict[tuple, list[dict]] = defaultdict(list)
    for j in slots:
        a, b = j["chunks"]
        by_family[(a["report_type"], a["year"])].append(j)
    keys = sorted(by_family.keys(), reverse=True)
    jobs, idx = [], 0
    while len(jobs) < need and any(idx < len(by_family[k]) for k in keys):
        for k in keys:
            if idx < len(by_family[k]) and len(jobs) < need:
                jobs.append(by_family[k][idx])
        idx += 1
    return jobs


# 管理层表述/展望类用词（有这些词的段落才谈得上「管理层观点」）
SUBJ_HINT = re.compile(r"主席|行政总裁|管理层|董事局|展望|策略|计划|相信|预期|预期|"
                       r"我们会|将继续|目标|致力|致股东|前景")


def build_subjective_jobs(chunks: list[dict], rng: random.Random, need: int) -> list[dict]:
    pool = [c for c in chunks
            if c["char_count"] >= 350 and not c["is_table"]
            and NARR_SEC.search(c.get("section") or "")
            and not SEC_EXCLUDE.search(c.get("section") or "")
            and not is_toc_page(c["text"])]
    groups: dict[tuple, list[dict]] = defaultdict(list)
    for c in pool:
        groups[(c["year"], c["report_type"])].append(c)
    for items in groups.values():
        items.sort(key=lambda c: (-len(SUBJ_HINT.findall(c["text"])), -c["char_count"], rng.random()))
    order = sorted(groups.keys(), key=lambda k: (k[0], k[1]))
    jobs, idx = [], 0
    while len(jobs) < need and any(idx < len(groups[k]) for k in order):
        for key in order:
            if idx < len(groups[key]) and len(jobs) < need:
                c = groups[key][idx]
                jobs.append({
                    "job_key": f"subjective::{c['chunk_id']}",
                    "qa_type": "subjective",
                    "chunks": [c],
                    "focus": sorted(segments_in(c["text"]))[:3],
                })
        idx += 1
    return jobs


# ---------------------------------------------------------------- Prompt

def meta_line(c: dict) -> str:
    return (f"年度={c['year']}｜报告类型={c['report_type']}｜章节={c['section']}"
            f"｜PDF页码={c['pdf_page']}｜来源文件={c['source_file']}")


def prompt_statement(job: dict) -> str:
    """三大报表共用：income / balance / cashflow。"""
    c = job["chunks"][0]
    cat = job["qa_type"]
    focus = "、".join(job["focus"]) or "该报表涉及的主要科目"
    return f"""【报告元数据】{meta_line(c)}

【source_text】
{c['text']}

【任务】生成 1 条关于{CATEGORY_DESC[cat]}的问答：
- question：模拟投资者询问该报表中某个**具体科目**的披露数值。**只问一个科目**，不要复合提问。
  **科目名称必须照抄 source_text 表格里的写法**，不得改写、不得归类（如原文写「债务净额」就不能写成「净负债」）。
  问句要写明报告期（用报告原文的期间说法，如「二零二三年十二月三十一日止年度」）与科目名。
  优先围绕这些内容：{focus}
  该报表常见科目范围：{STMT_FOCUS[cat]}
- answer：直接给出该科目本期披露的数值与单位，数字照抄 source_text 的写法。
  若原文同一行/同一张表里就有上期对照列，可以一并引用，但**不得计算任何比率**。
- **不要问需要计算才能回答的问题**（增长百分之多少、毛利率、占比等），除非 source_text 原文已披露该比率。
- source_text 里找不到可明确对应的科目、或表格多年度并列而无法确定数值归属 → skip。

只输出：{{"skip": false, "qa_type": "{cat}", "question": "...", "answer": "..."}}"""


def prompt_metric(job: dict) -> str:
    c = job["chunks"][0]
    ind = job["indicator"]
    return f"""【报告元数据】{meta_line(c)}

【source_text】
{c['text']}

【任务】生成 1 条财务指标问答。题目限定的指标是：**{ind}**，报告期是 **{c['year']} 年（{c['report_type']}）**。
- question：投资者询问该报告期内「{ind}」的披露数值。**指标名必须原样使用「{ind}」**，不得替换成同义词；问句要写明报告期。
- answer：给出 source_text 中该指标**对应 {c['year']} 那一列**的数值与单位，照抄原文写法（百分比、百分点等单位原样保留），并在答案里写明年度。
  若表中同一行已披露变幅/增减幅度，可一并说明「报告披露同比增长 X%」——这属于原文披露，可以直接引用。
- **严禁自行计算**：不得由两年数字推算增长率，不得由收入与溢利推算利润率，不得出现「计算得出」「据此」。
- 若 source_text 中找不到「{ind}」对应的数值，或多年并列表中无法确定该指标对应哪一列 → skip。

只输出：{{"skip": false, "qa_type": "metric", "question": "...", "answer": "..."}}"""


def prompt_compare(job: dict) -> str:
    a, b = job["chunks"][0], job["chunks"][1]     # a=本期, b=上年同期
    focus = "、".join(job["focus"]) or "该章节涉及的指标"
    return f"""【本期报告元数据】{meta_line(a)}

【source_text_A（本期）】
{a['text']}

【上年同期报告元数据】{meta_line(b)}

【source_text_B（上年同期）】
{b['text']}

【任务】生成 1 条 compare（同比/期间比较型）问答：
- question：投资者问本期较上年同期的变化。必须写明两个报告期（用报告原文的期间说法，如「二零二六年上半年」与「二零二五年上半年」），并且**只问 1 个具体财务指标或 1 个具体业务板块**，不要问「各方面有什么变化」这种笼统问题。指标名照抄原文。聚焦：{focus}。
- answer：控制在 **500 字以内**。先分别陈述 A、B 两个期间各自披露的事实，再指出变化方向（增加/减少/基本持平）。
- **只允许出现这两个期间的数字。** source_text 里若还出现了更早年度（如前年）的数字，一律不得写进 answer。严禁把两个期间的数字混成一句。
- **变幅口径**：只有在 source_text 原文（例如财务撮要表的「变幅」列）已披露增减百分比时，才可以引用该百分比并说明这是报告披露的数据；**未披露时只报两期数值，严禁自行计算增长率**。
- 两段 source_text 没有可对照的共同指标就 skip。

只输出：{{"skip": false, "qa_type": "compare", "question": "...", "answer": "..."}}"""


def prompt_subjective(job: dict) -> str:
    c = job["chunks"][0]
    focus = "、".join(job["focus"]) or "该章节涉及的业务"
    return f"""【报告元数据】{meta_line(c)}

【source_text】
{c['text']}

【任务】生成 1 条主观表述类问答（管理层观点，数量要少）：
- question：投资者问「管理层对XX如何表述」「公司对XX的策略/看法是什么」，围绕：{focus}。只问 source_text 中确有表述的内容。
- answer：只复述 source_text 中管理层/公司的原话口径，不得补充、不得推断因果、不得改写立场，也不要把它写成业绩数字题。
- 不得出现 source_text 里没有的预测、判断或结论。
- source_text 只是罗列数字、没有管理层表述 → skip。

只输出：{{"skip": false, "qa_type": "subjective", "question": "...", "answer": "..."}}"""


STMT_FOCUS = {
    "income": "收入、销售成本、毛利、营业溢利、财务支出净额、除税前溢利、所得税、"
              "股东应占溢利、每股盈利、投资物业公平值变动、其他全面收益",
    "balance": "资产总值、流动资产、非流动资产、投资物业、物业厂房及设备、无形资产、"
               "现金及现金等价物、负债总额、流动负债、借款、权益总额、非控股权益、应收款项",
    "cashflow": "营运业务、来自营运的现金、投资业务、融资业务、购买物业厂房及设备、"
                "增购投资物业、出售投资物业所得款项、已付利息、已收利息、已付税项、资本开支",
}

PROMPT_BUILDERS = {
    "income": prompt_statement,
    "balance": prompt_statement,
    "cashflow": prompt_statement,
    "metric": prompt_metric,
    "compare": prompt_compare,
    "subjective": prompt_subjective,
}


# ---------------------------------------------------------------- 单任务执行

def make_record(job: dict, obj: dict, seq: int, usage: dict, run_id: str = "") -> dict:
    chunks = job["chunks"]
    primary = chunks[0]
    qa_id = f"QA-{seq:04d}"
    sources = []
    for c in chunks:
        sources.append({
            "year": c["year"],
            "report_type": c["report_type"],
            "source_file": c["source_file"],
            "pdf_page": c["pdf_page"],
            "chunk_id": c["chunk_id"],
            "section": c["section"],
            "text": c["text"],
        })
    return {
        "id": qa_id,
        "job_key": job["job_key"],
        "qa_type": job["qa_type"],
        "question": str(obj.get("question", "")).strip(),
        "answer": str(obj.get("answer", "")).strip(),
        "year": primary["year"],
        "report_type": primary["report_type"],
        "source_file": primary["source_file"],
        "pdf_page": primary["pdf_page"],
        "section": primary["section"],
        "chunk_id": primary["chunk_id"],
        "source_text": primary["text"],
        "sources": sources,
        "focus": job.get("focus", []),
        "indicator": job.get("indicator", ""),
        # 空白 = 尚未人工抽检；人工确认正确填 TRUE，确认有问题填 FALSE。
        # 不要把「未检查」和「检查后错误」混为一谈，所以这里不预填 false。
        "verified": "",
        "error_type": "",
        "review_note": "",
        "gen_model": usage.get("_model", ""),
        "gen_usage": {k: v for k, v in usage.items() if not k.startswith("_")},
        "gen_run": run_id,
        "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    }


def run_job(cfg: dict, job: dict, seq: int, run_id: str = "") -> dict:
    """返回 {"status": "ok"|"skip"|"fail", ...}"""
    builder = PROMPT_BUILDERS[job["qa_type"]]
    user = builder(job)
    last_err = ""
    for attempt in range(1, RETRIES + 1):
        try:
            text, usage = api_call(cfg, SYSTEM_PROMPT, user, MAX_TOKENS)
            if usage.get("_stop_reason") == "max_tokens":
                text, usage = api_call(cfg, SYSTEM_PROMPT, user, MAX_TOKENS_RETRY)
            obj = extract_json(text)
            if obj is None:
                last_err = f"JSON 解析失败: {text[:120]}"
                time.sleep(1.5 * attempt)
                continue
            if obj.get("skip"):
                return {"status": "skip", "reason": str(obj.get("reason", ""))[:200],
                        "job_key": job["job_key"]}
            q = str(obj.get("question", "")).strip()
            a = str(obj.get("answer", "")).strip()
            if len(q) < 8 or len(a) < 15:
                last_err = f"问答过短: q={len(q)} a={len(a)}"
                time.sleep(1.0 * attempt)
                continue
            usage["_model"] = cfg["model"]
            rec = make_record(job, obj, seq, usage, run_id)
            allowed = " ".join(meta_line(c) for c in job["chunks"])
            rec["num_unsupported"] = check_numbers(
                a, " ".join(s["text"] for s in rec["sources"]), allowed)
            rec["selfcalc"] = selfcalc_suspect(rec)
            rec["answer_chars"] = len(a)
            return {"status": "ok", "record": rec, "job_key": job["job_key"]}
        except Exception as exc:                       # 网络/HTTP/超时
            last_err = f"{type(exc).__name__}: {exc}"
            time.sleep(2.0 * attempt)
    return {"status": "fail", "reason": last_err[:300], "job_key": job["job_key"]}


# ---------------------------------------------------------------- 主流程

def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="生成董秘问答 QA 数据集（只用 data/chunks/）")
    ap.add_argument("--only", choices=list(TARGETS) + ["statement"], default=None,
                    help="只生成某一类（income/balance/cashflow/metric/compare/subjective，"
                         "或 statement = 三大报表一起）")
    ap.add_argument("--limit", type=int, default=0, help="每类最多生成多少条（0=用默认目标）")
    ap.add_argument("--workers", type=int, default=DEFAULT_WORKERS, help="并发数")
    ap.add_argument("--seed", type=int, default=DEFAULT_SEED, help="抽样随机种子")
    ap.add_argument("--fresh", action="store_true", help="清空已有 qa_raw.jsonl 重新生成")
    ap.add_argument("--dry-run", action="store_true", help="只建任务池并打印样例 prompt，不调 API")
    ap.add_argument("--prepare-only", action="store_true", help="不调 API，只用已有 qa_raw.jsonl 出报告")
    ap.add_argument("--no-prepare", action="store_true", help="只生成，不出报告/抽检文件")
    ap.add_argument("--keep-all", action="store_true", help="不做每类条数裁剪，保留全部成功结果")
    return ap.parse_args()


def load_existing() -> tuple[list[dict], set[str]]:
    records, done = [], set()
    if QA_RAW.exists():
        with QA_RAW.open(encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    r = json.loads(line)
                except Exception:
                    continue
                records.append(r)
                if r.get("job_key"):
                    done.add(r["job_key"])
    return records, done


def generate(args, cfg, chunks) -> list[dict]:
    rng = random.Random(args.seed)
    types = tuple(TARGETS)
    targets = {t: (args.limit or TARGETS[t]) for t in types}
    # 任务池比目标略大，用来吸收 skip / 失败
    pools = {t: (int(targets[t] * 1.15) + 1 if args.limit else max(POOLS[t], targets[t]))
             for t in types}
    if args.only:
        keep = set(STATEMENT_CATS) if args.only == "statement" else {args.only}
        for t in types:
            if t not in keep:
                targets[t] = pools[t] = 0

    log("[1/4] 建任务池")
    jobs = []
    stmt_pools = {c: pools[c] for c in STATEMENT_CATS if pools.get(c)}
    if stmt_pools:
        jobs += build_statement_jobs(chunks, rng, stmt_pools)
    if pools.get("metric"):
        jobs += build_metric_jobs(chunks, rng, pools["metric"])
    if pools.get("compare"):
        jobs += build_compare_jobs(chunks, rng, pools["compare"])
    if pools.get("subjective"):
        jobs += build_subjective_jobs(chunks, rng, pools["subjective"])
    log(f"  任务池：{dict(Counter(j['qa_type'] for j in jobs))}  合计 {len(jobs)} 个任务")

    if args.dry_run:
        for j in jobs[:3]:
            log("\n" + "=" * 70)
            log(f"job_key = {j['job_key']}   type={j['qa_type']}   focus={j.get('focus')}")
            log("=" * 70)
            log(PROMPT_BUILDERS[j["qa_type"]](j)[:2000])
        log(f"\n[dry-run] 未调用 API。任务池共 {len(jobs)} 个任务。")
        return []

    existing, done = load_existing()
    if args.fresh and QA_RAW.exists():
        QA_RAW.unlink()
        existing, done = [], set()
        log("  --fresh：已清空旧结果")
    # 新 id 从已有最大编号往后接，避免重跑时重号
    max_num = 0
    for r in existing:
        digits = re.sub(r"\D", "", str(r.get("id", "")))
        if digits:
            max_num = max(max_num, int(digits))
    log(f"  已有 {len(existing)} 条记录、{len(done)} 个已完成任务（不会重复调 API）")

    todo = [j for j in jobs if j["job_key"] not in done]
    if not todo:
        log("  没有需要生成的新任务。")
        return existing

    log(f"[2/4] 调用 LLM 生成（并发 {args.workers}，待生成 {len(todo)} 个任务）")
    probe_endpoint(cfg)
    run_id = time.strftime("%Y%m%dT%H%M%S")
    log(f"  本次运行标识 run_id={run_id}")

    results, log_rows = [], []
    t0, finished = time.time(), 0
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futures = {ex.submit(run_job, cfg, job, max_num + 1 + i, run_id): job
                   for i, job in enumerate(todo)}
        for fut in as_completed(futures):
            job = futures[fut]
            finished += 1
            try:
                res = fut.result()
            except Exception as exc:
                res = {"status": "fail", "reason": f"{type(exc).__name__}: {exc}",
                       "job_key": job["job_key"]}
            log_rows.append({"job_key": res["job_key"], "qa_type": job["qa_type"],
                             "status": res["status"], "reason": res.get("reason", ""),
                             "run_id": run_id,
                             "at": time.strftime("%Y-%m-%dT%H:%M:%S")})
            if res["status"] == "ok":
                results.append(res["record"])
                with QA_RAW.open("a", encoding="utf-8") as fh:
                    fh.write(json.dumps(res["record"], ensure_ascii=False) + "\n")
            if finished % 10 == 0 or res["status"] != "ok":
                note = f"  ← {res.get('reason', '')[:70]}" if res["status"] != "ok" else ""
                log(f"    [{finished}/{len(todo)}] 成功 {len(results)}｜{res['status']}{note}")

    with QA_LOG.open("a", encoding="utf-8") as fh:
        for row in log_rows:
            fh.write(json.dumps(row, ensure_ascii=False) + "\n")

    # 收尾裁剪：先去掉与已保留问题重复的（同一个数字常在不同章节重复披露，
    # 容易生成出近乎一样的题），再按类别只保留目标条数（按 id 顺序，稳定可复现）。
    if not args.keep_all:
        kept = []
        cnt = Counter(r.get("qa_type") for r in existing)   # 已有条数计入总数上限
        seen_q = {normalize_q(r.get("question", "")) for r in existing}
        n_dup = 0
        for r in sorted(results, key=lambda r: r["id"]):
            nq = normalize_q(r.get("question", ""))
            if nq in seen_q:
                n_dup += 1
                continue
            if cnt[r["qa_type"]] < targets[r["qa_type"]]:
                kept.append(r)
                seen_q.add(nq)
                cnt[r["qa_type"]] += 1
        if n_dup or len(kept) != len(results):
            log(f"  裁剪去重：{dict(Counter(r['qa_type'] for r in results))} → "
                f"{dict(cnt)}（重复问题 {n_dup} 条，超额 {len(results) - len(kept) - n_dup} 条）")
            with QA_RAW.open("w", encoding="utf-8") as fh:
                for r in existing + kept:
                    fh.write(json.dumps(r, ensure_ascii=False) + "\n")
            results = kept

    ok = Counter(r["qa_type"] for r in results)
    log(f"[3/4] 生成结束：新增 {len(results)} 条 {dict(ok)}，用时 {time.time() - t0:.1f} 秒")
    return existing + results


# ---------------------------------------------------------------- 出报告

def check_one(rec: dict, report_index: dict, dup_key: str | None) -> list[str]:
    """对单条 QA 做自动检查，返回问题标签列表。"""
    problems = []
    srcs = rec.get("sources") or []
    if not srcs:
        problems.append("无来源")
        return problems
    for s in srcs:
        sf, page = s.get("source_file"), s.get("pdf_page")
        info = report_index.get(sf)
        if not sf:
            problems.append("source_file 为空")
            continue
        if info is None:
            problems.append(f"source_file 不在语料中：{sf}")
            continue
        sp = info.get("source_path")
        if not sp or not Path(sp).exists():
            problems.append(f"源 PDF 不存在：{sp}")
        try:
            page_i = int(page)
        except (TypeError, ValueError):
            problems.append(f"pdf_page 非法：{page!r}")
            continue
        tp = info.get("total_pages")
        if tp and not (1 <= page_i <= tp):
            problems.append(f"pdf_page 越界：{page_i}/{tp}")
        if info.get("year") and str(info["year"]) != str(s.get("year")):
            problems.append(f"year 与报告不符：{s.get('year')} vs {info['year']}")
        if info.get("report_type") and info["report_type"] != s.get("report_type"):
            problems.append(f"report_type 与报告不符：{s.get('report_type')} vs {info['report_type']}")
        if not (s.get("text") or "").strip():
            problems.append(f"source_text 为空：{s.get('chunk_id')}")
    if not (rec.get("source_text") or "").strip():
        problems.append("顶层 source_text 为空")
    if rec.get("qa_type") == "compare":
        years = {str(s.get("year")) for s in srcs}
        if len(srcs) < 2 or len(years) < 2:
            problems.append("同比题不足两个年度来源")
        else:
            types = {s.get("report_type") for s in srcs}
            if len(types) > 1:
                problems.append("同比题两个来源报告类型不一致")
            # 答案里出现第三个年度 = 把别的年度信息混进来了
            # （「2019冠状病毒病」这类专名里的数字不算年度）
            try:
                allowed = {str(int(rec["year"])), str(int(rec["year"]) - 1)}
                ans = re.sub(r"20\d{2}\s*冠状|COVID-?19|冠狀", "", rec.get("answer", ""))
                extra = {y for y in re.findall(r"(?<!\d)(20\d{2})(?!\d)", ans)} - allowed
                if extra:
                    problems.append("对比题答案出现第三个年度：" + "、".join(sorted(extra)))
            except (KeyError, TypeError, ValueError):
                pass
    if rec.get("num_unsupported"):
        problems.append("数字无原文依据：" + "、".join(rec["num_unsupported"][:6]))
    for s in (rec.get("selfcalc") or selfcalc_suspect(rec)):
        problems.append("疑似自行计算：" + s)
    if rec.get("answer_chars", 0) > 700:
        problems.append(f"答案过长（{rec['answer_chars']} 字，可能整表照抄）")
    if dup_key:
        problems.append(f"问题重复（与 {dup_key}）")
    return problems


def write_csv(records: list[dict]) -> None:
    """人工抽检表。verified / error_type / review_note 三列**默认全空**：
    空白 = 尚未抽检，人工确认正确才填 TRUE，确认有问题填 FALSE。"""
    existing_review = {}
    if QA_CSV.exists():
        with QA_CSV.open("r", encoding="utf-8-sig", newline="") as old:
            existing_review = {row["id"]: row for row in csv.DictReader(old)}
    for record in records:
        previous = existing_review.get(record["id"], {})
        for field in ("verified", "error_type", "review_note"):
            if previous.get(field):
                record[field] = previous[field]
    with QA_CSV.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["id", "question", "answer", "source_file", "pdf_page", "year",
                    "report_type", "verified", "error_type", "review_note"])
        for r in records:
            srcs = r.get("sources") or [{}]
            files = " ; ".join(str(s.get("source_file", "")) for s in srcs)
            pages = " ; ".join(str(s.get("pdf_page", "")) for s in srcs)
            w.writerow([r["id"], r["question"], r["answer"], files, pages,
                        r["year"], r["report_type"],
                        r.get("verified", ""), r.get("error_type", ""),
                        r.get("review_note", "")])


def write_sample(records: list[dict], rng: random.Random) -> list[dict]:
    """分层随机抽 30 条：六类各抽若干条，保证每类都被覆盖到。"""
    order = list(TARGETS)
    per = max(1, SAMPLE_SIZE // len(order))
    picked = []
    for t in order:
        pool = [r for r in records if r.get("qa_type") == t]
        rng.shuffle(pool)
        picked += pool[:per]
    rest = [r for r in records if r["id"] not in {p["id"] for p in picked}]
    rng.shuffle(rest)
    picked += rest[:max(0, SAMPLE_SIZE - len(picked))]

    lines = ["# 人工抽检样本（随机抽取）", "",
             f"共抽取 **{len(picked)}** 条，六类问题均有覆盖（"
             + "、".join(f"{CATEGORY_DESC[t]}{per} 条" for t in order)
             + " + 随机补足）。", "",
             "> `verified` / `error_type` / `review_note` 默认**全空**：空白＝尚未抽检；"
             "人工确认正确才填 `TRUE`，确认有问题填 `FALSE`，"
             "并在 `data/qa/qa_review.csv` 对应行填写 `error_type` / `review_note`。", ""]
    for i, r in enumerate(picked, 1):
        srcs = r.get("sources") or [{}]
        cat = r.get("qa_type")
        lines += [
            f"### QA-{i:03d}　（id: {r['id']}｜类型: {cat} {CATEGORY_DESC.get(cat, '')}）", "",
            "Question:", "", r["question"], "",
            "Answer:", "", r["answer"], "",
            "Source:",
            "",
        ]
        for j, s in enumerate(srcs, 1):
            label = f"（源 {j}）" if len(srcs) > 1 else ""
            lines.append(f"- {label}{s.get('source_file')}　PDF 第 {s.get('pdf_page')} 页"
                         f"　{s.get('section')}　chunk `{s.get('chunk_id')}`")
        lines += [
            "",
            f"Year: {r['year']}" + (f"（对比题另含 {len(srcs) - 1} 个来源，见上）"
                                    if len(srcs) > 1 else ""),
            "",
            f"Report: {r['report_type']}",
            "",
            f"PDF page: {r['pdf_page']}",
            "",
            "Source text:",
            "",
            "```",
        ]
        for j, s in enumerate(srcs, 1):
            if len(srcs) > 1:
                lines.append(f"—— 源 {j}：{s.get('year')} {s.get('report_type')} "
                             f"PDF 第 {s.get('pdf_page')} 页 ——")
            lines.append((s.get("text") or "").strip())
            lines.append("")
        lines += [
            "```", "",
            "Review:", "",
            "[ ] 正确　[ ] 部分正确　[ ] 编造　[ ] 来源错误　[ ] 问题错误　[ ] 其他", "",
            "---", "",
        ]
    SAMPLE_MD.write_text("\n".join(lines), encoding="utf-8")
    return picked


def main() -> int:
    args = parse_args()
    t_start = time.time()
    QA_DIR.mkdir(parents=True, exist_ok=True)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    log("=" * 72)
    log("阶段五：QA 数据生成（董秘问答）")
    log("=" * 72)
    chunks = load_chunks()
    report_index = load_report_index()
    log(f"  语料：{len(chunks)} 块 / 报告索引 {len(report_index)} 份")

    records: list[dict] = []
    if args.prepare_only:
        records, _ = load_existing()
        log(f"  --prepare-only：直接使用已有 {len(records)} 条")
    else:
        cfg = load_llm_config()
        records = generate(args, cfg, chunks)

    if args.dry_run:
        return 0

    if not args.no_prepare:
        log("[4/4] 输出抽检文件与报告")
        rng = random.Random(args.seed)
        records.sort(key=lambda r: r["id"])
        n_fixed = recompute_checks(records)
        log(f"  重算数字告警：{n_fixed} 条发生变化（qa_raw.jsonl 已同步）")
        # 自行计算嫌疑也按当前规则重算一遍，避免沿用旧批次存下的结论
        n_sc = 0
        for r in records:
            fresh = selfcalc_suspect(r)
            if fresh != (r.get("selfcalc") or []):
                n_sc += 1
            r["selfcalc"] = fresh
        if n_sc:
            log(f"  自行计算嫌疑重算：{n_sc} 条发生变化")

        # 重复问题标记：完全相同 + 近似（相似度 >= 0.90）
        problems: dict[str, list[str]] = {}
        seen_norm: dict[str, str] = {}
        norms = [(r["id"], normalize_q(r["question"])) for r in records]
        exact_dup, near_dup = 0, 0
        for rid, nq in norms:
            if nq in seen_norm:
                problems.setdefault(rid, []).append(f"问题重复（与 {seen_norm[nq]}）")
                exact_dup += 1
            else:
                seen_norm[nq] = rid
        buckets: dict[str, list[tuple[str, str]]] = defaultdict(list)
        for rid, nq in norms:
            buckets[nq[:12]].append((rid, nq))
        # 问的是不同科目/指标（如「每股盈利」vs「每股股息」）时，句式再像也不算重复
        subject = {r["id"]: frozenset(x for x in ([r.get("indicator")] + list(r.get("focus") or [])) if x)
                   for r in records}
        for key, items in buckets.items():
            for i in range(len(items)):
                for j in range(i + 1, len(items)):
                    if subject[items[i][0]] != subject[items[j][0]]:
                        continue
                    if SequenceMatcher(None, items[i][1], items[j][1]).ratio() >= 0.90:
                        problems.setdefault(items[j][0], []).append(f"问题疑似重复（与 {items[i][0]}）")
                        near_dup += 1

        for r in records:
            for p in check_one(r, report_index, None):
                problems.setdefault(r["id"], []).append(p)
            r["problems"] = problems.get(r["id"], [])

        write_csv(records)
        picked = write_sample(records, rng)
        write_reports(records, problems, exact_dup, near_dup, report_index, picked)

    log(f"\n完成，用时 {time.time() - t_start:.1f} 秒。")
    return 0


def pct(a: int, b: int) -> str:
    return f"{a / b * 100:.1f}%" if b else "—"


def recompute_checks(records: list[dict]) -> int:
    """重算每条的数字告警并回写 qa_raw.jsonl（不调 API，可反复执行）。"""
    changed = 0
    for r in records:
        srcs = r.get("sources") or []
        allowed = " ".join(
            f"年度={s.get('year')}｜报告类型={s.get('report_type')}｜章节={s.get('section')}"
            f"｜PDF页码={s.get('pdf_page')}｜来源文件={s.get('source_file')}" for s in srcs)
        new = check_numbers(r.get("answer", ""),
                            " ".join(s.get("text", "") for s in srcs), allowed)
        if new != (r.get("num_unsupported") or []):
            changed += 1
        r["num_unsupported"] = new
        r["answer_chars"] = len(r.get("answer", ""))
    if changed and QA_RAW.exists():
        records.sort(key=lambda r: r["id"])
        with QA_RAW.open("w", encoding="utf-8") as fh:
            for r in records:
                fh.write(json.dumps(r, ensure_ascii=False) + "\n")
    return changed


def write_reports(records, problems, exact_dup, near_dup, report_index, picked) -> None:
    total = len(records)
    by_year = Counter(r["year"] for r in records)
    by_type = Counter(r["report_type"] for r in records)
    by_qtype = Counter(r.get("qa_type") for r in records)

    # 四组比例（三大报表 / 财务指标 / 同比 / 主观）
    n_stmt = sum(by_qtype.get(c, 0) for c in STATEMENT_CATS)
    ratio_stmt = n_stmt / total if total else 0.0
    ratio_metric = by_qtype.get("metric", 0) / total if total else 0.0
    ratio_cmp = by_qtype.get("compare", 0) / total if total else 0.0
    ratio_subj = by_qtype.get("subjective", 0) / total if total else 0.0
    ok_stmt = 0.55 <= ratio_stmt <= 0.65
    ok_metric = 0.20 <= ratio_metric <= 0.25
    ok_cmp = 0.10 <= ratio_cmp <= 0.15
    ok_subj = ratio_subj <= 0.10

    # 字段完整率
    fields = ["id", "question", "answer", "year", "report_type", "source_file",
              "pdf_page", "section", "source_text"]
    filled = {f: sum(1 for r in records if str(r.get(f, "") or "").strip()
                     and str(r.get(f)) != "None") for f in fields}
    multi_src = sum(1 for r in records if len(r.get("sources") or []) > 1)

    bad_ids = [rid for rid, ps in problems.items() if ps]
    num_bad = [r["id"] for r in records if any("数字无原文依据" in p for p in r.get("problems", []))]
    cmp_bad = [r["id"] for r in records if r.get("qa_type") == "compare"
               and any("同比题" in p for p in r.get("problems", []))]

    def ids_with(tag: str) -> list[str]:
        return [r["id"] for r in records if any(tag in p for p in r.get("problems", []))]

    file_bad = sorted(set(ids_with("source_file") + ids_with("源 PDF")))
    page_bad = ids_with("pdf_page")
    empty_bad = ids_with("为空")
    yr_bad = ids_with("year 与报告不符")
    rt_bad = ids_with("report_type 与报告不符")
    selfcalc_bad = [r for r in records if (r.get("selfcalc") or selfcalc_suspect(r))]
    verified_filled = [r["id"] for r in records if str(r.get("verified", "") or "").strip()]

    runs = {r.get("gen_run") for r in records if r.get("gen_run")}
    log_rows = []
    if QA_LOG.exists():
        for line in QA_LOG.read_text(encoding="utf-8").splitlines():
            if line.strip():
                log_rows.append(json.loads(line))
    if runs:
        log_rows = [x for x in log_rows if x.get("run_id") in runs]

    # ---------------------------------------------------------- 生成报告
    stamps = sorted(r.get("generated_at", "") for r in records if r.get("generated_at"))
    span = f"{stamps[0]} ~ {stamps[-1]}" if stamps else "—"
    empty_srcs = sum(1 for r in records for s in (r.get("sources") or [])
                     if not (s.get("text") or "").strip())
    empty_src = sum(1 for r in records if not (r.get("source_text") or "").strip())
    flag = lambda ok: "✅ 达标" if ok else "⚠️ 未达标"

    L = ["# QA 数据生成报告", "",
         f"报告生成时间：{time.strftime('%Y-%m-%d %H:%M:%S')}", "",
         f"数据源：`data/chunks/all_chunks.jsonl`（3,246 块）　|　"
         f"生成时段：{span}", "",
         "> 本阶段只做「QA 数据生成 + 抽检准备」，未做 LoRA 训练、未改动 `vector_db/`、"
         "未重新解析 PDF、未重新 Embedding。", "",
         "> 本批训练目标是**财报披露口径**：三大财务报表 + 财务指标占绝大多数，"
         "主观表述类压到 10% 以内；所有数字与指标只允许照抄原文披露，严禁自行计算。", "",
         "## 1. 总数量", "",
         f"**{total} 条 QA**（目标约 300 条）", "",
         f"- 其中同比/期间比较（含两个报告期来源）{multi_src} 条", "",
         "## 2. 各类问题数量与比例", "",
         "| 类别 | 含义 | 条数 | 占比 | 目标区间 | 达标 |", "|---|---|---|---|---|---|"]
    band = {"income": "三大报表 55–65%（合计）", "balance": "三大报表 55–65%（合计）",
            "cashflow": "三大报表 55–65%（合计）", "metric": "20–25%",
            "compare": "10–15%", "subjective": "≤10%"}
    for t in TARGETS:
        n = by_qtype.get(t, 0)
        if t in STATEMENT_CATS:
            ok = ok_stmt
        elif t == "metric":
            ok = ok_metric
        elif t == "compare":
            ok = ok_cmp
        else:
            ok = ok_subj
        L.append(f"| `{t}` {CATEGORY_DESC[t]} | | {n} | {pct(n, total)} | {band[t]} | {flag(ok)} |")
    L += ["", "**四大分组占比**", "",
          "| 分组 | 条数 | 占比 | 目标 | 达标 |", "|---|---|---|---|---|",
          f"| 三大财务报表及财务数据（income+balance+cashflow） | {n_stmt} | {pct(n_stmt, total)} "
          f"| 55–65% | {flag(ok_stmt)} |",
          f"| 财务指标及经营财务数据（metric） | {by_qtype.get('metric', 0)} "
          f"| {pct(by_qtype.get('metric', 0), total)} | 20–25% | {flag(ok_metric)} |",
          f"| 同比/期间比较（compare） | {by_qtype.get('compare', 0)} "
          f"| {pct(by_qtype.get('compare', 0), total)} | 10–15% | {flag(ok_cmp)} |",
          f"| 主观表述/管理层观点（subjective） | {by_qtype.get('subjective', 0)} "
          f"| {pct(by_qtype.get('subjective', 0), total)} | ≤10% | {flag(ok_subj)} |", "",
          "## 3. 三大报表分别条数", "",
          "| 报表 | 类别 | 条数 | 占比 |", "|---|---|---|---|",
          f"| 损益表及收入数据 | `income` | {by_qtype.get('income', 0)} "
          f"| {pct(by_qtype.get('income', 0), total)} |",
          f"| 资产负债表数据 | `balance` | {by_qtype.get('balance', 0)} "
          f"| {pct(by_qtype.get('balance', 0), total)} |",
          f"| 现金流量表数据 | `cashflow` | {by_qtype.get('cashflow', 0)} "
          f"| {pct(by_qtype.get('cashflow', 0), total)} |",
          "", "## 4. 各年份数量", "",
          "| year | 条数 |", "|---|---|"]
    for y in sorted(by_year):
        L.append(f"| {y} | {by_year[y]} |")
    L += ["", "## 5. 年报 / 中报数量", "", "| report_type | 条数 | 占比 |", "|---|---|---|"]
    for t, n in by_type.most_common():
        L.append(f"| {t} | {n} | {pct(n, total)} |")
    L += ["", "## 6. 来源字段完整率", "", "| 字段 | 非空条数 | 完整率 |", "|---|---|---|"]
    for f in fields:
        L.append(f"| `{f}` | {filled[f]} | {pct(filled[f], total)} |")
    L += ["", "## 7. 是否存在重复问题", "",
          f"- 完全相同（归一化后）：**{exact_dup}** 条",
          f"- 高度相似（相似度 ≥ 0.90）：**{near_dup}** 条", "",
          "## 8. 是否存在 source_text 为空", "",
          f"- 顶层 `source_text` 为空：**{empty_src}** 条",
          f"- 任一 `sources[].text` 为空：**{empty_srcs}** 条", "",
          "## 9. verified 是否全部为空", "",
          f"- `verified` 非空的条数：**{len(verified_filled)}** 条"
          + ("（✅ 全部为空，等待人工抽检）" if not verified_filled
             else f"（⚠️ 见 {', '.join(verified_filled[:5])}）"), "",
          "> 口径：空白＝尚未人工抽检；`TRUE`＝人工确认正确；`FALSE`＝人工确认存在问题。"
          "本批**不预填 false**，`error_type` / `review_note` 同样留空。", "",
          "## 10. 疑似自行计算的条数", "",
          f"- 命中「推导性表述 / 索要原文未披露的比率 / 答案含原文没有的百分比」的条数："
          f"**{len(selfcalc_bad)}** 条"
          + ("（✅ 无）" if not selfcalc_bad else "（⚠️ 需人工复核，见 `qa_quality_check.md`）"),
          "",
          "> 本批要求学「披露口径」而非计算器：指标名照抄原文，比率只在原文已披露"
          "（如财务撮要表的「变幅」列）时引用。", "",
          "## 11. 人工抽检方法", "",
          "1. 打开 `data/qa/qa_review.csv`（Excel 可直接读，UTF-8 BOM）。",
          "2. 逐条对照 `source_text` 与 `answer`，按六种结论勾选：",
          "   - **正确**：answer 的事实、数字、口径全部能在 source_text 中找到；",
          "   - **部分正确**：主体正确但有个别表述超出原文；",
          "   - **编造**：出现 source_text 里没有的数字/事实/管理层观点；",
          "   - **来源错误**：source_file 或 pdf_page 指向别处；",
          "   - **问题错误**：问题本身不成立、问的是原文没有的指标、或问法有歧义；",
          "   - **其他**：重复、无意义等。",
          "3. 在 `verified` 列填 `TRUE`（正确）/`FALSE`（有问题）；"
          "`error_type` 填上面六类标签之一，`review_note` 写具体问题（错在哪个数字、哪句话）。"
          "**留空表示还没抽检**。",
          "4. 抽检结论先汇总到 `error_type` 分布，再决定是否整类重生成。", "",
          f"本次另外生成了 `output/qa_review_sample.md`（随机 {len(picked)} 条，六类均有覆盖），"
          "可直接打印勾选。", "",
          "## 12. 产出文件", "",
          "| 文件 | 内容 |", "|---|---|",
          "| `scripts/05_generate_qa.py` | 本阶段脚本（可重复运行，已生成的不会重复调 API） |",
          "| `data/qa/qa_raw.jsonl` | QA 全集（含完整溯源字段） |",
          "| `data/qa/qa_review.csv` | 人工抽检表（`verified` 等三列默认空白） |",
          "| `data/qa/qa_generation_log.jsonl` | 每个任务的执行结果（ok/skip/fail，追加写入） |",
          "| `output/qa_generation_report.md` | 本报告 |",
          "| `output/qa_quality_check.md` | 自动质量检查结果（13 项） |",
          "| `output/qa_review_sample.md` | 人工抽检样本 |", ""]
    REPORT_MD.write_text("\n".join(L), encoding="utf-8")

    # ---------------------------------------------------------- 质量检查
    n_stmt_cats = sum(1 for c in STATEMENT_CATS if by_qtype.get(c, 0) > 0)
    fin_cnt = n_stmt + by_qtype.get("metric", 0)
    fin_ratio = fin_cnt / total if total else 0.0
    Q = ["# QA 自动质量检查", "",
         f"检查时间：{time.strftime('%Y-%m-%d %H:%M:%S')}　|　样本 {total} 条", "",
         "## 结论速览", "",
         "| # | 检查项 | 结果 |", "|---|---|---|",
         f"| 1 | source_file 是否存在于语料/源 PDF | "
         f"{'✅ 全部通过' if not file_bad else f'⚠️ {len(file_bad)} 条异常'} |",
         f"| 2 | pdf_page 是否为有效数字且未越界 | "
         f"{'✅ 全部通过' if not page_bad else f'⚠️ {len(page_bad)} 条异常'} |",
         f"| 3 | source_text 是否为空 | "
         f"{'✅ 无空值' if not empty_bad else f'⚠️ {len(empty_bad)} 条异常'} |",
         f"| 4 | answer 是否含原文无法支持的数字 | "
         f"{'✅ 全部有原文依据' if not num_bad else f'⚠️ {len(num_bad)} 条需人工确认'} |",
         f"| 5 | question 是否重复 | "
         f"{'✅ 无重复' if not (exact_dup or near_dup) else f'⚠️ 完全重复 {exact_dup} / 高度相似 {near_dup}'} |",
         f"| 6 | 同比题是否有两个年度来源 | "
         f"{'✅ 全部满足' if not cmp_bad else f'⚠️ {len(cmp_bad)} 条不满足'} |",
         f"| 7 | year 与 source_file 是否匹配 | "
         f"{'✅ 一致' if not yr_bad else f'⚠️ {len(yr_bad)} 条不符'} |",
         f"| 8 | report_type 与 source_file 是否匹配 | "
         f"{'✅ 一致' if not rt_bad else f'⚠️ {len(rt_bad)} 条不符'} |",
         f"| 9 | 财务数据类比例（三大报表+指标） | {pct(fin_cnt, total)}"
         f"（{n_stmt}+{by_qtype.get('metric', 0)} 条，目标 75–90%） "
         f"{'✅ 达标' if 0.75 <= fin_ratio <= 0.90 else '⚠️ 未达标'} |",
         f"| 10 | 三大报表覆盖情况 | "
         f"{'✅ 损益表/资产负债表/现金流量表均有覆盖' if n_stmt_cats == 3 else f'⚠️ 仅覆盖 {n_stmt_cats}/3 张表'} |",
         f"| 11 | 主观类比例是否 ≤10% | {pct(by_qtype.get('subjective', 0), total)} "
         f"{'✅ 达标' if ok_subj else '⚠️ 超标'} |",
         f"| 12 | verified 是否全部为空 | "
         f"{'✅ 全部为空（等待人工抽检）' if not verified_filled else f'⚠️ {len(verified_filled)} 条非空'} |",
         f"| 13 | 是否存在明显模型自行计算的指标 | "
         f"{'✅ 未发现' if not selfcalc_bad else f'⚠️ {len(selfcalc_bad)} 条需人工确认'} |",
         "", "### 各项判定口径", "",
         "1. **source_file**：与 `data/extracted/*.json` 的 `source_file` 集合比对，"
         "并用其 `source_path` 确认源 PDF 在磁盘上存在。",
         "2. **pdf_page**：必须为整数，且落在该报告的 `total_pages` 范围内。",
         "3. **source_text**：检查顶层字段与每个 `sources[].text`。",
         "4. **数字**：把 answer 与 source_text 里的数字（含中文数字，如「四十九亿」→ 4.9e9、"
         "「二零二六」→ 2026）都换算成数值比对；answer 中的数字若既不能原文命中、"
         "也换算不出对应值，则列为「无原文依据」。**这是启发式筛查，不是判罪**，需人工复核。",
         "5. **重复**：问题归一化（去标点）后完全相同，或用 `difflib` 相似度 ≥ 0.90 判为高度相似。",
         "6. **同比**：`qa_type=compare` 必须有两个 source 且年份不同、报告类型相同；"
         "答案里出现第三个年度的数字也会被标记。",
         "7. **year / 8. report_type**：以「同一份 source_file 在全语料中的 year、report_type 唯一」"
         "为基准比对。",
         "9. **财务数据类比例**：`income` + `balance` + `cashflow` + `metric` 四条合计。"
         "本批目标是三大报表 55–65% 加财务指标 20–25%，即合计 75–90%。",
         "10. **三大报表覆盖**：损益表 / 资产负债表 / 现金流量表三类都要有样本，"
         "缺任何一张表都不算通过。",
         "11. **主观类比例**：`subjective` 必须 ≤10%，超过就要调整生成策略重跑。",
         "12. **verified 全空**：`verified` / `error_type` / `review_note` 默认留空。"
         "空白＝尚未抽检，`TRUE`＝人工确认正确，`FALSE`＝人工确认有问题。",
         "13. **自行计算**：三项启发式特征——答案含「计算/据此/推算/折算/得出」等推导性表述；"
         "问题索要的比率（增长率/毛利率/占比等）在原文中没有出现；答案出现百分比或百分点"
         "而原文未见相应披露。命中只代表**需要人工重点复核**。",
         "", "## 逐条问题明细", ""]

    if not bad_ids:
        Q += ["**没有发现问题。** 全部 " + str(total) + " 条通过逐条自动检查（第 1–8、13 项）。", ""]
    else:
        Q += [f"共 **{len(bad_ids)}** 条存在至少一项问题：", "",
              "| id | 类型 | 问题 |", "|---|---|---|"]
        for r in records:
            if r["id"] in bad_ids:
                Q.append(f"| {r['id']} | {r.get('qa_type')} | {'；'.join(r['problems'])} |")
        Q.append("")

    if log_rows:
        st = Counter(x["status"] for x in log_rows)
        Q += ["## 生成过程统计", "",
              "（下表为**本批运行**的统计；`data/qa/qa_generation_log.jsonl` 为追加写入，"
              "文件里另有历次运行的记录）", "",
              f"- 成功 `{st.get('ok', 0)}`　跳过（source 不足以支撑问题）`{st.get('skip', 0)}`　"
              f"失败 `{st.get('fail', 0)}`", ""]
        by_cat_st = Counter((x.get("qa_type"), x["status"]) for x in log_rows)
        Q += ["各类任务的 ok / skip / fail：", "",
              "| 类别 | ok | skip | fail |", "|---|---|---|---|"]
        for t in TARGETS:
            Q.append(f"| `{t}` | {by_cat_st.get((t, 'ok'), 0)} | {by_cat_st.get((t, 'skip'), 0)} "
                     f"| {by_cat_st.get((t, 'fail'), 0)} |")
        Q.append("")
        skips = [x for x in log_rows if x["status"] == "skip"][:10]
        if skips:
            Q += ["跳过的原因（前 10 条）：", ""]
            Q += [f"- `{x['job_key'][:80]}`：{x['reason'][:100]}" for x in skips]
            Q.append("")

    Q += ["## 复核建议", "",
          "1. 先看第 13 项（自行计算嫌疑）——本批最容易出问题的就是模型自作主张算比率。",
          "2. 数字类告警优先人工核对——这类最可能是真编造，也可能是中文数字写法差异造成的误报。",
          "3. 抽检按 `output/qa_review_sample.md` 走一遍，六类结论勾选："
          "正确 / 部分正确 / 编造 / 来源错误 / 问题错误 / 其他；"
          "`verified` 填 `TRUE` 或 `FALSE`（留空＝未抽检）。"
          "判为 `编造`/`来源错误` 的整条删除或按 chunk 重生成。",
          "4. 同比题要特别看：答案有没有把两个年度的数字混在一起、有没有引到第三个年度。",
          "5. 财务指标题要特别看：指标名是否与原文一致（不得把「债务净额」写成「净负债」）、"
          "数值是否取对了年份那一列。", ""]
    QUALITY_MD.write_text("\n".join(Q), encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
