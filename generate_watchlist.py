#!/usr/bin/env python3
"""
FCN Watchlist Generator
Input:  FCN_Results.xlsx (sorted by score desc, take top N)
Output: watchlist.json  (website sole data source)

Flow:
  1. Read Excel → top N stocks
  2. yfinance   → 10-week sparklines + latest price change
  3. DeepSeek API → bullets + structured analysis (skill spec format)
  4. Merge all  → watchlist.json
"""

import json, os, sys, re, time, argparse
from datetime import datetime, timedelta, time as dt_time

# Windows GBK console fix (only when running directly, not when imported)
if sys.platform == "win32" and hasattr(sys.stdout, 'buffer'):
    import io
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
import yfinance as yf
from openai import OpenAI
import json_repair

# ── Config ────────────────────────────────────────────────────────────────────
# 所有路径相对本脚本所在目录，团队成员 clone 到任何位置都能直接跑
_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH   = r"C:\Users\liy22223\Desktop\FCN筛选器v1\screener\FCN_Results.xlsx"   # 单文件模式（旧流程兜底，仅 Dave 机器）
# 三种策略类型 → 三个 Excel（不存在的文件自动跳过；同一标的出现在多个文件 = 多类型）
EXCEL_INPUTS = {
    "稳健": os.path.join(_DIR, "低波精选组.xlsx"),
    "进取": os.path.join(_DIR, "高波精选组.xlsx"),
    "热度": os.path.join(_DIR, "市场热度榜.xlsx"),
}
OUTPUT_PATH  = os.path.join(_DIR, "watchlist.json")
TOP_N        = 30   # 单文件模式每次取多少只
TOP_N_PER_TYPE = 20 # 三类型模式每个 Excel 取多少只
MODEL        = "deepseek-chat"   # 2026-08-17: deepseek-v4-pro 变推理模型、content返回空→改回 deepseek-chat
API_BASE     = "https://api.deepseek.com"
BATCH_SIZE   = 3
MAX_WORKERS  = 3

# ── 平台名清洗 ────────────────────────────────────────────────────────────────
# DeepSeek 常无视 prompt 里的禁令，把「富途/Yahoo Finance/yfinance」直接写进正文。
# 这些数据仅用于我方 fact-check，报告对外不得出现平台/供应商名，只保留底层出处
# （分析师共识 / 公司财报 / 交易所行情 / 投行评级 等）。此步为确定性兜底，幂等。
_PLATFORM_HEAD  = r'(?:富途|Futu|futu|Yahoo\s*Finance|Yahoo|yahoo|yfinance|雅虎财经|雅虎)'
_PLATFORM_SUFF  = (r'(?:实时|市场|财报|财务|数据快照|数据|快照|统计|汇总|提供|显示|'
                   r'中的|里的|的|信息|\s)*')
_PLATFORM_PHRASE = re.compile(_PLATFORM_HEAD + _PLATFORM_SUFF)
_SRC_PREFIX = r'(?:数据来源|来源|依据|注)'
_SRC_STOP = {'', '输入', '输入中的', '与', '及', '和', '根据', '来自', '此处',
             '数据', '注', '目标价'}

def _clean_src_paren(inner: str) -> str:
    s = _PLATFORM_PHRASE.sub('', inner)
    s = re.sub(r'[·/、]{2,}', '·', s)
    s = re.sub(r'(?:^|(?<=[:：]))[·/、,，\s]+', '', s)
    s = re.sub(r'[·/、,，\s]+$', '', s)
    s = re.sub(_SRC_PREFIX + r'\s*[:：]\s*$', '', s)
    s = re.sub(r'^(?:与|及|和|根据|来自|此处)+', '', s)
    s = s.strip(' ，,、·/与及和')
    core = re.sub(r'^' + _SRC_PREFIX + r'\s*[:：]\s*', '', s).strip(' ，,、·/与及和')
    return '' if core in _SRC_STOP else s

def strip_platform(text):
    if not isinstance(text, str) or not text:
        return text
    text = re.sub(r'（([^（）]*)）',
                  lambda m: (lambda c: f'（{c}）' if c else '')(_clean_src_paren(m.group(1))),
                  text)
    text = re.sub(r'\(([^()]*)\)',
                  lambda m: (lambda c: f'({c})' if c else '')(_clean_src_paren(m.group(1))),
                  text)
    text = _PLATFORM_PHRASE.sub('', text)                       # 行内平台短语
    text = re.sub(r'(?:根据|依据|据|依|基于)\s*(?=[，,：:、])', '', text)  # 悬空引导词
    text = re.sub(r'([。！？；;：:])\s*[，,、：:·/]+', r'\1', text)
    text = re.sub(r'(<p>|<li>|「|“)\s*[，,、：:·/]+', r'\1', text)
    text = re.sub(r'^[\s，,、：:；;·/]+', '', text)
    text = re.sub(r'\s+([。，、；：])', r'\1', text)
    text = re.sub(r'[，,]{2,}', '，', text)
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text

def sanitize_platform(o):
    if isinstance(o, dict):
        return {k: sanitize_platform(v) for k, v in o.items()}
    if isinstance(o, list):
        return [sanitize_platform(v) for v in o]
    return strip_platform(o)

# ── Optional field helpers (return None if cell is blank/NaN) ────────────────
def _opt_float(v):
    try:    return float(v) if v is not None and str(v).strip() not in ('', 'nan') else None
    except: return None

def _opt_int(v):
    # Handle values like "6M", "12M" by stripping trailing non-numeric characters
    if isinstance(v, str):
        v = v.strip().rstrip('Mm月').strip()
    f = _opt_float(v)
    return int(f) if f is not None else None

def _opt_str(v):
    s = str(v).strip() if v is not None else ''
    return s if s not in ('', 'nan', 'None') else None

# ── Code helpers ──────────────────────────────────────────────────────────────
def to_display_code(futu: str) -> str:
    """US.IREN → IREN.US   HK.09992 → 9992.HK"""
    mkt, sym = futu.split(".", 1)
    return f"{sym}.{mkt}"

def to_yf_code(futu: str) -> str:
    """US.IREN → IREN   HK.09992 → 9992.HK   (strips Futu leading zero, pads to 4 digits)"""
    mkt, sym = futu.split(".", 1)
    if mkt == "US":
        return sym
    # HK: Futu uses 5-digit (e.g. 09992), Yahoo Finance uses 4-digit (e.g. 9992)
    yf_sym = sym.lstrip("0").zfill(4)
    return f"{yf_sym}.HK"

# ── FCN terms from IV ─────────────────────────────────────────────────────────
def calc_fcn_terms(iv_pct: float) -> dict:
    if iv_pct >= 80:
        return dict(coupon=27.0, strike=72, ki=58, ko=100, kiType="美式敲入", tenor=6,  risk="高")
    if iv_pct >= 60:
        return dict(coupon=22.0, strike=78, ki=63, ko=100, kiType="欧式敲入", tenor=6,  risk="中")
    if iv_pct >= 40:
        return dict(coupon=18.0, strike=82, ki=68, ko=100, kiType="欧式敲入", tenor=6,  risk="中")
    if iv_pct >= 25:
        return dict(coupon=14.0, strike=88, ki=75, ko=100, kiType="欧式敲入", tenor=6,  risk="低")
    return     dict(coupon=12.0, strike=90, ki=78, ko=100, kiType="欧式敲入", tenor=6,  risk="低")

# ── yfinance helpers ──────────────────────────────────────────────────────────
def _drop_live_bar(hist):
    """If the last daily bar is today's still-trading session (before 16:00
    exchange local time, both NYSE and HKEX), drop it — we only want
    completed closes. 初始价格 = 最近一个已完结交易日的收盘价。"""
    if hist is None or len(hist) == 0:
        return hist
    idx_last = hist.index[-1]
    now = datetime.now(idx_last.tz) if idx_last.tz else datetime.now()
    if idx_last.date() == now.date() and now.time() < dt_time(16, 0):
        return hist.iloc[:-1]
    return hist

def get_price_history(yf_code: str) -> dict:
    """2-year daily closes + initial price (last completed session close).
    Returns {'dates': [...], 'closes': [...], 'initialPrice': x, 'initialPriceDate': 'YYYY-MM-DD'}"""
    try:
        hist = yf.Ticker(yf_code).history(period="2y", interval="1d")
        hist = _drop_live_bar(hist)
        closes = hist["Close"].dropna()
        if len(closes) == 0:
            return {}
        return {
            "dates":  [d.strftime("%Y-%m-%d") for d in closes.index],
            "closes": [round(float(p), 3) for p in closes],
            "initialPrice":     round(float(closes.iloc[-1]), 3),
            "initialPriceDate": closes.index[-1].strftime("%Y-%m-%d"),
        }
    except Exception as e:
        print(f"  ⚠  yfinance history {yf_code}: {e}")
        return {}

def get_sparkline(yf_code: str) -> list:
    try:
        hist = yf.Ticker(yf_code).history(period="12wk", interval="1wk")
        closes = hist["Close"].dropna().tolist()[-10:]
        return [round(p, 2) for p in closes]
    except Exception as e:
        print(f"  ⚠  yfinance sparkline {yf_code}: {e}")
        return []

def get_price_change(yf_code: str) -> float:
    try:
        hist = yf.Ticker(yf_code).history(period="5d", interval="1d")
        if len(hist) < 2:
            return 0.0
        prev, curr = hist["Close"].iloc[-2], hist["Close"].iloc[-1]
        return round((curr - prev) / prev * 100, 2)
    except:
        return 0.0

def get_realtime_context(yf_code: str, display_code: str) -> str:
    """Fetch current company info + recent news from Yahoo Finance.
    Returns a compact text block to prepend to the model prompt as ground truth."""
    lines = [f"[实时数据 · {display_code} · 抓取自 Yahoo Finance]"]
    try:
        t = yf.Ticker(yf_code)
        info = t.info or {}

        # Company facts
        fields = [
            ("longName",          "公司全名"),
            ("sector",            "板块"),
            ("industry",          "行业"),
            ("country",           "注册地"),
            ("longBusinessSummary","业务简介"),
        ]
        for key, label in fields:
            val = info.get(key, "")
            if val:
                summary = val[:200] + "…" if len(val) > 200 else val
                lines.append(f"{label}：{summary}")

        # Key financials from info
        fin_map = [
            ("marketCap",            "市值(USD)"),
            ("totalRevenue",         "年营收(USD)"),
            ("revenueGrowth",        "营收同比"),
            ("grossMargins",         "毛利率"),
            ("trailingEps",          "EPS(TTM)"),
            ("forwardPE",            "Forward P/E"),
            ("recommendationKey",    "分析师评级"),
            ("targetMeanPrice",      "目标均价"),
        ]
        fin_lines = []
        for key, label in fin_map:
            val = info.get(key)
            if val is not None:
                if isinstance(val, float) and val < 10:
                    fin_lines.append(f"{label}={val:.1%}" if "率" in label or "同比" in label else f"{label}={val:.2f}")
                else:
                    fin_lines.append(f"{label}={val:,}" if isinstance(val, (int, float)) else f"{label}={val}")
        if fin_lines:
            lines.append("财务快照：" + " | ".join(fin_lines))

        # Recent news headlines (last 8)
        news = t.news or []
        if news:
            lines.append("近期新闻（最新8条）：")
            for n in news[:8]:
                title = n.get("title", "")
                pub   = n.get("providerPublishTime", "")
                if title:
                    lines.append(f"  · {title}")
    except Exception as e:
        lines.append(f"(Yahoo Finance 获取失败: {e})")

    return "\n".join(lines)

# ── Method B: Futu real-time context ─────────────────────────────────────────

# Confirmed field IDs (cross-checked against AAPL Q2-2026 and Tencent FY2025)
# gross_profit removed: field 8003/5003 returns EBIT (operating income), NOT gross profit
# Displaying operating margin as "毛利率" was misleading the AI analysis
_FIN_FIELDS = {
    'US': {'revenue': 8001, 'net_income': 8037, 'operating_cf': 8015},
    'HK': {'revenue': 5001, 'net_income': 5051, 'operating_cf': 5015},
}

def _get_fin_field(report: dict, fid: int):
    """Return (value_float, yoy_float) for a field_id in one Futu report."""
    for item in report.get('item_list', []):
        if item.get('field_id') == fid:
            v = item.get('data')
            y = item.get('yoy')
            try:    v = float(v) if v is not None else None
            except: v = None
            try:    y = float(y) if y is not None else None
            except: y = None
            return v, y
    return None, None

def _fmt_money(v: float, currency: str) -> str:
    """Format large numbers: 1.23B USD or 45.6亿HKD."""
    if currency == 'HKD':
        yi = v / 1e8
        return f"{yi:.1f}亿HKD" if abs(yi) < 1000 else f"{yi/100:.2f}万亿HKD"
    else:
        b = v / 1e9
        return f"{b:.2f}B USD" if abs(b) >= 1 else f"{v/1e6:.0f}M USD"

def _get_futu_financials(ctx, futu_code: str, is_hk: bool) -> list:
    """Pull income statement + cashflow from Futu, return formatted lines."""
    from futu import RET_OK
    lines = []
    mkt   = 'HK' if is_hk else 'US'
    fids  = _FIN_FIELDS[mkt]
    curr  = 'HKD' if is_hk else 'USD'
    # HK: annual reports (type=7), US: quarterly TTM combo (type=9)
    fin_type   = 7 if is_hk else 9
    num_period = 2 if is_hk else 4
    period_lbl = '年度' if is_hk else '季度'

    # ── Income statement ──────────────────────────────────────────────────────
    ret, inc = ctx.get_financials_statements(
        futu_code, statement_type=1, financial_type=fin_type, num=num_period
    )
    time.sleep(1.5)
    if ret == RET_OK and inc and inc.get('report_list'):
        rows = []
        for rpt in inc['report_list']:
            period = str(rpt.get('report_date', ''))[:7]
            rev, rev_yoy = _get_fin_field(rpt, fids['revenue'])
            ni,  ni_yoy  = _get_fin_field(rpt, fids['net_income'])

            parts = [f"[{period}]"]
            if rev:
                yoy = f" YoY{rev_yoy:+.1f}%" if rev_yoy is not None else ""
                parts.append(f"营收={_fmt_money(rev, curr)}{yoy}")
            if ni:
                yoy = f" YoY{ni_yoy:+.1f}%" if ni_yoy is not None else ""
                parts.append(f"净利={_fmt_money(ni, curr)}{yoy}")
            if len(parts) > 1:
                rows.append("  " + " | ".join(parts))

        if rows:
            lines.append(f"财务报表（富途 最近{num_period}期{period_lbl}）：")
            lines.extend(rows)

    # ── Operating cash flow ───────────────────────────────────────────────────
    ret2, cf = ctx.get_financials_statements(
        futu_code, statement_type=3, financial_type=fin_type, num=2
    )
    time.sleep(1.5)
    if ret2 == RET_OK and cf and cf.get('report_list'):
        cf_parts = []
        for rpt in cf['report_list'][:2]:
            period = str(rpt.get('report_date', ''))[:7]
            ocf, _ = _get_fin_field(rpt, fids['operating_cf'])
            if ocf is not None:
                cf_parts.append(f"[{period}]{_fmt_money(ocf, curr)}")
        if cf_parts:
            lines.append("经营现金流：" + " | ".join(cf_parts))

    # ── Revenue breakdown by segment ─────────────────────────────────────────
    ret3, rbk = ctx.get_financials_revenue_breakdown(futu_code)
    time.sleep(1.5)
    if ret3 == RET_OK and rbk is not None:
        items = rbk if isinstance(rbk, list) else (rbk.get('breakdown_list') or [])
        segs = []
        for item in items[:6]:
            name = item.get('name') or item.get('segment_name', '')
            pct  = item.get('percentage') or item.get('pct')
            if name and pct:
                try:    segs.append(f"{name}:{float(pct):.0f}%")
                except: pass
        if segs:
            lines.append("收入构成：" + " | ".join(segs))

    # ── Earnings beat/miss history ────────────────────────────────────────────
    ret4, em = ctx.get_financials_earnings_price_move(futu_code, period_count=4)
    time.sleep(1.5)
    if ret4 == RET_OK and em is not None:
        records = em.to_dict('records') if hasattr(em, 'to_dict') else []
        beats = total = 0
        for row in records:
            if row.get('day_offset') == 1:
                c = float(row.get('close_price') or 0)
                p = float(row.get('last_close_price') or 0)
                if p > 0:
                    total += 1
                    if c > p: beats += 1
        if total > 0:
            lines.append(f"近{total}次财报发布后次日股价：{beats}涨/{total-beats}跌")

    return lines


def get_futu_context(futu_code: str, display_code: str) -> str:
    """Fetch analyst consensus + rating changes + financial statements from Futu."""
    try:
        from futu import OpenQuoteContext, RET_OK
        ctx = OpenQuoteContext(host='127.0.0.1', port=11111)
    except Exception as e:
        return f"(Futu 连接失败: {e})"

    is_hk = futu_code.startswith('HK.')
    lines = [f"[富途实时数据 · {display_code} · {datetime.today().strftime('%Y-%m-%d')}]"]
    try:
        # ── Analyst consensus ─────────────────────────────────────────────────
        ret, data = ctx.get_research_analyst_consensus(futu_code)
        if ret == RET_OK and data:
            avg = data.get('average')
            buy, hold, sell = data.get('buy_cnt',0), data.get('hold_cnt',0), data.get('sell_cnt',0)
            total_cnt = (buy or 0) + (hold or 0) + (sell or 0)
            if avg:
                lines.append(
                    f"分析师共识：目标价均值={avg} | 区间=[{data.get('low')}, {data.get('high')}] | "
                    f"买入{buy}/持有{hold}/卖出{sell}（共{total_cnt}家）"
                )

        # ── Recent rating changes (90 days) ───────────────────────────────────
        from datetime import timedelta
        cutoff = (datetime.today() - timedelta(days=90)).strftime('%Y-%m-%d')
        ret2, data2 = ctx.get_research_rating_summary(
            futu_code, rating_dimension_type=1, num=15, next_key=None, uid=None
        )
        if ret2 == RET_OK and data2:
            rating_map = {1:'强力买入', 2:'买入', 3:'持有', 4:'卖出', 5:'强力卖出'}
            changes = []
            for inst in (data2.get('inst_rating_summary_list') or []):
                items = inst.get('rating_item_list') or []
                if not items: continue
                latest = items[0]
                if (latest.get('recommendation_date_str') or '') < cutoff: continue
                firm   = inst.get('inst_name_simplified') or inst.get('inst_name', '')
                date_s = latest.get('recommendation_date_str', '')
                curr_r = rating_map.get(latest.get('rating', 0), '?')
                target = latest.get('target_price')
                s = f"{firm}({date_s})→{curr_r}"
                if target: s += f" 目标价={target}"
                if len(items) >= 2:
                    prev_r = rating_map.get(items[1].get('rating', 0), '?')
                    if curr_r != prev_r:
                        arrow = '↑' if (items[0].get('rating',0) or 0) < (items[1].get('rating',0) or 0) else '↓'
                        s += f"({arrow}从{prev_r})"
                changes.append(s)
            if changes:
                lines.append("近90天评级变动：")
                for c in changes[:8]: lines.append(f"  · {c}")

        # ── Financial statements ──────────────────────────────────────────────
        fin_lines = _get_futu_financials(ctx, futu_code, is_hk)
        lines.extend(fin_lines)

    except Exception as e:
        lines.append(f"(Futu 数据获取失败: {e})")
    finally:
        try: ctx.close()
        except Exception: pass

    return "\n".join(lines)


# ── DeepSeek prompts ──────────────────────────────────────────────────────────
SYSTEM_PROMPT = """\
你是一位专业的私人银行 FCN（固定票息票据）结构化产品分析师，服务对象是香港/新加坡的高净值客户。
你的输出将装配进「报价解读」报告：Part B 读懂公司（B 卡片组）、Part C 市场观点（看多/看空依据的客观汇总 + 接货价位体检）、Part D 风险提示。

核心写作原则：
- 结论先行：每个板块的第一句必须是核心判断，不以背景铺垫开头
- 数据具体：财务数字必须注明年度/季度，无法核实的绝对不引用
- 白话优先：技术术语紧跟白话解释，格式「术语（白话：……）」
- 客观专业：面向私人银行 HNW 客户，第三人称，语气直白有力

严格禁止：
- 捏造或杜撰任何数据；模型记忆中的数字不得直接当作事实引用
- 空洞表述（如"未来发展潜力巨大"）
- bullet 超过 40 字
- 正文板块缺少 topic sentence
- 输出 schema 之外的 HTML 标签或任何内联样式（仅允许 <p> <b> <ul> <li> 与 class="conc"）
- 【全局】正文任何位置都不得出现数据平台/供应商名称（如 富途/Futu、Yahoo Finance/雅虎、yfinance 等）。输入中标注的这些来源仅供你核对事实，输出时一律归因到数据的“底层出处”：分析师共识、公司财报 / 业绩说明会、交易所、市场行情等；无需点名数据从哪个平台取得。

【Part C 专属 · 合规要求】（仅约束 Part C 的看多/看空两卡，不影响 tagline/bullets/B/D 等其他字段）：
- Part C 是“第三方观点与客观事实的汇总”，不是我方推荐。只陈述可核实的事实与数据，并归因来源（如“分析师共识/公司财报/市场”）。
- 严禁在 Part C 出现任何主观结论或买卖暗示词：我们认为、值得（接）、建议、应该、看好、看淡、推荐、买入、卖出、低估、高估（除非明确归因为某第三方观点）。
- 看多依据与看空顾虑应数量相当、力度对称，避免一边倒；每条尽量带数字与来源。
"""

BATCH_PROMPT = """\
请为以下 {n} 只股票生成投资分析内容。

数据优先级规则（严格遵守）：
1. 标注「富途实时数据」的财务数字 → 必须原样引用，不得用训练记忆覆盖
2. 标注「Yahoo Finance」的公司信息 → 用于核实公司背景
3. 训练数据知识 → 仅用于行业背景、竞争格局等无实时数据的部分
4. 无法确认的推断 → 在 data_quality.model_inferences 中注明

股票数据：
{stock_data}

---

每只股票输出以下 JSON 对象（严格按 schema）。B/C 卡片正文是 HTML 片段，**仅允许 <p> <b> <ul> <li> 标签与 class="conc"**，禁止其他标签、属性与内联样式：

{{
  "ticker": "与输入一致的股票代码",
  "name": "公司常用名：仅当存在官方或媒体广泛使用的中文名时用中文（如：亚马逊、英伟达、戴尔科技）；没有通用中文名的公司直接原样返回英文常用名（如：Coherent、Snowflake、Palantir、CoreWeave）。严禁自行直译/音译造名（如「相干公司」「雪花公司」），严禁中英混搭（如「Palantir 科技」），不要带 Inc./Corp./Ltd. 等后缀",
  "name_en": "Company English Name",
  "sector": "行业分类（如：AI半导体、互联网、新能源、金融等）",
  "tagline": "一句话定位（≤16字，将拼在报告大标题后，如「AI 存储周期的最大受益者」）",
  "intro": "公司速览（120-160字，纯文本不含HTML）：成立时间、总部、上市市场与代码、当前市值、核心业务一句白话定位",
  "bullets": [
    "投资要点①（≤40字，结论先行，含具体数据，非空洞表述）",
    "投资要点②（≤40字，与①完全独立，不重复不互补）"
  ],
  "B": [
    ["公司是做什么的", "<p>150-250字：白话讲清公司靠什么赚钱、给谁提供什么价值；控股背景、市值规模、行业地位。第一句必须是直接点明商业本质的结论句</p>"],
    ["行业格局与竞争位置", "<p>180-280字：行业痛点→公司解法→为什么客户选这家；市场份额、竞争壁垒，可用「术语（白话：……）」</p>"],
    ["最新业绩与财务体质", "<p>180-280字：最新年度/季度营收（绝对值+同比）、净利润、毛利率走势、经营现金流。标注「富途实时数据」的数字必须原样引用并注明期间</p>"],
    ["市场怎么看", "<p>120-220字：分析师共识目标价与现价差距、近期评级变动方向、买卖比例；引用输入中的富途分析师数据</p>"]
  ],
  "C": [
    ["市场看多依据 · Bulls Say", "<ul><li>客观依据①（≤60字，必须含数据并归因来源，如「分析师共识目标价较现价 +28%」「FY25Q3 毛利率回升至 39%（公司财报）」）</li><li>依据②</li><li>依据③</li></ul>"],
    ["市场看空顾虑 · Bears Say", "<ul><li>客观顾虑①（≤60字，具体到数据/事件并归因，如「存储行业强周期，过去8季最大单日跌幅 −14%」）</li><li>顾虑②</li><li>顾虑③</li></ul>"]
  ],
  "health_conc": "接货价位一句话结论（≤45字，结合输入给出的行权价折扣与敲入缓冲，如「7.8 折接货+37%缓冲，结构留有余地」）",
  "c4_prose": "100-160字：现价处于什么位置（结合输入的最大跌幅/均线数据）、接货价相当于回到什么水平、这个折扣对该标的波动性而言厚不厚",
  "risks": [
    "<b>风险名</b>——结合该标的具体化的说明（如财报日期、竞争对手、政策节点），≥4条，每条≤70字"
  ],
  "data_quality": {{
    "verified": ["已核实数据项"],
    "broker_views": ["投行推论（来源：XX）"],
    "model_inferences": ["分析推断内容（依据：训练数据）"]
  }}
}}

输出格式：JSON 数组 [{{...}}, {{...}}]，不输出任何其他文字。
"""

# ── DeepSeek API call ─────────────────────────────────────────────────────────
def extract_json_array(text: str) -> list:
    text = text.strip()
    # Strip markdown fences
    text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.MULTILINE)
    text = re.sub(r'\s*```\s*$', '', text, flags=re.MULTILINE)
    # Find outermost array
    m = re.search(r'\[[\s\S]*\]', text)
    raw = m.group() if m else text
    # Try strict parse first, fall back to json_repair
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        repaired = json_repair.repair_json(raw, return_objects=True)
        if isinstance(repaired, list):
            return repaired
        raise ValueError(f"Cannot parse JSON even after repair: {raw[:200]}")

def analyze_batch(client: OpenAI, batch: list) -> list:
    stock_blocks = []
    for i, s in enumerate(batch):
        # ── Method A: structured FCN screener fields ──────────────────────────
        lines = [
            f"{i+1}. 代码={s['display_code']} | 名称={s['name']} | 市场={s['market']}",
            f"   现价={s['price']:.2f} {s['currency']} | 市值={s['market_cap']:.1f}B | IV={s['iv_pct']:.1f}%({s['iv_src']})",
            f"   分析师目标价上涨空间={s['analyst_upside']:.1%} | 综合评分={s['display_score']:.1f}",
        ]
        if s.get('catalyst_raw') is not None:
            lines.append(f"   催化剂评分(0-1)={s['catalyst_raw']:.3f} | 期权OI={int(s['option_oi'] or 0):,}")
        if s.get('max_drop') is not None:
            lines.append(
                f"   过去8季最大单日跌幅={s['max_drop']:.1%} | "
                f"现价/50日均线={s['sma50_ratio']:.3f} | "
                f"50日均线斜率={s['sma50_slope']:.4f}%/日"
            )
        # ── Method B: Futu analyst consensus + rating changes ─────────────────
        futu_ctx = s.get('futu_context', '')
        if futu_ctx:
            lines.append(futu_ctx)
        # ── Yahoo Finance: business summary + news ────────────────────────────
        yf_ctx = s.get('yf_context', '')
        if yf_ctx:
            lines.append(yf_ctx)

        stock_blocks.append("\n".join(lines))
    stock_data = "\n\n".join(stock_blocks)
    prompt = BATCH_PROMPT.format(n=len(batch), stock_data=stock_data)

    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                max_tokens=8192,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",   "content": prompt}
                ]
            )
            return extract_json_array(resp.choices[0].message.content)
        except Exception as e:
            wait = 2 ** attempt * 6
            if attempt < 2:
                print(f"  ⚠  Batch retry {attempt+1}/3 in {wait}s: {e}")
                time.sleep(wait)
            else:
                print(f"  ❌  Batch failed: {e}")
                return []
    return []

# ── 每日市场热点（美股 + 港股均用 Google News RSS → DeepSeek 综述，免 key）────────
def _google_news(query, hl, gl, ceid, limit=40):
    """通用 Google News RSS 抓取。失败返回 []。"""
    try:
        import feedparser, urllib.parse
    except ImportError:
        print("  ⚠  未安装 feedparser，跳过新闻抓取"); return []
    url = (f"https://news.google.com/rss/search?q={urllib.parse.quote(query)}"
           f"&hl={hl}&gl={gl}&ceid={ceid}")
    try:
        entries = feedparser.parse(url).entries[:limit]
    except Exception as e:
        print(f"  ⚠  新闻获取失败({gl}): {e}"); return []
    out = []
    for e in entries:
        src = getattr(getattr(e, "source", None), "title", "") or ""
        out.append({"title": getattr(e, "title", ""),
                    "summary": re.sub(r"<[^>]+>", "", getattr(e, "summary", ""))[:280],
                    "source": src, "url": getattr(e, "link", "")})
    return [x for x in out if x["title"]]

def _fetch_us_news(limit=40):
    """美股近期市场新闻（Google News 英文 RSS）。"""
    return _google_news("US stock market OR S&P 500 OR Nasdaq OR Federal Reserve OR earnings",
                        "en-US", "US", "US:en", limit)

def _fetch_hk_news(limit=40):
    """港股近期市场新闻（Google News 中文 RSS）。"""
    return _google_news("恒生指数 OR 港股 OR 香港股市 OR 港股通",
                        "zh-HK", "HK", "HK:zh-Hant", limit)

HOTSPOT_SYSTEM = (
    "你是港美股市场分析师，为结构化产品客户撰写每周市场简报。"
    "只能依据用户提供的真实新闻列表，严禁使用你自己记忆中的任何信息，"
    "严禁编造未在列表中出现的事件、公司、数字或日期。只输出 JSON。"
)

HOTSPOT_PROMPT = """下面是最近 1-2 个交易日抓取的真实新闻（美股、港股均来自 Google News）。

【美股新闻】
{us_news}

【港股新闻】
{hk_news}

请分别为【美股】和【港股】各挑选 2 条最重要的"市场级/板块级"新闻（不要选个股研报、机构持仓变动这类碎片新闻），输出如下 JSON：

{{
  "us": {{"items": [
    {{
      "headline": "中文标题，结论先行，≤30字",
      "facts": "事实陈述，40-90字，只能用上方新闻列表中出现的信息，可含具体数字",
      "stance": "bullish | bearish | neutral",
      "impact": "板块层面的利好/利空判断 + 传导机制，40-90字",
      "sources": [{{"title": "来源名 · 简述", "url": "列表中对应新闻的真实 url"}}]
    }}
  ]}},
  "hk": {{"items": [ ...同上 2 条... ]}}
}}

硬性要求：
1. facts 与 sources 的信息必须来自上方列表，url 必须是列表中真实出现的链接，不得虚构。
2. impact 只讲对"板块"或"新闻主角公司"的事实性影响（如何利好/利空及其逻辑）；
   严禁出现"票息""敲入""敲出""FCN""结构化产品""票据"等任何衍生品定价词汇，
   严禁预测我们推荐标的的任何衍生品条款变化。
3. stance 取该新闻对相关板块的总体方向：利多=bullish，利空=bearish，多空分化/中性=neutral。
4. 每个市场恰好 2 条。只输出 JSON，不要任何额外文字。"""

def get_market_hotspot(client, today):
    """抓取美股+港股新闻 → DeepSeek 综述 → 返回 marketHotspot dict；任一环节失败返回 None。"""
    us_raw = _fetch_us_news()
    hk_raw = _fetch_hk_news()
    if not us_raw and not hk_raw:
        print("  ⚠  美股、港股新闻均为空，跳过每日市场热点"); return None
    if client is None:
        print("  ⚠  无 DeepSeek client（--no-ai），跳过每日市场热点"); return None

    def _fmt(items, n=30):
        return "\n".join(
            f"- {x['title']}（{x.get('source','')}{('·'+x['sentiment']) if x.get('sentiment') else ''}）"
            f" {x.get('summary','')}\n  url: {x['url']}"
            for x in items[:n]) or "（无）"

    prompt = HOTSPOT_PROMPT.format(us_news=_fmt(us_raw), hk_news=_fmt(hk_raw))
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model=MODEL, max_tokens=8000,   # 3500 会截断中文 JSON → 解析必挂
                messages=[{"role": "system", "content": HOTSPOT_SYSTEM},
                          {"role": "user",   "content": prompt}])
            txt = resp.choices[0].message.content.strip()
            txt = re.sub(r'^```(?:json)?\s*', '', txt, flags=re.MULTILINE)
            txt = re.sub(r'\s*```\s*$', '', txt, flags=re.MULTILINE)
            m = re.search(r'\{[\s\S]*\}', txt)
            raw = m.group() if m else txt
            # Try strict parse first, fall back to json_repair (同个股分析路径)
            try:
                data = json.loads(raw)
            except json.JSONDecodeError:
                data = json_repair.repair_json(raw, return_objects=True)
                if not isinstance(data, dict):
                    raise ValueError(f"Cannot parse hotspot JSON even after repair: {raw[:200]}")
            us_items = (data.get("us") or {}).get("items") or []
            hk_items = (data.get("hk") or {}).get("items") or []
            # 完整性校验：必须美股/港股各 2 条，残缺（多为截断修复所致）视为失败重试
            if len(us_items) < 2 or len(hk_items) < 2:
                raise ValueError(f"items incomplete: us={len(us_items)} hk={len(hk_items)} (need 2+2)")
            print(f"  ✅  每日热点：美股 {len(us_items)} 条 · 港股 {len(hk_items)} 条")
            return {"generatedAt": today.strftime("%Y-%m-%d"),
                    "date": f"{today.year}/{today.month}/{today.day}",
                    "us": {"items": us_items}, "hk": {"items": hk_items}}
        except Exception as e:
            print(f"  ⚠  热点综述 attempt {attempt+1} 失败: {e}"); time.sleep(6)
    print("  ❌  每日市场热点生成失败（3 次尝试）"); return None

# ── 市场热点图：板块过去一周表现（富途行业板块原生数据）──────────────────────────
# 11 个 GICS 板块 → 富途行业子板块代码；板块过去 5 个交易日涨跌幅按成交额加权聚合。
# 美股 + 港股均取富途 INDUSTRY 板块的日 K（板块本身是可报价标的），免 ETF/篮子代理。
_HEATMAP_SECTORS = [
    ("科技",   "Technology"),     ("通信",   "Communication"),
    ("可选消费","Consumer Disc."), ("必需消费","Consumer Staples"),
    ("医疗",   "Health Care"),     ("金融",   "Financials"),
    ("工业",   "Industrials"),     ("能源",   "Energy"),
    ("材料",   "Materials"),       ("公用",   "Utilities"),
    ("地产",   "Real Estate"),
]
_HEATMAP_US = {
    "科技":   ["2015","2016","2470","2508","2252","2492","2275","2072"],
    "通信":   ["2004","2088","2253","2491","2497"],
    "可选消费":["2431","2225","2468","2106","2276","2494","2046","2240","2460"],
    "必需消费":["2108","2095","2459","2264","2003","2427"],
    "医疗":   ["2069","2513","2280","2011","2246","2220"],
    "金融":   ["2481","2456","2260","2249","2490","2484","2512","2261"],
    "工业":   ["2089","2102","2267","2500","2463","2474","2219","2090"],
    "能源":   ["2058","2224","2060","2226","2257"],
    "材料":   ["2020","2068","2110","2101","2510","2034","2237"],
    "公用":   ["2472","2488","2489","2458","2462"],
    "地产":   ["2140","2141","2466","2457","2482","2038","2511"],
}
_HEATMAP_HK = {
    "科技":   ["1013","1360","1100","1359","1274","1052","1053","23364","23363"],
    "通信":   ["23360","1054","1029","1027","1026"],
    "可选消费":["23361","1083","1040","1041","1049","1277","1270","1069","1071","1034","1056"],
    "必需消费":["1010","1070","1080","1072","1356","1062","1001","23850"],
    "医疗":   ["1050","1067","1284","1012","1086","1357"],
    "金融":   ["1079","1003","1068","1030","1004","23362","1007"],
    "工业":   ["1063","1065","1066","1005","1076","1074","1073","1025","1095"],
    "能源":   ["1042","1043","1044","1016","1358"],
    "材料":   ["1046","1075","1077","1078","1084","1028","1033","1006"],
    "公用":   ["1051","1045","1039"],
    "地产":   ["1019","1020","1311","1090","1089"],
}

def _heatmap_one_market(ctx, mkt: str, mapping: dict) -> list:
    """订阅板块日 K + 快照成交额 → 各 GICS 板块成交额加权 5 日涨跌幅。"""
    from futu import RET_OK, KLType, AuType, SubType
    codes = [f"{mkt}.LIST{n}" for ns in mapping.values() for n in ns]
    ctx.subscribe(codes, [SubType.K_DAY])
    time.sleep(2)
    # 成交额（批量快照，单次上限 400）用作板块内子行业权重
    turnover = {}
    for i in range(0, len(codes), 400):
        ret, snap = ctx.get_market_snapshot(codes[i:i+400])
        if ret == RET_OK and snap is not None:
            for _, row in snap.iterrows():
                turnover[row["code"]] = float(row.get("turnover") or 0)
    # 5 日涨跌幅（实时日 K，不消耗历史 K 线额度）
    ret5 = {}
    win = None        # 实际 K 线窗口 (起始日, 结束日)，用于 weekRange 标签
    for c in codes:
        ret, df = ctx.get_cur_kline(c, 6, KLType.K_DAY, AuType.QFQ)
        if ret == RET_OK and df is not None and len(df) >= 2:
            cl = df["close"]
            base_idx = -6 if len(df) >= 6 else 0
            base = cl.iloc[base_idx]
            if base:
                ret5[c] = (cl.iloc[-1] / base - 1) * 100
                if win is None and "time_key" in df.columns:
                    win = (str(df["time_key"].iloc[base_idx])[:10],
                           str(df["time_key"].iloc[-1])[:10])
    out = []
    for cn, en in _HEATMAP_SECTORS:
        num = den = 0.0; used = 0
        for n in mapping.get(cn, []):
            c = f"{mkt}.LIST{n}"
            if c in ret5:
                w = turnover.get(c, 0) or 1.0
                num += ret5[c] * w; den += w; used += 1
        if den > 0:
            out.append({"name": cn, "name_en": en,
                        "change": round(float(num / den), 2), "n": used})
    out.sort(key=lambda x: x["change"], reverse=True)
    return out, win

def get_sector_heatmap(today):
    """美股 + 港股各 GICS 板块过去一周表现（富途行业板块）。失败返回 None。"""
    print("[+] 生成市场热点图（富途行业板块 → 板块周涨跌幅，成交额加权）...")
    try:
        from futu import OpenQuoteContext
        ctx = OpenQuoteContext(host="127.0.0.1", port=11111)
    except Exception as e:
        print(f"  ⚠  Futu 连接失败，跳过市场热点图: {e}"); return None
    win = None
    try:
        us, us_win = _heatmap_one_market(ctx, "US", _HEATMAP_US)
        hk, hk_win = _heatmap_one_market(ctx, "HK", _HEATMAP_HK)
        win = us_win or hk_win
    except Exception as e:
        print(f"  ⚠  市场热点图获取失败: {e}"); us = hk = []
    finally:
        try: ctx.close()
        except Exception: pass
    if not us and not hk:
        print("  ⚠  美股、港股板块数据均为空，跳过市场热点图"); return None
    # weekRange 取自实际 K 线窗口（最近 5 个交易日），而非当周一~周五
    if win:
        _fmt = lambda s: "{}/{}/{}".format(s[:4], int(s[5:7]), int(s[8:10]))
        week_range = f"{_fmt(win[0])}-{_fmt(win[1])}"
    else:
        mon = today - timedelta(days=today.weekday()); fri = mon + timedelta(days=4)
        wk = lambda d: f"{d.year}/{d.month}/{d.day}"; week_range = f"{wk(mon)}-{wk(fri)}"
    print(f"  ✅  市场热点图：美股 {len(us)} 个 · 港股 {len(hk)} 个板块  ({week_range})")
    return {"generatedAt": today.strftime("%Y-%m-%d"),
            "weekRange": week_range, "us": us, "hk": hk}

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Generate FCN watchlist.json from Excel")
    parser.add_argument("--excel",        default=None, help="单文件模式（不打类型标签）")
    parser.add_argument("--excel-stable", default=None, help="低波精选 Excel 路径")
    parser.add_argument("--excel-growth", default=None, help="高波精选 Excel 路径")
    parser.add_argument("--excel-hot",    default=None, help="市场热度型 Excel 路径")
    parser.add_argument("--output",       default=OUTPUT_PATH)
    parser.add_argument("--top",          type=int, default=None,
                        help=f"每个数据源取前 N 只（默认：单文件 {TOP_N}，三类型每类 {TOP_N_PER_TYPE}）")
    parser.add_argument("--dry-run",      action="store_true", help="First 5 stocks only")
    parser.add_argument("--no-sparkline", action="store_true", help="Skip yfinance")
    parser.add_argument("--no-ai",        action="store_true", help="Skip DeepSeek API")
    parser.add_argument("--week",         default=None,        help="Override week e.g. W24")
    args = parser.parse_args()

    # key 优先级：环境变量 > 脚本目录下 deepseek_key.txt（已 gitignore，绝不能提交到公开仓库）
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        _key_file = os.path.join(_DIR, "deepseek_key.txt")
        if os.path.exists(_key_file):
            with open(_key_file, encoding="utf-8") as kf:
                api_key = kf.read().strip()
    if not api_key and not args.no_ai:
        print("❌  未找到 DeepSeek key：请设置环境变量 DEEPSEEK_API_KEY，"
              "或在本目录创建 deepseek_key.txt 并粘贴 key"); sys.exit(1)

    client = OpenAI(api_key=api_key, base_url=API_BASE) if api_key else None
    today    = datetime.today()
    iso_wk   = today.isocalendar()[1]
    week_str = args.week or f"W{iso_wk:02d} {today.year}"

    # ── 1. Read Excel(s) ──────────────────────────────────────────────────────
    if args.excel:                       # 显式单文件模式
        sources = [(None, args.excel)]
        per_n = 5 if args.dry_run else (args.top or TOP_N)
    else:                                # 三类型模式（默认）
        sources = [
            ("稳健", args.excel_stable or EXCEL_INPUTS["稳健"]),
            ("进取", args.excel_growth or EXCEL_INPUTS["进取"]),
            ("热度", args.excel_hot    or EXCEL_INPUTS["热度"]),
        ]
        if not any(Path(p).exists() for _, p in sources):
            print(f"  三类型 Excel 均不存在，回退单文件模式: {EXCEL_PATH}")
            sources = [(None, EXCEL_PATH)]
            per_n = 5 if args.dry_run else (args.top or TOP_N)
        else:
            per_n = 2 if args.dry_run else (args.top or TOP_N_PER_TYPE)

    def _parse_terms_excel(df, n):
        """条款式 Excel（区域/代码/名称/行业/执行价/敲入价/敲入类型/敲出价/票息）。
        无市场数据 → price/mktcap 由 yfinance 运行时补，条款为小数（0.8→80%）。"""
        out = []
        for _, row in df.head(n).iterrows():
            region = str(row['区域']).strip()
            # 代码可能是 'LRCX US' / '6651 HK' / 'IREN.US' / 'IREN' / '9992' → 取纯代码
            raw    = str(row['代码']).strip().split('.')[0].split()[0]
            if region.startswith('港'):
                mkt, futu = 'HK', 'HK.' + raw.zfill(5)
            else:
                mkt, futu = 'US', 'US.' + raw.upper()
            industry = None
            for c in ('子板块', '行业', '板块'):
                if c in df.columns and _opt_str(row.get(c)):
                    industry = _opt_str(row.get(c)); break
            pct = lambda v: round(v * 100, 2) if v is not None and v < 2 else v
            out.append({
                "futu_code":      futu,
                "display_code":   to_display_code(futu),
                "yf_code":        to_yf_code(futu),
                "name":           str(row.get('名称', raw)),
                "market":         mkt,
                "price":          0.0,        # 运行时由 yfinance 补
                "market_cap":     0.0,        # 运行时由 yfinance 补
                "currency":       "USD" if mkt == "US" else "HKD",
                "iv_pct":         0.0,        # 条款式输入无 IV
                "analyst_upside": 0.0,
                "display_score":  0.0,        # 无评分体系 → 前端自动隐藏评分 UI
                "avg_vol":        0.0,
                "industry_hint":  industry,
                "manual_coupon":  pct(_opt_float(row.get('票息'))),
                "manual_strike":  pct(_opt_float(row.get('执行价'))),
                "manual_ki":      pct(_opt_float(row.get('敲入价'))),
                "manual_ko":      pct(_opt_float(row.get('敲出价'))),
                "manual_ki_type": _opt_str(row.get('敲入类型')),
                "manual_tenor":   None,       # 统一 6 个月
                "catalyst_raw":   None, "max_drop": None, "sma50_ratio": None,
                "sma50_slope":    None, "option_oi": None, "iv_src": 'manual',
            })
        return out

    def _parse_excel(path, n):
        """读一个筛选结果 Excel → stock dict 列表（不含类型标签）。"""
        df = pd.read_excel(path)
        if '代码' in [str(c).strip() for c in df.columns]:
            return _parse_terms_excel(df, n)     # 条款式格式
        col_map = {}
        for col in df.columns:
            c = col.replace('\n', ' ').strip()
            if c == 'Code':                            col_map[col] = 'code'
            elif c == 'Name':                          col_map[col] = 'name'
            elif c == 'Mkt':                           col_map[col] = 'market'
            elif c == 'Price':                         col_map[col] = 'price'
            elif 'Mkt Cap' in c:                       col_map[col] = 'market_cap'
            elif 'IV 6M' in c:                         col_map[col] = 'iv6m'
            elif 'Analyst' in c and 'Upside' in c:     col_map[col] = 'analyst_upside'
            elif 'DISPLAY' in c and 'SCORE' in c:      col_map[col] = 'display_score'
            elif 'Avg Vol' in c:                       col_map[col] = 'avg_vol'
            elif c in ('Coupon%', 'Coupon', '票息%'):  col_map[col] = 'manual_coupon'
            elif c in ('Strike%', 'Strike', '行权价%'): col_map[col] = 'manual_strike'
            elif c in ('KI%', 'KI', '敲入价%'):        col_map[col] = 'manual_ki'
            elif c in ('KI Type', 'KIType', '敲入类型'): col_map[col] = 'manual_ki_type'
            elif c in ('Tenor', 'Tenor(M)', '期限'):   col_map[col] = 'manual_tenor'
            elif c in ('KO%', 'KO', '敲出价%', '敲出价'): col_map[col] = 'manual_ko'
            elif 'Catalyst' in c:                      col_map[col] = 'catalyst_raw'
            elif 'Max 1D Drop' in c:                   col_map[col] = 'max_drop'
            elif 'Price/50DMA' in c:                   col_map[col] = 'sma50_ratio'
            elif '50DMA Slope' in c:                   col_map[col] = 'sma50_slope'
            elif 'Option OI' in c:                     col_map[col] = 'option_oi'
            elif c == 'IV Src':                        col_map[col] = 'iv_src'
        df = df.rename(columns=col_map)
        df = df.head(n)
        out = []
        for _, row in df.iterrows():
            futu   = str(row['code']).strip()
            mkt    = str(row.get('market', 'US')).strip()
            iv_raw = float(row.get('iv6m', 0) or 0)
            iv_pct = round(iv_raw * 100 if iv_raw < 2 else iv_raw, 2)
            out.append({
                "futu_code":      futu,
                "display_code":   to_display_code(futu),
                "yf_code":        to_yf_code(futu),
                "name":           str(row.get('name', futu)),
                "market":         mkt,
                "price":          round(float(row.get('price', 0) or 0), 2),
                "market_cap":     round(float(row.get('market_cap', 0) or 0), 2),
                "currency":       "USD" if mkt == "US" else ("HKD" if mkt == "HK" else "CNY"),
                "iv_pct":         iv_pct,
                "analyst_upside": float(row.get('analyst_upside', 0) or 0),
                "display_score":  float(row.get('display_score', 0) or 0),
                "avg_vol":        float(row.get('avg_vol', 0) or 0),
                "manual_coupon":  _opt_float(row.get('manual_coupon')),
                "manual_strike":  _opt_float(row.get('manual_strike')),
                "manual_ki":      _opt_float(row.get('manual_ki')),
                "manual_ki_type": _opt_str(row.get('manual_ki_type')),
                "manual_tenor":   _opt_int(row.get('manual_tenor')),
                # Method A: extra screener fields for LLM context
                "catalyst_raw":   _opt_float(row.get('catalyst_raw')),
                "max_drop":       _opt_float(row.get('max_drop')),
                "sma50_ratio":    _opt_float(row.get('sma50_ratio')),
                "sma50_slope":    _opt_float(row.get('sma50_slope')),
                "option_oi":      _opt_float(row.get('option_oi')),
                "iv_src":         _opt_str(row.get('iv_src')) or 'futu',
                "manual_ko":      _opt_float(row.get('manual_ko')),
            })
        return out

    print("\n[1/4] Reading Excel(s)...")
    stocks, by_code, excel_names = [], {}, []
    for stype, path in sources:
        if not Path(path).exists():
            print(f"  ⚠  跳过{('「' + stype + '」') if stype else ''}：文件不存在 {path}")
            continue
        rows = _parse_excel(path, per_n)
        excel_names.append(Path(path).name)
        print(f"  {stype or '默认'}: {Path(path).name} → {len(rows)} 只")
        for s in rows:
            key = s['futu_code']
            if key in by_code:
                if stype and stype not in by_code[key]['types']:
                    by_code[key]['types'].append(stype)
            else:
                s['types'] = [stype] if stype else []
                by_code[key] = s
                stocks.append(s)
    if not stocks:
        print("❌  没有读到任何标的，请检查 Excel 路径"); sys.exit(1)
    print(f"  Loaded {len(stocks)} stocks（去重后；同标的多类型已合并）")

    # ── 2. yfinance sparklines ────────────────────────────────────────────────
    if not args.no_sparkline:
        print(f"\n[2/4] Fetching sparklines via yfinance...")
        for i, s in enumerate(stocks):
            s['sparkline']        = get_sparkline(s['yf_code'])
            _pc = get_price_change(s['yf_code'])
            s['priceChange']      = 0.0 if (_pc is None or _pc != _pc) else _pc  # NaN护栏(_pc!=_pc): 否则浏览器JSON.parse挂→整站空白
            s['price_history']    = get_price_history(s['yf_code'])
            # 条款式 Excel 不含价格/市值 → 运行时补
            if not s['price'] and s['price_history'].get('initialPrice'):
                s['price'] = s['price_history']['initialPrice']
            if not s['market_cap']:
                try:
                    fi = yf.Ticker(s['yf_code']).fast_info
                    mc = getattr(fi, 'market_cap', None)
                    if mc: s['market_cap'] = round(mc / 1e9, 2)
                except Exception:
                    pass
            s['yf_context']       = get_realtime_context(s['yf_code'], s['display_code'])
            s['futu_context']     = get_futu_context(s['futu_code'], s['display_code'])
            print(f"  [{i+1:2d}/{len(stocks)}] {s['display_code']:<14} "
                  f"{len(s['sparkline'])} weeks  Δ{s['priceChange']:+.1f}%")
            time.sleep(0.35)
    else:
        for s in stocks:
            s['sparkline']        = []
            s['priceChange']      = 0.0
            s['price_history']    = {}
            s['yf_context']       = get_realtime_context(s['yf_code'], s['display_code'])
            s['futu_context']     = get_futu_context(s['futu_code'], s['display_code'])
        print("\n[2/4] Sparklines skipped (--no-sparkline)")

    # ── 3. DeepSeek analysis ──────────────────────────────────────────────────
    def _validate_ana(ana: dict) -> tuple[bool, list]:
        """Check all required report-schema fields are present and non-empty."""
        issues = []
        if not ana.get("tagline"):                  issues.append("no tagline")
        if not ana.get("intro"):                    issues.append("no intro")
        B = ana.get("B") or []
        if len(B) < 3:                              issues.append(f"B={len(B)}<3")
        C = ana.get("C") or []
        if len(C) < 2:                              issues.append(f"C={len(C)}<2")
        if not ana.get("health_conc"):              issues.append("no health_conc")
        if not ana.get("c4_prose"):                 issues.append("no c4_prose")
        if len(ana.get("risks") or []) < 3:         issues.append("risks<3")
        return len(issues) == 0, issues

    analysis_map = {}
    if not args.no_ai:
        print(f"\n[3/4] Generating analysis ({MODEL})...")
        batches   = [stocks[i:i+BATCH_SIZE] for i in range(0, len(stocks), BATCH_SIZE)]
        completed = 0

        def _run(bi, batch):
            return bi, batch, analyze_batch(client, batch)

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as exe:
            futs = {exe.submit(_run, i, b): i for i, b in enumerate(batches)}
            for fut in as_completed(futs):
                _, batch, results = fut.result()
                for stock, ana in zip(batch, results):
                    analysis_map[stock['display_code']] = ana
                    completed += 1
                    print(f"  ✅ [{completed:2d}/{len(stocks)}] "
                          f"{stock['display_code']} {stock['name']} — 完成")

        # ── 单股重试：修复 batch 中格式不完整的股票 ──────────────────────────
        retry_stocks = []
        for s in stocks:
            ana = analysis_map.get(s['display_code'], {})
            ok, issues = _validate_ana(ana)
            if not ok:
                print(f"  ⚠  {s['display_code']} 分析不完整 ({issues})，单股重试...")
                retry_stocks.append(s)

        for s in retry_stocks:
            for attempt in range(3):
                results = analyze_batch(client, [s])
                if results:
                    ok, issues = _validate_ana(results[0])
                    if ok:
                        analysis_map[s['display_code']] = results[0]
                        print(f"  ✅  {s['display_code']} 重试成功")
                        break
                    print(f"  ⚠  {s['display_code']} 重试 {attempt+1}/3 仍不完整: {issues}")
                else:
                    print(f"  ⚠  {s['display_code']} 重试 {attempt+1}/3 无输出")
                time.sleep(6)
            else:
                print(f"  ❌  {s['display_code']} 重试耗尽，保留空分析")
    else:
        print("\n[3/4] DeepSeek skipped (--no-ai)")

    # ── 4. Merge and write ────────────────────────────────────────────────────
    def _strip_md_bold(v):
        """Recursively remove **bold** markdown markers from AI output."""
        if isinstance(v, str):  return re.sub(r'\*\*([^*]+)\*\*', r'<b>\1</b>', v)
        if isinstance(v, list): return [_strip_md_bold(x) for x in v]
        if isinstance(v, dict): return {k: _strip_md_bold(x) for k, x in v.items()}
        return v

    print(f"\n[4/4] Writing {args.output}...")
    score_max = max((s['display_score'] for s in stocks), default=100) or 100

    output_stocks = []
    for rank, s in enumerate(stocks):
        code = s['display_code']
        ana  = analysis_map.get(code, {})
        fcn  = calc_fcn_terms(s['iv_pct'])
        # Manual Excel values override IV-estimated terms
        if s['manual_coupon']  is not None: fcn['coupon'] = s['manual_coupon']
        if s['manual_strike']  is not None: fcn['strike'] = s['manual_strike']
        if s['manual_ki']      is not None: fcn['ki']     = s['manual_ki']
        if s['manual_ko']      is not None: fcn['ko']     = s['manual_ko']
        if s['manual_ki_type'] is not None: fcn['kiType'] = s['manual_ki_type']
        if s['manual_tenor']   is not None: fcn['tenor']  = s['manual_tenor']
        # 条款式输入无 IV → 风险等级按实际票息推导（票息是波动率的市场定价）
        if not s['iv_pct'] and fcn['coupon']:
            fcn['risk'] = '高' if fcn['coupon'] >= 30 else ('中' if fcn['coupon'] >= 15 else '低')

        # 条款式输入无评分体系 → score/scoreBreakdown 置 null，前端自动隐藏
        has_score = bool(s['display_score'])
        score_10 = round(s['display_score'] / score_max * 10, 1) if has_score else None
        score_bd = {
            "fundamental": min(round(s['display_score'] / score_max * 10, 1), 10),
            "volatility":  min(round(s['iv_pct'] / 120 * 10, 1), 10),
            "liquidity":   min(round(min(s['avg_vol'], 5000) / 5000 * 10, 1), 10),
            "momentum":    min(round(max(s['analyst_upside'] * 20, 0), 1), 10),
        } if has_score else None

        if rank == 0:     tag = "本周精选"
        elif s['iv_pct'] >= 70: tag = "高 IV"
        elif fcn['risk'] == "低": tag = "稳健"
        else:             tag = ""

        output_stocks.append({
            "code":           code,
            "name":           (s['name'] if any('一' <= _c <= '鿿' for _c in str(s['name'])) else (ana.get("name") or s['name'])),  # Excel名称为中文时优先保留(不被DeepSeek英文名覆盖)
            "name_en":        ana.get("name_en", ""),
            "industry":       s.get('industry_hint') or ana.get("sector", ""),
            "market":         s['market'],
            "types":          s.get('types', []),
            "price":          s['price'],
            "marketCap":      s['market_cap'],
            "currency":       s['currency'],
            "priceChange":    s['priceChange'],
            "sparkline":      s['sparkline'],
            "initialPrice":     s['price_history'].get('initialPrice'),
            "initialPriceDate": s['price_history'].get('initialPriceDate'),
            "priceHistory":     {"dates":  s['price_history'].get('dates', []),
                                 "closes": s['price_history'].get('closes', [])},
            "iv30":           s['iv_pct'],
            "score":          score_10,
            "scoreBreakdown": score_bd,
            "risk":           fcn['risk'],
            "tag":            tag,
            "coupon":         fcn['coupon'],
            "strike":         fcn['strike'],
            "ki":             fcn['ki'],
            "ko":             fcn['ko'],
            "kiType":         fcn['kiType'],
            "tenor":          fcn['tenor'],
            "bullets":        ana.get("bullets", []),
            "report":         _strip_md_bold({k: ana.get(k) for k in
                               ("tagline", "intro", "B", "C",
                                "health_conc", "c4_prose", "risks")}),
            "data_quality":   ana.get("data_quality", {}),
        })

    if any(s.get('types') for s in output_stocks):
        # 三类型模式：每个类型精选 2 只
        featured, _seen = [], set()
        for t in ("稳健", "进取", "热度"):
            picked = [s['code'] for s in output_stocks
                      if t in (s.get('types') or []) and s['code'] not in _seen][:2]
            _seen.update(picked); featured.extend(picked)
    else:
        featured = [s['code'] for s in output_stocks[:5]]

    # ── 每日市场热点：生成失败则保留上一版（板块永不留白）──
    print("\n[+] 生成每日市场热点（美股 + 港股 Google News → DeepSeek）...")
    hotspot = get_market_hotspot(client, today)
    if hotspot is None and os.path.exists(args.output):
        try:
            prev = json.load(open(args.output, encoding="utf-8"))
            hotspot = (prev.get("meta") or {}).get("marketHotspot")
            if hotspot:
                print("  ↪  沿用上一版每日市场热点")
        except Exception:
            pass

    # ── 市场热点图：生成失败则保留上一版（板块永不留白）──
    heatmap = get_sector_heatmap(today)
    if heatmap is None and os.path.exists(args.output):
        try:
            prev = json.load(open(args.output, encoding="utf-8"))
            heatmap = (prev.get("meta") or {}).get("sectorHeatmap")
            if heatmap:
                print("  ↪  沿用上一版市场热点图")
        except Exception:
            pass

    # ── 财经日历：脚本不抓取（investing.com 有反爬），保留上一版 ──
    # 数据由每周更新时用 WebFetch 抓 investing.com 写入 meta.calendar，此处仅负责不覆盖
    calendar = None
    if os.path.exists(args.output):
        try:
            _prevc = json.load(open(args.output, encoding="utf-8"))
            calendar = (_prevc.get("meta") or {}).get("calendar")
            if calendar:
                print("  ↪  保留现有财经日历（由每周 WebFetch investing.com 写入）")
        except Exception:
            pass

    # 对外正文一律剥离数据平台/供应商名（富途/Yahoo 等），只留底层出处
    output_stocks = [sanitize_platform(s) for s in output_stocks]

    watchlist = {
        "_generated": {
            "by":    "generate_watchlist.py",
            "model": MODEL,
            "at":    today.isoformat(),
            "count": len(output_stocks),
            "excel": " + ".join(excel_names),
        },
        "meta": {
            "week":        week_str,
            "theme":       "",
            "publishDate": today.strftime("%Y-%m-%d"),
            "nextUpdate":  (today + timedelta(days=7)).strftime("%Y-%m-%d"),
            "featuredIds": featured,
            "marketHotspot": hotspot,
            "sectorHeatmap": heatmap,
            "calendar": calendar,
        },
        "stocks": output_stocks
    }

    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(watchlist, f, ensure_ascii=False, indent=2)

    print(f"\n{'='*55}")
    print(f"✅  Done — {len(output_stocks)} stocks → {args.output}")
    print(f"   Week: {week_str}  |  Model: {MODEL}")
    print(f"   Top 3: {', '.join(s['code'] for s in output_stocks[:3])}")
    print(f"\n   Next step:")
    print(f"   git add watchlist.json && git commit -m \"{week_str} weekly update\" && git push")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()
