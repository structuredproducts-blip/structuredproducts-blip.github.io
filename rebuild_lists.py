# -*- coding: utf-8 -*-
"""
rebuild_lists.py — 每周把最新 FCN筛选结果_*.xlsx 重建成三个榜单 Excel(v2 格式)+ lists_input.xlsx。

v2 表头(9 列, 与交易台交接模板一致):
  区域 | 代码 | 名称 | 行业 | 执行价 | 敲入价 | 敲入类型 | 敲出价 | 票息(百分比 0.00%)
  - 执行价/敲入价: 留空, 由 strike_advisor + fill_terms 回填(小数)
  - 敲入类型: 欧式敲入(常量)   敲出价: 1.01(=101%, 常量)   票息: 留空待询价回填(格式已预置 0.00%)

源 sheet 顺序: 0 评分卡 / 1 稳健(→低波) / 2 进取(→高波) / 3 热度(→市场热度)。
旧档备份到 _filtered/。用法: python rebuild_lists.py
"""
import sys, os, glob, shutil, datetime as dt
sys.stdout.reconfigure(encoding="utf-8")
import pandas as pd, openpyxl

D = os.path.dirname(os.path.abspath(__file__))
BAK = os.path.join(D, "_filtered"); os.makedirs(BAK, exist_ok=True)
ts = dt.datetime.now().strftime("%Y%m%d_%H%M")

srcs = sorted(glob.glob(os.path.join(D, "FCN筛选结果_*.xlsx")))
if not srcs:
    raise SystemExit("未找到 FCN筛选结果_*.xlsx(筛选器应已交付到本目录)")
SRC = srcs[-1]
print(f"源文件: {os.path.basename(SRC)}")

HDR = ["区域", "代码", "名称", "行业", "执行价", "敲入价", "敲入类型", "敲出价", "票息"]
SHEET_BY_IDX = {1: "低波精选组.xlsx", 2: "高波精选组.xlsx", 3: "市场热度榜.xlsx"}

xl = pd.ExcelFile(SRC)
def col(df, name):
    for c in df.columns:
        if str(c).strip() == name:
            return c
    return None

for idx, outfn in SHEET_BY_IDX.items():
    df = xl.parse(xl.sheet_names[idx])
    c_code, c_name, c_mkt, c_sub = (col(df, "代码"), col(df, "名称"), col(df, "市场"), col(df, "子板块"))
    dst = os.path.join(D, outfn)
    if os.path.exists(dst):
        shutil.copy2(dst, os.path.join(BAK, outfn.replace(".xlsx", f"_backup_{ts}.xlsx")))
    wb = openpyxl.Workbook(); ws = wb.active; ws.append(HDR)
    n = 0
    for _, r in df.iterrows():
        code = r[c_code]
        if pd.isna(code) or str(code).strip() == "":
            continue
        mkt = str(r[c_mkt]).strip().upper()
        raw = str(code).strip()
        disp = str(int(raw)) if raw.isdigit() else raw.upper()
        name = "" if pd.isna(r[c_name]) else str(r[c_name])
        industry = "" if (c_sub is None or pd.isna(r[c_sub])) else str(r[c_sub])
        ws.append(["港股" if mkt == "HK" else "美股", f"{disp} {'HK' if mkt=='HK' else 'US'}",
                   name, industry, None, None, "欧式敲入", 1.01, None])
        n += 1
        ws.cell(ws.max_row, 9).number_format = "0.00%"   # 票息列预置百分比格式
    wb.save(dst)
    print(f"重建 {outfn}: {n} 行 (v2 格式)")

# lists_input.csv → lists_input.xlsx(供 fetch_iv / strike_advisor)
lix = os.path.join(D, "lists_input.xlsx")
if os.path.exists(lix):
    shutil.copy2(lix, os.path.join(BAK, f"lists_input_backup_{ts}.xlsx"))
lic = os.path.join(D, "lists_input.csv")
if os.path.exists(lic):
    pd.read_csv(lic).to_excel(lix, index=False)
    print(f"生成 lists_input.xlsx")
print("完成 — 下一步: python fetch_iv.py → python strike_advisor_pipeline.py → python fill_terms.py")
