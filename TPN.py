"""
TPN / PN 計算器 (Streamlit)
複製自 TPN_20260828.xlsx 的計算邏輯:
  - REE: Schofield-type 公式 (Sheet2)
  - 熱量需求 = REE x 壓力參數 x 0.85
  - PN 熱量 = Dex(g) x 3.4 + AA(g) x 4 + Fat(g) x 9
  - SIR (mg/kg/min) = Dex(g) x 1000 / 1440 / BW
  - AA、Fat (g/kg/day)
執行: streamlit run streamlit_app.py
"""
import io

import openpyxl
import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------
# 核心計算 (與 Excel 公式一致)
# ----------------------------------------------------------------------
DEX_KCAL_PER_G = 3.4
AA_KCAL_PER_G = 4.0
FAT_KCAL_PER_G = 9.0

STRESS_OPTIONS = {
    "1.3 — 營養良好、臥床、輕~中度壓力": 1.3,
    "1.5 — 正常活動且輕~中度壓力 / 臥床但重度壓力 / 低活動且營養不良需追趕": 1.5,
    "1.7 — 活動量大需追趕生長 / 活動量大且重度壓力": 1.7,
}


def calc_ree(age_y: float, sex: str, bw: float) -> float:
    """Sheet2 的 REE 公式 (年齡單位: 歲; 0.3 歲≈3 個月, 0.6 歲≈6 個月)。"""
    if age_y <= 0.3:
        return 89 * bw + 75
    if age_y <= 0.6:
        return 89 * bw - 44
    if age_y <= 1:
        return 89 * bw - 78
    if age_y < 3:
        return 89 * bw - 80
    if age_y < 10:
        return 22.7 * bw + 495 if sex == "M" else 22.5 * bw + 499
    return 17.5 * bw + 651 if sex == "M" else 12.2 * bw + 746


def calc_pn(df: pd.DataFrame, fat_conc: float, fat_vol: float, bw: float) -> dict:
    """df 欄位: Dex (%), AA (%), Volume (mL)。fat_conc 單位 g/mL (SMOF 20% = 0.2)。"""
    d = df.fillna(0)
    dex_g = (d["Dex (%)"] * d["Volume (mL)"] / 100).sum()
    aa_g = (d["AA (%)"] * d["Volume (mL)"] / 100).sum()
    fat_g = fat_conc * fat_vol
    kcal_dex = dex_g * DEX_KCAL_PER_G
    kcal_aa = aa_g * AA_KCAL_PER_G
    kcal_fat = fat_g * FAT_KCAL_PER_G
    return {
        "dex_g": dex_g,
        "aa_g": aa_g,
        "fat_g": fat_g,
        "kcal_dex": kcal_dex,
        "kcal_aa": kcal_aa,
        "kcal_fat": kcal_fat,
        "kcal": kcal_dex + kcal_aa + kcal_fat,
        "sir": dex_g * 1000 / 1440 / bw,  # mg/kg/min
        "aa_kg": aa_g / bw,  # g/kg/day
        "fat_kg": fat_g / bw,  # g/kg/day
        "volume": d["Volume (mL)"].sum() + fat_vol,
    }


# ----------------------------------------------------------------------
# 讀取 xlsx 預設值 (版面同 TPN_20260828.xlsx 的 Sheet1)
# ----------------------------------------------------------------------
def load_from_xlsx(file) -> dict | None:
    try:
        ws = openpyxl.load_workbook(file, data_only=True)["Sheet1"]
        rows = []
        for r in range(11, 17):
            rows.append(
                {
                    "Dex (%)": float(ws.cell(r, 1).value or 0),
                    "AA (%)": float(ws.cell(r, 2).value or 0),
                    "Volume (mL)": float(ws.cell(r, 3).value or 0),
                }
            )
        return {
            "age": float(ws["B1"].value),
            "sex": str(ws["B2"].value).strip().upper(),
            "bw": float(ws["B3"].value),
            "stress": float(ws["B4"].value),
            "fat_conc": float(ws["B18"].value or 0),
            "fat_vol": float(ws["C18"].value or 0),
            "bags": pd.DataFrame(rows),
        }
    except Exception as e:  # noqa: BLE001
        st.sidebar.error(f"無法讀取 xlsx:{e}")
        return None


DEFAULT_BAGS = pd.DataFrame(
    {
        "Dex (%)": [5, 7.5, 10, 12.5, 15, 20],
        "AA (%)": [1.5, 2.5, 2, 2, 3, 3],
        "Volume (mL)": [0, 0, 0, 0, 0, 0],
    }
).astype(float)

# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
st.set_page_config(page_title="PN 計算器", page_icon="💉", layout="wide")
st.title("💉 PN (全靜脈營養) 計算器")
st.caption("依 TPN_20260828.xlsx 的公式計算熱量、SIR、AA 與 Fat。")

# 上傳 xlsx 以預填
up = st.sidebar.file_uploader("上傳 TPN xlsx 預填資料 (選填)", type=["xlsx"])
loaded = load_from_xlsx(up) if up else None
if loaded:
    st.sidebar.success("已由 xlsx 載入預設值")


def default(key, fallback):
    return loaded[key] if loaded else fallback


# --- 病人資料 ---
st.sidebar.header("病人資料")
age = st.sidebar.number_input("年齡 (歲)", 0.0, 18.0, float(default("age", 10.0)), 0.1,
                              help="嬰兒請用小數,例如 3 個月 = 0.25")
sex = st.sidebar.radio("性別", ["M", "F"], index=0 if default("sex", "M") == "M" else 1,
                       horizontal=True)
bw = st.sidebar.number_input("體重 (kg)", 1.0, 150.0, float(default("bw", 34.5)), 0.1)

stress_default = default("stress", 1.7)
stress_keys = list(STRESS_OPTIONS)
idx = min(range(3), key=lambda i: abs(list(STRESS_OPTIONS.values())[i] - stress_default))
stress_label = st.sidebar.selectbox("壓力參數", stress_keys, index=idx)
stress = STRESS_OPTIONS[stress_label]
custom = st.sidebar.checkbox("自訂壓力參數")
if custom:
    stress = st.sidebar.number_input("壓力參數", 1.0, 2.5, float(stress), 0.05)
cal_factor = st.sidebar.number_input("熱量係數 (REE x ?)", 0.5, 1.5, 0.85, 0.05,
                                     help="Excel 中固定為 0.85")

# --- 目標熱量 ---
ree = calc_ree(age, sex, bw)
target_kcal = ree * stress * cal_factor

c1, c2, c3 = st.columns(3)
c1.metric("REE (kcal/d)", f"{ree:,.0f}")
c2.metric(f"熱量需求 (REE × {stress} × {cal_factor})", f"{target_kcal:,.0f} kcal/d")
c3.metric("每公斤目標熱量", f"{target_kcal / bw:,.1f} kcal/kg/d")

st.divider()

# --- PN 輸入 ---
left, right = st.columns([3, 2])
with left:
    st.subheader("PN 配方")
    st.caption("輸入各濃度組合的每日容量 (mL),未使用請填 0。可新增 / 修改濃度。")
    bags_init = loaded["bags"] if loaded else DEFAULT_BAGS
    bags = st.data_editor(
        bags_init,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Dex (%)": st.column_config.NumberColumn(min_value=0.0, max_value=50.0, step=0.5),
            "AA (%)": st.column_config.NumberColumn(min_value=0.0, max_value=15.0, step=0.5),
            "Volume (mL)": st.column_config.NumberColumn(min_value=0.0, step=50.0),
        },
        key="bags",
    )

with right:
    st.subheader("脂肪 Fat (SMOF)")
    fat_conc = st.number_input("濃度 (g/mL,20% = 0.2)", 0.0, 1.0,
                               float(default("fat_conc", 0.2)), 0.05)
    fat_vol = st.number_input("容量 (mL/day)", 0.0, 2000.0,
                              float(default("fat_vol", 400.0)), 50.0)

# --- 計算 ---
r = calc_pn(bags, fat_conc, fat_vol, bw)

st.divider()
st.subheader("計算結果")

m1, m2, m3, m4 = st.columns(4)
m1.metric("PN 熱量 (Kcal/day)", f"{r['kcal']:,.0f}",
          f"{r['kcal'] - target_kcal:+,.0f} vs 目標")
m2.metric("SIR (mg/kg/min)", f"{r['sir']:.2f}")
m3.metric("AA (g/kg/day)", f"{r['aa_kg']:.2f}")
m4.metric("Fat (g/kg/day)", f"{r['fat_kg']:.2f}")

pct = r["kcal"] / target_kcal * 100 if target_kcal else 0
st.progress(min(pct / 100, 1.0), text=f"達成目標熱量 {pct:.0f}%  "
                                      f"({r['kcal']:,.0f} / {target_kcal:,.0f} kcal)")

# 明細
detail = pd.DataFrame(
    {
        "成分": ["Dextrose", "Amino acid", "Fat", "合計"],
        "克數 (g/day)": [r["dex_g"], r["aa_g"], r["fat_g"], None],
        "熱量 (kcal/day)": [r["kcal_dex"], r["kcal_aa"], r["kcal_fat"], r["kcal"]],
        "熱量占比 (%)": [
            100 * r["kcal_dex"] / r["kcal"] if r["kcal"] else 0,
            100 * r["kcal_aa"] / r["kcal"] if r["kcal"] else 0,
            100 * r["kcal_fat"] / r["kcal"] if r["kcal"] else 0,
            100 if r["kcal"] else 0,
        ],
    }
)
st.dataframe(
    detail.style.format({"克數 (g/day)": "{:.1f}", "熱量 (kcal/day)": "{:,.0f}",
                         "熱量占比 (%)": "{:.0f}"}, na_rep="—"),
    hide_index=True,
    use_container_width=True,
)
st.caption(f"總液體量 (PN + Fat):{r['volume']:,.0f} mL/day "
           f"({r['volume'] / bw:,.0f} mL/kg/day)")

# --- 目標對照與提示 ---
st.subheader("目標範圍對照")
infant = age < 1
ref = pd.DataFrame(
    {
        "項目": ["SIR (mg/kg/min)", "AA (g/kg/day)", "Fat (g/kg/day)"],
        "目前": [round(r["sir"], 2), round(r["aa_kg"], 2), round(r["fat_kg"], 2)],
        "Initial": ["3 - 4", "1 - 2", "0.5 - 1"],
        "Increment": ["2 - 5", "1", "0.5"],
        "Maximum": ["12 - 14", "3 - 4 (infant)" if infant else "2 (child)",
                    "3 (infant)" if infant else "1 - 2 (child)"],
    }
)
st.dataframe(ref, hide_index=True, use_container_width=True)

max_aa = 4.0 if infant else 2.0
max_fat = 3.0 if infant else 2.0
warns = []
if r["sir"] > 14:
    warns.append(f"SIR {r['sir']:.1f} mg/kg/min 超過上限 12–14")
if r["aa_kg"] > max_aa:
    warns.append(f"AA {r['aa_kg']:.2f} g/kg/day 超過上限 {max_aa:g}")
if r["fat_kg"] > max_fat:
    warns.append(f"Fat {r['fat_kg']:.2f} g/kg/day 超過上限 {max_fat:g}")
if warns:
    st.warning("⚠️ " + ";".join(warns))
else:
    st.success("SIR / AA / Fat 皆在上限範圍內")

# --- 匯出 ---
summary = pd.DataFrame(
    {
        "項目": ["年齡(歲)", "性別", "體重(kg)", "壓力參數", "REE", "目標熱量(kcal/d)",
                 "PN熱量(kcal/d)", "SIR(mg/kg/min)", "AA(g/kg/day)", "Fat(g/kg/day)"],
        "數值": [age, sex, bw, stress, round(ree, 1), round(target_kcal, 1),
                 round(r["kcal"], 1), round(r["sir"], 2), round(r["aa_kg"], 2),
                 round(r["fat_kg"], 2)],
    }
)
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as xw:
    summary.to_excel(xw, sheet_name="Summary", index=False)
    bags.to_excel(xw, sheet_name="PN", index=False)
    detail.to_excel(xw, sheet_name="Detail", index=False)
st.download_button("⬇️ 下載計算結果 (xlsx)", buf.getvalue(),
                   file_name="PN_result.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("※ 本工具僅供輔助計算,實際處方請由臨床醫師/營養師確認。")
