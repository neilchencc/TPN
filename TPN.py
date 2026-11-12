"""
TPN / PN 計算器 (Streamlit)
複製自 TPN_20260828.xlsx 的計算邏輯:
  - 能量估算公式 (Sheet2): <3 歲為 DRI EER, 3 歲以上為 REE
  - 熱量需求 = 估算值 x 壓力參數 x 0.85
  - PN 熱量 = Dex(g) x 3.4 + AA(g) x 4 + SMOFlipid 20%(mL) x 2
  - SIR (mg/kg/min) = Dex(g) x 1000 / 1440 / BW
  - AA、Fat (g/kg/day)
執行: streamlit run streamlit_app.py
"""
import base64
from pathlib import Path

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------
# 核心計算 (與 Excel 公式一致)
# ----------------------------------------------------------------------
DEX_KCAL_PER_G = 3.4
AA_KCAL_PER_G = 4.0
# SMOFlipid 20%: 0.2 g/mL, 2 kcal/mL (含甘油、磷脂質; 即約 10 kcal/g)
FAT_G_PER_ML = 0.2
FAT_KCAL_PER_ML = 2.0

# 警示門檻
KCAL_OVER_LIMIT = 1.10   # PN 熱量超過目標 110% 提醒
KCAL_UNDER_LIMIT = 0.90  # 低於目標 90% 提醒
FLUID_OVER_LIMIT = 1.10  # 總液體量超過維持量 110% 提醒

STRESS_OPTIONS = {
    "1.3 — 營養良好、臥床、輕~中度壓力": 1.3,
    "1.5 — 正常活動且輕~中度壓力 / 臥床但重度壓力 / 低活動且營養不良需追趕": 1.5,
    "1.7 — 活動量大需追趕生長 / 活動量大且重度壓力": 1.7,
}


def calc_ree(age_y: float, sex: str, bw: float) -> float:
    """Sheet2 的公式 (年齡單位: 歲; 0.25 歲 = 3 個月, 0.5 歲 = 6 個月)。"""
    if age_y <= 0.25:
        return 89 * bw + 75
    if age_y <= 0.5:
        return 89 * bw - 44
    if age_y <= 1:
        return 89 * bw - 78
    if age_y < 3:
        return 89 * bw - 80
    if age_y < 10:
        return 22.7 * bw + 495 if sex == "M" else 22.5 * bw + 499
    return 17.5 * bw + 651 if sex == "M" else 12.2 * bw + 746


def energy_label(age_y: float) -> str:
    """<3 歲公式為 DRI EER(已含生長/活動), 3 歲以上為 REE。"""
    return "EER" if age_y < 3 else "REE"


def maintenance_fluid(bw: float) -> float:
    """Holliday-Segar 每日維持液體量 (mL/day)。"""
    if bw <= 10:
        return 100 * bw
    if bw <= 20:
        return 1000 + 50 * (bw - 10)
    return 1500 + 20 * (bw - 20)


def calc_pn(df: pd.DataFrame, fat_vol: float, bw: float) -> dict:
    """df 欄位: Dex (%), AA (%), Volume (mL)。fat_vol 為 SMOFlipid 20% 的 mL/day。"""
    d = df.fillna(0)
    dex_g = (d["Dex (%)"] * d["Volume (mL)"] / 100).sum()
    aa_g = (d["AA (%)"] * d["Volume (mL)"] / 100).sum()
    fat_g = FAT_G_PER_ML * fat_vol
    kcal_dex = dex_g * DEX_KCAL_PER_G
    kcal_aa = aa_g * AA_KCAL_PER_G
    kcal_fat = fat_vol * FAT_KCAL_PER_ML
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
_logo = Path(__file__).parent / "logo.png"
st.set_page_config(page_title="靜脈營養計算器",
                   page_icon=str(_logo) if _logo.exists() else "💉", layout="wide",
                   initial_sidebar_state="expanded")
# 隱藏側邊欄的收合按鈕 («),讓病人資料欄固定顯示
st.markdown(
    """
    <style>
    [data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarHeader"] button { display: none !important; }
    /* 縮小上方空白 */
    [data-testid="stSidebarHeader"] { display: none !important; }
    [data-testid="stSidebarUserContent"] { padding-top: 1rem !important; }
    .block-container,
    [data-testid="stMainBlockContainer"] { padding-top: 1.5rem !important; }
    header[data-testid="stHeader"] { background: transparent; }
    </style>
    """,
    unsafe_allow_html=True,
)
LOGO_PATH = Path(__file__).parent / "logo.png"
logo_html = ""
if LOGO_PATH.exists():
    logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode()
    logo_html = (f"<img src='data:image/png;base64,{logo_b64}' "
                 "style='height:64px; vertical-align:middle; margin-right:12px;'>")
st.markdown(
    f"<h1 style='margin-bottom:0; display:flex; align-items:center; flex-wrap:wrap;'>"
    f"{logo_html}靜脈營養計算器&nbsp;"
    "<span style='font-size:0.9rem; font-weight:400; color:gray;'>"
    #"Neil JJC HMMCH 製作</span></h1>",
    unsafe_allow_html=True,
)

# --- 病人資料 ---
st.sidebar.header("病人資料")
age = st.sidebar.number_input("年齡 (歲)", 0.0, 18.0, 10.0, 0.1,
                              help="嬰兒請用小數,例如 3 個月 = 0.25")
sex = st.sidebar.radio("性別", ["M", "F"], index=0, horizontal=True)
bw = st.sidebar.number_input("體重 (kg)", 1.0, 150.0, 34.5, 0.1)

stress_keys = list(STRESS_OPTIONS)
stress_label = st.sidebar.selectbox("壓力參數", stress_keys, index=2)
stress = STRESS_OPTIONS[stress_label]
custom = st.sidebar.checkbox("自訂壓力參數")
if custom:
    stress = st.sidebar.number_input("壓力參數", 1.0, 2.5, float(stress), 0.05)
cal_factor = st.sidebar.number_input(
    "熱量係數 (REE x ?)", 0.5, 1.5, 0.85, 0.05,
    help=(
        "PN 熱量需求通常比腸道營養 (EN) 低約 10–15%:\n\n"
        "- PN 直接進入血液,無消化吸收、無糞便流失\n"
        "- 食物熱效應較低\n"
        "- 可減少過度餵食造成的高血糖、肝功能異常\n\n"
        "建議 0.85–0.9 (Excel 中固定為 0.85)。"
        "若同時有部分腸道營養或單位慣例不打折,可改為 0.9–1.0。"
    ),
)
st.sidebar.caption("熱量係數:PN 約為 EN 需求的 85–90%,預設 0.85")

# --- 目標熱量 ---
ree = calc_ree(age, sex, bw)
elabel = energy_label(age)
target_kcal = ree * stress * cal_factor
maint = maintenance_fluid(bw)

c1, c2 = st.columns(2)
c1.metric(f"{elabel} (kcal/d)", f"{ree:,.0f}")
c2.metric(f"熱量需求 ({elabel} × {stress} × {cal_factor})", f"{target_kcal:,.0f} kcal/d")
if elabel == "EER":
    st.info("3 歲以下使用 DRI EER 公式,已包含生長與活動所需能量;"
            "再乘壓力參數與熱量係數可能重複加成,請依單位慣例確認。")

st.divider()

# --- PN 輸入 ---
left, right = st.columns([3, 2])
with left:
    st.subheader("PN 配方")
    st.caption("輸入各濃度組合的每日容量 (mL),未使用請填 0。可新增 / 修改濃度。")
    bags = st.data_editor(
        DEFAULT_BAGS,
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
    st.subheader("脂肪 Fat (SMOFlipid 20%)")
    st.caption("0.2 g/mL、2 kcal/mL(約 10 kcal/g)")
    fat_vol = st.number_input("容量 (mL/day)", 0.0, 2000.0, 400.0, 50.0)

# --- 計算 ---
r = calc_pn(bags, fat_vol, bw)

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
        "成分": ["Dextrose", "Amino acid", "Fat (SMOFlipid)", "合計"],
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
fluid_pct = r["volume"] / maint * 100 if maint else 0
st.caption(f"總液體量 (PN + Fat):{r['volume']:,.0f} mL/day "
           f"({r['volume'] / bw:,.0f} mL/kg/day);"
           f"Holliday-Segar 維持量 {maint:,.0f} mL/day,占 {fluid_pct:.0f}%")

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
if target_kcal > 0:
    ratio = r["kcal"] / target_kcal
    if ratio > KCAL_OVER_LIMIT:
        warns.append(f"PN 熱量 {r['kcal']:,.0f} kcal 為目標的 {ratio * 100:.0f}%,超過目標 "
                     f"{(KCAL_OVER_LIMIT - 1) * 100:.0f}% 以上")
    elif ratio < KCAL_UNDER_LIMIT and r["kcal"] > 0:
        warns.append(f"PN 熱量僅達目標的 {ratio * 100:.0f}%")
if maint > 0 and r["volume"] > maint * FLUID_OVER_LIMIT:
    warns.append(f"總液體量 {r['volume']:,.0f} mL 為維持量的 {fluid_pct:.0f}%,請留意液體過多")
if warns:
    for w in warns:
        st.warning("⚠️ " + w)
else:
    st.success("熱量、液體量與 SIR / AA / Fat 皆在建議範圍內")

st.caption("※ 本工具僅供輔助計算,實際處方請由臨床醫師/營養師確認。")
st.caption("Neil JJC HMMCH 製作")
