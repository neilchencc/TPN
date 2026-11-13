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
FLUID_OVER_LIMIT = 1.10   # 總液體量超過維持量 110% 提醒 (紅底)
FLUID_UNDER_LIMIT = 0.90  # 低於維持量 90% 提醒 (紅底);介於兩者之間為綠底

# 壓力參數: 選單只顯示數值,說明另外列在選單下方
STRESS_OPTIONS = {
    1.3: "營養良好、臥床、輕~中度壓力",
    1.5: "正常活動且輕~中度壓力 / 臥床但重度壓力 / 低活動且營養不良需追趕",
    1.7: "活動量大需追趕生長 / 活動量大且重度壓力",
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
# 標題列固定在畫面上方: 主畫面往下捲動時 logo 與標題不動
try:
    _dark = st.context.theme.type == "dark"  # Streamlit >= 1.46
except Exception:
    _dark = False
_header_bg = "#0e1117" if _dark else "#ffffff"
st.markdown(
    f"""
    <style>
    [data-testid="stElementContainer"]:has(.app-title),
    .element-container:has(.app-title) {{
        position: sticky;
        top: 0;
        z-index: 999;
        background-color: {_header_bg};
        padding-bottom: 0.25rem;
        border-bottom: 1px solid rgba(128, 128, 128, 0.25);
    }}
    </style>
    <h1 class='app-title' style='margin-bottom:0; display:flex; align-items:center; flex-wrap:wrap;'>{logo_html}靜脈營養計算器</h1>
    """,
    unsafe_allow_html=True,
)

# --- 病人資料 ---
st.sidebar.header("病人資料")
col_age, col_sex = st.sidebar.columns([3, 2])
age_years = col_age.number_input("年齡 (歲)", 0, 18, 10, 1, format="%d",
                                 help="未滿 1 歲請填 0,再輸入月齡")
sex = col_sex.radio("性別", ["M", "F"], index=0, horizontal=True)
age = float(age_years)
if age_years == 0:
    # 嬰兒公式依月齡區分 (≤3、≤6、≤12 個月),需另外輸入月齡
    age_months = st.sidebar.number_input("月齡 (月)", 0, 11, 6, 1, format="%d")
    age = age_months / 12
bw = float(st.sidebar.number_input("體重 (kg)", 1, 150, 35, 1, format="%d"))

stress_keys = list(STRESS_OPTIONS)
stress = st.sidebar.selectbox("壓力參數", stress_keys, index=2,
                              format_func=lambda v: f"{v:g}")
st.sidebar.caption("\n\n".join(f"**{k:g}**:{v}" for k, v in STRESS_OPTIONS.items()))
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

# --- 版面: 左側 PN 配方輸入、右側計算結果,同一畫面顯示 ---
left, right = st.columns([2, 3], gap="large")
with left:
    st.subheader("PN 配方")
    st.caption("輸入各濃度組合的每日容量 (mL),未使用請填 0。濃度可直接修改。")
    bags = st.data_editor(
        DEFAULT_BAGS,
        num_rows="fixed",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Dex (%)": st.column_config.NumberColumn(min_value=0.0, max_value=50.0, step=0.5),
            "AA (%)": st.column_config.NumberColumn(min_value=0.0, max_value=15.0, step=0.5),
            "Volume (mL)": st.column_config.NumberColumn(min_value=0, step=50, format="%d"),
        },
        key="bags",
    )

    st.markdown("**脂肪 Fat (SMOFlipid 20%)**")
    fat_vol = st.number_input("容量 (mL/day)", 0, 2000, 400, 50, format="%d",
                              help="0.2 g/mL、2 kcal/mL(約 10 kcal/g)")

# --- 計算 ---
r = calc_pn(bags, fat_vol, bw)
pct = r["kcal"] / target_kcal * 100 if target_kcal else 0
fluid_pct = r["volume"] / maint * 100 if maint else 0

# --- 超標 / 不足判斷: 直接顯示在各指標下方 (紅字) ---
infant = age < 1
max_sir = 14.0
# 建議下限 = Initial 起始量的下緣
min_sir, min_aa, min_fat = 3.0, 1.0, 0.5
max_aa = 4.0 if infant else 2.0
max_fat = 3.0 if infant else 2.0


def check(value: float, low: float, high: float) -> dict:
    """低於建議起始量或超過上限時,回傳 metric 的紅字 delta;範圍內則不顯示。"""
    if value > high:
        return {"delta": f"超過 {value - high:.2f}", "delta_color": "inverse"}
    if value <= 0:
        return {"delta": "-未給予", "delta_color": "normal"}
    if value < low:
        return {"delta": f"-{low - value:.2f} 低於建議 {low:g}", "delta_color": "normal"}
    return {}


kcal_ratio = r["kcal"] / target_kcal if target_kcal else 0
if kcal_ratio > KCAL_OVER_LIMIT:
    kcal_delta = {"delta": f"超過目標 {r['kcal'] - target_kcal:,.0f}", "delta_color": "inverse"}
elif 0 < kcal_ratio < KCAL_UNDER_LIMIT:
    kcal_delta = {"delta": f"-{target_kcal - r['kcal']:,.0f} 不足", "delta_color": "normal"}
else:
    kcal_delta = {"delta": f"{r['kcal'] - target_kcal:+,.0f} vs 目標", "delta_color": "off"}

fluid_ok = maint > 0 and FLUID_UNDER_LIMIT <= r["volume"] / maint <= FLUID_OVER_LIMIT

sir_delta = check(r["sir"], min_sir, max_sir)
aa_delta = check(r["aa_kg"], min_aa, max_aa)
fat_delta = check(r["fat_kg"], min_fat, max_fat)
all_ok = not (sir_delta or aa_delta or fat_delta
              or kcal_delta["delta_color"] != "off" or not fluid_ok)

with right:
    st.subheader("計算結果")
    m0a, m0b, m1 = st.columns(3)
    # 水分量: 自訂卡片,介於建議量 90–110% 為綠底,超出或不足為紅底
    _fg, _bg = ("#177233", "#dff3e4") if fluid_ok else ("#b42318", "#fde4e1")
    m0a.markdown(
        f"""
        <div title="PN + Fat;{r['volume'] / bw:,.0f} mL/kg/day">
          <div style="font-size:0.875rem; opacity:0.8;">水分量 (mL/d)</div>
          <div style="font-size:2.25rem; line-height:1.3;">{r['volume']:,.0f}</div>
          <span style="display:inline-block; padding:2px 10px; border-radius:999px;
                font-size:0.875rem; color:{_fg}; background:{_bg};">
            建議量的 {fluid_pct:.0f}%(建議 {maint:,.0f} mL)
          </span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    m0b.metric("熱量需求 (kcal/d)", f"{target_kcal:,.0f}",
               help=f"{elabel} × {stress:g} × {cal_factor:g}")
    m1.metric("PN 熱量 (kcal/d)", f"{r['kcal']:,.0f}", **kcal_delta)
    m2, m3, m4 = st.columns(3)
    m2.metric("SIR (mg/kg/min)", f"{r['sir']:.2f}", **sir_delta,
              help=f"建議 {min_sir:g} – {max_sir:g}")
    m3.metric("AA (g/kg/day)", f"{r['aa_kg']:.2f}", **aa_delta,
              help=f"建議 {min_aa:g} – {max_aa:g}")
    m4.metric("Fat (g/kg/day)", f"{r['fat_kg']:.2f}", **fat_delta,
              help=f"建議 {min_fat:g} – {max_fat:g}")

    st.progress(min(pct / 100, 1.0), text=f"達成目標熱量 {pct:.0f}%  "
                                          f"({r['kcal']:,.0f} / {target_kcal:,.0f} kcal)")
    if all_ok:
        st.success("熱量、液體量與 SIR / AA / Fat 皆在建議範圍內")

    # 目標範圍對照
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
    st.caption(f"熱量需求 = {elabel} {ree:,.0f} × 壓力參數 {stress:g} × 熱量係數 {cal_factor:g}")
    st.caption(f"總液體量 (PN + Fat):{r['volume']:,.0f} mL/day "
               f"({r['volume'] / bw:,.0f} mL/kg/day);"
               f"Holliday-Segar 維持量 {maint:,.0f} mL/day,占 {fluid_pct:.0f}%")
    if elabel == "EER":
        st.info("3 歲以下使用 DRI EER 公式,已包含生長與活動所需能量;"
                "再乘壓力參數與熱量係數可能重複加成,請依單位慣例確認。")

st.divider()

# --- 熱量明細 (原本在達成目標熱量下方) ---
st.subheader("熱量明細")
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

st.caption("※ 本工具僅供輔助計算,實際處方請由臨床醫師/營養師確認。")
st.caption("Neil JJC HMMCH 製作")
