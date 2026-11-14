"""
TPN / PN 計算器 (Streamlit)
複製自 TPN_20260828.xlsx 的計算邏輯:
  - 能量估算公式 (Sheet2): <3 歲為 DRI EER, 3 歲以上為 REE
  - 熱量需求 = 估算值 x 壓力參數 x 0.85
  - PN 熱量 = Dex(g) x 3.4 + AA(g) x 4 + SMOFlipid 20%(mL) x 2
  - SIR (mg/kg/min) = Dex(g) x 1000 / 1440 / BW
  - AA、Fat (g/kg/day)
帳號 / 床位:
  - 左側欄可註冊、登入帳號 (密碼以 PBKDF2-SHA256 + salt 雜湊儲存)
  - 登入後可新增 / 更名 / 移除床位 (病患),並儲存該床位的輸入資料
  - 資料存於與本檔同資料夾的 tpn_data.db (SQLite)
執行: streamlit run streamlit_app.py
"""
import base64
import hashlib
import hmac
import json
import math
import secrets
import sqlite3
from contextlib import contextmanager
from datetime import datetime
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
# 帳號 / 床位資料庫 (SQLite)
# ----------------------------------------------------------------------
DB_PATH = Path(__file__).parent / "tpn_data.db"
PBKDF2_ITER = 200_000


@contextmanager
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    with db() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS users (
                        username TEXT PRIMARY KEY,
                        salt     TEXT NOT NULL,
                        pw_hash  TEXT NOT NULL,
                        created  TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS beds (
                        id       INTEGER PRIMARY KEY AUTOINCREMENT,
                        username TEXT NOT NULL
                                 REFERENCES users(username) ON DELETE CASCADE,
                        name     TEXT NOT NULL,
                        data     TEXT NOT NULL,
                        updated  TEXT,
                        UNIQUE (username, name))""")


def _hash_pw(pw: str, salt_hex: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt_hex),
                               PBKDF2_ITER).hex()


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def create_user(username: str, pw: str):
    """成功回傳 None,失敗回傳錯誤訊息。"""
    salt = secrets.token_hex(16)
    try:
        with db() as c:
            c.execute("INSERT INTO users VALUES (?, ?, ?, ?)",
                      (username, salt, _hash_pw(pw, salt), _now()))
    except sqlite3.IntegrityError:
        return "此帳號已存在"
    return None


def verify_user(username: str, pw: str) -> bool:
    with db() as c:
        row = c.execute("SELECT salt, pw_hash FROM users WHERE username = ?",
                        (username,)).fetchone()
    if not row:
        return False
    return hmac.compare_digest(_hash_pw(pw, row[0]), row[1])


def list_beds(username: str) -> list:
    with db() as c:
        return c.execute("SELECT id, name FROM beds WHERE username = ? "
                         "ORDER BY name COLLATE NOCASE", (username,)).fetchall()


def get_bed(username: str, bed_id: int):
    with db() as c:
        row = c.execute("SELECT data FROM beds WHERE id = ? AND username = ?",
                        (bed_id, username)).fetchone()
    return json.loads(row[0]) if row else None


def add_bed(username: str, name: str, data: dict) -> int:
    with db() as c:
        cur = c.execute("INSERT INTO beds (username, name, data, updated) "
                        "VALUES (?, ?, ?, ?)",
                        (username, name, json.dumps(data, ensure_ascii=False), _now()))
        return cur.lastrowid


def rename_bed(username: str, bed_id: int, name: str):
    with db() as c:
        c.execute("UPDATE beds SET name = ?, updated = ? WHERE id = ? AND username = ?",
                  (name, _now(), bed_id, username))


def delete_bed(username: str, bed_id: int):
    with db() as c:
        c.execute("DELETE FROM beds WHERE id = ? AND username = ?", (bed_id, username))


def save_bed(username: str, bed_id: int, data: dict):
    with db() as c:
        c.execute("UPDATE beds SET data = ?, updated = ? WHERE id = ? AND username = ?",
                  (json.dumps(data, ensure_ascii=False), _now(), bed_id, username))


# ----------------------------------------------------------------------
# 表單狀態 (載入 / 儲存床位用)
# ----------------------------------------------------------------------
STRESS_KEYS = list(STRESS_OPTIONS)
DEFAULT_FORM = {
    "age": 10.0,
    "sex": "M",
    "bw": 34.5,
    "stress_label": STRESS_KEYS[2],
    "custom": False,
    "stress_custom": 1.7,
    "cal_factor": 0.85,
    "fat_vol": 400.0,
}
BAG_COLS = list(DEFAULT_BAGS.columns)
ss = st.session_state


def editor_key() -> str:
    return f"bags_{ss.bags_ver}"


def apply_form(data: dict):
    """把床位資料 (或預設值) 填入各輸入欄位。須在 widget 建立前呼叫 (callback 內)。"""
    for k, v in DEFAULT_FORM.items():
        ss[k] = data.get(k, v)
    if ss.stress_label not in STRESS_OPTIONS:
        ss.stress_label = DEFAULT_FORM["stress_label"]
    if ss.sex not in ("M", "F"):
        ss.sex = "M"
    bags = data.get("bags")
    df = pd.DataFrame(bags, columns=BAG_COLS) if bags else DEFAULT_BAGS.copy()
    ss.bags_df = df.astype(float)
    ss.bags_ver = ss.get("bags_ver", 0) + 1  # 換新 key,讓 data_editor 重新載入


def current_bags() -> pd.DataFrame:
    """依 data_editor 的編輯紀錄還原目前表格 (可在 data_editor 繪製前取得)。"""
    df = ss.bags_df.copy().reset_index(drop=True)
    delta = ss.get(editor_key())
    if not delta:
        return df
    for idx, changes in delta.get("edited_rows", {}).items():
        for col, val in changes.items():
            if col in df.columns:
                df.loc[int(idx), col] = val
    deleted = [int(i) for i in delta.get("deleted_rows", [])]
    if deleted:
        df = df.drop(index=deleted, errors="ignore")
    added = delta.get("added_rows", [])
    if added:
        df = pd.concat([df, pd.DataFrame(added, columns=BAG_COLS)], ignore_index=True)
    return df.reset_index(drop=True)


def _num(v):
    if v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(f) else round(f, 6)


def form_snapshot(bags: pd.DataFrame) -> dict:
    data = {k: ss[k] for k in DEFAULT_FORM}
    for k in ("age", "bw", "stress_custom", "cal_factor", "fat_vol"):
        data[k] = _num(data[k])
    data["custom"] = bool(data["custom"])
    data["bags"] = [{c: _num(row.get(c)) for c in BAG_COLS}
                    for row in bags.to_dict("records")]
    return data


def flash(kind: str, msg: str):
    ss.flash = (kind, msg)


# ---- callbacks (在 widget 建立前執行,可安全修改欄位值) ----
def cb_login():
    u, pw = ss.login_user.strip(), ss.login_pw
    if verify_user(u, pw):
        ss.user = u
        ss.bed_sel = None
        ss.saved_snapshot = None
        flash("success", f"歡迎,{u}")
    else:
        flash("error", "帳號或密碼錯誤")
    ss.login_pw = ""


def cb_register():
    u, pw, pw2 = ss.reg_user.strip(), ss.reg_pw, ss.reg_pw2
    if len(u) < 3:
        flash("error", "帳號至少 3 個字元")
    elif len(pw) < 6:
        flash("error", "密碼至少 6 個字元")
    elif pw != pw2:
        flash("error", "兩次密碼不一致")
    else:
        err = create_user(u, pw)
        if err:
            flash("error", err)
        else:
            ss.user = u
            ss.bed_sel = None
            ss.saved_snapshot = None
            ss.reg_user = ""
            flash("success", f"帳號 {u} 已建立並登入")
    ss.reg_pw = ss.reg_pw2 = ""


def cb_logout():
    ss.user = None
    ss.bed_sel = None
    ss.saved_snapshot = None
    apply_form({})
    flash("info", "已登出")


def cb_select_bed():
    bid = ss.bed_sel
    if bid is None:
        ss.saved_snapshot = None
        return
    data = get_bed(ss.user, bid)
    if data is None:
        ss.bed_sel = None
        flash("error", "找不到該床位")
        return
    apply_form(data)
    ss.saved_snapshot = form_snapshot(ss.bags_df)
    ss.rename_name = dict(list_beds(ss.user)).get(bid, "")


def cb_add_bed():
    name = ss.new_bed_name.strip()
    if not name:
        flash("error", "請輸入床位名稱")
        return
    if not ss.get("copy_current"):
        apply_form({})
    data = form_snapshot(current_bags())
    try:
        bid = add_bed(ss.user, name, data)
    except sqlite3.IntegrityError:
        flash("error", f"床位「{name}」已存在")
        return
    if ss.get("copy_current"):
        # 以目前畫面建立:把編輯中的表格固定為新床位的內容
        apply_form(data)
    ss.bed_sel = bid
    ss.saved_snapshot = data
    ss.rename_name = name
    ss.new_bed_name = ""
    flash("success", f"已新增床位「{name}」")


def cb_rename_bed():
    name = ss.rename_name.strip()
    if not name or ss.bed_sel is None:
        flash("error", "請輸入新名稱")
        return
    try:
        rename_bed(ss.user, ss.bed_sel, name)
    except sqlite3.IntegrityError:
        flash("error", f"床位「{name}」已存在")
        return
    flash("success", f"已更名為「{name}」")


def cb_delete_bed():
    if ss.bed_sel is None:
        return
    name = dict(list_beds(ss.user)).get(ss.bed_sel, "")
    delete_bed(ss.user, ss.bed_sel)
    ss.bed_sel = None
    ss.saved_snapshot = None
    ss.del_confirm = False
    apply_form({})
    flash("success", f"已移除床位「{name}」")


def cb_save():
    if ss.bed_sel is None:
        return
    data = form_snapshot(current_bags())
    save_bed(ss.user, ss.bed_sel, data)
    ss.saved_snapshot = data
    flash("success", f"已儲存 ({datetime.now():%H:%M:%S})")


def cb_custom_toggle():
    if ss.custom:
        ss.stress_custom = STRESS_OPTIONS[ss.stress_label]


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
    f"{logo_html}靜脈營養計算器</h1>",
    unsafe_allow_html=True,
)

# --- 狀態初始化 ---
init_db()
if "user" not in ss:
    ss.user = None
    ss.bed_sel = None
    ss.saved_snapshot = None
    apply_form({})
# 未顯示的 widget (例如未勾選自訂壓力參數時) 其值會被 Streamlit 清掉,
# 每次執行前重新指定一次以保留。
for _k, _v in DEFAULT_FORM.items():
    ss[_k] = ss.get(_k, _v)

# --- 帳號 / 床位 ---
with st.sidebar:
    if "flash" in ss:
        kind, msg = ss.pop("flash")
        getattr(st, kind)(msg)

    if ss.user is None:
        st.header("帳號")
        tab_in, tab_reg = st.tabs(["登入", "註冊"])
        with tab_in:
            with st.form("login_form"):
                st.text_input("帳號", key="login_user")
                st.text_input("密碼", type="password", key="login_pw")
                st.form_submit_button("登入", on_click=cb_login,
                                      use_container_width=True, type="primary")
        with tab_reg:
            with st.form("reg_form"):
                st.text_input("帳號 (至少 3 字元)", key="reg_user")
                st.text_input("密碼 (至少 6 字元)", type="password", key="reg_pw")
                st.text_input("確認密碼", type="password", key="reg_pw2")
                st.form_submit_button("建立帳號", on_click=cb_register,
                                      use_container_width=True)
        st.caption("未登入仍可使用計算器,但無法儲存病患資料。")
    else:
        cu, co = st.columns([3, 2])
        cu.markdown(f"👤 **{ss.user}**")
        co.button("登出", on_click=cb_logout, use_container_width=True)

        st.header("床位 / 病患")
        beds = list_beds(ss.user)
        bed_names = dict(beds)
        if ss.bed_sel not in bed_names:
            ss.bed_sel = None
        st.selectbox("選擇床位", list(bed_names), index=None,
                     format_func=lambda i: bed_names.get(i, "?"),
                     placeholder="尚未選擇 (新增或選擇床位)" if beds else "尚無床位,請先新增",
                     key="bed_sel", on_change=cb_select_bed)

        if ss.bed_sel is not None:
            st.button("💾 儲存目前資料", on_click=cb_save,
                      use_container_width=True, type="primary")
            if ss.saved_snapshot is not None and \
                    form_snapshot(current_bags()) != ss.saved_snapshot:
                st.warning("有未儲存的變更")
            else:
                st.caption("✅ 已與儲存資料一致")

        with st.expander("➕ 新增床位", expanded=not beds):
            st.text_input("床位名稱", key="new_bed_name",
                          placeholder="例如:PICU-3 王小明")
            st.checkbox("以目前畫面資料建立", key="copy_current",
                        help="不勾選則以預設值建立空白床位")
            st.button("新增", on_click=cb_add_bed, use_container_width=True)

        if ss.bed_sel is not None:
            with st.expander("✏️ 更名 / 🗑️ 移除"):
                st.text_input("新名稱", key="rename_name")
                st.button("更名", on_click=cb_rename_bed, use_container_width=True)
                st.divider()
                st.checkbox(f"確認移除「{bed_names[ss.bed_sel]}」", key="del_confirm")
                st.button("移除床位", on_click=cb_delete_bed, use_container_width=True,
                          disabled=not ss.get("del_confirm"))

    st.divider()

# --- 病人資料 ---
st.sidebar.header("病人資料")
age = st.sidebar.number_input("年齡 (歲)", min_value=0.0, max_value=18.0, step=0.1,
                              key="age", help="嬰兒請用小數,例如 3 個月 = 0.25")
sex = st.sidebar.radio("性別", ["M", "F"], horizontal=True, key="sex")
bw = st.sidebar.number_input("體重 (kg)", min_value=1.0, max_value=150.0, step=0.1,
                             key="bw")

stress_label = st.sidebar.selectbox("壓力參數", STRESS_KEYS, key="stress_label")
stress = STRESS_OPTIONS[stress_label]
custom = st.sidebar.checkbox("自訂壓力參數", key="custom", on_change=cb_custom_toggle)
if custom:
    stress = st.sidebar.number_input("壓力參數", min_value=1.0, max_value=2.5,
                                     step=0.05, key="stress_custom")
cal_factor = st.sidebar.number_input(
    "熱量係數 (REE x ?)", min_value=0.5, max_value=1.5, step=0.05, key="cal_factor",
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
        ss.bags_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Dex (%)": st.column_config.NumberColumn(min_value=0.0, max_value=50.0, step=0.5),
            "AA (%)": st.column_config.NumberColumn(min_value=0.0, max_value=15.0, step=0.5),
            "Volume (mL)": st.column_config.NumberColumn(min_value=0.0, step=50.0),
        },
        key=editor_key(),
    )

with right:
    st.subheader("脂肪 Fat (SMOFlipid 20%)")
    st.caption("0.2 g/mL、2 kcal/mL(約 10 kcal/g)")
    fat_vol = st.number_input("容量 (mL/day)", min_value=0.0, max_value=2000.0,
                              step=50.0, key="fat_vol")

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
