"""
TPN / PN 計算器 (Streamlit)
複製自 TPN_20260828.xlsx 的計算邏輯:
  - 能量估算公式 (Sheet2): <3 歲為 DRI EER, 3 歲以上為 REE
  - 熱量需求 = 估算值 x 壓力參數 x 0.85
  - PN 熱量 = Dex(g) x 3.4 + AA(g) x 4 + SMOFlipid 20%(mL) x 2
  - SIR (mg/kg/min) = Dex(g) x 1000 / 1440 / BW
  - AA、Fat (g/kg/day)
新增:
  - 帳號註冊 / 登入 (帳號、密碼皆需 4 個字以上,密碼以 PBKDF2 雜湊儲存)
  - 每個帳號可建立 / 刪除病患 (自訂名稱)
  - 儲存目前輸入的資料,於歷史紀錄中查看、載入、刪除
  - 資料存在同資料夾的 SQLite 檔案 tpn_data.db
執行: streamlit run TPN.py
"""
import base64
import hashlib
import hmac
import json
import secrets
import sqlite3
import tempfile
from contextlib import closing
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("Asia/Taipei")
except Exception:  # 系統沒有時區資料時改用本機時間
    TZ = None

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


BAG_COLS = ["Dex (%)", "AA (%)", "Volume (mL)"]
DEFAULT_BAGS = pd.DataFrame(
    {
        "Dex (%)": [5, 7.5, 10, 12.5, 15, 20],
        "AA (%)": [1.5, 2.5, 2, 2, 3, 3],
        "Volume (mL)": [0, 0, 0, 0, 0, 0],
    }
).astype(float)

# ----------------------------------------------------------------------
# 資料庫 (SQLite): 帳號、病患、紀錄
# ----------------------------------------------------------------------
def _pick_db_path() -> Path:
    """找一個可以寫入的位置放資料庫。
    Streamlit Cloud 上程式所在資料夾可能是唯讀的,寫不進去就改放家目錄或暫存資料夾。"""
    candidates = [Path(__file__).parent, Path.home(), Path(tempfile.gettempdir())]
    for folder in candidates:
        path = folder / "tpn_data.db"
        try:
            with closing(sqlite3.connect(path)) as c, c:
                c.execute("CREATE TABLE IF NOT EXISTS _write_test (x INTEGER)")
                c.execute("INSERT INTO _write_test VALUES (1)")
                c.execute("DROP TABLE _write_test")
            return path
        except (sqlite3.Error, OSError):
            continue
    raise RuntimeError("找不到可寫入的位置來存放資料庫")


DB_PATH = _pick_db_path()
MIN_LEN = 4  # 帳號、密碼最少字數
PBKDF2_ROUNDS = 200_000


def db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    with closing(db()) as c, c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                username   TEXT NOT NULL UNIQUE COLLATE NOCASE,
                pw_hash    TEXT NOT NULL,
                salt       TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS patients (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id    INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                name       TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE (user_id, name)
            );
            CREATE TABLE IF NOT EXISTS records (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                patient_id INTEGER NOT NULL REFERENCES patients(id) ON DELETE CASCADE,
                created_at TEXT NOT NULL,
                note       TEXT,
                data       TEXT NOT NULL
            );
            """
        )


def now_str() -> str:
    return datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")


def _hash_pw(password: str, salt_hex: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(salt_hex),
                               PBKDF2_ROUNDS).hex()


def create_user(username: str, password: str):
    """成功回傳 (user_id, None);失敗回傳 (None, 錯誤訊息)。"""
    username = username.strip()
    if len(username) < MIN_LEN:
        return None, f"帳號至少需要 {MIN_LEN} 個字"
    if len(password) < MIN_LEN:
        return None, f"密碼至少需要 {MIN_LEN} 個字"
    salt = secrets.token_hex(16)
    try:
        with closing(db()) as c, c:
            cur = c.execute(
                "INSERT INTO users (username, pw_hash, salt, created_at) VALUES (?, ?, ?, ?)",
                (username, _hash_pw(password, salt), salt, now_str()),
            )
            return cur.lastrowid, None
    except sqlite3.IntegrityError:
        return None, "此帳號已被使用"


def verify_user(username: str, password: str):
    """帳密正確回傳 {'id', 'username'},否則 None。"""
    with closing(db()) as c:
        row = c.execute("SELECT * FROM users WHERE username = ?", (username.strip(),)).fetchone()
    if row and hmac.compare_digest(row["pw_hash"], _hash_pw(password, row["salt"])):
        return {"id": row["id"], "username": row["username"]}
    return None


def list_patients(user_id: int) -> list:
    with closing(db()) as c:
        return c.execute("SELECT id, name FROM patients WHERE user_id = ? ORDER BY created_at, id",
                         (user_id,)).fetchall()


def create_patient(user_id: int, name: str):
    try:
        with closing(db()) as c, c:
            cur = c.execute("INSERT INTO patients (user_id, name, created_at) VALUES (?, ?, ?)",
                            (user_id, name, now_str()))
            return cur.lastrowid, None
    except sqlite3.IntegrityError:
        return None, f"已經有名為「{name}」的病患"


def delete_patient(user_id: int, patient_id: int) -> None:
    # 紀錄會透過 ON DELETE CASCADE 一併刪除
    with closing(db()) as c, c:
        c.execute("DELETE FROM patients WHERE id = ? AND user_id = ?", (patient_id, user_id))


def add_record(user_id: int, patient_id: int, note: str, data: dict) -> bool:
    with closing(db()) as c, c:
        owner = c.execute("SELECT 1 FROM patients WHERE id = ? AND user_id = ?",
                          (patient_id, user_id)).fetchone()
        if not owner:
            return False
        c.execute("INSERT INTO records (patient_id, created_at, note, data) VALUES (?, ?, ?, ?)",
                  (patient_id, now_str(), note, json.dumps(data, ensure_ascii=False)))
        return True


def list_records(user_id: int, patient_id: int) -> list:
    with closing(db()) as c:
        return c.execute(
            """SELECT r.id, r.created_at, r.note, r.data FROM records r
               JOIN patients p ON p.id = r.patient_id
               WHERE r.patient_id = ? AND p.user_id = ?
               ORDER BY r.created_at DESC, r.id DESC""",
            (patient_id, user_id),
        ).fetchall()


def delete_record(user_id: int, record_id: int) -> None:
    with closing(db()) as c, c:
        c.execute(
            """DELETE FROM records WHERE id = ? AND patient_id IN
               (SELECT id FROM patients WHERE user_id = ?)""",
            (record_id, user_id),
        )


init_db()

# ----------------------------------------------------------------------
# 輸入欄位的狀態 (用 session_state 才能從歷史紀錄載入)
# ----------------------------------------------------------------------
DEFAULT_INPUTS = {
    "age_years": 10,
    "age_months": 6,
    "sex": "M",
    "bw": 35,
    "stress_sel": 1.7,
    "custom": False,
    "stress_custom": 1.7,
    "cal_factor": 0.85,
    "fat_vol": 400,
}
INT_KEYS = {"age_years", "age_months", "bw", "fat_vol"}
FLOAT_KEYS = {"stress_sel", "stress_custom", "cal_factor"}

ss = st.session_state
for _k, _v in DEFAULT_INPUTS.items():
    ss.setdefault(_k, _v)
ss.setdefault("bags_df", DEFAULT_BAGS.copy())
ss.setdefault("bags_ver", 0)  # 改變 data_editor 的 key 以套用載入的配方
ss.setdefault("flash", [])


def flash(kind: str, msg: str) -> None:
    """在 callback 裡記下訊息,下次畫面更新時以 toast 顯示。"""
    ss.flash.append((kind, msg))


def apply_inputs(inputs: dict) -> None:
    """把一組輸入值套用到所有欄位 (只能在 callback 中呼叫)。"""
    for k, default in DEFAULT_INPUTS.items():
        v = inputs.get(k, default)
        if k in INT_KEYS:
            v = int(v)
        elif k in FLOAT_KEYS:
            v = float(v)
        elif k == "custom":
            v = bool(v)
        if k == "stress_sel" and v not in STRESS_OPTIONS:
            v = DEFAULT_INPUTS["stress_sel"]
        ss[k] = v
    bags = inputs.get("bags")
    ss.bags_df = (pd.DataFrame(bags)[BAG_COLS].astype(float) if bags
                  else DEFAULT_BAGS.copy())
    ss.bags_ver += 1


# --- callbacks ---
def cb_logout() -> None:
    for k in ["user", "patient_id"]:
        ss.pop(k, None)
    apply_inputs(DEFAULT_INPUTS)


def cb_add_patient() -> None:
    name = ss.get("new_patient_name", "").strip()
    if not name:
        flash("error", "請輸入病患名稱")
        return
    pid, err = create_patient(ss.user["id"], name)
    if err:
        flash("error", err)
        return
    ss.patient_id = pid
    ss.new_patient_name = ""
    apply_inputs(DEFAULT_INPUTS)  # 新病患從預設值開始
    flash("success", f"已新增病患「{name}」")


def cb_delete_patient(pid: int, name: str) -> None:
    delete_patient(ss.user["id"], pid)
    ss.pop("patient_id", None)
    flash("success", f"已刪除病患「{name}」及其紀錄")


def cb_select_patient() -> None:
    """切換病患時自動載入該病患最近一筆紀錄。"""
    recs = list_records(ss.user["id"], ss.patient_id)
    if recs:
        apply_inputs(json.loads(recs[0]["data"])["inputs"])
        flash("info", f"已載入最近紀錄 ({recs[0]['created_at']})")
    else:
        apply_inputs(DEFAULT_INPUTS)


def cb_save_record(pid: int, payload: dict) -> None:
    note = ss.get("note", "").strip()
    if add_record(ss.user["id"], pid, note, payload):
        ss.note = ""
        flash("success", "已儲存紀錄")
    else:
        flash("error", "儲存失敗:找不到此病患")


def cb_load_record(inputs: dict, created_at: str) -> None:
    apply_inputs(inputs)
    flash("info", f"已載入 {created_at} 的紀錄")


def cb_delete_record(rid: int) -> None:
    delete_record(ss.user["id"], rid)
    flash("success", "已刪除紀錄")


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

# 顯示 callback 留下的訊息
_ICONS = {"success": "✅", "error": "⚠️", "info": "ℹ️"}
for _kind, _msg in ss.flash:
    st.toast(_msg, icon=_ICONS.get(_kind))
ss.flash = []

# ----------------------------------------------------------------------
# 登入 / 建立帳號 (未登入時只顯示這一頁)
# ----------------------------------------------------------------------
if "user" not in ss:
    _, mid, _ = st.columns([1, 2, 1])
    with mid:
        tab_login, tab_reg = st.tabs(["登入", "建立帳號"])
        with tab_login:
            with st.form("login_form"):
                lu = st.text_input("帳號")
                lp = st.text_input("密碼", type="password")
                if st.form_submit_button("登入", type="primary", use_container_width=True):
                    user = verify_user(lu, lp)
                    if user:
                        ss.user = user
                        st.rerun()
                    else:
                        st.error("帳號或密碼錯誤")
        with tab_reg:
            with st.form("register_form"):
                ru = st.text_input("帳號", help=f"至少 {MIN_LEN} 個字")
                rp = st.text_input("密碼", type="password", help=f"至少 {MIN_LEN} 個字")
                rp2 = st.text_input("再次輸入密碼", type="password")
                if st.form_submit_button("建立帳號", type="primary", use_container_width=True):
                    if rp != rp2:
                        st.error("兩次輸入的密碼不一致")
                    else:
                        uid, err = create_user(ru, rp)
                        if err:
                            st.error(err)
                        else:
                            ss.user = {"id": uid, "username": ru.strip()}
                            st.rerun()
    st.caption("※ 本工具僅供輔助計算,實際處方請由臨床醫師/營養師確認。")
    st.stop()

user = ss.user

# ----------------------------------------------------------------------
# 側邊欄: 帳號、病患管理、病人資料
# ----------------------------------------------------------------------
sb = st.sidebar
ac1, ac2 = sb.columns([3, 2], vertical_alignment="center")
ac1.markdown(f"👤 **{user['username']}**")
ac2.button("登出", on_click=cb_logout, use_container_width=True)

sb.header("病患")
patients = {p["id"]: p["name"] for p in list_patients(user["id"])}
if ss.get("patient_id") not in patients:
    if patients:
        ss.patient_id = next(iter(patients))
    else:
        ss.pop("patient_id", None)

if patients:
    sb.selectbox("目前病患", list(patients), format_func=lambda i: patients[i],
                 key="patient_id", on_change=cb_select_patient)
else:
    sb.info("尚未建立病患,請先新增")
pid = ss.get("patient_id")

with sb.expander("➕ 新增病患", expanded=not patients):
    with st.form("new_patient_form"):
        st.text_input("病患名稱", key="new_patient_name", max_chars=50,
                      placeholder="例: 床號 / 代號")
        st.form_submit_button("新增", on_click=cb_add_patient, use_container_width=True)

if pid:
    with sb.popover("🗑 刪除此病患", use_container_width=True):
        st.warning(f"將刪除「{patients[pid]}」及其所有歷史紀錄,無法復原。")
        st.button("確認刪除", type="primary", key="confirm_del_patient",
                  on_click=cb_delete_patient, args=(pid, patients[pid]))

sb.divider()
sb.header("病人資料")
col_age, col_sex = sb.columns([3, 2])
age_years = col_age.number_input("年齡 (歲)", min_value=0, max_value=18, step=1, format="%d",
                                 key="age_years", help="未滿 1 歲請填 0,再輸入月齡")
sex = col_sex.radio("性別", ["M", "F"], horizontal=True, key="sex")
age = float(age_years)
if age_years == 0:
    # 嬰兒公式依月齡區分 (≤3、≤6、≤12 個月),需另外輸入月齡
    age_months = sb.number_input("月齡 (月)", min_value=0, max_value=11, step=1, format="%d",
                                 key="age_months")
    age = age_months / 12
bw = float(sb.number_input("體重 (kg)", min_value=1, max_value=150, step=1, format="%d",
                           key="bw"))

stress = sb.selectbox("壓力參數", list(STRESS_OPTIONS), key="stress_sel",
                      format_func=lambda v: f"{v:g}")
sb.caption("\n\n".join(f"**{k:g}**:{v}" for k, v in STRESS_OPTIONS.items()))
custom = sb.checkbox("自訂壓力參數", key="custom")
if custom:
    stress = sb.number_input("壓力參數 (自訂)", min_value=1.0, max_value=2.5, step=0.05,
                             key="stress_custom")
cal_factor = sb.number_input(
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
sb.caption("熱量係數:PN 約為 EN 需求的 85–90%,預設 0.85")

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
        ss.bags_df,
        num_rows="fixed",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Dex (%)": st.column_config.NumberColumn(min_value=0.0, max_value=50.0, step=0.5),
            "AA (%)": st.column_config.NumberColumn(min_value=0.0, max_value=15.0, step=0.5),
            "Volume (mL)": st.column_config.NumberColumn(min_value=0, step=50, format="%d"),
        },
        key=f"bags_{ss.bags_ver}",
    )

    st.markdown("**脂肪 Fat (SMOFlipid 20%)**")
    fat_vol = st.number_input("容量 (mL/day)", min_value=0, max_value=2000, step=50,
                              format="%d", key="fat_vol",
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

    # --- 儲存紀錄 ---
    payload = {
        "inputs": {
            **{k: ss[k] for k in DEFAULT_INPUTS},
            "bags": {c: [float(x) for x in bags[c].fillna(0)] for c in BAG_COLS},
        },
        "results": {
            "elabel": elabel,
            "ree": float(ree),
            "stress": float(stress),
            "target_kcal": float(target_kcal),
            "kcal": float(r["kcal"]),
            "pct": float(pct),
            "sir": float(r["sir"]),
            "aa_kg": float(r["aa_kg"]),
            "fat_kg": float(r["fat_kg"]),
            "volume": float(r["volume"]),
            "fluid_pct": float(fluid_pct),
        },
    }
    if pid:
        sc1, sc2 = st.columns([3, 1], vertical_alignment="bottom")
        sc1.text_input("備註 (選填)", key="note", max_chars=200,
                       placeholder="例: Day 3 增量")
        sc2.button("💾 儲存紀錄", type="primary", use_container_width=True,
                   on_click=cb_save_record, args=(pid, payload))
    else:
        st.info("在左側新增病患後,即可儲存這次的資料")

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

# ----------------------------------------------------------------------
# 歷史紀錄
# ----------------------------------------------------------------------
if pid:
    st.divider()
    st.subheader(f"歷史紀錄:{patients[pid]}")
    recs = list_records(user["id"], pid)
    if not recs:
        st.caption("尚無紀錄。調整好配方後,按上方「💾 儲存紀錄」即可存檔。")
    else:
        def _age_text(inp: dict) -> str:
            if int(inp.get("age_years", 0)) == 0:
                return f"{int(inp.get('age_months', 0))} 個月"
            return f"{int(inp['age_years'])} 歲"

        rows = []
        for rec in recs:
            d = json.loads(rec["data"])
            inp, res = d["inputs"], d["results"]
            rows.append({
                "時間": rec["created_at"][:16],
                "備註": rec["note"] or "",
                "年齡": _age_text(inp),
                "體重 (kg)": inp["bw"],
                "PN 熱量": round(res["kcal"]),
                "目標熱量": round(res["target_kcal"]),
                "達成 (%)": round(res["pct"]),
                "SIR": round(res["sir"], 2),
                "AA": round(res["aa_kg"], 2),
                "Fat": round(res["fat_kg"], 2),
                "水分 (mL)": round(res["volume"]),
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

        st.caption("展開單筆紀錄可載入到上方計算器,或刪除該筆紀錄。")
        for rec in recs:
            d = json.loads(rec["data"])
            inp, res = d["inputs"], d["results"]
            label = (f"{rec['created_at'][:16]}　PN {res['kcal']:,.0f} / "
                     f"目標 {res['target_kcal']:,.0f} kcal")
            if rec["note"]:
                label += f"　｜{rec['note']}"
            with st.expander(label):
                stress_txt = (f"{float(inp['stress_custom']):g} (自訂)" if inp.get("custom")
                              else f"{float(inp['stress_sel']):g}")
                st.markdown(
                    f"**年齡** {_age_text(inp)}　**性別** {inp['sex']}　"
                    f"**體重** {inp['bw']} kg　**壓力參數** {stress_txt}　"
                    f"**熱量係數** {float(inp['cal_factor']):g}　"
                    f"**SMOFlipid** {inp['fat_vol']} mL"
                )
                used = pd.DataFrame(inp["bags"])[BAG_COLS]
                used = used[used["Volume (mL)"] > 0]
                if len(used):
                    st.dataframe(used, hide_index=True, use_container_width=True)
                else:
                    st.caption("(未使用 PN 袋)")
                b1, b2, _ = st.columns([1, 1, 3])
                b1.button("載入", key=f"load_{rec['id']}", use_container_width=True,
                          on_click=cb_load_record, args=(inp, rec["created_at"][:16]))
                with b2.popover("🗑 刪除", use_container_width=True):
                    st.write("確定要刪除這筆紀錄?")
                    st.button("確認刪除", key=f"del_{rec['id']}", type="primary",
                              on_click=cb_delete_record, args=(rec["id"],))

st.caption("※ 本工具僅供輔助計算,實際處方請由臨床醫師/營養師確認。")
st.caption("Neil JJC HMMCH 製作")
