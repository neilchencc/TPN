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
  - 標題旁的按鈕: 登入 / 註冊;登入後同位置可選擇病患、儲存、看歷史紀錄、管理病患
  - 病患、歷史紀錄都可單獨刪除 (刪除前會再確認)
  - 計算器不登入也能使用
  - 需要 Streamlit >= 1.40 (st.dialog、具 key 的 container)
  - 資料存在 SQLite 檔案 tpn_calc.db (位置見 _pick_db_path)
單位帳號 (仿照 NICU 計算器的 PICU / NBC):
  - 登入時可選 6A / 8A8B (密碼預設 hmmch,可在 .streamlit/secrets.toml 以
    [passwords] 6A = "..." 覆寫) 或個人帳號
  - 單位沒有預設床位,床位需自行新增,也可刪除 (連同該床紀錄);選單依床號自然排序
  - 登入後同位置可選擇床位、儲存、看歷史紀錄、新增 / 管理床位
執行: streamlit run TPN.py
"""
import base64
import hashlib
import hmac
import html
import json
import secrets
import functools
import os
import re
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
DB_FILENAME = "tpn_calc.db"


def _pick_db_path() -> Path:
    """找一個可以正常寫入的位置放資料庫。
    Streamlit Cloud 的程式資料夾 (/mount/src/...) 是掛載的檔案系統,SQLite 在上面可能無法寫入,
    所以優先放家目錄或暫存資料夾,最後才試程式資料夾。
    可用環境變數 TPN_DB_DIR 指定資料夾。"""
    candidates = []
    if os.environ.get("TPN_DB_DIR"):
        candidates.append(Path(os.environ["TPN_DB_DIR"]))
    candidates += [Path.home() / ".tpn_calc", Path(tempfile.gettempdir()) / "tpn_calc",
                   Path(__file__).parent]
    for folder in candidates:
        path = folder / DB_FILENAME
        try:
            folder.mkdir(parents=True, exist_ok=True)
            with closing(sqlite3.connect(path)) as c:
                # 模擬實際寫入 (含 journal 檔),測完 rollback 不留痕跡
                c.execute("BEGIN IMMEDIATE")
                c.execute("CREATE TABLE IF NOT EXISTS _write_test (x TEXT)")
                c.execute("INSERT INTO _write_test VALUES ('ok')")
                c.rollback()
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
            -- 單位帳號 (6A / 8A8B): 自訂床位與各床紀錄
            CREATE TABLE IF NOT EXISTS unit_beds (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                unit       TEXT NOT NULL,
                bed        TEXT NOT NULL COLLATE NOCASE,
                created_at TEXT NOT NULL,
                UNIQUE (unit, bed)
            );
            CREATE TABLE IF NOT EXISTS unit_records (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                unit       TEXT NOT NULL,
                bed        TEXT NOT NULL,
                created_at TEXT NOT NULL,
                note       TEXT,
                data       TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_unit_records ON unit_records (unit, bed, id);
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
        return c.execute(
            """SELECT p.id, p.name, COUNT(r.id) AS n FROM patients p
               LEFT JOIN records r ON r.patient_id = p.id
               WHERE p.user_id = ? GROUP BY p.id ORDER BY p.created_at, p.id""",
            (user_id,),
        ).fetchall()


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


# ----------------------------------------------------------------------
# 單位帳號 (6A / 8A8B): 沒有預設床位,床位自行新增 / 刪除
# ----------------------------------------------------------------------
UNITS = ["6A", "8A8B"]
DEFAULT_PASSWORD = "hmmch"
BED_MAX_LEN = 20


def get_password(unit: str) -> str:
    """優先使用 st.secrets["passwords"][unit],否則用預設密碼。"""
    try:
        return str(st.secrets["passwords"][unit])
    except Exception:
        return DEFAULT_PASSWORD


def check_unit_login(unit: str, password: str) -> bool:
    return unit in UNITS and hmac.compare_digest(password.encode(),
                                                 get_password(unit).encode())


def bed_sort_key(bed: str):
    """依床號自然排序: 數字部分以數值比較 (2 < 10 < 10-1 < 12),英文不分大小寫。"""
    return [(0, int(t), "") if t.isdigit() else (1, 0, t.lower())
            for t in re.split(r"(\d+)", str(bed)) if t]


def list_beds(unit: str) -> list:
    with closing(db()) as c:
        beds = [r["bed"] for r in c.execute("SELECT bed FROM unit_beds WHERE unit = ?", (unit,))]
    return sorted(beds, key=bed_sort_key)


def add_bed(unit: str, bed: str):
    """成功回傳 None,失敗回傳錯誤訊息。"""
    bed = bed.strip()
    if not bed:
        return "請輸入床號"
    if len(bed) > BED_MAX_LEN:
        return f"床號最多 {BED_MAX_LEN} 個字"
    try:
        with closing(db()) as c, c:
            c.execute("INSERT INTO unit_beds (unit, bed, created_at) VALUES (?, ?, ?)",
                      (unit, bed, now_str()))
    except sqlite3.IntegrityError:
        return f"已經有床位「{bed}」"
    return None


def delete_bed(unit: str, bed: str) -> None:
    """刪除床位及其所有紀錄。"""
    with closing(db()) as c, c:
        c.execute("DELETE FROM unit_records WHERE unit = ? AND bed = ?", (unit, bed))
        c.execute("DELETE FROM unit_beds WHERE unit = ? AND bed = ?", (unit, bed))


def unit_add_record(unit: str, bed: str, note: str, data: dict) -> bool:
    with closing(db()) as c, c:
        if not c.execute("SELECT 1 FROM unit_beds WHERE unit = ? AND bed = ?",
                         (unit, bed)).fetchone():
            return False
        c.execute("INSERT INTO unit_records (unit, bed, created_at, note, data) "
                  "VALUES (?, ?, ?, ?, ?)",
                  (unit, bed, now_str(), note, json.dumps(data, ensure_ascii=False)))
        return True


def unit_list_records(unit: str, bed: str) -> list:
    with closing(db()) as c:
        return c.execute(
            """SELECT id, created_at, note, data FROM unit_records
               WHERE unit = ? AND bed = ? ORDER BY created_at DESC, id DESC""",
            (unit, bed),
        ).fetchall()


def unit_delete_record(unit: str, record_id: int) -> None:
    with closing(db()) as c, c:
        c.execute("DELETE FROM unit_records WHERE id = ? AND unit = ?", (record_id, unit))


def unit_latest(unit: str) -> dict:
    """{bed: (created_at, data)} 每床最近一筆。"""
    with closing(db()) as c:
        rows = c.execute("SELECT bed, created_at, data FROM unit_records WHERE unit = ? "
                         "ORDER BY created_at, id", (unit,)).fetchall()
    return {r["bed"]: (r["created_at"], json.loads(r["data"])) for r in rows}


init_db()

# ----------------------------------------------------------------------
# 輸入欄位的狀態 (用 session_state 才能從歷史紀錄載入)
# ----------------------------------------------------------------------
DEFAULT_INPUTS = {
    "age_years": None,  # 年齡、月齡、體重不給預設值,需自行輸入
    "age_months": None,
    "sex": "M",
    "bw": None,
    "stress_sel": 1.3,
    "custom": False,
    "stress_custom": 1.3,
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
ss.pop("_rerun_app", None)


def flash(kind: str, msg: str, seconds: float = 0) -> None:
    """在 callback 裡記下訊息,下次整頁更新時顯示。
    seconds > 0: 右上角小提示,顯示指定秒數後自動消失;0: 一般 toast。"""
    ss.flash.append((kind, msg, seconds))


def apply_inputs(inputs: dict) -> None:
    """排定要載入的輸入值。
    真正填入欄位是在下一次整頁執行的最前面 (任何輸入欄位畫出來之前),
    這樣不論是從主畫面或從對話框 (歷史紀錄) 按的,左側病人資料與 PN 配方表格都會正確更新。"""
    ss._pending_inputs = dict(inputs)


def _fill_inputs_now(inputs: dict) -> None:
    """把一組輸入值實際寫入所有欄位,只在腳本最前面呼叫。"""
    for k, default in DEFAULT_INPUTS.items():
        v = inputs.get(k, default)
        if v is None:
            pass
        elif k in INT_KEYS:
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
    ss.bags_ver += 1  # 換一個新的 data_editor,表格才會顯示載入的配方


def target_records(target: tuple) -> list:
    """target: ("unit", 單位, 床號) 或 ("acct", 病患 id)。回傳紀錄 (由新到舊)。"""
    if target[0] == "unit":
        return unit_list_records(target[1], target[2])
    return list_records(ss.user["id"], target[1])


def load_latest(target: tuple) -> None:
    """載入床位 / 病患最近一筆紀錄;沒有紀錄就回到預設值。"""
    recs = target_records(target)
    if recs:
        apply_inputs(json.loads(recs[0]["data"])["inputs"])
        flash("info", f"已載入最近紀錄 ({recs[0]['created_at'][:16]})")
    else:
        apply_inputs(DEFAULT_INPUTS)


def request_app_rerun() -> None:
    """對話框裡的 callback 用:之後整頁重新整理並關閉對話框。"""
    ss._rerun_app = True


def rerun_if_requested() -> None:
    if ss.pop("_rerun_app", False):
        st.rerun()


def db_safe(func):
    """資料庫出錯時不讓整個 app 當掉,改把真正的錯誤訊息顯示在畫面上。"""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except sqlite3.Error as e:
            ss._db_err = f"資料庫錯誤 ({type(e).__name__}):{e}\n\n資料庫位置:{DB_PATH}"
    return wrapper


def show_db_error() -> None:
    err = ss.pop("_db_err", None)
    if err:
        st.error(err)


# --- callbacks: 登入 (單位帳號 6A / 8A8B 或個人帳號) ---
PERSONAL = "個人帳號"
LOGIN_KINDS = UNITS + [PERSONAL]
LOGIN_KIND_LABELS = {u: f"{u} (單位帳號)" for u in UNITS}
LOGIN_KIND_LABELS[PERSONAL] = "個人帳號 (自行管理病患)"


def _clear_login() -> None:
    for k in ["unit", "bed_sel", "user", "patient_id", "confirm_del_patient",
              "confirm_del_record", "confirm_del_bed"]:
        ss.pop(k, None)


@db_safe
def cb_login() -> None:
    kind = ss.get("li_kind", LOGIN_KINDS[0])
    pw = ss.get("li_pw", "")
    ss.li_pw = ""
    if kind in UNITS:
        if not check_unit_login(kind, pw):
            ss._auth_err = "密碼錯誤"
            return
        _clear_login()
        ss.unit = kind
        beds = list_beds(kind)
        if beds:
            ss.bed_sel = beds[0]
            load_latest(("unit", kind, beds[0]))
        else:
            apply_inputs(DEFAULT_INPUTS)
        flash("success", f"已登入單位 {kind}" + ("" if beds else ",請先新增床位"))
        request_app_rerun()
        return
    user = verify_user(ss.get("li_user", ""), pw)
    if not user:
        ss._auth_err = "帳號或密碼錯誤"
        return
    _clear_login()
    ss.user = user
    plist = list_patients(user["id"])
    if plist:
        ss.patient_id = plist[0]["id"]
        load_latest(("acct", plist[0]["id"]))
    flash("success", f"歡迎,{user['username']}")
    request_app_rerun()


@db_safe
def cb_register() -> None:
    u, p, p2 = ss.get("rg_user", ""), ss.get("rg_pw", ""), ss.get("rg_pw2", "")
    if u.strip().upper() in (x.upper() for x in UNITS):
        ss._auth_err = f"「{u.strip()}」是單位帳號,請直接在「登入」頁選擇單位登入"
        return
    if p != p2:
        ss._auth_err = "兩次輸入的密碼不一致"
        return
    uid, err = create_user(u, p)
    if err:
        ss._auth_err = err
        return
    _clear_login()
    ss.user = {"id": uid, "username": u.strip()}
    ss.rg_pw = ss.rg_pw2 = ""
    flash("success", "帳號已建立,請新增第一位病患")
    request_app_rerun()


def cb_logout() -> None:
    _clear_login()
    apply_inputs(DEFAULT_INPUTS)
    flash("info", "已登出")


# --- callbacks: 床位 (單位帳號) ---
@db_safe
def cb_select_bed() -> None:
    load_latest(("unit", ss.unit, ss.bed_sel))


@db_safe
def cb_add_bed() -> None:
    bed = ss.get("new_bed", "").strip()
    err = add_bed(ss.unit, bed)
    if err:
        ss._bed_err = err
        return
    ss.new_bed = ""
    ss.bed_sel = bed
    apply_inputs(DEFAULT_INPUTS)  # 新床位從預設值開始
    flash("success", f"已新增床位 {bed}")
    request_app_rerun()


@db_safe
def cb_delete_bed(bed: str) -> None:
    unit = ss.unit
    n = len(unit_list_records(unit, bed))
    delete_bed(unit, bed)
    ss.pop("confirm_del_bed", None)
    if ss.get("bed_sel") == bed:
        beds = list_beds(unit)
        if beds:
            ss.bed_sel = beds[0]
            load_latest(("unit", unit, beds[0]))
        else:
            ss.pop("bed_sel", None)
            apply_inputs(DEFAULT_INPUTS)
    flash("success", f"已刪除床位 {bed}" + (f" 及 {n} 筆紀錄" if n else ""))
    request_app_rerun()


# --- callbacks: 病患 (個人帳號) ---
@db_safe
def cb_select_patient() -> None:
    load_latest(("acct", ss.patient_id))


@db_safe
def cb_add_patient() -> None:
    name = ss.get("new_patient_name", "").strip()
    if not name:
        ss._pat_err = "請輸入病患名稱"
        return
    pid, err = create_patient(ss.user["id"], name)
    if err:
        ss._pat_err = err
        return
    ss.patient_id = pid
    ss.new_patient_name = ""
    apply_inputs(DEFAULT_INPUTS)  # 新病患從預設值開始
    flash("success", f"已新增病患「{name}」")
    request_app_rerun()


@db_safe
def cb_choose_patient(pid: int) -> None:
    ss.patient_id = pid
    load_latest(("acct", pid))
    request_app_rerun()


@db_safe
def cb_delete_patient(pid: int, name: str) -> None:
    delete_patient(ss.user["id"], pid)
    ss.pop("confirm_del_patient", None)
    if ss.get("patient_id") == pid:
        ss.pop("patient_id", None)
    flash("success", f"已刪除病患「{name}」及其紀錄")
    request_app_rerun()


# --- callbacks: 紀錄 ---
@db_safe
def cb_save_record(target: tuple, payload: dict) -> None:
    note = ss.get("note", "").strip()
    if target[0] == "unit":
        ok = unit_add_record(target[1], target[2], note, payload)
    else:
        ok = add_record(ss.user["id"], target[1], note, payload)
    if ok:
        ss.note = ""
        flash("success", "已儲存紀錄", seconds=1)
    else:
        flash("error", "儲存失敗:找不到此" + ("床位" if target[0] == "unit" else "病患"))
    request_app_rerun()


@db_safe
def cb_load_record(inputs: dict, created_at: str) -> None:
    apply_inputs(inputs)
    flash("info", f"已載入 {created_at} 的紀錄")
    request_app_rerun()


@db_safe
def cb_delete_record(target: tuple, rid: int) -> None:
    if target[0] == "unit":
        unit_delete_record(target[1], rid)
    else:
        delete_record(ss.user["id"], rid)
    ss.pop("confirm_del_record", None)  # 對話框保持開啟,可以繼續刪


def cb_set(key: str, value) -> None:
    ss[key] = value


# ----------------------------------------------------------------------
# 對話框
# ----------------------------------------------------------------------
@st.dialog("登入 / 註冊")
def auth_dialog() -> None:
    rerun_if_requested()
    show_db_error()
    err = ss.pop("_auth_err", None)
    if err:
        st.error(err)
    tab_login, tab_reg = st.tabs(["登入", "建立帳號"])
    with tab_login:
        # 選單放在 form 外面,切換時才能立即顯示 / 隱藏帳號欄位
        kind = st.selectbox("登入類型", LOGIN_KINDS, key="li_kind",
                            format_func=lambda k: LOGIN_KIND_LABELS.get(k, k))
        with st.form("login_form", border=False):
            if kind == PERSONAL:
                st.text_input("帳號", key="li_user")
            st.text_input("單位密碼" if kind in UNITS else "密碼", type="password",
                          key="li_pw")
            st.form_submit_button("登入", type="primary", use_container_width=True,
                                  on_click=cb_login)
        if kind in UNITS:
            st.caption("單位帳號沒有預設床位,登入後請由「👤 單位」選單新增床位。")
        else:
            st.caption("個人帳號請先到「建立帳號」註冊,可自行管理病患。")
    with tab_reg:
        with st.form("register_form", border=False):
            st.text_input("帳號", key="rg_user", help=f"至少 {MIN_LEN} 個字")
            st.text_input("密碼", type="password", key="rg_pw", help=f"至少 {MIN_LEN} 個字")
            st.text_input("再次輸入密碼", type="password", key="rg_pw2")
            st.form_submit_button("建立帳號", type="primary", use_container_width=True,
                                  on_click=cb_register)
        st.caption(f"單位帳號 ({' / '.join(UNITS)}) 不需註冊。")


@st.dialog("管理病患")
def patients_dialog() -> None:
    rerun_if_requested()
    show_db_error()
    with st.form("new_patient_form", border=False):
        c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
        c1.text_input("新增病患", key="new_patient_name", max_chars=50,
                      placeholder="例: 床號 / 代號")
        c2.form_submit_button("新增", type="primary", use_container_width=True,
                              on_click=cb_add_patient)
    err = ss.pop("_pat_err", None)
    if err:
        st.error(err)

    st.divider()
    plist = list_patients(ss.user["id"])
    if not plist:
        st.caption("尚無病患")
    for p in plist:
        with st.container(border=True):
            c1, c2, c3 = st.columns([3, 1.3, 1.3], vertical_alignment="center")
            current = "　✅ 目前" if p["id"] == ss.get("patient_id") else ""
            c1.markdown(f"**{p['name']}**{current}  \n{p['n']} 筆紀錄")
            if ss.get("confirm_del_patient") == p["id"]:
                c2.button("確認刪除", type="primary", key=f"pdel_ok_{p['id']}",
                          use_container_width=True,
                          on_click=cb_delete_patient, args=(p["id"], p["name"]))
                c3.button("取消", key=f"pdel_no_{p['id']}", use_container_width=True,
                          on_click=cb_set, args=("confirm_del_patient", None))
                st.warning(f"將刪除「{p['name']}」及其 {p['n']} 筆紀錄,無法復原。")
            else:
                c2.button("選擇", key=f"psel_{p['id']}", use_container_width=True,
                          on_click=cb_choose_patient, args=(p["id"],))
                c3.button("🗑 刪除", key=f"pdel_{p['id']}", use_container_width=True,
                          on_click=cb_set, args=("confirm_del_patient", p["id"]))


@st.dialog("新增 / 管理床位")
def beds_dialog(unit: str) -> None:
    rerun_if_requested()
    show_db_error()
    with st.form("new_bed_form", border=False):
        c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
        c1.text_input(f"新增 {unit} 床位", key="new_bed", max_chars=BED_MAX_LEN,
                      placeholder="床號,例: 12 / 12-1 (請勿輸入姓名、病歷號)")
        c2.form_submit_button("新增", type="primary", use_container_width=True,
                              on_click=cb_add_bed)
    err = ss.pop("_bed_err", None)
    if err:
        st.error(err)
    st.caption("床位依床號排序;刪除床位會一併刪除該床所有紀錄。")

    st.divider()
    beds = list_beds(unit)
    if not beds:
        st.caption("尚無床位")
    for bed in beds:
        n = len(unit_list_records(unit, bed))
        with st.container(border=True):
            c1, c2, c3 = st.columns([3, 1.3, 1.3], vertical_alignment="center")
            current = "　✅ 目前" if bed == ss.get("bed_sel") else ""
            c1.markdown(f"**{bed}**{current}  \n{n} 筆紀錄")
            if ss.get("confirm_del_bed") == bed:
                c2.button("確認刪除", type="primary", key=f"bdel_ok_{bed}",
                          use_container_width=True, on_click=cb_delete_bed, args=(bed,))
                c3.button("取消", key=f"bdel_no_{bed}", use_container_width=True,
                          on_click=cb_set, args=("confirm_del_bed", None))
                st.warning(f"將刪除床位「{bed}」" + (f"及其 {n} 筆紀錄" if n else "")
                           + ",無法復原。")
            else:
                c3.button("🗑 刪除", key=f"bdel_{bed}", use_container_width=True,
                          on_click=cb_set, args=("confirm_del_bed", bed))


def _target_title(target: tuple, label: str) -> str:
    return f"{'床位' if target[0] == 'unit' else '病患'}:**{label}**"


@st.dialog("儲存紀錄")
def save_dialog(target: tuple, label: str, payload: dict) -> None:
    rerun_if_requested()
    show_db_error()
    res, inp = payload["results"], payload["inputs"]
    st.markdown(_target_title(target, label))
    st.caption(
        f"體重 {inp['bw']} kg ｜ PN {res['kcal']:,.0f} / 目標 {res['target_kcal']:,.0f} kcal "
        f"({res['pct']:.0f}%) ｜ SIR {res['sir']:.2f} ｜ AA {res['aa_kg']:.2f} ｜ "
        f"Fat {res['fat_kg']:.2f} ｜ 水分 {res['volume']:,.0f} mL"
    )
    st.text_input("備註 (選填)", key="note", max_chars=200, placeholder="例: Day 3 增量")
    st.button("💾 儲存", type="primary", use_container_width=True,
              on_click=cb_save_record, args=(target, payload))


def _age_text(inp: dict) -> str:
    if int(inp.get("age_years", 0)) == 0:
        return f"{int(inp.get('age_months', 0))} 個月"
    return f"{int(inp['age_years'])} 歲"


def _pn_text(inp: dict) -> str:
    """PN 配方摘要,例: D10/AA2 300 mL、D12.5/AA2 200 mL (只列容量 > 0 的組合)。"""
    bags = pd.DataFrame(inp.get("bags") or {c: [] for c in BAG_COLS})
    parts = [
        f"D{row['Dex (%)']:g}/AA{row['AA (%)']:g} {row['Volume (mL)']:g} mL"
        for _, row in bags.iterrows() if row["Volume (mL)"] > 0
    ]
    return "、".join(parts) if parts else "—"


def _text_width(text) -> int:
    """估計文字寬度:中文等全形字算 2,英數算 1。"""
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(text))


def _cell(text, header: bool = False) -> str:
    # 用固定最小高度 + flex 置中,讓文字和按鈕在同一水平線上
    weight = "font-weight:600;" if header else ""
    height = "2rem" if header else "2.5rem"
    return (f"<div style='display:flex; align-items:center; justify-content:center; "
            f"min-height:{height}; text-align:center; {weight} font-size:0.9rem; "
            f"line-height:1.3; word-break:break-word;'>{text}</div>")


@st.dialog("歷史紀錄", width="large")
def history_dialog(target: tuple, label: str) -> None:
    rerun_if_requested()
    show_db_error()
    st.markdown(_target_title(target, label))
    recs = target_records(target)
    if not recs:
        st.caption("尚無紀錄。調整好配方後,按上方「💾 儲存」即可存檔。")
        return

    rows = []
    for rec in recs:
        d = json.loads(rec["data"])
        inp, res = d["inputs"], d["results"]
        when = rec["created_at"][:16]
        rows.append({
            "_id": rec["id"],
            "_inp": inp,
            "時間": when,
            "備註": html.escape(rec["note"] or "") or "—",
            "年齡": _age_text(inp),
            "體重": f"{inp['bw']}",
            "PN 配方": _pn_text(inp),
            "SMOF (mL)": f"{int(inp.get('fat_vol', 0))}",
            "PN kcal": f"{res['kcal']:,.0f}",
            "目標": f"{res['target_kcal']:,.0f}",
            "達成%": f"{res['pct']:.0f}",
            "SIR": f"{res['sir']:.2f}",
            "AA": f"{res['aa_kg']:.2f}",
            "Fat": f"{res['fat_kg']:.2f}",
            "水分": f"{res['volume']:,.0f}",
        })
    cols = [c for c in rows[0] if not c.startswith("_")]

    # 欄寬 = 欄名與內容中最寬的那個 (按鈕欄給固定寬度)
    widths = [5, 5] + [
        max(_text_width(c), *(_text_width(r[c]) for r in rows)) + 2 for c in cols
    ]

    # 每一列底下加細線
    st.markdown(
        """
        <style>
        [class*="st-key-hrow_"] {
            border-bottom: 1px solid rgba(128, 128, 128, 0.25);
            padding: 0.3rem 0;
        }
        .st-key-hrow_header { border-bottom-width: 2px; }
        /* Streamlit 文字元件預設有負的下邊距,會讓分隔線壓到文字、文字偏下,這裡歸零 */
        [class*="st-key-hrow_"] [data-testid="stMarkdown"],
        [class*="st-key-hrow_"] [data-testid="stMarkdownContainer"],
        [class*="st-key-hrow_"] .stMarkdown {
            margin: 0 !important;
        }
        [class*="st-key-hrow_"] p { margin: 0 !important; }
        [class*="st-key-hrow_"] button { min-height: 2.5rem; padding: 0 0.25rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    with st.container(key="hrow_header"):
        hdr = st.columns(widths, gap="small", vertical_alignment="center")
        for c, name in zip(hdr, ["載入", "刪除"] + cols):
            c.markdown(_cell(name, header=True), unsafe_allow_html=True)

    for r in rows:
        rid = r["_id"]
        with st.container(key=f"hrow_{rid}"):
            cells = st.columns(widths, gap="small", vertical_alignment="center")
            cells[0].button("📥", key=f"rload_{rid}", use_container_width=True,
                            help="載入這筆資料到計算器",
                            on_click=cb_load_record, args=(r["_inp"], r["時間"]))
            cells[1].button("🗑", key=f"rdel_{rid}", use_container_width=True,
                            help="刪除這筆紀錄",
                            on_click=cb_set, args=("confirm_del_record", rid))
            for c, name in zip(cells[2:], cols):
                c.markdown(_cell(r[name]), unsafe_allow_html=True)

    rid = ss.get("confirm_del_record")
    ids = [r["_id"] for r in rows]
    if rid in ids:
        when = rows[ids.index(rid)]["時間"]
        with st.container(border=True):
            st.warning(f"確定要刪除 {when} 的紀錄?刪除後無法復原。")
            c1, c2, _ = st.columns([1, 1, 3])
            c1.button("確認刪除", type="primary", key=f"rdel_ok_{rid}",
                      use_container_width=True, on_click=cb_delete_record, args=(target, rid))
            c2.button("取消", key=f"rdel_no_{rid}", use_container_width=True,
                      on_click=cb_set, args=("confirm_del_record", None))


# 套用上一個動作排定的輸入值 (必須在任何輸入欄位建立之前)
if "_pending_inputs" in ss:
    _fill_inputs_now(ss.pop("_pending_inputs"))

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
try:
    _dark = st.context.theme.type == "dark"  # Streamlit >= 1.46
except Exception:
    _dark = False
_header_bg = "#0e1117" if _dark else "#ffffff"
# 標題列 (含登入 / 病患按鈕) 固定在畫面上方: 主畫面往下捲動時不動
st.markdown(
    f"""
    <style>
    /* 標題列右側: 全部排成一排不換行,按鈕寬度配合文字 */
    .st-key-hdr_ctrl [data-testid="stHorizontalBlock"] {{
        flex-wrap: nowrap !important;
        justify-content: flex-end;
        align-items: center;
        gap: 0.5rem;
    }}
    .st-key-hdr_ctrl [data-testid="stColumn"],
    .st-key-hdr_ctrl [data-testid="column"] {{
        flex: 0 0 auto !important;
        width: auto !important;
        min-width: 0 !important;
    }}
    /* 提示文字欄: 佔用剩下的空間,文字靠右 (緊貼病患選單左邊) */
    .st-key-hdr_ctrl [data-testid="stColumn"]:has(.st-key-notice),
    .st-key-hdr_ctrl [data-testid="column"]:has(.st-key-notice) {{
        flex: 1 1 auto !important;
        min-width: 0 !important;
        overflow: visible !important;
    }}
    .st-key-notice,
    .st-key-notice [data-testid="stElementContainer"],
    .st-key-notice [data-testid="stMarkdown"],
    .st-key-notice [data-testid="stMarkdownContainer"] {{
        overflow: visible !important;
    }}
    .st-key-hdr_ctrl, .st-key-hdr_ctrl [data-testid="stHorizontalBlock"] {{
        overflow: visible !important;
    }}
    /* 病患選單欄寬: 約 5 個英數 + 4 個中文字 */
    .st-key-hdr_ctrl [data-testid="stColumn"]:has(.st-key-pat_sel),
    .st-key-hdr_ctrl [data-testid="column"]:has(.st-key-pat_sel) {{
        flex: 0 0 10rem !important;
        width: 10rem !important;
    }}
    /* 床位選單含體重與時間,較寬 */
    .st-key-hdr_ctrl [data-testid="stColumn"]:has(.st-key-bed_box),
    .st-key-hdr_ctrl [data-testid="column"]:has(.st-key-bed_box) {{
        flex: 0 0 15rem !important;
        width: 15rem !important;
    }}
    .st-key-hdr_ctrl button {{ white-space: nowrap; }}
    .st-key-app_header {{
        position: sticky;
        top: 0;
        z-index: 999;
        background-color: {_header_bg};
        padding-bottom: 0.25rem;
        border-bottom: 1px solid rgba(128, 128, 128, 0.25);
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# 顯示 callback 留下的訊息
_ICONS = {"success": "✅", "error": "⚠️", "info": "ℹ️"}
_notes = ss.flash  # 提示文字統一顯示在標題列 (病患選單左邊)
ss.flash = []

NOTE_SECONDS = 3  # 一般提示顯示秒數 (有指定秒數者依指定,例如儲存 1 秒)


def render_notes(notes: list) -> None:
    """在目前位置顯示提示文字 (靠右對齊),數秒後淡出。"""
    parts = []
    for kind, msg, sec in notes:
        ss._note_n = ss.get("_note_n", 0) + 1  # 每次用新的動畫名稱,確保動畫重新播放
        anim = f"tpnFade{ss._note_n}"
        fg, bg = ("#b42318", "#fde4e1") if kind == "error" else ("#177233", "#dff3e4")
        parts.append(
            f"<style>@keyframes {anim} {{ 0%, 85% {{ opacity: 1; }} "
            f"100% {{ opacity: 0; visibility: hidden; }} }}</style>"
            f"<span style='display:inline-block; padding:0.35rem 0.8rem; "
            f"border-radius:0.5rem; font-size:0.9rem; font-weight:600; "
            f"white-space:normal; text-align:right; "
            f"color:{fg}; background:{bg}; animation:{anim} {sec or NOTE_SECONDS}s forwards;'>"
            f"{_ICONS.get(kind, '')} {html.escape(msg)}</span>"
        )
    # 外框只佔位置;提示本身以絕對定位貼齊病患選單左側、往左延伸,不會被欄寬切掉
    st.markdown(
        "<div style='position:relative; min-height:2.5rem;'>"
        "<div style='position:absolute; right:0; top:50%; transform:translateY(-50%); "
        "z-index:1000; width:max-content; max-width:min(32rem, 45vw); "
        "display:flex; flex-direction:column; align-items:flex-end; gap:0.25rem;'>"
        + "".join(parts) + "</div></div>",
        unsafe_allow_html=True,
    )

# --- 登入狀態: 單位帳號 (床位) 或個人帳號 (病患) ---
unit = ss.get("unit")
user = ss.get("user")
beds, bed_latest = [], {}
if unit:
    beds = list_beds(unit)
    bed_latest = unit_latest(unit)
    if ss.get("bed_sel") not in beds:
        if beds:
            ss.bed_sel = beds[0]
        else:
            ss.pop("bed_sel", None)
patients = {}
if user:
    patients = {p["id"]: p["name"] for p in list_patients(user["id"])}
    if ss.get("patient_id") not in patients:
        if patients:
            ss.patient_id = next(iter(patients))
        else:
            ss.pop("patient_id", None)
pid = ss.get("patient_id") if user else None
bed = ss.get("bed_sel") if unit else None
# 目前存檔對象: ("unit", 單位, 床號) 或 ("acct", 病患 id)
target = (("unit", unit, bed) if bed else None) if unit else (("acct", pid) if pid else None)
target_label = (f"{unit} · {bed}" if unit else patients.get(pid, "")) if target else ""
dialog_to_open = None


def bed_label(b: str) -> str:
    if b in bed_latest:
        t, d = bed_latest[b]
        bw_ = d.get("inputs", {}).get("bw")
        return f"{b} ● {bw_} kg · {t[5:16]}" if bw_ is not None else f"{b} ● {t[5:16]}"
    return f"{b} (尚無紀錄)"


# --- 標題列: 左邊標題,右邊登入 / 床位或病患功能 ---
with st.container(key="app_header"):
    h_title, h_ctrl = st.columns([5, 6], vertical_alignment="center")
    h_title.markdown(
        "<h1 class='app-title' style='margin-bottom:0; display:flex; align-items:center; "
        f"flex-wrap:wrap;'>{logo_html}靜脈營養計算器</h1>",
        unsafe_allow_html=True,
    )
    with h_ctrl.container(key="hdr_ctrl"):
        # 一排靠右: [提示文字] [床位 / 病患選單] [儲存] [紀錄] [帳號]
        if not (unit or user):
            c_note, c_login = st.columns(2, vertical_alignment="center")
            with c_note.container(key="notice"):
                render_notes(_notes)
            if c_login.button("🔑 登入 / 註冊",
                              help=f"{' / '.join(UNITS)} 單位帳號或個人帳號:"
                                   "登入後可儲存與查看歷史紀錄"):
                dialog_to_open = "auth"
        else:
            c_note, c_pat, c_save, c_hist, c_acct = st.columns(
                5, vertical_alignment="center")
            with c_note.container(key="notice"):
                render_notes(_notes)
            if unit:
                with c_pat.container(key="bed_box"):  # 寬度由 CSS 固定
                    if beds:
                        st.selectbox("床位", beds, key="bed_sel", format_func=bed_label,
                                     on_change=cb_select_bed, label_visibility="collapsed")
                    else:
                        st.caption("尚無床位,請由右側選單新增")
            else:
                with c_pat.container(key="pat_sel"):  # 寬度由 CSS 固定 (約 5 英數 + 4 中文)
                    if patients:
                        st.selectbox("目前病患", list(patients),
                                     format_func=lambda i: patients[i],
                                     key="patient_id", on_change=cb_select_patient,
                                     label_visibility="collapsed")
                    else:
                        st.caption("尚無病患")
            _ready_now = (ss.get("bw") is not None and ss.get("age_years") is not None
                          and (ss.get("age_years") != 0 or ss.get("age_months") is not None))
            if c_save.button("💾 儲存", disabled=not (target and _ready_now),
                             help="把目前輸入的資料存到這個" + ("床位" if unit else "病患")):
                dialog_to_open = "save"
            if c_hist.button("🕘 紀錄", disabled=not target,
                             help="查看、載入或刪除歷史紀錄"):
                dialog_to_open = "history"
            # 床位 / 病患管理與登出整合成一個按鈕
            with c_acct.popover(f"👤 {unit}" if unit else "👤 帳號"):
                if unit:
                    st.caption(f"單位帳號:**{unit}**")
                    if st.button("🛏️ 新增 / 管理床位", use_container_width=True):
                        dialog_to_open = "beds"
                else:
                    st.caption(f"登入身分:**{user['username']}**")
                    if st.button("👥 新增 / 管理病患", use_container_width=True):
                        dialog_to_open = "patients"
                st.button("🚪 登出", on_click=cb_logout, use_container_width=True)

show_db_error()

# ----------------------------------------------------------------------
# 側邊欄: 病人資料
# ----------------------------------------------------------------------
sb = st.sidebar
sb.header("病人資料")
col_age, col_sex = sb.columns([3, 2])
age_years = col_age.number_input("年齡 (歲)", min_value=0, max_value=18, step=1, format="%d",
                                 key="age_years", placeholder="請輸入",
                                 help="未滿 1 歲請填 0,再輸入月齡")
sex = col_sex.radio("性別", ["M", "F"], horizontal=True, key="sex")
age = None if age_years is None else float(age_years)
if age_years == 0:
    # 嬰兒公式依月齡區分 (≤3、≤6、≤12 個月),需另外輸入月齡
    age_months = sb.number_input("月齡 (月)", min_value=0, max_value=11, step=1, format="%d",
                                 key="age_months", placeholder="請輸入")
    age = None if age_months is None else age_months / 12
bw_in = sb.number_input("體重 (kg)", min_value=1, max_value=150, step=1, format="%d",
                        key="bw", placeholder="請輸入")
bw = None if bw_in is None else float(bw_in)
ready = age is not None and bw is not None  # 年齡與體重都填了才計算

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

if not ready:
    with right:
        st.subheader("計算結果")
        st.info("請先在左側輸入年齡與體重" + ("(未滿 1 歲請再輸入月齡)" if age_years == 0 else "")
                + ",即可計算熱量需求與 PN 結果。")
else:
    # --- 目標熱量 ---
    ree = calc_ree(age, sex, bw)
    elabel = energy_label(age)
    target_kcal = ree * stress * cal_factor
    maint = maintenance_fluid(bw)

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

# ----------------------------------------------------------------------
# 開啟標題列按下的對話框 (放最後,儲存時才拿得到計算結果)
# ----------------------------------------------------------------------
if dialog_to_open == "auth":
    auth_dialog()
elif dialog_to_open == "patients":
    patients_dialog()
elif dialog_to_open == "beds" and unit:
    beds_dialog(unit)
elif dialog_to_open == "history" and target:
    history_dialog(target, target_label)
elif dialog_to_open == "save" and target and ready:
    save_dialog(target, target_label, {
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
    })
