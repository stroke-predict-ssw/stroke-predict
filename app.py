import streamlit as st
import pandas as pd
import numpy as np
import joblib
import configparser
from sqlalchemy import create_engine, text
from cryptography.fernet import Fernet
import requests
import plotly.graph_objects as go
from urllib.parse import quote_plus
import re
import os
import sys

# ============================================
# 🛠️ Helper: หาไฟล์
# ============================================
def get_path(filename):
    search_paths = []
    
    # 1. Check directory of the Executable (if frozen)
    if getattr(sys, 'frozen', False):
        exe_path = sys.executable
        exe_folder = os.path.dirname(exe_path)
        search_paths.append(os.path.join(exe_folder, filename))
    
    # 2. Check Current Working Directory
    search_paths.append(os.path.join(os.getcwd(), filename))
    
    # 3. Check internal _MEIPASS (PyInstaller temp dir)
    if getattr(sys, 'frozen', False):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        search_paths.append(os.path.join(base_path, filename))
    else:
        # Development mode
        search_paths.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), filename))
    
    # Iterate and find first existing file
    for path in search_paths:
        if os.path.exists(path): return path
        
    # Return last attempted path to show in error (usually the internal one)
    return search_paths[-1]

# ============================================
# 🔑 Master Key
# ============================================
def load_key():
    key_path = get_path("secret.key")
    if os.path.exists(key_path):
        with open(key_path, "rb") as f:
            return f.read().strip()
    
    # Generate a new secure random key if it doesn't exist
    new_key = Fernet.generate_key()
    try:
        with open(key_path, "wb") as f: 
            f.write(new_key)
    except: pass
    return new_key

MASTER_KEY = load_key()
cipher_suite = Fernet(MASTER_KEY)

# ============================================
# ☁️ Cloud Logging
# ============================================
GOOGLE_SCRIPT_URL = "" 
ENABLE_CLOUD_LOG = True

# ============================================
# ⚙️ Config & Helpers
# ============================================
FBS_CODES_SQL = "'76', '1698'"
CHOL_CODES_SQL = "'102', '1691'"

def parse_lab_codes(code_string):
    if not code_string: return "''"
    codes = [c.strip() for c in code_string.split(",")]
    quoted_codes = [f"'{c}'" for c in codes]
    return ", ".join(quoted_codes)

def load_and_secure_config():
    global FBS_CODES_SQL, CHOL_CODES_SQL
    config = configparser.ConfigParser()
    try:
        config_path = get_path("config.ini")
        if os.path.exists(config_path):
            config.read(config_path, encoding="utf-8")
            raw_pass = config["Database"].get("password", "")
            final_db_pass = raw_pass
            if raw_pass:
                try:
                    decrypted = cipher_suite.decrypt(raw_pass.encode()).decode()
                    final_db_pass = decrypted
                except:
                    encrypted_str = cipher_suite.encrypt(raw_pass.encode()).decode()
                    config.set("Database", "password", encrypted_str)
                    try:
                        with open(config_path, "w", encoding="utf-8") as configfile:
                            config.write(configfile)
                    except: pass
                    final_db_pass = raw_pass

            try:
                raw_fbs = config["LabCodes"].get("fbs_codes", "76, 1698")
                raw_chol = config["LabCodes"].get("chol_codes", "102, 1691")
                FBS_CODES_SQL = parse_lab_codes(raw_fbs)
                CHOL_CODES_SQL = parse_lab_codes(raw_chol)
            except: pass
            return config, final_db_pass
        else:
            return config, ""
    except Exception as e:
        st.error(f"Config Error: {e}")
        return None, None

conf, DB_PASS = load_and_secure_config()

# Default Variables
DB_TYPE = "mysql"; DB_HOST = "localhost"; DB_PORT = "3306"; DB_NAME = "hos"; DB_USER = "sa"; HOSPITAL_NAME = "Hospital AI"; HCODE = "00000"

if conf:
    try:
        HOSPITAL_NAME = conf["General"].get("hospital_name", HOSPITAL_NAME)
        HCODE = conf["General"].get("hospital_code", HCODE)
        DB_TYPE = conf["Database"].get("db_type", DB_TYPE).lower()
        DB_HOST = conf["Database"].get("host", DB_HOST)
        DB_PORT = conf["Database"].get("port", DB_PORT)
        DB_NAME = conf["Database"].get("database", DB_NAME)
        DB_USER = conf["Database"].get("username", DB_USER)
        if "Cloud" in conf: GOOGLE_SCRIPT_URL = conf["Cloud"].get("google_script_url", GOOGLE_SCRIPT_URL)
    except: pass

st.set_page_config(page_title=f"Stroke Risk - {HOSPITAL_NAME}", page_icon="🧠", layout="wide")

# ============================================
# 🛠️ Load Model
# ============================================
try: cache_decorator = st.cache_resource
except: cache_decorator = st.cache(allow_output_mutation=True)

@cache_decorator
def load_model():
    try:
        model_path = get_path("stroke_model.pkl")
        model = joblib.load(model_path)
        # features = ["gender", "age", "cardio", "marry_status_numeric_map", "smoking_status", "drinking_status", "sbp", "dbp", "weight", "bmi", "waist", "avg_glocose_level", "cholesterol", "occupation", "education"]
        features = ['age', 'gender_numeric_map', 'marry_status_numeric_map', 'sbp', 'dbp', 'weight', 'bmi', 'waist', 'avg_glocose_level', 'cholesterol', 'smoking_status', 'drinking_status', 'occupation', 'education', 'cardio_numeric_map']
        return model, features, None
    except Exception as e:
        return None, None, str(e)

model, feature_names, load_error = load_model()

if model is None: 
    st.error(f"❌ ไม่สามารถโหลดไฟล์โมเดล (stroke_model.pkl) ได้: {load_error}")
    with st.expander("🛠️ Debug Info (คลิกเพื่อดูรายการไฟล์)"):
        try:
            st.write(f"Ref Name: stroke_model.pkl")
            if getattr(sys, 'frozen', False):
                exe_dir = os.path.dirname(sys.executable)
                st.write(f"📂 Exe Dir: {exe_dir}")
                st.write("📄 Files found:", os.listdir(exe_dir))
            st.write(f"📂 CWD: {os.getcwd()}")
            st.write("📄 Files in CWD:", os.listdir(os.getcwd()))
        except Exception as e:
            st.write(f"Debug Error: {e}")

def recalibrate_probability(model_prob):
    if model_prob == 0: return 0
    prior_train = 0.5; prior_real = 0.1
    model_odds = model_prob / (1 - model_prob)
    real_odds = (model_odds * (prior_real / (1 - prior_real)) / (prior_train / (1 - prior_train)))
    return real_odds / (1 + real_odds)

def get_db_connection():
    if not DB_PASS: return None
    try:
        encoded_user = quote_plus(DB_USER)
        encoded_pass = quote_plus(DB_PASS)
        conn_str = ""
        if DB_TYPE == "mysql": conn_str = f"mysql+pymysql://{encoded_user}:{encoded_pass}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
        elif DB_TYPE == "postgresql": conn_str = f"postgresql+psycopg2://{encoded_user}:{encoded_pass}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
        else: return None
        return create_engine(conn_str)
    except Exception as e:
        st.error(f"Connection Error: {e}"); return None

# ============================================
# 🔎 Search Logic
# ============================================
def fetch_patient_data(search_value, search_type='hn'):
    engine = get_db_connection()
    if not engine: return None
    regex_op = "~" if DB_TYPE == "postgresql" else "REGEXP"
    where_clause = "p.hn = :val" if search_type == 'hn' else "p.cid = :val"

    sql = text(f"""
    SELECT p.hn, p.cid, p.birthday, p.sex, p.marrystatus,
        (SELECT bps FROM opdscreen WHERE hn = p.hn ORDER BY vstdate DESC LIMIT 1) AS sbp,
        (SELECT bpd FROM opdscreen WHERE hn = p.hn ORDER BY vstdate DESC LIMIT 1) AS dbp, 
        (SELECT bw FROM opdscreen WHERE hn = p.hn ORDER BY vstdate DESC LIMIT 1) AS weight, 
        (SELECT height FROM opdscreen WHERE hn = p.hn AND height > 0 ORDER BY vstdate DESC LIMIT 1) AS height,
        (SELECT waist FROM opdscreen WHERE hn = p.hn AND waist > 0 ORDER BY vstdate DESC LIMIT 1) AS waist,
        (SELECT lab_order_result FROM lab_order lo JOIN lab_head lh ON lo.lab_order_number = lh.lab_order_number WHERE lh.hn = p.hn AND lo.lab_items_code IN ({FBS_CODES_SQL}) AND lo.lab_order_result {regex_op} '^[0-9]' ORDER BY lh.order_date DESC LIMIT 1) AS fbs,
        (SELECT lab_order_result FROM lab_order lo JOIN lab_head lh ON lo.lab_order_number = lh.lab_order_number WHERE lh.hn = p.hn AND lo.lab_items_code IN ({CHOL_CODES_SQL}) AND lo.lab_order_result {regex_op} '^[0-9]' ORDER BY lh.order_date DESC LIMIT 1) AS chol,
        (SELECT smoking_type_id FROM opdscreen WHERE hn = p.hn AND smoking_type_id IS NOT NULL ORDER BY vstdate DESC LIMIT 1) as smoke, 
        (SELECT drinking_type_id FROM opdscreen WHERE hn = p.hn AND drinking_type_id IS NOT NULL ORDER BY vstdate DESC LIMIT 1) as drink,
        EXISTS(SELECT 1 FROM ovstdiag WHERE hn = p.hn AND icd10 BETWEEN 'I20' AND 'I25') as has_cardio
    FROM patient p WHERE {where_clause} LIMIT 1
    """)
    try:
        with engine.connect() as conn:
            result = conn.execute(sql, {"val": search_value}).mappings().fetchone()
        if result:
            data = dict(result)
            def clean_lab(val):
                if not val: return 0.0
                nums = re.findall(r"[-+]?\d*\.\d+|\d+", str(val)); return float(nums[0]) if nums else 0.0
            data["fbs"] = clean_lab(data.get("fbs")); data["chol"] = clean_lab(data.get("chol"))
            if data.get("birthday"): data["age"] = int((pd.Timestamp.now() - pd.to_datetime(data["birthday"])).days / 365.25)
            else: data["age"] = 0
            return data
        return None
    except Exception as e: st.error(f"SQL Error: {e}"); return None

def send_to_google_sheet(hn, cid, age, risk_score, risk_level, input_data):
    if not ENABLE_CLOUD_LOG: return False
    try:
        # Prepare data with defaults to avoid 'undefined'
        safe_age = int(age) if age else 0
        
        # Flatten input data for easier column mapping if script allows
        # Also map numeric values back to readable strings for logging
        gender_str = "ชาย" if input_data.get("gender_numeric_map") == 0 else "หญิง"
        cardio_str = "มี" if input_data.get("cardio_numeric_map") == 1 else "ไม่มี"
        
        payload = {
            "hcode": str(HCODE),
            "hn": str(hn),
            # "cid": safe_cid, # Removed for privacy
            "age": safe_age, "Age": safe_age,
            "risk_score": float(risk_score),
            "risk_level": str(risk_level),
            
            # Flatten Details
            "gender": gender_str,
            "sbp": input_data.get("sbp", 0),
            "dbp": input_data.get("dbp", 0),
            "weight": input_data.get("weight", 0),
            "height": input_data.get("height", 0),
            "bmi": input_data.get("bmi", 0),
            "waist": input_data.get("waist", 0),
            "fbs": input_data.get("avg_glocose_level", 0),
            "chol": input_data.get("cholesterol", 0),
            "cardio": cardio_str,
            "smoke": input_data.get("smoking_status", 0),
            "drink": input_data.get("drinking_status", 0),
            
            "inputs": str(input_data)
        }
        requests.post(GOOGLE_SCRIPT_URL, json=payload, timeout=3); return True
    except: return False

def search_callback(source='hn'):
    search_val = st.session_state.get('search_hn_key', '').strip() if source == 'hn' else st.session_state.get('search_cid_key', '').strip()
    if search_val:
        data = fetch_patient_data(search_val, search_type=source)
        if data:
            st.session_state['search_msg'] = {'type': 'success', 'text': f"✅ พบข้อมูล ({source.upper()}: {search_val})"}
            st.session_state['form'].update({
                'hn_display': str(data.get('hn') or ''), 'cid': str(data.get('cid') or ''),
                'age': int(data.get('age') or 0), 'sbp': int(data.get('sbp') or 0), 'dbp': int(data.get('dbp') or 0),
                'weight': float(data.get('weight') or 0), 'height': float(data.get('height') or 0),
                'waist': int(data.get('waist') or 0), 'fbs': float(data.get('fbs') or 0), 'chol': float(data.get('chol') or 0),
                'gender_idx': 0 if str(data.get('sex')) == '1' else 1,
                'marry_idx': 0 if str(data.get('marrystatus')) in ['1','9','6'] else 1,
                'smoke_idx': 2 if str(data.get('smoke')) in ['3'] else (1 if str(data.get('smoke')) in ['2','5'] else 0),
                'cardio': bool(data.get('has_cardio'))
            })
        else:
            st.session_state['search_msg'] = {'type': 'error', 'text': f"❌ ไม่พบข้อมูลจาก {source.upper()}: {search_val}"}

# ============================================
# 🖥️ UI & Layout
# ============================================
st.title(f"🧠 Stroke Prediction: {HOSPITAL_NAME}")

if "form" not in st.session_state:
    st.session_state["form"] = {
        "hn_display": "", "age": 0, "sbp": 0, "dbp": 0, "weight": 0.0, "height": 0.0, "waist": 0,
        "fbs": 0.0, "chol": 0.0, "cid": "", "gender_idx": 0, "marry_idx": 0, "smoke_idx": 0, "drink_idx": 0, "cardio": False,
    }

# Search Bar
c_search1, c_search2, c_search3 = st.columns([2, 2, 1])
with c_search1: st.text_input("HN (Enter)", value=st.session_state['form']['hn_display'], key="search_hn_key", on_change=search_callback, args=('hn',))
with c_search2: st.text_input("CID (Enter)", value=st.session_state["form"]["cid"], key="search_cid_key", on_change=search_callback, args=('cid',))
with c_search3: st.write(""); st.write(""); st.button("📥 ดึงข้อมูล", on_click=search_callback, args=('hn',))

if 'search_msg' in st.session_state:
    msg = st.session_state['search_msg']
    if msg['type'] == 'success': st.success(msg['text'])
    else: st.error(msg['text'])
    del st.session_state['search_msg']

st.markdown("---")

# ============================================
# 🧪 Test Cases
# ============================================
# def set_test_case(data):
#     st.session_state["form"].update(data)
# 
# with st.expander("🛠️ Test Cases (กดเพื่อเติมข้อมูลทดสอบ)"):
#     t1, t2, t3, t4 = st.columns(4)
#     
#     with t1:
#         if st.button("🟢 Low Risk (วัยรุ่น)"):
#             set_test_case({
#                 "age": 25, "gender_idx": 0, "marry_idx": 0, "weight": 65.0, "height": 175.0,
#                 "sbp": 110, "dbp": 70, "waist": 78, "fbs": 85.0, "chol": 160.0,
#                 "smoke_idx": 0, "drink_idx": 0, "cardio": False
#             })
#             
#     with t2:
#         if st.button("🟡 Medium Risk (วัยทำงาน)"):
#             # Case จาก User: 30/ญ/38kg/SBP130/Chol280
#             set_test_case({
#                 "age": 30, "gender_idx": 1, "marry_idx": 0, "weight": 38.0, "height": 150.0,
#                 "sbp": 130, "dbp": 80, "waist": 63, "fbs": 100.0, "chol": 280.0,
#                 "smoke_idx": 0, "drink_idx": 0, "cardio": False
#             })
# 
#     with t3:
#         if st.button("🟠 High Risk (เริ่มมีอายุ)"):
#             set_test_case({
#                 "age": 55, "gender_idx": 0, "marry_idx": 1, "weight": 85.0, "height": 170.0,
#                 "sbp": 150, "dbp": 95, "waist": 102, "fbs": 130.0, "chol": 240.0,
#                 "smoke_idx": 1, "drink_idx": 2, "cardio": False 
#             })
# 
#     with t4:
#         if st.button("🔴 Very High (สูงอายุ)"):
#             set_test_case({
#                 "age": 72, "gender_idx": 0, "marry_idx": 1, "weight": 70.0, "height": 165.0,
#                 "sbp": 170, "dbp": 100, "waist": 95, "fbs": 180.0, "chol": 290.0,
#                 "smoke_idx": 2, "drink_idx": 0, "cardio": True
#             })

# 🟢 LAYOUT: แบ่ง 4 คอลัมน์ *ภายในฟอร์มเดียว* เพื่อแก้ปัญหา Nesting Error
# [Col1] [Col2] [Col3] | [Result_Col]
# ใช้ st.form ครอบทั้งหมด เพื่อให้จัด Layout ได้อิสระและปุ่มอยู่ใน Col2 ได้

# ตัวแปรสำหรับเก็บผลลัพธ์ (ประกาศไว้นอก Form ก่อน)
final_risk_score = 0
result_ready = False
clinical_boost = 0
bmi_calc = 0
prob_percent = 0
level_text = ""
input_data_pack = {}
submitted = False

with st.form("pred_form"):
    # สร้าง 4 คอลัมน์รวดเดียว (3 ช่องกรอก + 1 ช่องผลลัพธ์) ภายในฟอร์ม
    col1, col2, col3, res_col = st.columns([1, 1, 1, 1.3])
    
    # --- Col 1: ข้อมูลทั่วไป ---
    with col1:
        st.markdown("##### 👤 ทั่วไป")
        age = st.number_input("อายุ", value=st.session_state["form"]["age"])
        gender = st.selectbox("เพศ", ["ชาย", "หญิง"], index=st.session_state["form"]["gender_idx"])
        marry = st.selectbox("สถานะสมรส", ["โสด", "แต่งงาน"], index=st.session_state["form"]["marry_idx"])
        weight = st.number_input("น้ำหนัก (kg)", value=st.session_state["form"]["weight"])

    # --- Col 2: ร่างกาย + ปุ่ม ---
    with col2:
        st.markdown("##### 📏 ร่างกาย")
        sbp = st.number_input("SBP (บน)", value=st.session_state["form"]["sbp"])
        dbp = st.number_input("DBP (ล่าง)", value=st.session_state["form"]["dbp"])
        height = st.number_input("ส่วนสูง (cm)", value=st.session_state["form"]["height"])
        waist = st.number_input("รอบเอว (cm)", value=st.session_state["form"]["waist"])
        
        # ✅ ปุ่ม Submit อยู่ใน Col 2 ได้แล้ว (เพราะ Col นี้เกิดใน Form)
        st.write(""); st.write("")
        submitted = st.form_submit_button("🔮 ประเมินความเสี่ยง")

    # --- Col 3: แล็บ ---
    with col3:
        st.markdown("##### 🩸 แล็บ/อื่นๆ")
        fbs = st.number_input("FBS (น้ำตาล)", value=st.session_state["form"]["fbs"])
        chol = st.number_input("Chol. (ไขมัน)", value=st.session_state["form"]["chol"])
        smoke = st.selectbox("สูบบุหรี่", ["ไม่เคย", "เคยสูบ", "สูบปัจจุบัน"], index=st.session_state["form"]["smoke_idx"])
        drink = st.selectbox("ดื่มสุรา", ["ไม่ดื่ม", "นานๆ", "บางครั้ง", "ประจำ"], index=st.session_state["form"]["drink_idx"])
        st.write("") 
        cardio = st.checkbox("โรคหัวใจ (I20-25)?", value=st.session_state["form"]["cardio"])

    # --- ส่วนคำนวณ (Logic) ---
    # ต้องคำนวณภายใน Block ของ Form หรือส่งค่าออกไป
    if submitted:
        missing = []
        if age == 0: missing.append("อายุ")
        if sbp == 0: missing.append("SBP")
        if dbp == 0: missing.append("DBP")
        if weight == 0: missing.append("น้ำหนัก")
        if height == 0: missing.append("ส่วนสูง")
        if waist == 0: missing.append("รอบเอว")
        if fbs == 0: missing.append("FBS (น้ำตาล)")
        if chol == 0: missing.append("Cholesterol (ไขมัน)")
        
        if missing:
            st.error(f"⚠️ กรุณากรอก: {', '.join(missing)}")
        else:
            bmi_calc = weight / ((height / 100) ** 2)
            pred_age = age + 10
            input_data = {
                "gender_numeric_map": 0 if gender == "ชาย" else 1, "age": pred_age, "cardio_numeric_map": 1 if cardio else 0,
                "marry_status_numeric_map": 1 if marry == "แต่งงาน" else 0,
                "smoking_status": ["ไม่เคย", "เคยสูบ", "สูบปัจจุบัน"].index(smoke),
                "drinking_status": ["ไม่ดื่ม", "นานๆ", "บางครั้ง", "ประจำ"].index(drink),
                "sbp": sbp, "dbp": dbp, "weight": weight, "height": height, "bmi": bmi_calc, "waist": waist,
                "avg_glocose_level": fbs, "cholesterol": chol, "occupation": 1, "education": 4,
            }
            input_df = pd.DataFrame([input_data])[feature_names]
            raw_prob = model.predict_proba(input_df)[0][1]
            calibrated_prob = recalibrate_probability(raw_prob)
            prob_percent = calibrated_prob * 100

            clinical_boost = 0
            # if age > 60: clinical_boost += 10
            # if sbp >= 140: clinical_boost += 10
            # if sbp >= 160: clinical_boost += 10
            # if fbs >= 126: clinical_boost += 5
            # if smoke != "ไม่เคย": clinical_boost += 5
            # if drink != "ไม่ดื่ม": clinical_boost += 5
            # if bmi_calc >= 30: clinical_boost += 3
            # if cardio: clinical_boost += 20
            # 1. อายุ (Age) - ปัจจัยหลัก
            if pred_age >= 75:
                clinical_boost += 20
            elif pred_age >= 65:
                clinical_boost += 15
            elif pred_age >= 55:
                clinical_boost += 10
            elif pred_age >= 45:
                clinical_boost += 5
            else:
                clinical_boost += 2
            
            # 2. ความดันโลหิตซิสโตลิก (SBP) - สูงสุด 10 คะแนน
            if sbp >= 180:
                clinical_boost += 10  # Hypertensive crisis
            elif sbp >= 160:
                clinical_boost += 8   # Stage 2 HT
            elif sbp >= 140:
                clinical_boost += 6   # Stage 1 HT
            elif sbp >= 130:
                clinical_boost += 3   # Elevated (ปรับลดจาก 4)
            elif sbp >= 120:
                clinical_boost += 1   # (ปรับลดจาก 2)
            
            # 3. ความดันโลหิตไดแอสโตลิก (DBP) - สูงสุด 10 คะแนน
            if dbp >= 120:
                clinical_boost += 10
            elif dbp >= 100:
                clinical_boost += 8
            elif dbp >= 90:
                clinical_boost += 6
            elif dbp >= 85:
                clinical_boost += 4
            elif dbp >= 80:
                clinical_boost += 1   # (ปรับลดจาก 2)
            
            # 4. ระดับน้ำตาลในเลือด (Glucose) - สูงสุด 10 คะแนน
            if fbs >= 200:
                clinical_boost += 10  # DM ควบคุมไม่ได้
            elif fbs >= 126:
                clinical_boost += 8   # DM
            elif fbs >= 100:
                clinical_boost += 3   # Prediabetes (ปรับลดจาก 5)
            
            # 5. คอเลสเตอรอล (Total Cholesterol) - ปรับลดลงเล็กน้อย
            if chol >= 280:
                clinical_boost += 8   # (ปรับลดจาก 10)
            elif chol >= 240:
                clinical_boost += 6   # (ปรับลดจาก 8)
            elif chol >= 200:
                clinical_boost += 4   # (ปรับลดจาก 6)
            elif chol >= 180:
                clinical_boost += 2   # (ปรับลดจาก 3)

            # 6. การสูบบุหรี่ (Smoking) - ปัจจัยเสี่ยงสูง
            if smoke == "สูบปัจจุบัน":
                clinical_boost += 10
            elif smoke == "เคยสูบ":
                clinical_boost += 5

            # 7. ดัชนีมวลกาย (BMI)
            if bmi_calc >= 30:
                clinical_boost += 5  # Obesity
            elif bmi_calc >= 25:
                clinical_boost += 2  # Overweight

            # 8. ประวัติโรคหัวใจ (Cardio History) - ความเสี่ยงสูงมาก
            if cardio: 
                clinical_boost += 15                

            if model is None:
                st.error("⛔ Model is missing. Cannot predict.")
                st.stop()

            # ปรับสูตรคำนวณใหม่ (Dynamic Dampening)
            # 1. AI ต่ำ (<5%): เชื่อ Clinical เต็มที่ (x1.0) เพื่อเตือนกลุ่มเสี่ยง Lifestyle
            # 2. AI กลาง (5-30%): ลดทอนครึ่งนึง (x0.5) กันซ้ำซ้อน
            # 3. AI สูง (>30%): ลดทอนเยอะๆ (x0.2) เพราะ AI จับได้แล้ว
            
            boost_factor = 0.5
            if prob_percent < 5:
                boost_factor = 1.0
            elif prob_percent > 30:
                boost_factor = 0.2
            
            # final_risk_score = min(prob_percent + (clinical_boost * boost_factor), 99.9)
            final_risk_score = prob_percent
            result_ready = True
            input_data_pack = input_data

    # --- Col 4: แสดงผลลัพธ์ (Result) ---
    # เขียนใส่ res_col ที่จองไว้ตั้งแต่ต้น (ยังอยู่ใน Form แต่แสดงผลได้ปกติ)
    with res_col:
        st.markdown(f"### 📊 ผลลัพธ์ (อีก 10 ปี / อายุ {age+10})")
        
        if result_ready:
            # Gauge Chart
            fig = go.Figure(go.Indicator(
                mode="gauge+number",
                value=final_risk_score,
                domain={"x": [0, 1], "y": [0, 1]},
                gauge={
                    "axis": {"range": [0, 100]},
                    "bar": {"color": "rgba(0,0,0,0)"},
                    "steps": [
                        {"range": [0, 10], "color": "#28a745"},
                        {"range": [10, 20], "color": "#ffc107"},
                        {"range": [20, 30], "color": "#fd7e14"},
                        {"range": [30, 40], "color": "#dc3545"},
                        {"range": [40, 100], "color": "#8b0000"},
                    ],
                    "threshold": {"line": {"color": "black", "width": 5}, "thickness": 0.8, "value": final_risk_score},
                },
            ))
            fig.update_layout(height=220, margin=dict(l=10, r=10, t=0, b=0))
            st.plotly_chart(fig, use_container_width=True)

            # Advice Text
            if final_risk_score < 10:
                st.success("✅ **ความเสี่ยงต่ำ (Low Risk):** รักษาสุขภาพดีเยี่ยม หมั่นตรวจประจำปี")
                level_text = "Low"
            elif final_risk_score < 20:
                st.warning("⚠️ **ความเสี่ยงปานกลาง (Medium Risk):** ควรเริ่มคุมอาหาร ออกกำลังกาย และติดตามความดัน")
                level_text = "Medium"
            elif final_risk_score < 30:
                st.markdown('<div style="background-color:#ffeeba; padding:10px; border-radius:5px; border-left:5px solid #fd7e14; color:#fd7e14;"><b>🟠 ความเสี่ยงสูง (High Risk)</h4><p>ควรปรึกษาแพทย์เพื่อปรับเปลี่ยนพฤติกรรม และพิจารณาการรักษา</div>', unsafe_allow_html=True)
                level_text = "High"
            elif final_risk_score < 40:
                st.error("🚨 **ความเสี่ยงสูงมาก (Very High Risk):** ต้องพบแพทย์ด่วนเพื่อรับยาและติดตามอย่างใกล้ชิด!")
                level_text = "Very High"
            else:
                st.error("💀 **สูงอันตราย (Dangerous Risk):** เสี่ยงต่อการเกิดโรคหลอดเลือดสมองสูงมาก ต้องได้รับการดูแลทันที!")
                level_text = "Dangerous"
                
            # st.caption(f"BMI: {bmi_calc:.1f} | AI: {prob_percent:.1f}% | Clinical: {clinical_boost:.1f}%")
            st.caption(f"BMI: {bmi_calc:.1f}")

            # Send Cloud (No CID sent inside function)
            current_hn = st.session_state["form"]["hn_display"] or st.session_state.get('search_hn_key', '')
            current_cid = st.session_state["form"]["cid"]
            
            # send_to_google_sheet(hn, cid, current_age, pred_age, risk_score, level, input_data, bmi)
            send_to_google_sheet(current_hn, current_cid, age, final_risk_score, level_text, input_data_pack)
            
        else:
            # Placeholder
            st.info("รอผลการประเมิน...")
            st.markdown("""
            <div style="text-align: center; color: gray; font-size: 0.8em;">
                กรอกข้อมูลด้านซ้าย<br>แล้วกดปุ่มประเมิน
            </div>
            """, unsafe_allow_html=True)