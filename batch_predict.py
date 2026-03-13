
import pandas as pd
import numpy as np
import joblib
import configparser
import requests
import re
import os
import sys
import time
import logging
import traceback
import threading
import queue
from datetime import datetime
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus

# GUI Imports
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

# Try to import DateEntry for Calendar Picker
try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None

# ============================================
# ⚙️ CONFIG & GLOBALS
# ============================================

# GLOBAL QUEUE (Must be defined before usage)
gui_queue = queue.Queue()

def get_path(filename):
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

# Logging Setup
log_file = get_path("batch_log.txt")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s'
)

# Encrypted Password Key
from cryptography.fernet import Fernet

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

# Config Variables
CONFIG_FILE = "config.ini"
DB_HOST = ""; DB_PORT = ""; DB_USER = ""; DB_PASS = ""; DB_NAME = ""
DB_TYPE = "mysql"
HCODE = ""; HOSPITAL_NAME = ""
GOOGLE_SCRIPT_URL = ""
CENTRAL_CONFIG_URL = ""
FBS_CODES_SQL = "'76', '1698'"
CHOL_CODES_SQL = "'102', '1691'"

# --- DATE RANGE GLOBALS ---
# Default to "Yesterday"
yest = (datetime.now() - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
SEARCH_DATE_START = yest
SEARCH_DATE_END = yest

# Load Config
config_path = get_path(CONFIG_FILE)
config = configparser.ConfigParser()

try:
    if os.path.exists(config_path):
        config.read(config_path, encoding='utf-8')
        
        if "Database" in config:
            DB_TYPE = config["Database"].get("db_type", "mysql").lower()
            DB_HOST = config["Database"].get("host", "")
            DB_PORT = config["Database"].get("port", "")
            DB_USER = config["Database"].get("username", "") # ini uses 'username' often
            if not DB_USER: DB_USER = config["Database"].get("user", "")
            
            DB_NAME = config["Database"].get("database", "")
            if not DB_NAME: DB_NAME = config["Database"].get("db_name", "")
            
            # Decrypt Password
            raw_pass = config["Database"].get("password", "")
            try:
                # Try to decrypt assuming it's a token
                DB_PASS = cipher_suite.decrypt(raw_pass.encode()).decode()
            except:
                # Fallback to raw if decryption fails (maybe plain text)
                DB_PASS = raw_pass

        if "General" in config:
            HOSPITAL_NAME = config["General"].get("hospital_name", "")
            HCODE = config["General"].get("hospital_code", "")
        
        if "Cloud" in config:
            GOOGLE_SCRIPT_URL = config["Cloud"].get("google_script_url", "")
            CENTRAL_CONFIG_URL = config["Cloud"].get("central_config_url", "")
            
        if "LabCodes" in config:
            raw_fbs = config["LabCodes"].get("fbs_codes", "76, 1698")
            raw_chol = config["LabCodes"].get("chol_codes", "102, 1691")
            
            def parse_lab_codes(raw_str):
                codes = [c.strip() for c in raw_str.split(',')]
                return ", ".join([f"'{c}'" for c in codes])

            FBS_CODES_SQL = parse_lab_codes(raw_fbs)
            CHOL_CODES_SQL = parse_lab_codes(raw_chol)
            
except Exception as e:
    logging.error(f"Config Error: {e}")

# ============================================
# 🛠️ HELPER FUNCTIONS
# ============================================

def log_to_gui(msg, level="INFO"):
    """ Custom Logger to send text to GUI Queue """
    timestamp = datetime.now().strftime("%H:%M:%S")
    prefix = ""
    if level == "INFO": prefix = "[INFO]"
    elif level == "WARN": prefix = "[WARN]"
    elif level == "ERROR": prefix = "[ERROR]"
    elif level == "CRITICAL": prefix = "[FATAL]"
    elif level == "SUCCESS": prefix = "[DONE]"
    
    clean_msg = f"[{timestamp}] {prefix} {msg}"
    
    # Also log to file matching level
    if level == "ERROR": logging.error(msg)
    elif level == "CRITICAL": logging.critical(msg)
    else: logging.info(msg)
    
    print(clean_msg) # Console
    if gui_queue:
        gui_queue.put(("log", clean_msg))

def load_ai_model():
    try:
        model_path = get_path("stroke_model.pkl")
        model = joblib.load(model_path)
        # Fix: Sync with actual trained model feature order (from old version)
        features = ['age', 'gender_numeric_map', 'marry_status_numeric_map', 'sbp', 'dbp', 'weight', 'bmi', 'waist', 'avg_glocose_level', 'cholesterol', 'smoking_status', 'drinking_status', 'occupation', 'education', 'cardio_numeric_map']
        return model, features
    except Exception as e:
        log_to_gui(f"Model Load Error: {e}", "ERROR")
        return None, None

def recalibrate_probability(model_prob):
    if model_prob == 0: return 0
    prior_train = 0.5; prior_real = 0.1
    model_odds = model_prob / (1 - model_prob)
    real_odds = (model_odds * (prior_real / (1 - prior_real)) / (prior_train / (1 - prior_train)))
    return real_odds / (1 + real_odds)

# ============================================
# 🛠️ DB & DATA HELPERS
# ============================================

def fetch_remote_db_config(url):
    try:
        if not url: return None
        log_to_gui(f"Fetching DB Config from Cloud...", "INFO")
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            config_data = {}
            for line in response.text.splitlines():
                if ',' in line:
                    parts = line.split(',')
                    if len(parts) >= 2:
                        key = parts[0].strip().lower()
                        val = parts[1].strip()
                        config_data[key] = val
            return config_data
    except Exception as e:
        logging.error(f"Remote Config Error: {e}")
        return None
    return None

def get_engine():
    if not DB_PASS: return None
    try:
        port_txt = f":{DB_PORT}" if DB_PORT else ""
        
        if DB_TYPE == 'postgresql':
            connection_str = f"postgresql+psycopg2://{DB_USER}:{quote_plus(DB_PASS)}@{DB_HOST}{port_txt}/{DB_NAME}"
        else:
            connection_str = f"mysql+pymysql://{DB_USER}:{quote_plus(DB_PASS)}@{DB_HOST}{port_txt}/{DB_NAME}"
        
        engine = create_engine(connection_str)
        return engine
    except Exception as e:
        log_to_gui(f"DB Engine Error: {e}", "ERROR")
        return None

# Helper to handle '@' and other special chars in passwords
from urllib.parse import unquote_plus

def fix_connection_string(conn_str):
    """ Helper to handle '@' and other special chars in passwords """
    if not conn_str: return ""
    try:
        # Pattern: dialect+driver://user:pass@host:port/db
        if '://' not in conn_str: return conn_str
        
        prefix, rest = conn_str.split("://", 1)
        
        # Helper: Split from the RIGHT (Last @ is usually the host separator)
        if '@' in rest:
            credentials_part, host_part = rest.rsplit('@', 1)
            
            # credential_part = user:pass
            if ':' in credentials_part:
                user, raw_pass = credentials_part.split(':', 1)
                
                # Logic: Ensure password is properly URL-encoded
                # 1. Unquote first (in case it was already encoded)
                # 2. Quote it properly
                safe_pass = quote_plus(unquote_plus(raw_pass))
                
                new_conn = f"{prefix}://{user}:{safe_pass}@{host_part}"
                return new_conn
                     
    except Exception as e:
        log_to_gui(f"Conn String Fix Warn: {e}", "WARN")
        return conn_str # Fallback to original if error
        
    return conn_str

def fetch_valid_patients_iterator(chunk_size=1000, progress_callback=None):
    """ Fetch patients based on SEARCH_DATE_START and SEARCH_DATE_END globals """
    engine = get_engine()
    if not engine: return
    
    # Date Range Logic
    s_date = SEARCH_DATE_START
    e_date = SEARCH_DATE_END
    log_to_gui(f"ช่วงวันที่: {s_date} ถึง {e_date}")
    
    # 1. SQL Syntax Adapters
    if DB_TYPE == 'postgresql':
        regex_op = "~"
        # Cast vstdate to DATE for comparison
        date_condition = f"vstdate::DATE BETWEEN '{s_date}' AND '{e_date}'"
    else:
        regex_op = "REGEXP"
        date_condition = f"DATE(vstdate) BETWEEN '{s_date}' AND '{e_date}'"

    # 2. Base SQL (Adapted from Old Version - Subquery Style)
    # Note: Age is calculated in Python to avoid SQL Compat issues
    base_sql = f"""
            SELECT * FROM (
            SELECT p.hn, p.cid, p.birthday, p.sex as gender, p.marrystatus,
            (SELECT vstdate FROM opdscreen WHERE hn = p.hn ORDER BY vstdate DESC LIMIT 1) AS last_visit_date,
            CAST((SELECT bps FROM opdscreen WHERE hn = p.hn ORDER BY vstdate DESC LIMIT 1) AS DECIMAL(12,2)) AS sbp,
            CAST((SELECT bpd FROM opdscreen WHERE hn = p.hn ORDER BY vstdate DESC LIMIT 1) AS DECIMAL(12,2)) AS dbp,
            CAST((SELECT bw FROM opdscreen WHERE hn = p.hn ORDER BY vstdate DESC LIMIT 1) AS DECIMAL(12,2)) AS weight,
            CAST((SELECT height FROM opdscreen WHERE hn = p.hn AND height > 0 ORDER BY vstdate DESC LIMIT 1) AS DECIMAL(12,2)) AS height,
            CAST((SELECT waist FROM opdscreen WHERE hn = p.hn AND waist > 0 ORDER BY vstdate DESC LIMIT 1) AS DECIMAL(12,2)) AS waist,
            CAST((SELECT lab_order_result FROM lab_order lo JOIN lab_head lh ON lo.lab_order_number = lh.lab_order_number WHERE lh.hn = p.hn AND lo.lab_items_code IN ({FBS_CODES_SQL}) AND lo.lab_order_result {regex_op} '^[0-9]' ORDER BY lh.order_date DESC LIMIT 1) AS DECIMAL(12,2)) AS fbs,
            CAST((SELECT lab_order_result FROM lab_order lo JOIN lab_head lh ON lo.lab_order_number = lh.lab_order_number WHERE lh.hn = p.hn AND lo.lab_items_code IN ({CHOL_CODES_SQL}) AND lo.lab_order_result {regex_op} '^[0-9]' ORDER BY lh.order_date DESC LIMIT 1) AS DECIMAL(12,2)) AS chol,
            (SELECT smoking_type_id FROM opdscreen WHERE hn = p.hn AND smoking_type_id IS NOT NULL ORDER BY vstdate DESC LIMIT 1) as smoke,
            (SELECT drinking_type_id FROM opdscreen WHERE hn = p.hn AND drinking_type_id IS NOT NULL ORDER BY vstdate DESC LIMIT 1) as drink,
            (SELECT count(*) FROM ovstdiag WHERE hn = p.hn AND (icd10 BETWEEN 'I20' AND 'I25')) as cardio_history
            FROM patient p
            JOIN (
                select distinct o.hn
                from ovstdiag o
                where o.icd10 in (select code
                from icd101
                where code3 between 'I10' and 'I15')
                and {date_condition}
            ) active ON p.hn = active.hn
        ) a
        WHERE a.sbp IS NOT NULL AND a.sbp > 0
        AND a.weight IS NOT NULL AND a.weight > 0
        AND a.height IS NOT NULL AND a.height > 0
        AND a.waist IS NOT NULL AND a.waist > 0
        AND a.fbs IS NOT NULL AND a.fbs > 0
        AND a.chol IS NOT NULL AND a.chol > 0
        """
    
    try:
        with engine.connect() as conn:
            # Count First
            count_sql = f"SELECT COUNT(*) FROM ({base_sql}) as subquery"
            try:
                total_rows = conn.execute(text(count_sql)).scalar()
            except Exception as e:
                log_to_gui(f"เกิดข้อผิดพลาดในการนับจำนวน: {e}", "WARN")
                total_rows = 1000

            log_to_gui(f"กำลังดึงข้อมูล (ประมาณ {total_rows} รายการ)...")
            
            offset = 0
            while True:
                # Use LIMIT/OFFSET for paging
                final_sql = f"{base_sql} LIMIT {chunk_size} OFFSET {offset}"
                
                # Fetch as mappings
                result = conn.execute(text(final_sql)).mappings().fetchall()
                
                if not result:
                    break
                
                clean_chunk = []
                for row in result:
                    data = dict(row)
                    
                    # --- CLEANING LOGIC (From Old Version) ---
                    def clean_lab(val):
                        if not val: return 0.0
                        try:
                            # Extract number from string (e.g. "120 mg%")
                            nums = re.findall(r"[-+]?\d*\.\d+|\d+", str(val))
                            return float(nums[0]) if nums else 0.0
                        except: return 0.0

                    # Apply cleaning
                    data["fbs"] = clean_lab(data.get("fbs"))
                    data["chol"] = clean_lab(data.get("chol"))
                    
                    # Compute Age in Python
                    if data.get("birthday"): 
                        try:
                            data["age"] = int((pd.Timestamp.now() - pd.to_datetime(data["birthday"])).days / 365.25)
                        except:
                            data["age"] = 0
                    else: 
                        data["age"] = 0
                        
                    clean_chunk.append(data)
                
                # Yield DataFrame
                if clean_chunk:
                    yield pd.DataFrame(clean_chunk)
                    
                    # Update Progress (0-50% reserved for Fetching)
                    offset += chunk_size
                    if total_rows > 0 and progress_callback:
                        prog = min((offset / total_rows) * 50, 50) 
                        progress_callback(prog)
                
                if len(result) < chunk_size:
                    break
                    
    except Exception as e:
        log_to_gui(f"ดึงข้อมูลล้มเหลว: {e}", "ERROR")

# ============================================
# 🧠 MAIN PROCESS LOGIC
# ============================================
def process_data(update_progress_callback=None):
    """ Main Business Logic """
    try:
        # Load Model
        log_to_gui("กำลังโหลดโมเดล AI...")
        model, FEATURES = load_ai_model()
        
        if not model:
            log_to_gui("ไม่พบไฟล์โมเดล AI ยกเลิกการทำงาน", "ERROR")
            return

        # Prepare for processing
        log_to_gui(f"เริ่มกระบวนการสำหรับรหัส: {HCODE}")
        
        chunk_idx = 0
        batch_payloads = []
        
        # Iterate Chunks (Fetch uses 0-50% progress)
        for chunk in fetch_valid_patients_iterator(chunk_size=1000, progress_callback=update_progress_callback):
            chunk_idx += 1
            log_to_gui(f"กำลังประมวลผลชุดที่ {chunk_idx} (จำนวน {len(chunk)} ราย)...")
            
            # Note: chunk is already a DataFrame from fetch_valid_patients_iterator
            
            for _, raw_data in chunk.iterrows():
                try:
                    # --- LOGIC COPIED FROM OLD VERSION ---
                    age = int(raw_data.get('age', 0))
                    if age == 0: continue

                    sbp = float(raw_data.get('sbp') or 0)
                    dbp = float(raw_data.get('dbp') or 0)
                    weight = float(raw_data.get('weight') or 0)
                    height = float(raw_data.get('height') or 0)
                    waist = float(raw_data.get('waist') or 0)
                    fbs = float(raw_data.get('fbs') or 0)
                    chol = float(raw_data.get('chol') or 0)
                    
                    # Gender map
                    gender_code = str(raw_data.get('gender')) 
                    gender_val = 1 # Default Female
                    if gender_code in ['1', 'ชาย', 'Male', '0', 'male']: gender_val = 0
                    
                    # Marry
                    marry_code = str(raw_data.get('marrystatus'))
                    is_married = 0 if marry_code in ['1', '9', '6'] else 1
                    
                    # Smoke
                    smoke_code = str(raw_data.get('smoke'))
                    smoke_idx = 0
                    if smoke_code in ['3', '4']: smoke_idx = 2
                    elif smoke_code in ['2', '5']: smoke_idx = 1
                    
                    # Drink
                    drink_code = str(raw_data.get('drink'))
                    drink_idx = 0 
                    if drink_code == '4': drink_idx = 3 
                    elif drink_code == '3': drink_idx = 2 
                    elif drink_code == '2': drink_idx = 1 
        
                    cardio = raw_data.get('cardio_history', 0) > 0 
                    
                    bmi_calc = 0
                    if height > 0:
                        bmi_calc = weight / ((height / 100) ** 2)
                    
                    pred_age = age + 10 # Old code logic
                    
                    # --- INPUT DATA EXACTLY MATCHING FEATURES ---
                    input_data = {
                        "gender_numeric_map": gender_val, 
                        "age": pred_age, 
                        "cardio_numeric_map": 1 if cardio else 0,
                        "marry_status_numeric_map": is_married,
                        "smoking_status": smoke_idx,
                        "drinking_status": drink_idx,
                        "sbp": sbp, "dbp": dbp, "weight": weight, "bmi": bmi_calc, "waist": waist,
                        "avg_glocose_level": fbs, "cholesterol": chol, 
                        "occupation": 1, "education": 4, 
                    }
                    
                    # Create DataFrame for single prediction
                    input_df = pd.DataFrame([input_data])[FEATURES]

                    # Predict
                    model_prob = model.predict_proba(input_df)[0][1]
                    calibrated_prob = recalibrate_probability(model_prob)
                    prob_percent = calibrated_prob * 100
                    final_risk_score = prob_percent 

                    risk_level = "Dangerous"
                    if final_risk_score < 10: risk_level = "Low"
                    elif final_risk_score < 20: risk_level = "Medium"
                    elif final_risk_score < 30: risk_level = "High"
                    elif final_risk_score < 40: risk_level = "Very High"

                    gender_str = "ชาย" if gender_val == 0 else "หญิง"
                    cardio_str = "มี" if cardio else "ไม่มี"
                    
                    payload = {
                        "hcode": str(HCODE),
                        # "hospital": str(HOSPITAL_NAME), # Removed: Not in DB schema
                        "hn": str(raw_data.get('hn')),
                        "age": age, 
                        "risk_score": round(final_risk_score, 2),
                        "risk_level": str(risk_level),
                        "gender": gender_str,
                        "weight": float(weight),
                        "height": float(height),
                        "bmi": round(bmi_calc, 1),
                        "sbp": int(sbp), 
                        "dbp": int(dbp), 
                        "fbs": int(fbs), 
                        "chol": int(chol),
                        "waist": int(waist),
                        "smoke": int(smoke_idx), 
                        "drink": int(drink_idx), 
                        "cardio": cardio_str, 
                        "pred_age_10y": pred_age,
                        "visit_date": str(raw_data.get('last_visit_date') or ""),
                        "last_update": datetime.now()
                    }
                    batch_payloads.append(payload)
                except Exception as e:
                    pass

        # --- UPLOAD PHASE ---
        if batch_payloads:
            log_to_gui(f"ประมวลผลเสร็จสิ้น เตรียมส่งข้อมูล {len(batch_payloads)} รายการ", "SUCCESS")
            
            # --- UPLOAD TO MYSQL (ENCRYPTED + UPSERT) ---
            db_config = fetch_remote_db_config(CENTRAL_CONFIG_URL)
            
            if db_config and 'connection_string' in db_config:
                target_engine = None
                try:
                    log_to_gui(f"กำลังถอดรหัสข้อมูลเชื่อมต่อ...")
                    encrypted_conn = db_config['connection_string']
                    decrypted_conn = cipher_suite.decrypt(encrypted_conn.encode()).decode()
                    safe_conn_str = fix_connection_string(decrypted_conn)
                    target_engine = create_engine(safe_conn_str)
                except Exception as e:
                    log_to_gui(f"การเตรียมการเชื่อมต่อล้มเหลว: {e}", "ERROR")

                if target_engine:
                    tbl = db_config.get('table_name', 'center_db')
                    log_to_gui(f"กำลังนำข้อมูลเข้าตาราง: {tbl} (โหมดอัปเดต)...")
                    
                    try:
                        with target_engine.connect() as conn:
                            # --- BATCH OPTIMIZED UPSERT ---
                            batch_size_sql = 500
                            total = len(batch_payloads)
                            count_sent = 0
                            
                            # Prepare Data Lists first
                            data_values = []
                            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            for p in batch_payloads:
                                v_date_str = p.get('visit_date', '')
                                if not v_date_str: v_date_str = datetime.now().strftime("%Y-%m-%d")
                                try:
                                    v_dt = pd.to_datetime(v_date_str)
                                    month_key = v_dt.strftime("%Y-%m")
                                    v_date_fmt = v_dt.strftime("%Y-%m-%d")
                                except:
                                    month_key = datetime.now().strftime("%Y-%m")
                                    v_date_fmt = datetime.now().strftime("%Y-%m-%d")
                                
                                # Escape values for raw SQL string (simple safety)
                                def esc(v):
                                    if v is None: return "NULL"
                                    # Fix SyntaxError by moving replace out or using different quotes
                                    val = str(v).replace("'", "''")
                                    return f"'{val}'"

                                row_val = (
                                    esc(p['hcode']), esc(p['hn']), esc(v_date_fmt), esc(month_key),
                                    str(int(p['age'])), esc(p['gender']),
                                    str(float(p['weight'])), str(float(p['height'])), str(float(p['bmi'])),
                                    str(int(p['sbp'])), str(int(p['dbp'])), str(int(p['fbs'])), str(int(p['chol'])), str(float(p['waist'])),
                                    esc(p['smoke']), esc(p['drink']), esc(p['cardio']),
                                    str(float(p['risk_score'])), esc(p['risk_level']), str(int(p['pred_age_10y']))
                                )
                                data_values.append(f"({','.join(row_val)})")

                            # Execute in Chunks
                            for i in range(0, total, batch_size_sql):
                                chunk_vals = data_values[i:i + batch_size_sql]
                                values_str = ",".join(chunk_vals)
                                
                                # Use simple f-string with verifiable indentation
                                final_sql = (
                                    f"INSERT INTO {tbl} "
                                    f"(hcode, hn, visit_date, month_key, age, gender, "
                                    f"weight, height, bmi, sbp, dbp, fbs, chol, waist, "
                                    f"smoke, drink, cardio, "
                                    f"risk_score, risk_level, pred_age_10y) "
                                    f"VALUES {values_str} "
                                    f"ON DUPLICATE KEY UPDATE "
                                    f"visit_date = VALUES(visit_date), "
                                    f"age = VALUES(age), "
                                    f"gender = VALUES(gender), "
                                    f"weight = VALUES(weight), "
                                    f"height = VALUES(height), "
                                    f"bmi = VALUES(bmi), "
                                    f"sbp = VALUES(sbp), "
                                    f"dbp = VALUES(dbp), "
                                    f"fbs = VALUES(fbs), "
                                    f"chol = VALUES(chol), "
                                    f"waist = VALUES(waist), "
                                    f"smoke = VALUES(smoke), "
                                    f"drink = VALUES(drink), "
                                    f"cardio = VALUES(cardio), "
                                    f"risk_score = VALUES(risk_score), "
                                    f"risk_level = VALUES(risk_level), "
                                    f"pred_age_10y = VALUES(pred_age_10y), "
                                    f"updated_at = NOW()"
                                )
                                
                                conn.execute(text(final_sql))
                                count_sent += len(chunk_vals)
                                
                                if update_progress_callback:
                                     # Scale 50-100% (Upload Phase)
                                     progress = 50 + ((count_sent / total) * 50)
                                     update_progress_callback(progress)
                            
                            conn.commit()
                            log_to_gui(f"สำเร็จ! บันทึกข้อมูลเรียบร้อยแล้ว {count_sent} รายการ", "SUCCESS")
                            
                    except Exception as db_e:
                        log_to_gui(f"เกิดข้อผิดพลาดในการอัปโหลด: {db_e}", "ERROR")
            else:
                 log_to_gui("ไม่พบการตั้งค่าฐานข้อมูลปลายทาง ไม่ได้ส่งข้อมูล", "WARN")
                
        else:
            log_to_gui("ไม่พบข้อมูลสำหรับประมวลผล", "WARN")

    except Exception as fatal_e:
        log_to_gui(f"ข้อผิดพลาดร้ายแรง: {fatal_e}", "CRITICAL")
        log_to_gui(traceback.format_exc())

# ============================================
# 🖥️ GUI APPLICATION
# ============================================

import base64
from io import BytesIO
from tkinter import PhotoImage

import ctypes

class StrokeBatchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # 1. Set App ID for Windows Taskbar Icon (Fixes "feather" icon issue)
        try:
            myappid = 'myorg.strokeagent.batch.1.0' # arbitrary string
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except:
            pass

        self.title(f"{HOSPITAL_NAME} - ระบบ AI ประเมินความเสี่ยง Stroke")
        self.geometry("400x580") # Compact Size
        self.resizable(False, False)
        
        # Style
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure("TFrame", background="#f0f2f5")
        self.style.configure("TLabel", background="#f0f2f5", font=("Segoe UI", 9))
        self.style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"), foreground="#2c3e50")
        self.style.configure("Status.TLabel", font=("Segoe UI", 8), foreground="#7f8c8d")
        self.style.configure("TLabelframe", background="#f0f2f5", relief="flat")
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 9, "bold"), foreground="#34495e", background="#f0f2f5")
        
        # Custom Icon (File)
        self.setup_icon()
        
        # Variables
        self.date_start_var = tk.StringVar(value=(datetime.now() - pd.Timedelta(days=1)).strftime("%Y-%m-%d"))
        self.date_end_var = tk.StringVar(value=(datetime.now() - pd.Timedelta(days=1)).strftime("%Y-%m-%d"))
        
        self.agent_h_var = tk.StringVar(value="08")
        self.agent_m_var = tk.StringVar(value="30")
        self.agent_running = False
        self.stop_event = threading.Event()
        
        # Layout
        self.configure(bg="#f0f2f5")
        self.create_widgets()
        self.process_queue()

    def setup_icon(self):
        try:
            # Load from file (artificial-intelligence.ico)
            icon_path = get_path("artificial-intelligence.ico")
            
            # Set Window Icon
            self.iconbitmap(icon_path)
            
            # Force Taskbar Icon (sometimes needed)
            # self.wm_iconbitmap(icon_path) 
        except Exception as e:
            pass # Fallback to default Tk icon if missing

    def create_widgets(self):
        # 1. Header Area
        header_frame = ttk.Frame(self, padding="20 15 20 10")
        header_frame.pack(fill=tk.X)
        
        ttk.Label(header_frame, text=f"{HOSPITAL_NAME}", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header_frame, text=f"รหัสหน่วยบริการ: {HCODE}", style="Status.TLabel").pack(anchor="w")
        
        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=20, pady=(5,15))

        # 2. Manual Zone
        manual_frame = ttk.Labelframe(self, text=" สั่งทำงานด้วยตนเอง (Manual) ", padding=15)
        manual_frame.pack(fill=tk.X, padx=20, pady=5)
        
        # Date Inputs
        date_grid = ttk.Frame(manual_frame)
        date_grid.pack(fill=tk.X)
        
        ttk.Label(date_grid, text="วันที่เริ่ม:").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        if DateEntry:
            self.date_start = DateEntry(date_grid, width=12, background='#2980b9', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
            self.date_start.set_date(datetime.now() - pd.Timedelta(days=1))
        else:
            self.date_start = ttk.Entry(date_grid, textvariable=self.date_start_var, width=12)
        self.date_start.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(date_grid, text="ถึงวันที่:").grid(row=0, column=2, padx=5, pady=2, sticky='w')
        if DateEntry:
            self.date_end = DateEntry(date_grid, width=12, background='#2980b9', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
            self.date_end.set_date(datetime.now() - pd.Timedelta(days=1))
        else:
            self.date_end = ttk.Entry(date_grid, textvariable=self.date_end_var, width=12)
        self.date_end.grid(row=0, column=3, padx=5, pady=2)
        
        self.btn_run = ttk.Button(manual_frame, text="เริ่มประมวลผลทันที", command=self.start_manual_run)
        self.btn_run.pack(fill=tk.X, pady=(15, 0), ipady=3)

        # 3. Agent Zone
        agent_frame = ttk.Labelframe(self, text=" ตั้งค่าระบบอัตโนมัติ (Auto Agent) ", padding=15)
        agent_frame.pack(fill=tk.X, padx=20, pady=15)
        
        time_frame = ttk.Frame(agent_frame)
        time_frame.pack(fill=tk.X, pady=5)
        ttk.Label(time_frame, text="เวลาทำงาน (ทุกวัน):").pack(side=tk.LEFT)
        
        # Comboboxes for Time
        hours = [f"{i:02d}" for i in range(24)]
        minutes = [f"{i:02d}" for i in range(0, 60, 5)] # 00, 05, 10...
        
        self.cb_h = ttk.Combobox(time_frame, textvariable=self.agent_h_var, values=hours, width=3, state="readonly")
        self.cb_h.pack(side=tk.LEFT, padx=(10,2))
        ttk.Label(time_frame, text=":").pack(side=tk.LEFT)
        self.cb_m = ttk.Combobox(time_frame, textvariable=self.agent_m_var, values=minutes, width=3, state="readonly")
        self.cb_m.pack(side=tk.LEFT, padx=2)
        
        self.agent_btn = ttk.Button(agent_frame, text="เปิดระบบอัตโนมัติ", command=self.toggle_agent)
        self.agent_btn.pack(fill=tk.X, pady=(15, 0), ipady=3)
        
        self.agent_status_lbl = ttk.Label(agent_frame, text="สถานะ: ปิดใช้งาน", style="Status.TLabel")
        self.agent_status_lbl.pack(anchor="c", pady=(5,0))

        # 4. Terminal Log
        log_frame = ttk.Frame(self, padding="20 0 20 10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, state='disabled', 
                                              font=("Consolas", 8), bg="#ffffff", fg="#2d3436", relief="flat")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 5. Progress Bar
        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.pack(fill=tk.X, side=tk.BOTTOM)

    def start_manual_run(self):
        if self.agent_running:
            messagebox.showwarning("ระบบไม่ว่าง", "ระบบอัตโนมัติทำงานอยู่ กรุณาปิดระบบอัตโนมัติก่อน")
            return
            
        global SEARCH_DATE_START, SEARCH_DATE_END

        if DateEntry and isinstance(self.date_start, DateEntry):
            try:
                s_val = self.date_start.get_date()
                e_val = self.date_end.get_date()
                if s_val > e_val:
                    messagebox.showerror("ข้อผิดพลาด", "วันที่เริ่มต้น ต้องมาก่อน วันที่สิ้นสุด!")
                    return
                SEARCH_DATE_START = s_val.strftime("%Y-%m-%d")
                SEARCH_DATE_END = e_val.strftime("%Y-%m-%d")
            except Exception as e:
                messagebox.showerror("Error", f"วันที่ไม่ถูกต้อง: {e}")
                return
        else:
            SEARCH_DATE_START = self.date_start_var.get()
            SEARCH_DATE_END = self.date_end_var.get()
            try:
                 d1 = datetime.strptime(SEARCH_DATE_START, "%Y-%m-%d")
                 d2 = datetime.strptime(SEARCH_DATE_END, "%Y-%m-%d")
                 if d1 > d2: return
            except ValueError: return
        
        self.btn_run.config(state="disabled")
        self.log_text.config(bg="#f9f9f9")
        log_to_gui("เริ่มการทำงานด้วยตนเอง...")
        
        t = threading.Thread(target=self.run_process_thread)
        t.daemon = True
        t.start()
    
    def run_process_thread(self):
        try:
            self.gui_queue_put("reset_ui", "")
            process_data(update_progress_callback=self.update_progress)
        except Exception as e:
            log_to_gui(f"เกิดข้อผิดพลาด: {e}", "ERROR")
        finally:
            log_to_gui("เสร็จสิ้นการทำงาน")
            self.gui_queue_put("reset_ui", "")

    def toggle_agent(self):
        if not self.agent_running:
            # Start logic
            self.agent_running = True
            self.stop_event.clear()
            self.cb_h.config(state='disabled')
            self.cb_m.config(state='disabled')
            
            h_val = self.agent_h_var.get()
            m_val = self.agent_m_var.get()
            schedule_str = f"{int(h_val):02d}:{int(m_val):02d}"
            
            self.agent_btn.config(text="ปิดระบบอัตโนมัติ")
            self.agent_status_lbl.config(text=f"สถานะ: ทำงาน | เวลา: {schedule_str}", foreground="#27ae60")
            
            yest_str = (datetime.now() - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
            global SEARCH_DATE_START, SEARCH_DATE_END
            SEARCH_DATE_START = yest_str
            SEARCH_DATE_END = yest_str
            
            t = threading.Thread(target=self.run_agent_loop, args=(schedule_str,))
            t.daemon = True
            t.start()
            log_to_gui(f"เริ่มระบบอัตโนมัติ... เวลา: {schedule_str}")
        else:
            # Stop logic
            self.agent_running = False
            self.stop_event.set()
            self.cb_h.config(state='readonly')
            self.cb_m.config(state='readonly') # Changed to readonly
            
            self.agent_btn.config(text="เปิดระบบอัตโนมัติ")
            self.agent_status_lbl.config(text="สถานะ: ปิดใช้งาน", foreground="#7f8c8d")
            log_to_gui("ปิดระบบอัตโนมัติแล้ว")

    def run_agent_loop(self, schedule_time):
        while not self.stop_event.is_set():
            now = datetime.now()
            current_time = now.strftime("%H:%M")
            if current_time == schedule_time:
                log_to_gui(f"ถึงเวลาทำงานตามกำหนด {current_time}")
                # Update date range
                yest_str = (datetime.now() - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
                global SEARCH_DATE_START, SEARCH_DATE_END
                SEARCH_DATE_START = yest_str
                SEARCH_DATE_END = yest_str
                
                try:
                    process_data(update_progress_callback=self.update_progress)
                except Exception as e:
                    log_to_gui(f"Agent Error: {e}", "ERROR")
                
                log_to_gui("ทำงานรอบนี้เสร็จสิ้น รอรอบถัดไป")
                for _ in range(65):
                    if self.stop_event.is_set(): break
                    time.sleep(1)
            
            for _ in range(10): 
                if self.stop_event.is_set(): break
                time.sleep(1)

    # --- GUI Helpers ---
    def gui_queue_put(self, type_, msg):
        gui_queue.put((type_, msg))

    def update_progress(self, val):
        self.gui_queue_put("progress", val)

    def process_queue(self):
        while not gui_queue.empty():
            try:
                type_, msg = gui_queue.get_nowait()
                if type_ == "log":
                    self.log_text.config(state='normal')
                    self.log_text.insert(tk.END, msg + "\n")
                    self.log_text.see(tk.END)
                    self.log_text.config(state='disabled')
                elif type_ == "progress":
                    self.progress['value'] = float(msg)
                elif type_ == "reset_ui":
                    self.btn_run.config(state="normal")
                    self.progress['value'] = 0
            except queue.Empty:
                break
        
        self.after(100, self.process_queue)

if __name__ == "__main__":
    app = StrokeBatchApp()
    app.mainloop()
