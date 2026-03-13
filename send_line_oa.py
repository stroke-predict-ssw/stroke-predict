import os
import sys
import time
import queue
import logging
import threading
import traceback
import configparser
import requests
import sqlite3
import hashlib
import json
import pandas as pd
from datetime import datetime, timedelta
from urllib.parse import quote_plus, unquote_plus
from cryptography.fernet import Fernet
from sqlalchemy import create_engine, text

# GUI Imports
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None

import ctypes

# ============================================
# ⚙️ CONFIG & GLOBALS
# ============================================

gui_queue = queue.Queue()

def get_path(filename):
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

log_file = get_path("line_oa_log.txt")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s'
)

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

CONFIG_FILE = "config.ini"
HOS_HOST = ""; HOS_PORT = ""; HOS_USER = ""; HOS_PASS = ""; HOS_NAME = ""
HOS_TYPE = "postgresql"
HCODE = ""; HOSPITAL_NAME = ""; MOPH_USER = ""; MOPH_PASS_HASH = ""
CENTRAL_CONFIG_URL = ""

SEARCH_DATE_START = (datetime.now() - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
SEARCH_DATE_END = SEARCH_DATE_START

# ============================================
# 🛠️ LOAD SETTINGS
# ============================================
def load_config():
    global HOS_HOST, HOS_PORT, HOS_USER, HOS_PASS, HOS_NAME, HOS_TYPE
    global HCODE, HOSPITAL_NAME, MOPH_USER, MOPH_PASS_HASH, CENTRAL_CONFIG_URL
    config_path = get_path(CONFIG_FILE)
    config = configparser.ConfigParser()
    
    if os.path.exists(config_path):
        config.read(config_path, encoding='utf-8')
        
        if "General" in config:
            HOSPITAL_NAME = config["General"].get("hospital_name", "")
            HCODE = config["General"].get("hospital_code", "")
            MOPH_USER = config["General"].get("user", "")
            MOPH_PASS_HASH = config["General"].get("password_hash", "")
            
        if "Database" in config:
            HOS_TYPE = config["Database"].get("db_type", "mysql").lower()
            HOS_HOST = config["Database"].get("host", "")
            HOS_PORT = config["Database"].get("port", "")
            HOS_USER = config["Database"].get("username", "")
            if not HOS_USER: HOS_USER = config["Database"].get("user", "")
            HOS_NAME = config["Database"].get("database", "")
            
            raw_pass = config["Database"].get("password", "")
            try:
                HOS_PASS = cipher_suite.decrypt(raw_pass.encode()).decode()
            except:
                HOS_PASS = raw_pass
                
        if "Cloud" in config:
            CENTRAL_CONFIG_URL = config["Cloud"].get("central_config_url", "")

load_config()

# ============================================
# 🛠️ HELPER FUNCTIONS
# ============================================

def log_to_gui(msg, level="INFO"):
    timestamp = datetime.now().strftime("%H:%M:%S")
    prefix = ""
    if level == "INFO": prefix = "[INFO]"
    elif level == "WARN": prefix = "[WARN]"
    elif level == "ERROR": prefix = "[ERROR]"
    elif level == "SUCCESS": prefix = "[DONE]"
    
    clean_msg = f"[{timestamp}] {prefix} {msg}"
    
    if level == "ERROR": logging.error(msg)
    else: logging.info(msg)
    
    print(clean_msg)
    if gui_queue:
        gui_queue.put(("log", clean_msg))

import csv
import io

def fetch_central_config(url):
    try:
        log_to_gui("กำลังดึงการตั้งค่าจาก Cloud...")
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            # Force UTF-8 encoding for Thai characters
            response.encoding = 'utf-8'
            
            config_data = {}
            # Use csv reader to properly handle Google Sheet CSV escaping (like newlines and quotes)
            f = io.StringIO(response.text)
            reader = csv.reader(f)
            for parts in reader:
                if len(parts) >= 2:
                    key = parts[0].strip()
                    # Google sheets CSV might still split JSON by commas if user pastes it without quotes.
                    # Join all remaining parts with comma to reconstruct the string perfectly.
                    val = ",".join(parts[1:]).strip()
                    
                    if key.lower() == 'connection_string' or key.lower() == 'table_name':
                        config_data[key.lower()] = val
                    else:
                         # Keep original case for Risk Levels and Template
                         config_data[key] = val
                         if key == "flex_template":
                             log_to_gui(f"RAW TEMPLATE EXTRACTED: {val[:50]}...", "INFO")
            log_to_gui(f"Loaded config keys: {list(config_data.keys())}", "INFO")
            return config_data
    except Exception as e:
        log_to_gui(f"เกิดข้อผิดพลาดในการดึง Config: {e}", "ERROR")
    return None

def fix_connection_string(conn_str):
    if not conn_str: return ""
    try:
        if '://' not in conn_str: return conn_str
        prefix, rest = conn_str.split("://", 1)
        if '@' in rest:
            credentials_part, host_part = rest.rsplit('@', 1)
            if ':' in credentials_part:
                user, raw_pass = credentials_part.split(':', 1)
                safe_pass = quote_plus(unquote_plus(raw_pass))
                return f"{prefix}://{user}:{safe_pass}@{host_part}"
    except:
        pass
    return conn_str

def get_hos_engine():
    if not HOS_PASS: return None
    try:
        port_txt = f":{HOS_PORT}" if HOS_PORT else ""
        if HOS_TYPE == 'postgresql':
            conn_str = f"postgresql+psycopg2://{HOS_USER}:{quote_plus(HOS_PASS)}@{HOS_HOST}{port_txt}/{HOS_NAME}"
        else:
            conn_str = f"mysql+pymysql://{HOS_USER}:{quote_plus(HOS_PASS)}@{HOS_HOST}{port_txt}/{HOS_NAME}"
        return create_engine(conn_str)
    except Exception as e:
        log_to_gui(f"HOSxP Engine Error: {e}", "ERROR")
        return None

def setup_tracking_db():
    db_path = get_path("sent_tracking.db")
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sent_line (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cid TEXT,
            visit_date TEXT,
            sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(cid, visit_date)
        )
    """)
    conn.commit()
    conn.close()

def is_already_sent(cid, visit_date):
    db_path = get_path("sent_tracking.db")
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM sent_line WHERE cid = ? AND visit_date = ?", (cid, visit_date))
    result = cur.fetchone()
    conn.close()
    return result is not None

def mark_as_sent(cid, visit_date):
    db_path = get_path("sent_tracking.db")
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    try:
        cur.execute("INSERT OR IGNORE INTO sent_line (cid, visit_date) VALUES (?, ?)", (cid, visit_date))
        conn.commit()
    except Exception as e:
        log_to_gui(f"SQL Error marking sent: {e}", "ERROR")
    finally:
        conn.close()

def get_moph_token(force_new=False):
    global MOPH_PASS_HASH
    
    token_path = get_path("moph_token.txt")
    
    # 1. Attempt to use cached token if not forced
    if not force_new and os.path.exists(token_path):
        try:
            file_mtime = os.path.getmtime(token_path)
            # Token is valid for 24 hours, so refresh after 23 hours to be safe
            if time.time() - file_mtime < (23 * 3600):
                with open(token_path, 'r', encoding='utf-8') as f:
                    token = f.read().strip()
                    if token:
                        log_to_gui("ใช้งาน Token เดิมที่บันทึกไว้ (ยังไม่หมดอายุ)", "INFO")
                        return token
        except Exception:
            pass
            
    try:
        if force_new:
             log_to_gui("Token มีปัญหา กำลังขอ Token ใหม่...", "WARN")
        else:
             log_to_gui("กำลังขอ Token ใหม่จาก MOPH...")
             
        url = "https://cvp1.moph.go.th/token?Action=get_moph_access_token"
        
        # Check if already hashed (MD5 is 32 chars hex, SHA-256 is 64 chars hex)
        is_hashed = len(MOPH_PASS_HASH) in (32, 64) and all(c in "0123456789abcdefABCDEF" for c in MOPH_PASS_HASH)
        
        if is_hashed:
             md5_pass = MOPH_PASS_HASH.upper()
        else:
             md5_pass = hashlib.md5(MOPH_PASS_HASH.encode()).hexdigest().upper()
             
             # Save back to config.ini
             try:
                 config_path = get_path(CONFIG_FILE)
                 config = configparser.ConfigParser()
                 if os.path.exists(config_path):
                     config.read(config_path, encoding='utf-8')
                     if "General" in config:
                         config.set("General", "password_hash", md5_pass)
                         with open(config_path, 'w', encoding='utf-8') as configfile:
                             config.write(configfile)
                         MOPH_PASS_HASH = md5_pass # Update global
                         log_to_gui("เข้ารหัสรหัสผ่านและอัปเดตไฟล์ config.ini แล้ว", "INFO")
             except Exception as write_e:
                 log_to_gui(f"ไม่สามารถอัปเดต config.ini ได้: {write_e}", "WARN")
        
        payload = {
            "hospital_code": HCODE,
            "user": MOPH_USER,
            "password_hash": md5_pass
        }
        
        response = requests.post(url, json=payload, timeout=10)
        if response.status_code == 200:
            token = response.text.replace('"', '').strip()
            log_to_gui("ได้รับ Token ใหม่สำเร็จ", "SUCCESS")
            
            # Save token to file
            try:
                with open(token_path, 'w', encoding='utf-8') as f:
                    f.write(token)
            except Exception as e:
                log_to_gui(f"ไม่สามารถบันทึกไฟล์ moph_token.txt ได้: {e}", "WARN")
                
            return token
        else:
            log_to_gui(f"ไม่สามารถขอ Token ได้ (HTTP {response.status_code}): {response.text}", "ERROR")
    except Exception as e:
        log_to_gui(f"Error Token: {e}", "ERROR")
    return None

def send_line_flex(token, cid, patient_name, visit_date_th, risk_level, risk_score, sbp, dbp, chol, age, image_url, flex_template_str=None, advice_text=""):
    url = "https://morpromt2c.moph.go.th/api/send-message/send-now"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    # Custom colors and texts based on Risk Level
    bg_color = "#EAF4E1"  # Default Green
    box_color = "#27ae60"
    risk_text_en = risk_level.lower()
    
    if "medium" in risk_text_en:
        bg_color = "#FEF5E7"
        box_color = "#F39C12"
    elif "high" in risk_text_en and "very" not in risk_text_en:
        bg_color = "#FDEDEC"
        box_color = "#E74C3C"
    elif "very high" in risk_text_en or "dangerous" in risk_text_en:
        bg_color = "#F9EBEA"
        box_color = "#C0392B"
        
    BP_TEXT = f"{sbp}/{dbp} mmHg"
    if sbp > 140 or dbp > 90: BP_TEXT += " (สูง)"
    
    CHOL_TEXT = f"{chol} mg/dL"
    if chol > 200: CHOL_TEXT += " (สูง)"
        
    # Build the flex content dynamically or use fallback
    flex_json_obj = None
    if flex_template_str:
        try:
             # Replace string placeholders
             replaced_str = flex_template_str.replace("{patient_name}", str(patient_name))
             replaced_str = replaced_str.replace("{visit_date_th}", str(visit_date_th))
             replaced_str = replaced_str.replace("{bg_color}", str(bg_color))
             replaced_str = replaced_str.replace("{box_color}", str(box_color))
             replaced_str = replaced_str.replace("{risk_level}", str(risk_level))
             replaced_str = replaced_str.replace("{risk_score}", str(risk_score))
             replaced_str = replaced_str.replace("{bp_text}", str(BP_TEXT))
             replaced_str = replaced_str.replace("{chol_text}", str(CHOL_TEXT))
             replaced_str = replaced_str.replace("{age}", str(age))
             
             safe_advice = str(advice_text).replace('\n', '\\n').replace('\r', '').replace('"', '\\"')
             replaced_str = replaced_str.replace("{advice_text}", safe_advice)
             
             flex_json_obj = json.loads(replaced_str)
        except Exception as e:
             log_to_gui(f"รูปแบบ Template JSON ผิดพลาด: {e} (ไม่สามารถส่งข้อความได้)", "WARN")
             flex_json_obj = None

    if not flex_json_obj:
        log_to_gui(f"ยกเลิกการส่ง CID {cid} เนื่องจากไม่มีรูปแบบ Template JSON ที่ถูกต้อง", "WARN")
        return False
            
    messages = []
    
    # 1. Flex Message
    
    # 2. Flex Message
    messages.append({
        "type": "flex",
        "altText": "ผลประเมินความเสี่ยงโรคหลอดเลือดสมอง 10 ปีข้างหน้า",
        "contents": flex_json_obj
    })
    
    # 3. Image Message (Only if valid URL from Google Sheet)
    if image_url and len(image_url.strip()) > 10:
        messages.append({
            "type": "image",
            "originalContentUrl": str(image_url).strip(),
            "previewImageUrl": str(image_url).strip()
        })
            
    payload = {
        "datas": [str(cid)],
        "messages": messages
    }
    
    try:
         resp = requests.post(url, headers=headers, json=payload, timeout=10)
         
         # Log the raw response for debugging purposes
         if "ทดสอบ ระบบ" in patient_name:
             log_to_gui(f"API Response: {resp.text}", "INFO")
             
         if resp.status_code == 200:
             try:
                 resp_data = resp.json()
                 # MOPH APIs sometimes return 200 OK but with MessageCode != 200 or result = false
                 msg_code = str(resp_data.get("MessageCode", resp_data.get("message_code", "")))
                 if msg_code not in ["", "200"]:
                     log_to_gui(f"ส่ง Line มีข้อผิดพลาดจาก API: {resp.text}", "WARN")
                     if msg_code in ["401", "403"] or "expire" in resp.text.lower() or "invalid" in resp.text.lower():
                         token_path = get_path("moph_token.txt")
                         if os.path.exists(token_path):
                             os.remove(token_path)
                             log_to_gui("ลบ Token เดิมทิ้งเนื่องจากหมดอายุหรือไม่ถูกต้อง", "INFO")
                         return "TOKEN_EXPIRED"
                     return False
             except:
                 pass
             return True
         else:
             log_to_gui(f"ส่ง Line ล้มเหลว CID {cid} (HTTP {resp.status_code}): {resp.text}", "WARN")
             # Delete token if unauthorized
             if resp.status_code in [401, 403] or "token" in resp.text.lower() or "expire" in resp.text.lower():
                 token_path = get_path("moph_token.txt")
                 if os.path.exists(token_path):
                     os.remove(token_path)
                     log_to_gui("ลบ Token เดิมทิ้งเนื่องจากหมดอายุหรือไม่ถูกต้อง", "INFO")
                 return "TOKEN_EXPIRED"
             return False
    except Exception as e:
         log_to_gui(f"Network error on send Line {cid}: {e}", "ERROR")
         return False

# ============================================
# 🧠 MAIN PROCESS LOGIC
# ============================================
def process_data(update_progress_callback=None):
    setup_tracking_db()
    
    config_data = fetch_central_config(CENTRAL_CONFIG_URL)
    if not config_data or 'connection_string' not in config_data:
        log_to_gui("ไม่สามารถดึง Config จาก Cloud ได้ หยุกการทำงาน", "ERROR")
        return
        
    cloud_table = config_data.get('table_name', 'center_db')
    
    # Decrypt cloud DB connection
    try:
        encrypted_conn = config_data['connection_string']
        decrypted_conn = cipher_suite.decrypt(encrypted_conn.encode()).decode()
        safe_conn_str = fix_connection_string(decrypted_conn)
        cloud_engine = create_engine(safe_conn_str)
    except Exception as e:
        log_to_gui(f"การเชื่อมต่อ Cloud DB ล้มเหลว: {e}", "ERROR")
        return

    # Check HOSxP connection
    hos_engine = get_hos_engine()
    if not hos_engine:
        log_to_gui("เชื่อมต่อ HOSxP ไม่สำเร็จ", "ERROR")
        return

    # Get Token
    token = get_moph_token()
    if not token:
        log_to_gui("ไม่สามารถดึง Token ได้ ยกเลิกกระบวนการ", "ERROR")
        return

    log_to_gui(f"กำลังดึงข้อมูลทำนายของวันที่ {SEARCH_DATE_START} ถึง {SEARCH_DATE_END}...")
    
    # Convert dates to MySQL format suitable format
    query = f"""
        SELECT hn, visit_date, risk_score, risk_level, sbp, dbp, chol, age 
        FROM {cloud_table} 
        WHERE hcode = '{HCODE}' 
        AND visit_date >= '{SEARCH_DATE_START}' 
        AND visit_date <= '{SEARCH_DATE_END}'
    """
    
    try:
        with cloud_engine.connect() as conn:
            predictions = conn.execute(text(query)).mappings().fetchall()
    except Exception as e:
        log_to_gui(f"ดึงข้อมูลจาก Cloud ล้มเหลว: {e}", "ERROR")
        return
        
    total_preds = len(predictions)
    log_to_gui(f"พบข้อมูลประเมินความเสี่ยงทั้งหมด {total_preds} รายการ")
    
    if total_preds == 0:
        return
        
    success_count = 0
    skip_count = 0
    
    with hos_engine.connect() as hos_conn:
        for i, row in enumerate(predictions):
            try:
                 hn = row['hn']
                 visit_date = str(row['visit_date'])
                 risk_level = row['risk_level']
                 
                 # 1. Query HOSxP for Name and CID
                 patient_sql = f"SELECT cid, fname, lname FROM patient WHERE hn = '{hn}' LIMIT 1"
                 p_res = hos_conn.execute(text(patient_sql)).mappings().fetchone()
                 
                 if not p_res:
                     log_to_gui(f"ไม่พบ HN {hn} ใน HOSxP", "WARN")
                     continue
                     
                 cid = p_res['cid']
                 fname = p_res['fname']
                 lname = p_res['lname']
                 full_name = f"{fname} {lname}"
                 
                 # 2. Check if already sent
                 if is_already_sent(cid, visit_date):
                     skip_count += 1
                     continue
                     
                 # 3. Prepare Image URL
                 image_url = config_data.get(risk_level, "") if config_data else ""
                 if not image_url or not isinstance(image_url, str) or len(image_url.strip()) < 10:
                     image_url = "https://img5.pic.in.th/file/secure-sv1/Gemini_Generated_Image_pv29qxpv29qxpv29.png" # Fallback
                     
                 # Formart Date to Thai
                 try:
                     vd = datetime.strptime(visit_date, "%Y-%m-%d")
                     visit_date_th = f"{vd.day} {['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'][vd.month-1]} {vd.year+543}"
                 except:
                     visit_date_th = visit_date
                     
                 # 3.5 Prepare Flex Template & Advice from Config (if available)
                 flex_template_str = config_data.get('flex_template', "") if config_data else ""
                 
                 # Fetch advice for specific risk level, fallback to general default
                 advice_key = f"{risk_level}_Advice".replace(" ", "_")
                 default_advice = "ควรดูแลสุขภาพและปฏิบัติตามคำแนะนำของแพทย์"
                 advice_text = config_data.get(advice_key, default_advice) if config_data else default_advice
                     
                 # 4. Send Line
                 res = send_line_flex(
                     token=token,
                     cid=cid,
                     patient_name=full_name,
                     visit_date_th=visit_date_th,
                     risk_level=risk_level,
                     risk_score=row['risk_score'],
                     sbp=row['sbp'],
                     dbp=row['dbp'],
                     chol=row['chol'],
                     age=row['age'],
                     image_url=image_url,
                     flex_template_str=flex_template_str,
                     advice_text=advice_text
                 )
                 
                 # Automatic Token Refresh Retry
                 if res == "TOKEN_EXPIRED":
                     log_to_gui("Token มีปัญหา กำลังขอใหม่เพื่อส่งซ้ำ...", "INFO")
                     token = get_moph_token(force_new=True)
                     if token:
                         res = send_line_flex(
                             token=token,
                             cid=cid,
                             patient_name=full_name,
                             visit_date_th=visit_date_th,
                             risk_level=risk_level,
                             risk_score=row['risk_score'],
                             sbp=row['sbp'],
                             dbp=row['dbp'],
                             chol=row['chol'],
                             age=row['age'],
                             image_url=image_url,
                             flex_template_str=flex_template_str,
                             advice_text=advice_text
                         )
                 
                 if res == True:
                     mark_as_sent(cid, visit_date)
                     success_count += 1
                     
            except Exception as e:
                 log_to_gui(f"Error on patient {hn}: {e}", "ERROR")

            if update_progress_callback:
                 update_progress_callback((i + 1) / total_preds * 100)
                 
    log_to_gui(f"---- สรุปผลการส่ง ----", "SUCCESS")
    log_to_gui(f"ส่ง Line สำเร็จ: {success_count} รายการ", "SUCCESS")
    log_to_gui(f"ข้าม (ส่งไปแล้ว): {skip_count} รายการ", "INFO")


# ============================================
# 🖥️ GUI APPLICATION
# ============================================

class LineOAAgentApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        try:
            myappid = 'myorg.strokeagent.lineoa.1.0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except:
            pass

        self.title(f"{HOSPITAL_NAME} - ระบบส่ง Line OA อัตโนมัติ")
        self.geometry("450x680")
        self.resizable(False, False)
        
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure("TFrame", background="#f0f2f5")
        self.style.configure("TLabel", background="#f0f2f5", font=("Segoe UI", 9))
        self.style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"), foreground="#2c3e50")
        self.style.configure("Status.TLabel", font=("Segoe UI", 8), foreground="#7f8c8d")
        self.style.configure("TLabelframe", background="#f0f2f5", relief="flat")
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 9, "bold"), foreground="#34495e", background="#f0f2f5")
        
        self.setup_icon()
        
        self.date_start_var = tk.StringVar(value=(datetime.now() - pd.Timedelta(days=1)).strftime("%Y-%m-%d"))
        self.date_end_var = tk.StringVar(value=(datetime.now() - pd.Timedelta(days=1)).strftime("%Y-%m-%d"))
        
        self.agent_h_var = tk.StringVar(value="08")
        self.agent_m_var = tk.StringVar(value="30")
        self.agent_running = False
        self.stop_event = threading.Event()
        
        self.configure(bg="#f0f2f5")
        self.create_widgets()
        self.process_queue()

    def setup_icon(self):
        try:
            icon_path = get_path("artificial-intelligence.ico")
            self.iconbitmap(icon_path)
        except: pass

    def create_widgets(self):
        header_frame = ttk.Frame(self, padding="20 15 20 10")
        header_frame.pack(fill=tk.X)
        
        ttk.Label(header_frame, text=f"{HOSPITAL_NAME} (Line OA)", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header_frame, text=f"รหัสหน่วยบริการ: {HCODE}", style="Status.TLabel").pack(anchor="w")
        
        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=20, pady=(5,15))

        manual_frame = ttk.Labelframe(self, text=" สั่งทำงานด้วยตนเอง (Manual) ", padding=15)
        manual_frame.pack(fill=tk.X, padx=20, pady=5)
        
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
        
        self.btn_run = ttk.Button(manual_frame, text="เริ่มส่งข้อความ", command=self.start_manual_run)
        self.btn_run.pack(fill=tk.X, pady=(15, 0), ipady=3)

        # ====== โซนทดสอบ (Test Mode) ======
        test_frame = ttk.Labelframe(self, text=" ทดสอบระบบ (Test Mode) ", padding=15)
        test_frame.pack(fill=tk.X, padx=20, pady=5)
        
        input_test_frame = ttk.Frame(test_frame)
        input_test_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(input_test_frame, text="ส่งเข้าบัตร ปชช:").pack(side=tk.LEFT, padx=5)
        self.test_cid_var = tk.StringVar(value="")
        self.test_cid_entry = ttk.Entry(input_test_frame, textvariable=self.test_cid_var, width=15)
        self.test_cid_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        btn_test_frame = ttk.Frame(test_frame)
        btn_test_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.btn_test = ttk.Button(btn_test_frame, text="เทสระบบ (จำลอง)", command=self.start_test_run)
        self.btn_test.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        self.btn_real_test = ttk.Button(btn_test_frame, text="เทสข้อมูลจริง", command=self.start_real_test_run)
        self.btn_real_test.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)

        # ====== โซนอัตโนมัติ (Auto Mode) ======
        agent_frame = ttk.Labelframe(self, text=" ตั้งค่าระบบอัตโนมัติ (Auto Agent) ", padding=15)
        agent_frame.pack(fill=tk.X, padx=20, pady=15)
        
        time_frame = ttk.Frame(agent_frame)
        time_frame.pack(fill=tk.X, pady=5)
        ttk.Label(time_frame, text="เวลาทำงาน (ทุกวัน):").pack(side=tk.LEFT)
        
        hours = [f"{i:02d}" for i in range(24)]
        minutes = [f"{i:02d}" for i in range(0, 60, 5)]
        
        self.cb_h = ttk.Combobox(time_frame, textvariable=self.agent_h_var, values=hours, width=3, state="readonly")
        self.cb_h.pack(side=tk.LEFT, padx=(10,2))
        ttk.Label(time_frame, text=":").pack(side=tk.LEFT)
        self.cb_m = ttk.Combobox(time_frame, textvariable=self.agent_m_var, values=minutes, width=3, state="readonly")
        self.cb_m.pack(side=tk.LEFT, padx=2)
        
        self.agent_btn = ttk.Button(agent_frame, text="เปิดระบบอัตโนมัติ", command=self.toggle_agent)
        self.agent_btn.pack(fill=tk.X, pady=(15, 0), ipady=3)
        
        self.agent_status_lbl = ttk.Label(agent_frame, text="สถานะ: ปิดใช้งาน", style="Status.TLabel")
        self.agent_status_lbl.pack(anchor="c", pady=(5,0))

        log_frame = ttk.Frame(self, padding="20 0 20 10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, state='disabled', 
                                              font=("Consolas", 8), bg="#ffffff", fg="#2d3436", relief="flat")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.pack(fill=tk.X, side=tk.BOTTOM)

    def start_manual_run(self):
        if self.agent_running:
            messagebox.showwarning("ระบบไม่ว่าง", "ระบบอัตโนมัติทำงานอยู่ กรุณาปิดปุ่ม Auto ก่อน")
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
        
        self.btn_run.config(state="disabled")
        self.btn_test.config(state="disabled")
        self.log_text.config(bg="#f9f9f9")
        log_to_gui("เริ่มประมวลผลการส่ง Line OA (Manual)...")
        
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

    def start_test_run(self):
        cid = self.test_cid_var.get().strip()
        if not cid or len(cid) != 13:
            messagebox.showerror("Error", "กรุณากรอกเลขบัตรประชาชน 13 หลักให้ถูกต้อง")
            return
            
        self.btn_test.config(state="disabled")
        self.btn_run.config(state="disabled")
        log_to_gui(f"เริ่มทดสอบส่ง Line หา CID: {cid}...")
        t = threading.Thread(target=self.run_test_thread, args=(cid,))
        t.daemon = True
        t.start()
        
    def run_test_thread(self, cid):
        try:
            self.gui_queue_put("reset_ui", "")
            token = get_moph_token()
            if not token:
                log_to_gui("ทดสอบล้มเหลว: ไม่สามารถดึง Token ได้ โปรดตรวจสอบ user/pass ใน config.ini", "ERROR")
                return
                
            config_data = fetch_central_config(CENTRAL_CONFIG_URL)
            flex_template_str = config_data.get('flex_template', "") if config_data else ""
            
            test_cases = [
                {"level": "Low", "score": 5.0, "sbp": 120, "dbp": 80, "chol": 180, "age": 45},
                {"level": "Medium", "score": 15.0, "sbp": 135, "dbp": 85, "chol": 210, "age": 55},
                {"level": "High", "score": 25.5, "sbp": 160, "dbp": 95, "chol": 220, "age": 60},
                {"level": "Very High", "score": 35.0, "sbp": 170, "dbp": 100, "chol": 250, "age": 65},
                {"level": "Dangerous", "score": 45.0, "sbp": 180, "dbp": 110, "chol": 280, "age": 70}
            ]
            
            success_count = 0
            for case in test_cases:
                r_level = case["level"]
                image_url = config_data.get(r_level, "") if config_data else ""
                if not image_url or len(image_url.strip()) < 10:
                    image_url = ""
                
                advice_key = f"{r_level}_Advice".replace(" ", "_")
                advice_text = config_data.get(advice_key, "ควรดูแลสุขภาพตามคำแนะนำของแพทย์") if config_data else "ควรดูแลสุขภาพตามคำแนะนำของแพทย์"
                
                log_to_gui(f"กำลังส่งทดสอบระดับ: {r_level}...", "INFO")
                res = send_line_flex(
                    token=token,
                    cid=cid,
                    patient_name=f"ทดสอบ ({r_level})",
                    visit_date_th="1 มีนาคม 2569",
                    risk_level=r_level,
                    risk_score=case["score"],
                    sbp=case["sbp"],
                    dbp=case["dbp"],
                    chol=case["chol"],
                    age=case["age"],
                    image_url=image_url,
                    flex_template_str=flex_template_str,
                    advice_text=advice_text
                )
                
                if res == "TOKEN_EXPIRED":
                    log_to_gui("Token มีปัญหา กำลังขอใหม่เพื่อส่งซ้ำ...", "INFO")
                    token = get_moph_token(force_new=True)
                    if token:
                        res = send_line_flex(
                            token=token,
                            cid=cid,
                            patient_name=f"ทดสอบ ({r_level})",
                            visit_date_th="1 มีนาคม 2569",
                            risk_level=r_level,
                            risk_score=case["score"],
                            sbp=case["sbp"],
                            dbp=case["dbp"],
                            chol=case["chol"],
                            age=case["age"],
                            image_url=image_url,
                            flex_template_str=flex_template_str,
                            advice_text=advice_text
                        )
                        
                if res == True:
                    success_count += 1
                time.sleep(1) # delay between each test message to prevent spam rate limiting
                
            if success_count == len(test_cases):
                log_to_gui(f"ทดสอบส่งสำเร็จครบ {success_count} ระดับ! กรุณาเช็คใน Line หมอพร้อม", "SUCCESS")
                messagebox.showinfo("Success", f"ส่งข้อความทดสอบสำเร็จ {success_count} ข้อความ!")
            else:
                log_to_gui(f"ส่งข้อความทดสอบสำเร็จ {success_count}/{len(test_cases)}", "WARN")
        except Exception as e:
            log_to_gui(f"Error test: {e}", "ERROR")
        finally:
            self.gui_queue_put("reset_ui", "")

    def start_real_test_run(self):
        cid = self.test_cid_var.get().strip()
        if not cid or len(cid) != 13:
            messagebox.showerror("Error", "กรุณากรอกเลขบัตรประชาชน 13 หลักให้ถูกต้อง")
            return
            
        self.btn_test.config(state="disabled")
        try:
            self.btn_real_test.config(state="disabled")
        except: pass
        self.btn_run.config(state="disabled")
        
        log_to_gui(f"เริ่มดึงข้อมูลจริงของ CID: {cid}...", "INFO")
        t = threading.Thread(target=self.run_real_test_thread, args=(cid,))
        t.daemon = True
        t.start()
        
    def run_real_test_thread(self, cid):
        try:
            self.gui_queue_put("reset_ui", "")
            
            import sys
            import os
            sys.path.append(os.path.dirname(os.path.abspath(__file__)))
            import batch_predict
            import pandas as pd
            
            hos_engine = get_hos_engine()
            if not hos_engine:
                log_to_gui("เชื่อมต่อ HOSxP ไม่สำเร็จ ล้มเหลวการทดสอบ", "ERROR")
                return
                
            config_data = fetch_central_config(CENTRAL_CONFIG_URL)
            if not config_data:
                log_to_gui("ไม่สามารถดึง Config จาก Cloud ได้", "ERROR")
                return
                
            log_to_gui(f"ค้นหาและคำนวณสดสำหรับ CID: {cid}...", "INFO")
            
            config_path = get_path("config.ini")
            config = configparser.ConfigParser()
            config.read(config_path, encoding='utf-8')
            db_type = config["Database"]["db_type"].lower() if "Database" in config else "mysql"
            regex_op = "~" if db_type == "postgresql" else "REGEXP"
            fbs_codes = batch_predict.FBS_CODES_SQL if hasattr(batch_predict, 'FBS_CODES_SQL') else "'76','1698'"
            chol_codes = batch_predict.CHOL_CODES_SQL if hasattr(batch_predict, 'CHOL_CODES_SQL') else "'102','1691'"
            
            query = f"""
            SELECT p.hn, p.cid, p.fname, p.lname, p.birthday, p.sex as gender, p.marrystatus,
            (SELECT vstdate FROM opdscreen WHERE hn = p.hn ORDER BY vstdate DESC LIMIT 1) AS last_visit_date,
            CAST((SELECT bps FROM opdscreen WHERE hn = p.hn ORDER BY vstdate DESC LIMIT 1) AS DECIMAL(12,2)) AS sbp,
            CAST((SELECT bpd FROM opdscreen WHERE hn = p.hn ORDER BY vstdate DESC LIMIT 1) AS DECIMAL(12,2)) AS dbp,
            CAST((SELECT bw FROM opdscreen WHERE hn = p.hn ORDER BY vstdate DESC LIMIT 1) AS DECIMAL(12,2)) AS weight,
            CAST((SELECT height FROM opdscreen WHERE hn = p.hn AND height > 0 ORDER BY vstdate DESC LIMIT 1) AS DECIMAL(12,2)) AS height,
            CAST((SELECT waist FROM opdscreen WHERE hn = p.hn AND waist > 0 ORDER BY vstdate DESC LIMIT 1) AS DECIMAL(12,2)) AS waist,
            CAST((SELECT lab_order_result FROM lab_order lo JOIN lab_head lh ON lo.lab_order_number = lh.lab_order_number WHERE lh.hn = p.hn AND lo.lab_items_code IN ({fbs_codes}) AND lo.lab_order_result {regex_op} '^[0-9]' ORDER BY lh.order_date DESC LIMIT 1) AS DECIMAL(12,2)) AS fbs,
            CAST((SELECT lab_order_result FROM lab_order lo JOIN lab_head lh ON lo.lab_order_number = lh.lab_order_number WHERE lh.hn = p.hn AND lo.lab_items_code IN ({chol_codes}) AND lo.lab_order_result {regex_op} '^[0-9]' ORDER BY lh.order_date DESC LIMIT 1) AS DECIMAL(12,2)) AS chol,
            (SELECT smoking_type_id FROM opdscreen WHERE hn = p.hn AND smoking_type_id IS NOT NULL ORDER BY vstdate DESC LIMIT 1) as smoke,
            (SELECT drinking_type_id FROM opdscreen WHERE hn = p.hn AND drinking_type_id IS NOT NULL ORDER BY vstdate DESC LIMIT 1) as drink,
            (SELECT count(*) FROM ovstdiag WHERE hn = p.hn AND (icd10 BETWEEN 'I20' AND 'I25')) as cardio_history
            FROM patient p WHERE p.cid = '{cid}' LIMIT 1
            """
            
            with hos_engine.connect() as hos_conn:
                p_res = hos_conn.execute(text(query)).mappings().fetchone()
                if not p_res:
                    log_to_gui(f"ไม่พบข้อมูลผู้ป่วยที่มี CID {cid} ในระบบ HOSxP", "WARN")
                    return
            
            raw_data = dict(p_res)
            full_name = f"{raw_data.get('fname','')} {raw_data.get('lname','')}"
            
            model, FEATURES = batch_predict.load_ai_model()
            if not model:
                log_to_gui("โหลดโมเดล AI ล้มเหลว", "ERROR")
                return
                
            def calc_age(bdate):
                if not bdate: return 50
                try: return int((datetime.now().date() - bdate).days / 365.25)
                except: return 50
                
            age = calc_age(raw_data.get('birthday'))
            sbp = float(raw_data.get('sbp') or 0)
            dbp = float(raw_data.get('dbp') or 0)
            weight = float(raw_data.get('weight') or 0)
            height = float(raw_data.get('height') or 0)
            waist = float(raw_data.get('waist') or 0)
            fbs = float(raw_data.get('fbs') or 0)
            chol = float(raw_data.get('chol') or 0)
            
            gender_code = str(raw_data.get('gender')) 
            gender_val = 0 if gender_code in ['1', 'ชาย', 'Male', '0', 'male'] else 1
            is_married = 0 if str(raw_data.get('marrystatus')) in ['1', '9', '6'] else 1
            
            smoke_code = str(raw_data.get('smoke'))
            smoke_idx = 2 if smoke_code in ['3', '4'] else (1 if smoke_code in ['2', '5'] else 0)
            drink_code = str(raw_data.get('drink'))
            drink_idx = 3 if drink_code == '4' else (2 if drink_code == '3' else (1 if drink_code == '2' else 0))
            
            cardio = raw_data.get('cardio_history', 0) > 0 
            bmi_calc = (weight / ((height / 100) ** 2)) if height > 0 else 0
            pred_age = age + 10
            
            input_data = {
                "gender_numeric_map": gender_val, "age": pred_age, "cardio_numeric_map": 1 if cardio else 0,
                "marry_status_numeric_map": is_married, "smoking_status": smoke_idx, "drinking_status": drink_idx,
                "sbp": sbp, "dbp": dbp, "weight": weight, "bmi": bmi_calc, "waist": waist,
                "avg_glocose_level": fbs, "cholesterol": chol, "occupation": 1, "education": 4, 
            }
            
            input_df = pd.DataFrame([input_data])[FEATURES]
            model_prob = model.predict_proba(input_df)[0][1]
            prob_percent = batch_predict.recalibrate_probability(model_prob) * 100
            final_risk_score = round(prob_percent, 2)
            
            risk_level = "Dangerous"
            if final_risk_score < 10: risk_level = "Low"
            elif final_risk_score < 20: risk_level = "Medium"
            elif final_risk_score < 30: risk_level = "High"
            elif final_risk_score < 40: risk_level = "Very High"
            
            log_to_gui(f"คำนวณเสร็จสิ้น คะแนน: {final_risk_score}% ({risk_level})", "SUCCESS")
            
            visit_date = str(raw_data.get('last_visit_date') or datetime.now().strftime("%Y-%m-%d"))
            try:
                vd = datetime.strptime(visit_date, "%Y-%m-%d")
                visit_date_th = f"{vd.day} {['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'][vd.month-1]} {vd.year+543}"
            except:
                visit_date_th = visit_date
                
            image_url = config_data.get(risk_level, "") if config_data else ""
            if not image_url or len(image_url.strip()) < 10:
                image_url = ""
                
            flex_template_str = config_data.get('flex_template', "") if config_data else ""
            
            advice_key = f"{risk_level}_Advice".replace(" ", "_")
            default_advice = "ควรดูแลสุขภาพและปฏิบัติตามคำแนะนำของแพทย์"
            advice_text = config_data.get(advice_key, default_advice) if config_data else default_advice
            
            token = get_moph_token()
            if not token:
                log_to_gui("ทดสอบล้มเหลว: ไม่สามารถดึง Token ได้ โปรดตรวจสอบ user/pass ใน config.ini", "ERROR")
                return
                
            log_to_gui(f"กำลังส่งข้อมูลจริงไปยัง Line สำหรับ CID {cid} (ดึงเมื่อ {visit_date_th}, ระดับ {risk_level})", "INFO")
            
            res = send_line_flex(
                token=token,
                cid=cid,
                patient_name=f"{full_name} (ข้อมูลจริง)",
                visit_date_th=visit_date_th,
                risk_level=risk_level,
                risk_score=final_risk_score,
                sbp=int(sbp),
                dbp=int(dbp),
                chol=int(chol),
                age=age,
                image_url=image_url,
                flex_template_str=flex_template_str,
                advice_text=advice_text
            )
            
            if res == "TOKEN_EXPIRED":
                log_to_gui("Token มีปัญหา กำลังขอใหม่เพื่อส่งซ้ำ...", "INFO")
                token = get_moph_token(force_new=True)
                if token:
                    res = send_line_flex(
                        token=token,
                        cid=cid,
                        patient_name=f"{full_name} (ข้อมูลจริง)",
                        visit_date_th=visit_date_th,
                        risk_level=risk_level,
                        risk_score=final_risk_score,
                        sbp=int(sbp),
                        dbp=int(dbp),
                        chol=int(chol),
                        age=age,
                        image_url=image_url,
                        flex_template_str=flex_template_str,
                        advice_text=advice_text
                    )
                    
            if res == True:
                log_to_gui("ทดสอบส่งข้อมูลจริงสำเร็จ! กรุณาเช็คใน Line หมอพร้อม", "SUCCESS")
                messagebox.showinfo("Success", "ส่งข้อความทดสอบข้อมูลจริงสำเร็จ!")
            else:
                log_to_gui("ส่งข้อความทดสอบล้มเหลว", "ERROR")

        except Exception as e:
             log_to_gui(f"Error real test: {e}", "ERROR")
        finally:
             self.gui_queue_put("reset_ui", "")

    def toggle_agent(self):
        if not self.agent_running:
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
            log_to_gui(f"เริ่มระบบอัตโนมัติ... เวลาส่ง Line: {schedule_str}")
        else:
            self.agent_running = False
            self.stop_event.set()
            self.cb_h.config(state='readonly')
            self.cb_m.config(state='readonly')
            
            self.agent_btn.config(text="เปิดระบบอัตโนมัติ")
            self.agent_status_lbl.config(text="สถานะ: ปิดใช้งาน", foreground="#7f8c8d")
            log_to_gui("ปิดระบบอัตโนมัติแล้ว")

    def run_agent_loop(self, schedule_time):
        while not self.stop_event.is_set():
            now = datetime.now()
            current_time = now.strftime("%H:%M")
            if current_time == schedule_time:
                log_to_gui(f"ถึงเวลาส่ง Line ตามกำหนด {current_time}")
                # Update date range to yesterday
                yest_str = (datetime.now() - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
                global SEARCH_DATE_START, SEARCH_DATE_END
                SEARCH_DATE_START = yest_str
                SEARCH_DATE_END = yest_str
                
                try:
                    process_data(update_progress_callback=self.update_progress)
                except Exception as e:
                    log_to_gui(f"Agent Error: {e}", "ERROR")
                
                log_to_gui("ส่งข้อความรอบนี้เสร็จสิ้น รอรอบถัดไปพรุ่งนี้")
                for _ in range(65):
                    if self.stop_event.is_set(): break
                    time.sleep(1)
            
            for _ in range(10): 
                if self.stop_event.is_set(): break
                time.sleep(1)

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
                    self.btn_test.config(state="normal")
                    try:
                        self.btn_real_test.config(state="normal")
                    except: pass
                    self.progress['value'] = 0
            except queue.Empty:
                break
        
        self.after(100, self.process_queue)

if __name__ == "__main__":
    app = LineOAAgentApp()
    app.mainloop()
