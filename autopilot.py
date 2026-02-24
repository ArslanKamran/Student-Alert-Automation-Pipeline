import imaplib
import email
import os
import shutil
import json
import time
import re
import os
import shutil
import tempfile
import logging
import requests
from dotenv import load_dotenv

from email.header import decode_header
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
#      LOGGING SETUP
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler("automation.log"),
        logging.StreamHandler()
    ]
)

# ==========================================
#      USER CONFIGURATION
# ==========================================
load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
IMAP_SERVER = "imap.gmail.com"

# Set to True to completely hide Chrome, set to False to watch it upload live.
HEADLESS_MODE = os.getenv("HEADLESS_MODE", "True").lower() == "true"

SITE_USERNAME = os.getenv("SITE_USERNAME")
SITE_PASSWORD = os.getenv("SITE_PASSWORD")
LOGIN_URL = "https://app1.student-alert.com/"

OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "gemma3:12b"

# ==========================================
#      HELPER FUNCTIONS
# ==========================================

def clean_filename(filename):
    if not filename: return "unnamed_file"
    decoded_header = decode_header(filename)
    header_parts = []
    for content, encoding in decoded_header:
        if isinstance(content, bytes):
            header_parts.append(content.decode(encoding or "utf-8"))
        else:
            header_parts.append(content)
    return "".join(header_parts)

def get_upload_title(subject_text):
    lower_subj = subject_text.lower()
    match = re.search(r"week\s*#?(\d+)", lower_subj)
    if match: return f"Week {match.group(1)}"
    if "revision" in lower_subj: return "Revision"
    return None

def map_folder_to_dropdown_text(folder_name):
    clean = folder_name.replace("Class_", "").replace("Class ", "")
    if "Reception" in clean: return "RECEPTION - A"
    if "Nursery" in clean: return "NURSERY - A"
    if "Prep" in clean: return "PREP - A"
    
    suffix = "A"
    number = ''.join(filter(str.isdigit, clean))
    if "Matric" in clean: suffix = "M"
    elif "Cambridge" in clean: suffix = "C"
    elif "M" in clean.upper() and "CAMBRIDGE" not in clean.upper(): suffix = "M"
    elif "C" in clean.upper() and "MATRIC" not in clean.upper(): suffix = "C"
    elif not number: return None 
    if number in ["1", "2", "3", "4", "5", "6", "7"]: suffix = "A"
    return f"{number} - {suffix}"

def force_select_hidden_dropdown(driver, element_id, visible_text):
    try:
        select_element = driver.find_element(By.ID, element_id)
        driver.execute_script("arguments[0].style.display = 'block';", select_element)
        driver.execute_script("arguments[0].style.visibility = 'visible';", select_element)
        driver.execute_script("arguments[0].style.height = 'auto';", select_element)
        driver.execute_script("arguments[0].style.opacity = '1';", select_element)
        time.sleep(0.5)
        
        dropdown = Select(select_element)
        for opt in dropdown.options:
            if visible_text.lower() in opt.text.lower():
                dropdown.select_by_visible_text(opt.text)
                driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", select_element)
                logging.info(f"Selected: '{opt.text}'")
                return True
        logging.warning(f"Dropdown option containing '{visible_text}' not found")
        return False
    except Exception as e:
        logging.warning(f"Could not force select dropdown: {e}")
        return False

def classify_file_locally(filename):
    """Sorts based on filename patterns (Pass 1)"""
    name = filename.lower().replace("-", " ").replace("_", " ")
    
    if re.search(r"\b10\s*m\b", name) or "10 matric" in name: return ["Class_10_Matric"]
    if re.search(r"\b10\s*c\b", name) or "10 cambridge" in name: return ["Class_10_Cambridge"]
    if re.search(r"\b9\s*m\b", name) or "9 matric" in name: return ["Class_9_Matric"]
    if re.search(r"\b9\s*c\b", name) or "9 cambridge" in name: return ["Class_9_Cambridge"]
    if re.search(r"\b8\s*m\b", name) or "8 matric" in name: return ["Class_8_Matric"]
    if re.search(r"\b8\s*c\b", name) or "8 cambridge" in name: return ["Class_8_Cambridge"]

    if "class 7" in name or "grade 7" in name: return ["Class_7"]
    if "class 6" in name or "grade 6" in name: return ["Class_6"]
    if "class 5" in name or "grade 5" in name: return ["Class_5"]
    if "class 4" in name or "grade 4" in name: return ["Class_4"]
    if "class 3" in name or "grade 3" in name: return ["Class_3"]
    if "class 2" in name or "grade 2" in name: return ["Class_2"]
    if "class 1" in name or "grade 1" in name: return ["Class_1"]
    
    if "prep" in name: return ["Class_Prep"]
    if "nursery" in name: return ["Class_Nursery"]
    if "reception" in name: return ["Class_Reception"]

    return None

def ask_ai_batch(file_list):
    if not file_list: return {}
    logging.info(f"Sending {len(file_list)} files to Ollama for sorting...")
    
    prompt = f"""
    You are a file sorting assistant.
    Valid Folders: ["Class_Reception", "Class_Nursery", "Class_Prep", "Class_1", "Class_2", "Class_3", "Class_4", "Class_5", "Class_6", "Class_7", "Class_8_Matric", "Class_9_Matric", "Class_10_Matric", "Class_8_Cambridge", "Class_9_Cambridge", "Class_10_Cambridge"]
    INPUT LIST: {json.dumps(file_list)}
    INSTRUCTIONS: 1. Match filenames to folders. 2. Ignore week numbers. 3. Return JSON exactly matching this format: {{"file_0": ["folder_name"]}}. Do not return any other text or reasoning.
    """
    
    try:
        response = requests.post(OLLAMA_URL, json={
            "model": OLLAMA_MODEL,
            "prompt": prompt,
            "stream": False,
            "format": "json"
        })
        response.raise_for_status()
        
        text = response.json().get("response", "")
        # Safely extract JSON if the model outputs markdown blocks
        if "```json" in text: text = text.split("```json")[1]
        if "```" in text: text = text.split("```")[0]
        return json.loads(text.strip())
    except requests.exceptions.ConnectionError:
        logging.error("CRITICAL: Could not connect to Ollama. Make sure the Ollama app is running locally!")
        return {}
    except Exception as e:
        logging.error(f"AI sorting failed: {e}")
        return {}

# ==========================================
#      PROCESS LOGIC
# ==========================================

def process_email_and_upload(email_subject, upload_title, email_msg, driver=None):
    logging.info(f"Processing work for {upload_title} ({email_subject})...")
    
    safe_title = upload_title.replace(" ", "_")
    with tempfile.TemporaryDirectory() as temp_work_dir:
        base_dir = os.path.join(temp_work_dir, f"Batch_Run_{safe_title}")
        os.makedirs(base_dir)
        download_dir = os.path.join(temp_work_dir, "_temp")
        os.makedirs(download_dir)
        zip_dir = os.path.join(temp_work_dir, f"Zips_Batch_{safe_title}")
        os.makedirs(zip_dir)

        # 1. DOWNLOAD
        files_to_process = []
        file_map = {}
        
        for part in email_msg.walk():
            if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None: continue
            fname = clean_filename(part.get_filename())
            if not fname: continue
            
            file_id = f"file_{len(file_map)}"
            temp_path = os.path.join(download_dir, f"{file_id}_{fname}") 
            with open(temp_path, "wb") as f: f.write(part.get_payload(decode=True))
            
            files_to_process.append({"id": file_id, "filename": fname, "context": email_subject})
            file_map[file_id] = {"path": temp_path, "real_name": fname}

        if not files_to_process:
            logging.warning("No attachments found in email.")
            return False, driver

        # 2. SORT (Hybrid)
        folders_created = set()
        files_for_ai = []
        ai_results = {}

        logging.info("Running local pattern match...")
        for item in files_to_process:
            fid = item['id']
            local_folders = classify_file_locally(item['filename'])
            if local_folders:
                logging.info(f"Matched {item['filename']} -> {local_folders[0]}")
                for folder in local_folders:
                    dest = os.path.join(base_dir, folder)
                    os.makedirs(dest, exist_ok=True)
                    shutil.copy2(file_map[fid]['path'], os.path.join(dest, file_map[fid]['real_name']))
                    folders_created.add(folder)
            else:
                files_for_ai.append(item)

        if files_for_ai:
            ai_results = ask_ai_batch(files_for_ai)
            for item in files_for_ai:
                fid = item['id']
                targets = ai_results.get(fid, [])
                for folder in targets:
                    dest = os.path.join(base_dir, folder)
                    os.makedirs(dest, exist_ok=True)
                    shutil.copy2(file_map[fid]['path'], os.path.join(dest, file_map[fid]['real_name']))
                    folders_created.add(folder)

        if not folders_created:
            logging.warning("No recognizable folders found or created.")
            return False, driver

        # 3. ZIP
        zips_to_upload = []
        for folder in folders_created:
            zip_name = f"{folder} {upload_title}"
            shutil.make_archive(os.path.join(zip_dir, zip_name), 'zip', os.path.join(base_dir, folder))
            zips_to_upload.append(f"{zip_name}.zip")

        # 4. UPLOAD
        if not driver:
            logging.info("Starting Browser...")
            chrome_options = Options()
            if HEADLESS_MODE:
                chrome_options.add_argument("--headless=new")
                
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            driver = webdriver.Chrome(options=chrome_options)
            
            try:
                driver.get(LOGIN_URL)
                time.sleep(3)
                driver.find_element(By.ID, "txtUsername").send_keys(SITE_USERNAME)
                driver.find_element(By.ID, "txtPassword").send_keys(SITE_PASSWORD)
                driver.find_element(By.NAME, "commit").click()
                logging.info("Login Success.")
                time.sleep(5)
            except Exception as e:
                logging.error(f"Login failed: {e}")
                return False, driver

        try:
            for zip_file in zips_to_upload:
                logging.info(f"Uploading {zip_file}...")
                full_zip_path = os.path.abspath(os.path.join(zip_dir, zip_file))
                class_name = zip_file.replace(f" {upload_title}.zip", "")
                target_dropdown = map_folder_to_dropdown_text(class_name)

                if not target_dropdown: continue

                driver.get("https://app1.student-alert.com/index.php/administrator/homeworks/add_homework")
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "sectionId")))
                except:
                    logging.error("Page load timeout for add_homework")
                    continue

                if not force_select_hidden_dropdown(driver, "sectionId", target_dropdown): continue
                force_select_hidden_dropdown(driver, "subjectId", "Weekly Work Files")
                
                driver.find_element(By.NAME, "Title").send_keys(upload_title)
                
                date_box = driver.find_element(By.ID, "DueDate")
                date_box.click()
                time.sleep(0.5)
                date_box.send_keys(Keys.ARROW_DOWN)
                date_box.send_keys(Keys.ENTER)

                driver.find_element(By.ID, "upload").send_keys(full_zip_path)

                try:
                    wait = WebDriverWait(driver, 60)
                    ok_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'OK')]")))
                    driver.execute_script("arguments[0].click();", ok_btn)
                    logging.info("Upload Popup Clicked OK.")
                    time.sleep(1)
                except Exception as e:
                    logging.error(f"Popup 'OK' not found or upload timed out: {e}")
                    continue # Skip saving if upload failed

                save_btn = driver.find_element(By.ID, "btnSave")
                driver.execute_script("arguments[0].click();", save_btn)
                
                try:
                    WebDriverWait(driver, 15).until(EC.staleness_of(save_btn))
                    logging.info("Uploaded and saved.")
                except:
                    logging.warning("Save operation took longer than 15s or didn't trigger navigation.")
                    
                time.sleep(2) 

        except Exception as e:
            logging.error(f"Upload flow encounted an error: {e}")
        finally:
            # Explicitly clear temporary files after locks are released
            try: shutil.rmtree(temp_work_dir)
            except: pass

        return True, driver

# ==========================================
#      MAIN EXECUTION (RUN ONCE)
# ==========================================
def load_processed_ids():
    if not os.path.exists("processed_log.txt"): return set()
    with open("processed_log.txt", "r") as f: return set(line.strip() for line in f)

def save_processed_id(msg_id):
    with open("processed_log.txt", "a") as f: f.write(f"{msg_id}\n")

if __name__ == "__main__":
    print("--- STARTING ONE-CLICK BATCH RUN ---")
    
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")
        
        import sys
        
        status, messages = mail.search(None, '(OR SUBJECT "Week" SUBJECT "Revision")')
        
        if not messages or not messages[0]:
            logging.info("Done: No emails matching 'Week' or 'Revision' found in inbox.")
            mail.logout()
            time.sleep(1)
            sys.exit(0)

        email_ids = messages[0].split()
        
        processed_log = load_processed_ids()
        work_count = 0
        driver = None
        
        # Check ALL emails found
        for eid in email_ids: 
            eid_str = eid.decode()
            if eid_str in processed_log: continue 
            
            res, msg_data = mail.fetch(eid, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])
            subject_raw = decode_header(msg["Subject"])[0][0]
            subject = subject_raw.decode() if isinstance(subject_raw, bytes) else subject_raw
            
            upload_title = get_upload_title(subject)
            if upload_title:
                logging.info(f"*** New Email: {subject} ***")
                success, driver = process_email_and_upload(subject, upload_title, msg, driver)
                if success:
                    save_processed_id(eid_str)
                    work_count += 1
            else:
                save_processed_id(eid_str)

        mail.logout()
        
        if driver:
            driver.quit()
        
        if work_count == 0:
            logging.info("Done: No new work found in inbox.")
        else:
            logging.info(f"Done: Successfully processed {work_count} emails.")
            
    except Exception as e:
        logging.error(f"Global execution error: {e}", exc_info=True)

    input("\nPress Enter to close window...")