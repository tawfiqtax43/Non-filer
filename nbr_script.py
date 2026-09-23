import os
import time
import json
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

OUTPUT_EXCEL = "Missing_Return_Report.xlsx"
PROGRESS_FILE = "nbr_progress.json"

def find_excel_file():
    """ফোল্ডারে থাকা যেকোনো .xlsx ফাইল খুঁজে বের করবে"""
    excel_files = glob.glob("*.xlsx")
    # আউটপুট ফাইল বাদ দিয়ে মেইন ফাইল খোঁজা
    valid_files = [f for f in excel_files if f not in [OUTPUT_EXCEL, "Missing_Return_Report_Backup.xlsx"]]
    if not valid_files:
        return None
    return valid_files[0]

def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--window-size=1920,1080")
    options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def check_single_tin_year(driver, tin, assessment_year):
    driver.get("https://etaxnbr.gov.bd/#/submission-verification")
    wait = WebDriverWait(driver, 10)
    tin_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text']")))
    
    tin_input.clear()
    tin_input.send_keys(str(tin).strip())

    year_select = driver.find_element(By.XPATH, "//select")
    year_select.send_keys(assessment_year)

    verify_btn = driver.find_element(By.XPATH, "//button[contains(text(),'Verify')]")
    driver.execute_script("arguments[0].click();", verify_btn)

    time.sleep(1.8)
    page_text = driver.page_source
    if "No matching return was found" in page_text:
        return True 
    return False 

def check_with_retry(tin, assessment_year, max_retries=3):
    for attempt in range(1, max_retries + 1):
        driver = None
        try:
            driver = create_driver()
            res = check_single_tin_year(driver, tin, assessment_year)
            return res
        except Exception:
            pass
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
        if attempt < max_retries:
            time.sleep(attempt * 2)
    return None

def save_progress(data_list):
    if data_list:
        df = pd.DataFrame(data_list)
        try:
            df.to_excel(OUTPUT_EXCEL, index=False)
        except PermissionError:
            df.to_excel("Missing_Return_Report_Backup.xlsx", index=False)
            
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(data_list, f, ensure_ascii=False, indent=4)

def run_nbr_automation():
    excel_file = find_excel_file()
    if not excel_file:
        print("\n[ERROR] ফোল্ডারে কোনো এক্সেল (.xlsx) ফাইল পাওয়া যায়নি!")
        print("দয়া করে TIN সম্বলিত এক্সেল ফাইলটি এই ফোল্ডারে রেখে আবার চেষ্টা করুন।")
        return

    print(f"\nপাওয়া গেছে এক্সেল ফাইল: '{excel_file}'")
    df = pd.read_excel(excel_file)
    
    # TIN কলামটি খুঁজে নেওয়া
    tin_column = None
    for col in df.columns:
        if 'tin' in str(col).lower():
            tin_column = col
            break
            
    if not tin_column:
        print("[ERROR] এক্সেলে 'TIN' নামের কোনো কলাম পাওয়া যায়নি!")
        return

    # সাল নির্ধারণ (Issue Date না থাকলে বর্তমান সাল অনুযায়ী গত ৪-৫ বছর চেক করবে)
    # বর্তমান বছর ২০২৬ হলে ২০২১-২০২২ থেকে ২০২৪-২০২৫ পর্যন্ত চেক করবে
    years_to_check = ["2024-2025", "2023-2024", "2022-2023", "2021-2022"]

    final_data = []
    processed_tins = set()

    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                saved_data = json.load(f)
                if saved_data:
                    print(f"\n[INFO] আগে প্রসেস করা {len(saved_data)} টি ডাটা পাওয়া গেছে।")
                    choice = input("আপনি কি আগের কাজ থেকে শুরু (Resume) করতে চান? (y/n): ").strip().lower()
                    if choice == 'y':
                        final_data = saved_data
                        processed_tins = {str(row["TIN"]).strip() for row in final_data}
                        print(f"✓ {len(processed_tins)} টি TIN স্কিপ করে বাকিগুলো প্রসেস করা হচ্ছে...\n")
        except Exception:
            final_data = []

    total_rows = len(df)
    print(f"মোট প্রসেস করা হবে: {total_rows} টি ডাটা...\n")

    try:
        for idx, row in df.iterrows():
            tin = str(row[tin_column]).strip()
            name = row.get("Name", "N/A")

            if tin in processed_tins:
                continue

            print(f"[{idx+1}/{total_rows}] প্রসেস হচ্ছে - TIN: {tin}...")

            missing_years = []
            for assessment_year in years_to_check:
                is_missing = check_with_retry(tin, assessment_year)
                if is_missing is True:
                    missing_years.append(assessment_year)

            final_data.append({
                "TIN": tin,
                "Name": name,
                "যেসব সালের রিটার্ন দেয় নাই": ", ".join(missing_years) if missing_years else "সব সালের রিটার্ন জমা দেওয়া আছে"
            })

            save_progress(final_data)
            print(f"   ✓ সেভ করা হয়েছে। অনুপস্থিত বছর: {', '.join(missing_years) if missing_years else 'নাই'}")
            time.sleep(1)

        print(f"\nসম্পূর্ণ প্রসেস সফলভাবে শেষ হয়েছে! ফাইল সেভ করা হয়েছে: {OUTPUT_EXCEL}")
        if os.path.exists(PROGRESS_FILE):
            os.remove(PROGRESS_FILE)

    except KeyboardInterrupt:
        print("\n\n[WARNING] ইউজার কর্তৃক প্রসেস থামানো হয়েছে (Ctrl+C)!")
        save_progress(final_data)
        print(f"✓ এ পর্যন্ত {len(final_data)} টি ডাটা '{OUTPUT_EXCEL}' ফাইলে সুরক্ষিত আছে।")

if __name__ == "__main__":
    run_nbr_automation()
