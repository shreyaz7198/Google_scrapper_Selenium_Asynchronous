import sys
import os
import time
import asyncio
import sqlite3
import aiohttp
from pathlib import Path
import pandas as pd

# Selenium Drivers & Core Modules
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# ==========================================
# 🛠️ SECTION 1: STORAGE & INITIALIZATION
# ==========================================

def get_default_storage_path(file_name: str) -> Path:
    """Gets a safe, cross-platform default path in the Downloads folder."""
    downloads_dir = Path.home() / "Downloads"
    if not downloads_dir.exists():
        downloads_dir = Path.home()
    return downloads_dir / file_name

def init_excel_file(path: Path):
    """Creates a structured Excel file with headers if it doesn't exist."""
    if not path.exists():
        headers = ["Timestamp", "Keyword", "Place", "Name", "Contact Number", "Address", "Plus Code", "Location", "Website", "Timings"]
        df = pd.DataFrame(columns=headers)
        df.to_excel(path, index=False)
        print(f"✅ Created fresh Excel template at: {path}")

def init_sqlite_db(path: Path):
    """Creates a SQLite database and the leads table if it doesn't exist."""
    conn = sqlite3.connect(path)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS leads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
            keyword TEXT,
            place TEXT,
            name TEXT,
            contact_number TEXT,
            address TEXT,
            plus_code TEXT,
            location TEXT,
            website TEXT,
            timings TEXT
        )
    """)
    conn.commit()
    conn.close()
    print(f"✅ Initialized SQLite Database template at: {path}")

# ==========================================
# 🎛️ SECTION 2: USER INTERFACE & FLOW
# ==========================================

def display_google_sheets_instructions() -> str:
    """Shows clear step-by-step instructions for Apps Script Web App integration."""
    print("\n" + "="*60)
    print("📋 GOOGLE SHEETS SETUP INSTRUCTIONS FOR NON-TECH USERS")
    print("="*60)
    print("1. Open Google Sheets (Ensure you are logged into ONE Google account only).")
    print("2. Create a brand new Spreadsheet (or open an existing one).")
    print("3. In the top menu, go to: Extensions -> Apps Script.")
    print("4. Erase any code inside the editor window and paste the exact code below:\n")
    
    apps_script_code = """function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp", "Keyword", "Place", "Name", "Contact Number", 
        "Address", "Plus Code", "Location", "Website", "Timings"
      ]);
    }
    
    var timestamp   = new Date();
    var keyword     = data.keyword || "";
    var place       = data.place || "";
    var name        = data.name || "";
    var contact     = data.contact_number || "";
    var address     = data.address || "";
    var plusCode    = data.plus_code || "";
    var location    = data.location || "";
    var website     = data.website || "";
    var timings     = data.timings || "";
    
    sheet.appendRow([timestamp, keyword, place, name, contact, address, plusCode, location, website, timings]);
    
    return ContentService.createTextOutput(JSON.stringify({"status": "success", "message": "Row added successfully"}))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}"""
    print(apps_script_code)
    print("\n" + "-"*60)
    print("5. Click 'Save' (the floppy disk icon).")
    print("6. Click the blue 'Deploy' button -> Select 'New deployment'.")
    print("7. Click the Gear icon (Select type) -> Choose 'Web app'.")
    print("8. Change 'Who has access' to: 'Anyone'.")
    print("9. Click 'Deploy', authorize permissions if asked, and COPY the generated 'Web app URL'.")
    print("="*60 + "\n")
    
    web_app_url = ""
    while not web_app_url.startswith("http"):
        web_app_url = input("🔗 Paste your deployed Google Web App URL here to proceed: ").strip()
        if not web_app_url.startswith("http"):
            print("⚠️ Invalid URL. It should start with http:// or https://")
            
    return web_app_url

def get_search_criteria():
    """Validates and collects keywords and locations from the user."""
    print("\n--- Search Target Setup ---")
    while True:
        keywords_raw = input("🔍 Enter Search Keywords (comma-separated): ").strip()
        keywords = [k.strip() for k in keywords_raw.split(",") if k.strip()]
        if keywords: break
        print("⚠️ You must enter at least one keyword!")

    while True:
        places_raw = input("📍 Enter Target Places/Locations (comma-separated): ").strip()
        places = [p.strip() for p in places_raw.split(",") if p.strip()]
        if places: break
        print("⚠️ You must enter at least one place!")
        
    return keywords, places

def run_setup_flow():
    print("=========================================")
    print("🚀 Welcome to the Modern Lead Scraper CLI")
    print("=========================================")
    print("How would you like your data to be generated?")
    print("1. Excel file in your system")
    print("2. Google Sheet")
    print("3. Database file (.db)")
    print("Press any other key to Exit.")
    
    choice = input("Select an option: ").strip()
    config = {"type": None, "path_or_url": None, "keywords": [], "places": []}
    
    if choice == "1":
        config["type"] = "excel"
        config["path_or_url"] = get_default_storage_path("maps_scraped_leads.xlsx")
        init_excel_file(config["path_or_url"])
    elif choice == "3":
        config["type"] = "sqlite"
        config["path_or_url"] = get_default_storage_path("maps_scraped_leads.db")
        init_sqlite_db(config["path_or_url"])
    elif choice == "2":
        config["type"] = "google_sheets"
        config["path_or_url"] = display_google_sheets_instructions()
    else:
        print("❌ Exiting application. Goodbye!")
        sys.exit(0)
        
    config["keywords"], config["places"] = get_search_criteria()
    return config

# ==========================================
# 💾 SECTION 3: ASYNC DATA STORAGE ROUTER
# ==========================================

async def save_scraped_data_async(config, payload, session=None):
    """Cleans up target artifacts and saves fields cleanly using async persistence routing."""
    def clean_text(text):
        if not text: return ""
        for char in ["", "", "", "🌍", "📞", "🕒"]:
            text = text.replace(char, "")
        return text.strip()

    cleaned_payload = {
        "keyword": payload["keyword"], "place": payload["place"], "name": clean_text(payload["name"]),
        "contact_number": clean_text(payload["contact_number"]), "address": clean_text(payload["address"]),
        "plus_code": clean_text(payload["plus_code"]), "location": payload["location"],
        "website": payload["website"], "timings": clean_text(payload["timings"])
    }

    if config["type"] == "excel":
        file_path = config["path_or_url"]
        excel_payload = {
            "Timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Keyword": cleaned_payload["keyword"], "Place": cleaned_payload["place"], "Name": cleaned_payload["name"],
            "Contact Number": cleaned_payload["contact_number"], "Address": cleaned_payload["address"],
            "Plus Code": cleaned_payload["plus_code"], "Location": cleaned_payload["location"],
            "Website": cleaned_payload["website"], "Timings": cleaned_payload["timings"]
        }
        df_new = pd.DataFrame([excel_payload])
        if file_path.exists():
            try:
                df_old = pd.read_excel(file_path)
                headers = ["Timestamp", "Keyword", "Place", "Name", "Contact Number", "Address", "Plus Code", "Location", "Website", "Timings"]
                df_old = df_old.reindex(columns=headers)
                df_final = pd.concat([df_old, df_new], ignore_index=True)
            except Exception:
                df_final = df_new
        else:
            df_final = df_new
        df_final.to_excel(file_path, index=False)
        
    elif config["type"] == "sqlite":
        conn = sqlite3.connect(config["path_or_url"])
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO leads (keyword, place, name, contact_number, address, plus_code, location, website, timings)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (cleaned_payload["keyword"], cleaned_payload["place"], cleaned_payload["name"], cleaned_payload["contact_number"], 
              cleaned_payload["address"], cleaned_payload["plus_code"], cleaned_payload["location"], cleaned_payload["website"], cleaned_payload["timings"]))
        conn.commit()
        conn.close()
        
    elif config["type"] == "google_sheets" and session:
        try:
            async with session.post(config["path_or_url"], json=cleaned_payload) as response:
                if response.status != 200:
                    text = await response.text()
                    print(f"⚠️ Google Sheets warning response: {text}")
        except Exception as e:
            print(f"❌ Failed to broadcast to Google Sheets async: {e}")

# ==========================================
# 🎭 SECTION 4: ENGINE 4 (SELENIUM ASYNC POOL)
# ==========================================

async def selenium_consumer_worker(worker_id, queue, config, session):
    """Asynchronous consumer worker that operates an isolated headless Chrome context."""
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0")
    chrome_options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    
    driver = await asyncio.to_thread(webdriver.Chrome, options=chrome_options)
    
    try:
        while True:
            task = await queue.get()
            if task is None:
                queue.task_done()
                break
                
            url, keyword, place = task
            try:
                await asyncio.to_thread(driver.get, url)
                await asyncio.sleep(2)  
                
                try: name = driver.find_element(By.TAG_NAME, "h1").text
                except: name = "Unknown Name"
                
                try: address = driver.find_element(By.CSS_SELECTOR, "button[data-item-id='address']").text
                except: address = ""
                
                try: contact = driver.find_element(By.CSS_SELECTOR, "button[data-item-id*='phone:tel']").text
                except: contact = ""
                
                try: website = driver.find_element(By.CSS_SELECTOR, "a[data-item-id='authority']").get_attribute("href")
                except: website = ""
                
                try: plus_code = driver.find_element(By.CSS_SELECTOR, "button[data-item-id='oloc']").text
                except: plus_code = ""
                
                try: timings = driver.find_element(By.CSS_SELECTOR, "div[aria-label*='Hours']").get_attribute("aria-label")
                except: timings = ""
                
                payload = {
                    "keyword": keyword, "place": place, "name": name, "contact_number": contact,
                    "address": address, "plus_code": plus_code, "location": driver.current_url, "website": website, "timings": timings
                }
                
                print(f"   ✨ [Selenium Worker {worker_id}] Scraped: {payload['name']}")
                await save_scraped_data_async(config, payload, session)
                
            except Exception:
                pass
            finally:
                queue.task_done()
    finally:
        await asyncio.to_thread(driver.quit)

async def run_async_selenium_scraper(config):
    print("\n🚀 Starting Phase 4: Selenium Asynchronous Concurrent Engine...")
    queue = asyncio.Queue(maxsize=30)
    
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0")
    chrome_options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    
    feed_driver = await asyncio.to_thread(webdriver.Chrome, options=chrome_options)
    
    async with aiohttp.ClientSession() as session:
        num_workers = 3
        workers = [asyncio.create_task(selenium_consumer_worker(i+1, queue, config, session)) for i in range(num_workers)]
        
        for keyword in config["keywords"]:
            for place in config["places"]:
                search_query = f"{keyword} {place}"
                print(f"\n⚡ [Producer] Fetching query feed: '{search_query}'")
                
                formatted_query = search_query.replace(" ", "+")
                await asyncio.to_thread(feed_driver.get, f"https://www.google.com/maps/search/{formatted_query}")
                await asyncio.sleep(4)
                
                processed_urls = set()
                scroll_attempts_without_new_data = 0
                max_scrolls = 8
                scrolls = 0
                
                while scrolls < max_scrolls:
                    elements = feed_driver.find_elements(By.CSS_SELECTOR, "a[href*='/maps/place/']")
                    current_urls = [el.get_attribute("href") for el in elements if el.get_attribute("href")]
                    new_urls = [url for url in current_urls if url not in processed_urls]
                    
                    if new_urls:
                        scroll_attempts_without_new_data = 0
                        print(f"🔄 [Producer] Found {len(new_urls)} new profiles. Appending to queue...")
                        for url in new_urls:
                            processed_urls.add(url)
                            await queue.put((url, keyword, place))
                    else:
                        scroll_attempts_without_new_data += 1
                        
                    try:
                        panel = feed_driver.find_element(By.CSS_SELECTOR, "div[role='feed']")
                        feed_driver.execute_script("arguments[0].scrollBy(0, arguments[0].scrollHeight);", panel)
                    except:
                        break
                        
                    await asyncio.sleep(2.5)
                    if scroll_attempts_without_new_data >= 3:
                        break
                    scrolls += 1
                    
        await queue.join()
        for _ in range(num_workers): await queue.put(None)
        await asyncio.gather(*workers)
        
    await asyncio.to_thread(feed_driver.quit)
    print("\n🏁 Phase 4 Async Processing Complete. System clean.")

# ==========================================
# 🚀 SECTION 5: MASTER RUNNER
# ==========================================

if __name__ == "__main__":
    session_config = run_setup_flow()
    asyncio.run(run_async_selenium_scraper(session_config))