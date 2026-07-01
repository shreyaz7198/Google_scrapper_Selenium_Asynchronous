# Google_scrapper_Selenium_Asynchronous

## What This Tool Does
Think of this script as a digital assistant. It automatically **opens a hidden web browser**, visits **Google Maps**, types in **any business you want to find** (like "Dentists" or "Cafes"), and scans the results. It copies down their **Names, Phone Numbers, Addresses, Websites, and Working Hours**, and neatly **saves them for you** line-by-line.

## Behind the Scenes (What makes it work)
* **Python:** The underlying engine that runs our instructions.
* **Pandas and Openpyxl:** The digital accountants that organize your data and build standard Excel spreadsheets.
* **Selenium 4 Multi-Threading:** Spreads blocking web page load commands onto independent background thread tasks safely.
* **Aiohttp:** The ultra-fast delivery tool that lets the background browsers send multiple lines of data to your Google Sheet at the exact same time without slowing down.

## Step-by-Step Guide: How to Download and Use
### Step 1: Install the Machinery
Before you can type commands into your terminal, your computer needs to have the actual software machinery installed. If you skip this, your computer will give you a "command not found" error.

* **Install Git** (The tool that downloads code from GitHub)
  * **Windows:** Go to [git-scm.com](https://git-scm.com/), download the installer, and click "Next" through the setup prompts (the default settings are perfect).
  * **Mac:** Open your Terminal app, type `git`, and press Enter. A pop-up box will appear. Click **Install** to let your Mac set it up automatically.
* **Install Python** (The engine that runs our code)
  * Go to [python.org/downloads](https://www.python.org/downloads/) and click the yellow button to download the latest version for your computer.
  * Run the downloaded installer file.
  
> ⚠️ **CRUCIAL FOR WINDOWS USERS:** On the very first screen of the installer window, look at the bottom and check the box that says **"Add python.exe to PATH"**. If you do not check this box, your computer will not understand your Python commands later!

#### How to confirm it worked:
Close your Terminal or Command Prompt window, open it back up completely fresh, and run:
```bash
python --version
git --version
```
 
### Step 2: Download the Tool to Your Computer
* Open your internet browser and go to project's GitHub webpage
[Google_scrapper_python](https://github.com/shreyaz7198/Google_scrapper_python)
* Click the green Code button and copy the web address link shown there.
* Open the Terminal app (if you use a Mac) or the Command Prompt / PowerShell app (if you use Windows).
* Type this command, then press Enter on your keyboard to copy the files to your computer:
```bash
git clone https://github.com/shreyaz7198/Google_scrapper_python.git
```
* Move your terminal inside the newly downloaded folder by typing:
```bash
cd Google_scrapper_python
```
### Step 3: Set Up a Safe Virtual environment
* Before running the script, we need to create an isolated environment. This ensures the script's settings don't mix up or interfere with any other files on your computer.
* Create your digital(virtual) environment by running:

```bash
python -m venv venv
```
* Turn the environment on so your computer uses it: **Windows users**: Type 
```bash
venv\Scripts\Activate.ps1
``` 
and press Enter.

* Turn the environment on so your computer uses it: **Mac users**: Type source 
```bash
venv/bin/activate
``` 
and press Enter.

#### (You will know it worked because (venv) will now appear at the start of your text line).
* Update your computer's helper tool:
```bash
python -m pip install --upgrade pip
```
* Install all the automated puzzle pieces required for this script:
```bash
pip install -r requirements.txt
```

### Step 4: Run the Program
* Start the script by typing:
```bash
python generator.py
```
* A welcome menu will pop up on your screen asking **where you want to save your final results**:
  * **Type 1:** Creates a regular Excel file on your computer.
  * **Type 2:** Sends the data straight to a live online Google Sheet.
  * **Type 3:** Saves everything inside a secure local database file.

### Step 5: Setting Up Your Google Sheet (Only if you chose Option 2)
* To send data straight to an online spreadsheet, we need to give our tool a secret link so it knows exactly where your sheet lives.
* Go to your internet browser and open a brand-new, empty Google Sheet.
* At the top menu, click Extensions -> Apps Script.
* Delete any text currently sitting inside the editor box and paste this exact block of code:
```
function doPost(e) {
try {
var data = JSON.parse(e.postData.contents);
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
if (sheet.getLastRow() === 0) {
  sheet.appendRow(["Timestamp", "Keyword", "Place", "Name", "Contact Number", "Address", "Plus Code", "Location", "Website", "Timings"]);
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
return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
} catch (error) {
return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()})).setMimeType(ContentService.MimeType.JSON);
}
}
```
* Click the small Save icon (the floppy disk).
* Click the blue Deploy button in the top right -> choose New deployment.
* Click the Gear icon next to "Select type" -> choose Web app.
* Change the bottom setting labeled "Who has access" to Anyone. (If you skip this, your sheet will lock the script out).
* Click Deploy, click "Authorize Access" if Google asks for permission, and copy the long Web App URL link it gives you.
* Paste that long link directly into your terminal window where the script is waiting, and press Enter.

### Step 6: Tell the Tool What to Look For
* The tool will ask you to type your search **keywords** (for example: Dentists, Cafes, Bakeries).
* Next, it will ask you to type the target cities or **locations** (for example: Nagpur, London).
* Press Enter, and the tool will quietly begin gathering leads for you in the background!

### Step 7: Where to See Your Collected Data
* If you chose Excel (**Option 1**): Go look inside your computer's normal Downloads folder. You will find a spreadsheet named maps_scraped_leads.xlsx full of your data.
* If you chose Google Sheet (**Option 2**): Keep your Google Sheet open in your browser tab—you will literally watch rows of data type themselves out on your screen in real-time.
* If you chose Database (**Option 3**): A file named maps_scraped_leads.db will appear inside your Downloads folder. Read below to see how to open it.

## How to View Your Database (.db) File
* Database files are tightly locked, organized data safes. You cannot open them using Microsoft Word or Excel.
* Download a free, safe program called DB Browser for SQLite (https://sqlitebrowser.org/) and install it on your computer.
* Open the app and click the Open Database button at the top of the window.
* Go to your computer's Downloads folder and select your maps_scraped_leads.db file.
* Click on the tab near the top named Browse Data. Your database will open up like a beautiful, clean, searchable spreadsheet!
