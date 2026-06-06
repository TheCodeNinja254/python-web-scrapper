# 🏫 Manitoba Child Care Facility Scraper

Hello, quick intro here. The following is a technical description of the steps to take from a very green start: initial setup to scraping data and refreshing an excel sheet with the data.

I have added a disclaimer since this code scrapes code from a government site. The disclaimer indicates that this is for research purposes only.

This tool automatically collects information on **all licensed child care facilities in Manitoba** from the [Manitoba Child Care Search](https://childcaresearch.gov.mb.ca/) website and saves everything into a neatly formatted **Excel spreadsheet** (`.xlsx`).

Each time you run it, it refreshes the spreadsheet with the latest data — including any newly added facilities.

---

## 📋 What You'll Need

Before you start, you need to install two free programs on your Windows computer:

| Program | What it is | Download |
|---|---|---|
| **Python** | The language this tool is written in | [python.org](https://www.python.org/downloads/) |
| **PyCharm Community** | A free app for running Python projects | [jetbrains.com/pycharm](https://www.jetbrains.com/pycharm/download/) |
| **Git** | A tool for downloading code from GitHub | [git-scm.com](https://git-scm.com/download/win) |

Don't worry — all three are free and the steps below will walk you through each one.

---

## 🪟 Step 1 — Install Python

1. Go to [python.org/downloads](https://www.python.org/downloads/) and click the big yellow **"Download Python"** button.
2. Run the downloaded file (it will be named something like `python-3.x.x-amd64.exe`).
3. ⚠️ **Important:** On the first screen of the installer, tick the checkbox that says **"Add Python to PATH"** before clicking Install Now.

   ![Add Python to PATH checkbox](https://docs.python.org/3/_images/win_installer.png)

4. Click **"Install Now"** and wait for it to finish.
5. Click **"Close"** when done.

**To check it worked:** Press `Windows key + R`, type `cmd`, press Enter. In the black window that opens, type:
```
python --version
```
You should see something like `Python 3.12.3`. If you do, Python is installed! ✅

---

## 🖥️ Step 2 — Install PyCharm

You will need this to make further edits to the code: 

1. Go to [jetbrains.com/pycharm/download](https://www.jetbrains.com/pycharm/download/).
2. Scroll down to **"PyCharm Community Edition"** (this is the free version) and click **Download**.
3. Run the downloaded installer.
4. Click **Next** through all the steps — the default options are fine.
5. When asked about **"Create Desktop Shortcut"**, tick that box so it's easy to find later.
6. Click **Install**, then **Finish**.

---

## 📦 Step 3 — Install Git

Git is what lets you download ("clone") this project from GitHub.

1. Go to [git-scm.com/download/win](https://git-scm.com/download/win) — the download should start automatically.
2. Run the installer.
3. Click **Next** through all the steps — the default options are perfectly fine.
4. Click **Install**, then **Finish**.

**To check it worked:** Open Command Prompt again (`Windows key + R` → `cmd` → Enter) and type:
```
git --version
```
You should see something like `git version 2.x.x`. ✅

---

## 📥 Step 4 — Download This Project from GitHub

This step copies all the project files onto your computer.

1. Press `Windows key + R`, type `cmd`, and press Enter to open Command Prompt.
2. Decide where you want to save the project. For example, your Documents folder. Type this to go there:
   ```
   cd C:\Users\YourName\Documents
   ```
   *(Replace `YourName` with your actual Windows username)*

3. Now download the project by typing:
   ```
   git clone https://github.com/TheCodeNinja254/python-web-scrapper.git
   ```
4. Press Enter. You'll see files being downloaded. When it finishes, a new folder will appear in your Documents.

5. Move into the project folder:
   ```
   cd YOUR-REPO-NAME
   ```

---

## 🔧 Step 5 — Open the Project in PyCharm

1. Open **PyCharm** (find it on your Desktop or in the Start Menu).
2. On the Welcome screen, click **"Open"**.
3. Browse to your Documents folder, find the project folder you just downloaded, and click **OK**.
4. The project will open. You'll see the file `mb_childcare_scraper.py` in the left panel.

---

## ⚙️ Step 6 — Set Up a Virtual Environment

A virtual environment keeps this project's tools separate from the rest of your computer. PyCharm makes this easy.

1. In PyCharm, go to **File → Settings** (or press `Ctrl + Alt + S`).
2. In the left panel, click **Project: your-project-name → Python Interpreter**.
3. Click the gear icon ⚙️ at the top right, then click **"Add..."**.
4. Select **"Virtualenv Environment"**, make sure **"New environment"** is selected, and click **OK**.
5. Click **OK** again to close Settings.

PyCharm has now created a clean, isolated Python environment for this project. ✅

---

## 📚 Step 7 — Install Dependencies

The project needs two extra Python libraries. Let's install them using PyCharm's built-in terminal.

1. In PyCharm, open the Terminal at the bottom of the screen. Click **Terminal** in the bottom toolbar (or go to **View → Tool Windows → Terminal**).
2. You should see a prompt like `(venv) C:\Users\...`. The `(venv)` part means your virtual environment is active. ✅
3. Type this and press Enter:
   ```
   pip install playwright openpyxl
   ```
   Wait for it to finish downloading and installing.

4. Now install the browser that the scraper uses. Type this and press Enter:
   ```
   python -m playwright install chromium
   ```
   This downloads a small version of the Chrome browser. It may take a minute or two.

5. Now install requests. Type this and press Enter:
   ```
   pip install requests
   ```
---

## ▶️ Step 8 — Run the Scraper

You're all set! Here's how to run the tool.

### Option A — Quick Test First (Recommended)

Before running the full scrape, do a quick test with just 10 facilities to make sure everything is working:

1. In the PyCharm Terminal, type:
   ```
   python mb_childcare_scraper.py --test
   ```
2. You'll see messages like:
   ```
   🔍  Starting Manitoba Child Care scraper...
       Found 847 facility IDs
       [TEST MODE] Limiting to 10 facilities
     [  1/10] Scraping facility 9648...
   ```
3. When it finishes, a file called `manitoba_childcare.xlsx` will appear in your project folder. Open it in Excel to check the data looks correct.

### Option B — Full Manitoba Scrape

Once the test works, run the full scrape:

```
python mb_childcare_scraper.py
```

> ⏱️ **This will take 30–60 minutes** depending on your internet speed. There are hundreds of facilities and the tool visits each one individually. You can leave it running and come back later.

When it completes, you'll see:
```
✅  Scraped 847 facilities  (0 errors)
✅  Saved 847 facilities → manitoba_childcare.xlsx
```

---

## 📊 What's in the Excel File?

Open `manitoba_childcare.xlsx` in Microsoft Excel. You'll find:

| Sheet | Contents |
|---|---|
| **Facilities** | All scraped data — one row per facility |
| **▶ How to Refresh** | Quick reference instructions |

Each row in the **Facilities** sheet contains:

- Facility name, type, and ID
- Full address, region, area, and neighbourhood
- Phone number, email, and website
- Available spaces (infant, preschool, nursery, school age)
- Operating hours (weekdays, weekends, evenings, overnight)
- Current job openings (CCA, ECE II, ECE III)
- Direct link to the facility's profile page
- Date and time the data was scraped

---

## 🔄 Refreshing the Data

Run the scraper again any time you want updated information:

```
python mb_childcare_scraper.py
```

It will **overwrite** the Facilities sheet with fresh data. Any new day cares added to the government site will automatically appear.

### 🕐 Automatic Daily Refresh (Optional)

If you want the spreadsheet to update automatically every day without you doing anything:

1. Press `Windows key + S` and search for **"Task Scheduler"**, then open it.
2. Click **"Create Basic Task..."** in the right panel.
3. Give it a name like `Manitoba Childcare Refresh` and click **Next**.
4. Choose **Daily** and click **Next**. Set your preferred time (e.g. 6:00 AM).
5. Choose **"Start a program"** and click **Next**.
6. In the **Program/script** box, type:
   ```
   python
   ```
7. In the **Add arguments** box, type the full path to the scraper, for example:
   ```
   C:\Users\YourName\Documents\YOUR-REPO-NAME\mb_childcare_scraper.py
   ```
8. Click **Next**, then **Finish**.

The spreadsheet will now refresh itself every morning. ✅

---

## ❓ Troubleshooting

**"Python is not recognized as a command"**
→ Python wasn't added to PATH during installation. Uninstall Python and reinstall it, making sure to tick **"Add Python to PATH"** on the first screen.

**"No module named playwright" or "No module named openpyxl"**
→ The dependencies aren't installed. Make sure you ran Step 7 in the PyCharm Terminal (with `(venv)` showing in the prompt).

**The scraper runs but the Excel file is mostly empty**
→ The website may have changed its layout. Try the `--test` flag first and check what data comes through. Feel free to open an issue on GitHub.

**The scraper stops halfway through**
→ This can happen with slow internet. Just run it again — it will start fresh. You can also check your internet connection and try again.

**Excel shows the file but it won't open**
→ Make sure the file isn't already open in Excel when the scraper runs. Close Excel first, then run the script.

---

## 📁 Project Files

```
your-project-folder/
│
├── mb_childcare_scraper.py    ← The main scraper script
├── manitoba_childcare.xlsx    ← Output file (created when you run the scraper)
└── README.md                  ← This guide
```

---

## 📜 Data Notice

This tool collects publicly available data from the [Manitoba Child Care Search](https://childcaresearch.gov.mb.ca/) portal, which is operated by the Government of Manitoba. The data belongs to the Government of Manitoba. This tool is intended for personal research and informational use only.

For questions about the data itself, contact: **mbchildcaresearch@gov.mb.ca**

---

*Built with [Python](https://python.org), [Playwright](https://playwright.dev/python/), and [openpyxl](https://openpyxl.readthedocs.io/).*