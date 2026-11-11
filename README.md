# College Web Scraper & Two-Sheet Excel Exporter (`web tool`)

A self-contained web dashboard and live internet scraper that collects **Government** and **Private** colleges (Polytechnic, ITI, B.Voc & D.Voc, Engineering, Medical/Nursing/Pharmacy, Management, Graduation/PG, or any custom course) along with **College Name, Contact No, Email ID, District, Pincode, University/Board, Official Website, and Course Offered**, and exports them into a two-sheet Excel file (`Government` and `Private`).

## 1. Run Locally on Windows
1. Double-click **`start_dashboard.bat`** (or run `python app.py` inside this folder).
2. It automatically opens **`http://127.0.0.1:8765`** in your browser (avoiding port 8080 which is used by EDB Postgres).
3. Choose the category/course, district, or enter any custom course keyword or college URLs, click **1. Scrape & Collect Data**, and click **2. Download Excel (.xlsx)**.

## 2. Host for Free on the Internet
This folder is pre-configured with standard WSGI (`app:app`), `requirements.txt`, `Procfile`, and `render.yaml` so you can host it for free on:
- **Render.com (Free Tier):** Push this `web tool` folder to a GitHub repository -> New Web Service on Render -> Connect Repo -> Render automatically detects `render.yaml` / `Procfile` and deploys for free.
- **PythonAnywhere (Free Forever Tier):** Upload `app.py` and `seed_data.json`, install `openpyxl`, and point the WSGI configuration file to `from app import app as application`.
- **Railway / Koyeb / Hugging Face Spaces:** Connect the repository and it runs automatically via `Procfile`.
