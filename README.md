# GBTS Daily Report Generator

![Python](https://img.shields.io/badge/python-3.8%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

A Python script to automatically generate the **GBTS Daily Report** in PDF by downloading logbook data from SharePoint and MPDS, processing it, and creating a comprehensive report.

---

## 📝 Features

- Download logbooks and data from SharePoint Online using Microsoft Graph API.
- Skip report generation on holidays and weekends automatically.
- Analyze today’s and tomorrow’s simulator activities.
- Generate a well-formatted PDF including:
  - Simulator status
  - Completed and scheduled training
  - RTMS, MPDS, and upcoming activity overview
  - Legend and summary tables
- Optional email notification of report.

---

## ⚙️ Requirements

- Python 3.8+
- Python libraries:
  - `tkinter` (for optional GUI)
  - `tkcalendar`
  - `reportlab`
  - `python-dotenv`
  - `requests`
- SharePoint Online access with **Client ID, Client Secret, and Tenant ID**.
- `.env` file with paths and credentials.

---

## 📂 Project Structure
project/
│
├─ main.py # Main script to generate the report
├─ GBTS_Daily_Report.pdf # Generated PDF
├─ env.env # Environment variables (dest_file, credentials)
├─ images/
│ └─ ajt_official.ico # Logo for the report
└─ modules/
├─ data.py # Handles downloading and reading files
├─ log_definition.py # Custom logging
└─ pdf_dev.py # PDF generation


---

## 🚀 Installation

 Clone the repository:
```bash
git clone https://github.com/<username>/gbts-daily-report.git
cd gbts-daily-report
---

## 🚀 Installation

1. Clone the repository:
```bash
git clone https://github.com/<username>/gbts-daily-report.git
cd gbts-daily-report

## 🖥️ Usage
- From Terminal
  python main.py --today 2025-10-21 --tomorrow 2025-10-22

- With GUI (Tkinter)
Run the GUI to select dates: python gui.py


