# Gmail-to-Excel-Automation
This Python script reads emails from your Gmail inbox and exports key details  (From, Subject, Date) into an Excel file (gmail_emails.xlsx).
# 📧 Gmail to Excel Automation using Python

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python)
![License](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/Status-Active-success)
![Contributions](https://img.shields.io/badge/Contributions-Welcome-orange)

## 🧭 Overview

This Python project automates the process of **reading Gmail messages** and **exporting key email details** (From, Subject, Date) into an Excel spreadsheet.

It uses the official **Gmail API** from Google for secure, OAuth 2.0–based access.  
Each run appends new emails to your Excel log, making it ideal for **email tracking, business correspondence logs**, or **automated data extraction** workflows.

---

## ✨ Features

- ✅ Connects securely to your Gmail account using OAuth 2.0  
- ✅ Extracts **From**, **Subject**, and **Date** from recent messages  
- ✅ Saves or appends email data to `gmail_emails.xlsx`  
- ✅ Stores a persistent `token.json` so re-login is not required each time  
- ✅ Written in clean, modular, well-documented Python  
- ✅ Easily extendable for automation, analytics, or reporting

---

## ⚙️ Prerequisites

- A Google account (Gmail)
- Python **3.9 or later**
- Gmail API enabled in your Google Cloud project

---

## 🚀 Setup Instructions

### 1️⃣ Clone the repository

```bash
git clone https://github.com/<your-username>/gmail-to-excel.git
cd gmail-to-excel
```

### 2️⃣ Install dependencies

```bash
pip install -r requirements.txt
```

### 3️⃣ Enable Gmail API and download credentials

1. Go to Google Cloud Console → Credentials
2. Create a new **OAuth 2.0 Client ID**
  * Application type: **Desktop App**
3. Download the `credentials.json`  file
→ Save it in your project root directory (same folder as the script)
4. Add your Gmail address as a Test User under
OAuth consent screen → Test users.

### ▶️ Run the Script
```bash
python gmail_to_excel.py
```
* On first run, a browser window will open for Google sign-in.
* Approve access to your Gmail (this creates `token.json` automatically).
* The script then reads your emails and generates:
```bash
gmail_emails.xlsx
```
Each row will contain:
* **From (Sender)** 
* **Subject**
* **Date**

**Once running, the tool will continue updating your Excel file with each new batch of emails — no extra steps needed.**
