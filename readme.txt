# Texas Court Legacy Crawler (Demo)

This repository contains a **legacy demonstration script** used in an earlier phase of a Texas court data extraction project.  
It does **not** represent the current architecture, code quality, or tooling used in production.

---

##⚠️ Status: Legacy / Archived

- Uses **Internet Explorer WebDriver** (now deprecated)
- Relies on an **AutoIt** script to automate GUI clicks to download `.xls` files
- Script logic migrated into newer pipelines with modern browsers + fully automated data I/O
- Maintained only for reference and documentation

---

## 🧩 What This Code Shows

The main script provides examples for:

- Generating request URLs programmatically for:
  - County × Year × Month crawling ranges
- Automating report retrieval via Selenium
- Renaming downloaded reports using county metadata
- Tracking missing downloads for retry
- Cleaning up partial download artifacts

This is useful only for **historical understanding** of the workflow.

---

## 🔧 Requirements (Legacy)

- Python 3.x
- Selenium (IE WebDriver mode)
- pandas
- AutoIt (external executable for saving files)
- Ability to run IE in Edge Compatibility Mode on Windows

⚠️ Because IE is deprecated, setup is not recommended going forward.

