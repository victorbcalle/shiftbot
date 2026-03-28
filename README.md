## Shiftbot
Serverless Python bot powered by Playwright and Google Cloud APIs. Autonomously scrapes corporate shift portals, syncs schedules with Google Calendar, and generates real-time payroll forecasts in Google Sheets. Deployed via GitHub Actions with secure OAuth 2.0 and CI/CD cron jobs for zero-maintenance automation.

Autonomous Shift Sync & Payroll Engine

> **Transforming legacy corporate portals into modern, automated workflows.**

## The "Why": Building to Solve Real Problems
Necessity is the mother of invention. Like many professionals doing shift work, I found myself constantly logging into a clunky corporate portal just to check my schedule. Writing down shifts manually was tedious, and keeping track of variable payroll data (night shifts, holidays, split shifts) was a logistical headache prone to human error.

I built this project to scratch my own itch: **I wanted a system that worked for me, not the other way around.** What started as a personal script has evolved into a robust, serverless automation engine. It proves that even legacy systems without native APIs can be seamlessly integrated into modern cloud ecosystems.

## What Problem Does This Solve?
This bot acts as a **24/7 personal assistant and accountant**. It eliminates manual data entry and provides total financial clarity by doing the following autonomously:
1. **Never Miss a Shift:** Automatically extracts the latest schedule from the company portal and syncs it directly to Google Calendar.
2. **Financial Forecasting:** Calculates exact gross payroll estimates before the month ends, accounting for base rates, night hours, holiday multipliers, and transport bonuses.
3. **Data Preservation:** Maintains a historical record in Google Sheets, ensuring that past shifts and manually inputted overtime are never overwritten.

## The "How": Technical Architecture
This project is built with scalability, security, and zero-maintenance in mind. It runs entirely in the cloud without the need for a dedicated server.

* **Core Language:** `Python 3.11`
* **Headless Automation (Scraping):** `Playwright` is used to navigate the corporate portal, handle secure logins, and extract HTML table data dynamically.
* **Cloud Infrastructure:** `Google Cloud Platform (GCP)`. Integrates natively via OAuth 2.0 with:
  * *Google Calendar API:* For real-time schedule injection and event color-coding.
  * *Google Sheets API:* For dynamic matrix reconstruction and payroll calculations.
  * *Google Drive API:* For automated folder and file lifecycle management.
* **CI/CD & Serverless Execution:** Deployed on `GitHub Actions`. A Cron Job triggers the pipeline every hour, creating a truly autonomous, "set-it-and-forget-it" architecture.
* **Security:** Strict separation of concerns. Passwords and API tokens are completely decoupled from the codebase, stored safely via GitHub Secrets and injected at runtime as environment variables.

---
*Disclaimer: This is a showcase project built for educational and portfolio purposes. It operates in a private cloud environment and does not distribute third-party data.*
