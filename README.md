# Sheets Auto-Fill Engine

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-Enabled-4285F4.svg)](https://developers.google.com/apps-script)

A lightweight, data-driven Google Sheets extension that allows users to create complex, dynamic auto-fill logic without writing a single line of code. 

Instead of hardcoding `if/else` statements into a fragile `onEdit` trigger, this tool provides a custom sidebar GUI where users can define rules. These rules are compiled into a JSON array, stored securely in the document's properties, and executed flawlessly by a single, highly optimized background engine.

## Developer Note

I was just throwing together a few Google Sheets earlier today and found that there's no easy way to automatically add a date to a cell, and have it not update every time you make a change to the file. The easiest way is to use the Apps Script editor, which is not ideal for anyone who's not a developer. Two hours of prompting later and I've got a solution that works very similarly to the Data Validation feature, which can auto fill not only dates, but a few other things with a variety of trigger types. Still a work in progress, and has zero styling right now, but it works for me, so I thought it could work for someone else. 👍

## ✨ Features

* **Visual Rule Builder:** Create, edit, clone, toggle, and delete rules from a responsive sidebar UI.
* **Row-Based Logic Engine:** Highly optimized engine only evaluates the specific rows being edited, avoiding heavy full-sheet scans. Safely handles multi-cell pastes and bulk deletions.
* **Granular Sheet Scopes:** Apply rules globally across "All Sheets" or restrict them to a specific tab.
* **Smart Detection Types:**
  * Fill and Update (updates timestamp/value on any change)
  * Fill When Full (requires all cells in detection range to have data)
  * Fill When Full, Clear when Empty
  * Fill Once (prevents overwriting existing data)
  * Specific Value Match (triggers on exact text matches)
  * Is Checked (specifically watches checkbox columns)
* **Versatile Fill Payloads:**
  * Current Date/Timestamp
  * Static Text or Numbers
  * Unique IDs (UUIDs) for instant ticketing
  * Editor's Email (Audit logging)
  * Data Copying (Pull value from another column in the same row)
  * Smart Checkboxes (Safely inserts and toggles boolean values)
* **Portable Configurations:** Import and Export your entire rule set as raw JSON to share across different spreadsheets.

## 🛠 Architecture

This project is built using vanilla JavaScript and Google Apps Script's `HtmlService`. It utilizes a **data-driven interpreter pattern**. The frontend `Sidebar.html` acts as the compiler, turning user inputs into a structured JSON payload. The backend `Code.gs` acts as the virtual machine, listening for Google Sheets' `onEdit` event, pulling the JSON ruleset from `PropertiesService`, and executing the corresponding logic.

## 🚀 Getting Started

To install this tool on your own Google Sheet before it is available on the Google Workspace Marketplace:

1. Open your target Google Sheet.
2. Navigate to **Extensions > Apps Script**.
3. Replace the contents of the default `Code.gs` file with the `Code.gs` from this repository.
4. Create a new HTML file in the Apps Script editor by clicking the **+** icon next to Files, and name it `Sidebar` (this will create `Sidebar.html`).
5. Paste the contents of `Sidebar.html` from this repository into that file.
6. Save the project (Ctrl+S / Cmd+S).
7. Refresh your Google Sheet. You will now see a new menu item called **Auto-Fill Rules** near the Help tab.

### Initialization

For the engine to track edits outside of the sidebar's active session, you must authorize the background trigger:
1. Click **Auto-Fill Rules > Setup Trigger (Run Once)**.
2. Grant the necessary Google Workspace permissions when prompted.

## 💻 Usage

1. Open the sidebar via **Auto-Fill Rules > Configure Auto-Fill Rules**.
2. Click **+ Create New Rule**.
3. Define your **Fill Range** (Target) and **Detection Range** (Condition). *Note: Use standard A1 notation (e.g., `A2:A` or `C:C`). Starting and ending rows between the two ranges must match to prevent offset errors.*
4. Select your Fill Type and Detection Type.
5. Click **Save Rule**.

If you make manual adjustments to the sheet layout or import a large dataset and need the rules to evaluate everything retroactively, click the **⚡ Recalculate Active Sheet** button in the sidebar.

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! 
1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

Distributed under the [MIT License](LICENSE).
