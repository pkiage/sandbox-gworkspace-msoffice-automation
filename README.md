# sandbox‑gworkspace‑msoffice‑automation

This repo is a personal experimental “sandbox” of scripts that automate Google Workspace and Microsoft Office tasks:
- Streamline repetitive Google Workspace chores with Apps Script
- Batch‑process files in Google Colab using the Google Drive & Docs APIs
- Audit Excel workbooks and PowerPoint decks with VBA macros

> **Why?**  
> Automation on the *right things* in the *right way* can save a *lot* of clicks and time in achieving intended outcomes

### [xkcd: Is It Worth The Time?](https://xkcd.com/1205/)

Evaluate if task worth automating to achieve intended outcome

<img src="https://imgs.xkcd.com/comics/is_it_worth_the_time_2x.png" width="700" alt="xkcd: Is It Worth The Time?">

### [xkcd: Automation](https://xkcd.com/1319/)

Focus on core task outcome and what's needed to achieve it

<img src="https://imgs.xkcd.com/comics/automation_2x.png" width="700" alt="xkcd: Automation">

---

## Folder structure

```code
sandbox-gworkspace-msoffice-automation/
├── .gitignore
├── LICENSE
├── README.md
│
├── .github/
│   ├── ISSUE_TEMPLATE/
│   │   ├── bug_report.md
│   │   └── feature_request.md
│   └── pull_request_template.md
│
├── GoogleAppScripts-Sheets/
│   ├── gdrive-create_folders.js
│   ├── gdrive-links_from_folder.js
│   ├── gmail-attachments_from_domain.js
│   ├── gmail-list_day_email.js
│   ├── gsheets-fuzzy_names_companies.js
│   ├── gsheets-text_split.js
│   └── gsheets-update_timezones.js
│
├── GoogleColab/
│   ├── get_comments-ms_excel.ipynb
│   ├── get_comments-ms_ppt.ipynb
│   ├── get_comments-ms_word.ipynb
│   ├── merge_pdfs.ipynb
│   └── text_to_csv.ipynb
│
├── Macro-Excel/
│   ├── fill_color_checks.bas
│   ├── font_check.bas
│   ├── pivot_table-locations.bas
│   └── pivot_table-same_cache.bas
│
├── Macro-PPT/
    ├── font_check.bas
    ├── get_titles.bas
    └── header_consistency_check.bas
```

## Quick start

Ensure to inspect code before running

1. **Clone**
   ```bash
   git clone https://github.com/<your‑org>/sandbox-gworkspace-msoffice-automation.git
   cd sandbox-gworkspace-msoffice-automation

2. **Google Apps Script**
- Open the Apps Script editor in a Sheet or standalone project
- Copy‑paste any .js file, set project scopes, and run
- Grant one‑time OAuth consent

3. **Google Colab**
Google colab used for quick sandboxing with minimum setup and can be used on most systems

- colab.new to create new notebook and paste in
- Follow cell prompts

Or
- Click [open in colab](https://openincolab.com/) to open in notebook in your own Colab Virtual Machine (VM)
- File > Save a copy in drive

1. **Microsoft Macros**
- Enable the Developer tab
- Import .bas modules (Alt + F11 then import file)
- Run the desired Sub. (Alt + F8)

## Prerequisites
- Google Account: Access to Drive, Sheets, Gmail, etc.
- Python: 3.8+ (handled by Colab)
- Microsoft Office: 2016 desktop or later with VBA enabled

## Contributing
PRs are welcome! Feel free to:
- Add new utility scripts
- Improve error handling and 
- Expand/create docs, examples, and tests


## License

Licensed under the [Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0).

© 2025 [The Contributors](CONTRIBUTORS.md)
