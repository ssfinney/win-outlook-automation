# Win Outlook Automation (Local-Only)

Local, compliance-friendly Outlook “triage” automation for **Windows 11 + Classic Outlook (M365)**.

This project:
- Scans recent Inbox items
- Assigns an Outlook **Category** (`Urgent/Action/Waiting/FYI/Noise`)
- Flags Urgent/Action for follow-up
- Optionally moves obvious noise to a **Read Later** folder
- Writes an **Excel triage report** you can label to train a lightweight model
- (Optional) Trains a local sklearn model to learn *your* inbox patterns — **no cloud calls**

> Designed for locked-down environments: runs entirely on your PC and writes to `OneDrive\AI_Outlook\...` (or local OneDrive folder).

---

## Repo contents

- `outlook_triage.py`  
  Main script: triage + categorize + report export.

- `train_model.py`  
  Training script: learns from labeled rows in the triage reports and saves `model/triage_model.joblib`.

- `OutlookTriage_Task.xml`  
  Windows Task Scheduler task to run `outlook_triage.py` on a schedule.

- `TrainModel_Task.xml`  
  Windows Task Scheduler task to run `train_model.py` (typically weekly or manual).

---

## What it creates on disk

By default, the scripts use:

`%OneDrive%\AI_Outlook\`
- `config\vip_senders.csv`  (one email per line)
- `data\inbox_scored_YYYY-mm-dd_HHMM.csv`
- `outputs\triage_report_YYYY-mm-dd_HHMM.xlsx`
- `model\triage_model.joblib`

If `%OneDrive%` is not set, it falls back to `C:\Users\<you>\OneDrive\AI_Outlook`.

---

## Requirements

### 1) Outlook
- **Classic Outlook for Windows** (COM automation).
- You must be signed in and able to open your Inbox normally.

### 2) Python
- Recommended: **Python 3.12 (64-bit)**

> Note: `outlook_triage.py` prints a warning that **pywin32 may not be compatible with Python 3.13** in your environment. If you hit that, use Python 3.12.

### 3) Python packages
- `pywin32`
- `pandas`
- `openpyxl`
- `scikit-learn`
- `joblib`

If your environment blocks internet access, install from an internal package source or pre-downloaded wheels.

---

## Install (Windows 11)

### Step A — Create a venv
Open PowerShell in the repo folder:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
```

### Step B — Install dependencies

**Online (if allowed):**
```powershell
pip install pywin32 pandas openpyxl scikit-learn joblib
```

**Offline (common in locked-down environments):**
1. On a machine that *can* download packages, download wheels:
   ```powershell
   mkdir wheels
   pip download -d wheels pywin32 pandas openpyxl scikit-learn joblib
   ```
2. Copy the `wheels` folder to your work PC.
3. Install locally:
   ```powershell
   pip install --no-index --find-links .\wheels pywin32 pandas openpyxl scikit-learn joblib
   ```

### Step C — Verify pywin32 registration
Sometimes Outlook COM needs this after install:

```powershell
python -m pywin32_postinstall -install
```

---

## First run

```powershell
.\.venv\Scripts\Activate.ps1
python outlook_triage.py
```

Expected output includes:
- How many emails were processed
- Path to the Excel report and CSV log
- Path to `vip_senders.csv`

In Outlook, you should see Categories applied to recent emails and flags on Urgent/Action.

---

## Configure VIPs (high priority senders)

Edit:

`%OneDrive%\AI_Outlook\config\vip_senders.csv`

One SMTP address per line, e.g.
```
vip1@company.com
vip2@company.com
```

VIP senders get a scoring boost (`+50`).

---

## The training loop (make it learn your inbox)

1. Run `outlook_triage.py` to generate a report:
   - `outputs\triage_report_*.xlsx`
2. Open the Excel report, sheet **All Scored**
3. Fill the `label` column with one of:
   - `Urgent`, `Action`, `Waiting`, `FYI`, `Noise`
4. Save the workbook
5. Train the model:
   ```powershell
   python train_model.py
   ```
6. Run triage again:
   ```powershell
   python outlook_triage.py
   ```

### Guardrails
Even with a model trained, **rules override the model** when risk is high:
- Rule `Urgent` always stays `Urgent`
- Strong `Noise` stays `Noise` when score < 0

This prevents the model from “down-ranking” critical items.

---

## Scheduling (Task Scheduler)

### Import the provided tasks
1. Open **Task Scheduler**
2. Choose **Import Task…**
3. Import:
   - `OutlookTriage_Task.xml`
   - `TrainModel_Task.xml`
4. Edit each task:
   - Set **User account** to you
   - Confirm **Triggers** (time of day)
   - Update **Actions** paths:
     - Python: `<repo>\.venv\Scripts\python.exe`
     - Script: `<repo>\outlook_triage.py` or `<repo>\train_model.py`
5. Ensure “Run only when user is logged on” (Outlook COM is usually safest this way)

> Outlook COM automation often fails in non-interactive sessions. If a scheduled run doesn’t work, run it while logged in.

---

## Tuning knobs (in `outlook_triage.py`)

Top-level settings:
- `DAYS_BACK` (default 7)
- `MAX_ITEMS` (default 500)
- `MOVE_NOISE_TO_READ_LATER` (default True)
- `KEYWORD_WEIGHTS` (add your high-value business terms)
- `NOISE_PATTERNS` (catch newsletters/promos)

Categories/folders:
- Categories: `Urgent/Action/Waiting/FYI/Noise`
- Folder: `Read Later` (created under Inbox if missing)

---

## Safety notes

- This script **modifies your mailbox** (Categories, Flags, and optionally moves mail).
- Start with a smaller window (`DAYS_BACK=2`) and `MOVE_NOISE_TO_READ_LATER=False` until you trust it.
- Keep your reports; they are your audit trail (`outputs\triage_report_*.xlsx`).

---

## Troubleshooting

### “pywin32 not installed or incompatible”
- Use **Python 3.12** 64-bit
- Reinstall `pywin32`
- Run:
  ```powershell
  python -m pywin32_postinstall -install
  ```

### “Access is denied” / security prompts
- Your org may restrict programmatic access to Outlook.
- Try running Outlook and the script **as the same user**.
- If prompts appear, your security policy may require admin changes (Group Policy / Trust Center).

### Categories not appearing
- Outlook Categories exist per mailbox/profile.
- The script sets `mail_item.Categories = "<Category>"`. If you want color-coded categories, define them in Outlook manually once.

---

## Roadmap ideas (optional)
- Thread-level triage via `ConversationID`
- “Resurface” logic that pulls older high-scored emails into a daily digest
- Local keyword + sender model features from subject/body snippets (currently subject + headers + rule features)

---

## License
MIT
