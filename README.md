# 📸 OneDrive Photo Importer

A command-line agent that **downloads every photo from your Microsoft OneDrive to a local folder** and **deletes each photo from the cloud after a safe download** — with resume support, date filtering, quality preservation, and dry-run preview.

Built with Python + Microsoft Graph API.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [How It Works](#how-it-works)
- [Prerequisites](#prerequisites)
- [Azure App Registration Setup](#azure-app-registration-setup)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Command Reference](#command-reference)
- [Supported File Formats](#supported-file-formats)
- [Project Structure](#project-structure)
- [Troubleshooting](#troubleshooting)
- [Future Improvements](#future-improvements)

---

## Overview

If you have thousands of photos sitting on OneDrive and want to migrate them to local storage without losing quality or manually downloading them one by one — this tool does it automatically.

It connects to your OneDrive via the **Microsoft Graph API**, walks your entire folder tree, downloads every image in original quality, and deletes it from the cloud only after confirming the download succeeded. A progress file ensures no photo is ever lost if the process is interrupted.

---

## Features

- 🔍 **Recursive folder walk** — finds photos in every subfolder automatically
- 🖼️ **Original quality** — uses pre-signed blob URLs (`@microsoft.graph.downloadUrl`), bypassing Graph API transcoding
- 📅 **Date filtering** — process photos by exact date or date range
- 🔢 **Limit flag** — test with a single photo before committing to the full run
- ♻️ **Resume support** — saves progress after each photo; picks up exactly where it left off if interrupted
- 🧪 **Dry-run mode** — preview what would be downloaded/deleted without making any changes
- 🛡️ **Safe delete** — photo is only deleted from OneDrive after a successful local download
- 📄 **Logging** — every action written to `onedrive_importer.log`
- ⚡ **Rate-limit handling** — automatic retry with backoff on HTTP 429 responses
- 🗂️ **Format support** — JPG, PNG, HEIC, HEIF, RAW, DNG, TIFF, WEBP, and more

---

## How It Works

```
┌─────────────────────────────────────────────────────────────────┐
│                        AGENT FLOW                               │
├─────────────────────────────────────────────────────────────────┤
│                                                                  │
│  1. AUTHENTICATE                                                 │
│     └─ Device Code Flow (open URL, enter code, done once)       │
│     └─ Token cached locally for future runs                     │
│                                                                  │
│  2. DISCOVER PHOTOS                                              │
│     └─ Recursively walk all OneDrive folders                     │
│     └─ Detect images by facet (image/photo) + file extension    │
│     └─ Apply date filters if specified                          │
│     └─ Skip already-processed photos (resume mode)              │
│                                                                  │
│  3. FOR EACH PHOTO:                                              │
│     ├─ Fetch pre-signed blob download URL                       │
│     ├─ Stream download → local folder (512 KB chunks)           │
│     ├─ Verify file is not empty                                 │
│     ├─ DELETE from OneDrive (only after confirmed download)      │
│     └─ Save progress to .onedrive_progress.json                 │
│                                                                  │
│  4. SUMMARY                                                      │
│     └─ Report downloaded / deleted / failed counts              │
│                                                                  │
└─────────────────────────────────────────────────────────────────┘
```

---

## Prerequisites

- Python 3.8 or higher
- A Microsoft account with OneDrive (personal — Outlook / Hotmail / Live)
- Internet connection
- ~5 minutes for the one-time Azure setup

---

## Azure App Registration Setup

This is a **one-time setup** that takes about 5 minutes. You are registering a free app identity so the script can authenticate with Microsoft on your behalf.

### Step 1 — Create the App

1. Go to [https://portal.azure.com](https://portal.azure.com) and sign in with your Microsoft account
2. In the search bar, type **"App registrations"** and click it
3. Click **"+ New registration"**
4. Fill in:
   - **Name:** `OneDrive Photo Importer` (or anything you like)
   - **Supported account types:** `Personal Microsoft accounts only`
   - **Redirect URI:** leave blank for now
5. Click **Register**

### Step 2 — Copy Your Client ID

On the Overview page, copy the **Application (client) ID** — you'll paste this into the script.

```
Application (client) ID:  xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx  ← copy this
```

### Step 3 — Configure Authentication

1. In the left sidebar, click **Authentication**
2. Click **"+ Add a platform"** → choose **"Mobile and desktop applications"**
3. Check this redirect URI:
   ```
   https://login.microsoftonline.com/common/oauth2/nativeclient
   ```
4. Click **Configure**
5. Scroll down to **"Advanced settings"**
6. Set **"Allow public client flows"** to **Yes**
7. Click **Save**

### Step 4 — Add API Permission

1. In the left sidebar, click **API permissions**
2. Click **"+ Add a permission"**
3. Choose **Microsoft Graph**
4. Choose **Delegated permissions**
5. Search for `Files.ReadWrite` → check it
6. Click **Add permissions**

> ✅ You do **not** need to click "Grant admin consent" for personal accounts.

---

## Installation

```bash
# 1. Clone or download this project
git clone https://github.com/yourname/onedrive-photo-importer
cd onedrive-photo-importer

# 2. Create a virtual environment (recommended)
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt
```

---

## Configuration

Open `onedrive_photo_importer.py` and set these two lines near the top:

```python
# ═══════════════════════════════════════════
#   CONFIGURATION  ← Edit these two lines
# ═══════════════════════════════════════════
CLIENT_ID       = "YOUR_CLIENT_ID_HERE"    # ← Paste your Azure Application (client) ID
DOWNLOAD_FOLDER = "./onedrive_photos"      # ← Change to your preferred local folder
# ═══════════════════════════════════════════
```

---

## Usage

### First Run — Authenticate

The first time you run any command, Microsoft will ask you to log in:

```
To sign in, use a web browser to open the page https://www.microsoft.com/link
and enter the code ABC123XY to authenticate.
```

Open the link, enter the code, sign in with your Microsoft account. The token is cached locally so you won't need to do this again unless the cache is deleted.

---

### Recommended Workflow

```bash
# Step 1 — Preview everything first (no changes made)
python onedrive_photo_importer.py --dry-run

# Step 2 — Test with a single photo
python onedrive_photo_importer.py --limit 1

# Step 3 — Verify the photo downloaded correctly and is gone from OneDrive

# Step 4 — Run the full import
python onedrive_photo_importer.py
```

---

### Full Run (all photos)

```bash
# Prevent your PC from sleeping mid-import (Windows)
powercfg /change standby-timeout-ac 0

# Run the importer
python onedrive_photo_importer.py

# Restore sleep settings after
powercfg /change standby-timeout-ac 30
```

---

### Resume After Interruption

If the script is interrupted (power cut, Ctrl+C, crash), just run:

```bash
python onedrive_photo_importer.py --resume
```

It reads `.onedrive_progress.json` and skips every photo that was already successfully downloaded and deleted.

---

## Command Reference

| Flag | Description | Example |
|---|---|---|
| *(none)* | Download all photos and delete from OneDrive | `python onedrive_photo_importer.py` |
| `--dry-run` | List all matching photos, no download or delete | `--dry-run` |
| `--no-delete` | Download photos but keep them on OneDrive | `--no-delete` |
| `--resume` | Skip photos already completed in a previous run | `--resume` |
| `--folder PATH` | Set a custom download destination folder | `--folder D:\MyPhotos` |
| `--limit N` | Process only N photos (useful for testing) | `--limit 1` |
| `--date YYYY-MM-DD` | Only process photos from this exact date | `--date 2024-03-16` |
| `--date-from YYYY-MM-DD` | Process photos from this date onwards | `--date-from 2024-01-01` |
| `--date-to YYYY-MM-DD` | Process photos up to this date | `--date-to 2024-12-31` |

### Combining Flags

```bash
# Preview all photos from 2024 before importing
python onedrive_photo_importer.py --date-from 2024-01-01 --date-to 2024-12-31 --dry-run

# Test one photo from a specific date
python onedrive_photo_importer.py --date 2024-03-16 --limit 1

# Download all photos from 2023 but keep them on OneDrive (backup only)
python onedrive_photo_importer.py --date-from 2023-01-01 --date-to 2023-12-31 --no-delete

# Resume a previous full import
python onedrive_photo_importer.py --resume
```

---

## Supported File Formats

The script detects images both by Microsoft Graph API facets (`image`, `photo`) and by file extension as a fallback — so formats the Graph API doesn't tag correctly (like HEIC) are still caught.

| Category | Extensions |
|---|---|
| Standard | `.jpg` `.jpeg` `.png` `.gif` `.bmp` `.webp` |
| Apple / Samsung | `.heic` `.heif` |
| High quality | `.tiff` `.tif` |
| RAW / Camera | `.raw` `.dng` `.cr2` `.nef` `.arw` `.orf` `.rw2` |

---

## Project Structure

```
onedrive-photo-importer/
│
├── onedrive_photo_importer.py    # Main script
├── requirements.txt              # Python dependencies
├── README.md                     # This file
│
├── onedrive_photos/              # Created on first run — downloaded photos go here
│
├── .onedrive_token_cache.json    # Cached auth token (auto-created, keep private)
├── .onedrive_progress.json       # Resume progress tracker (auto-created)
└── onedrive_importer.log         # Full run log (auto-created)
```

> ⚠️ **Never commit** `.onedrive_token_cache.json` to version control — it contains your Microsoft auth token. Add it to `.gitignore`:
> ```
> .onedrive_token_cache.json
> .onedrive_progress.json
> onedrive_importer.log
> ```

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `ValueError: You cannot use any scope value that is reserved` | `offline_access` passed manually to MSAL | Remove `offline_access` from `SCOPES` — MSAL adds it automatically |
| `AADSTS50059: No tenant-identifying information` | Missing `authority` in MSAL app constructor | Add `authority="https://login.microsoftonline.com/consumers"` |
| `Search Query cannot be empty` | Old search-based discovery approach | Use recursive `walk_folder()` instead (already fixed in current version) |
| `Could not get direct download URL` | `$select` strips the download URL annotation | Remove `$select` from the metadata fetch — fetch the full item |
| Date filter returns 0 matches for HEIC files | Graph API doesn't return `image`/`photo` facet for HEIC | Extension-based fallback detection (already fixed in current version) |
| `--limit` applied before `--date` filter | Limit ran first, grabbed wrong photos | Always apply date/range filters before limit (already fixed) |
| Photos downloaded in lower quality | Using `/content` endpoint causes transcoding | Use `@microsoft.graph.downloadUrl` for original blob (already fixed) |

---

## Future Improvements

### 🚀 Performance
- **Parallel downloads** — use `concurrent.futures.ThreadPoolExecutor` to download multiple photos simultaneously, reducing total runtime significantly
- **Batch delete** — Microsoft Graph supports batch requests (up to 20 operations per call); grouping deletes would cut API call overhead by 20×
- **Delta API** — use `/me/drive/root/delta` instead of a full walk on subsequent runs to only fetch newly added photos

### 🗂️ Organization
- **Preserve folder structure** — recreate the OneDrive folder hierarchy locally instead of dumping everything into a single flat folder
- **Sort into subfolders by date** — automatically organize downloaded photos into `YYYY/MM/DD/` subfolders
- **Duplicate detection** — hash-based deduplication to skip photos already present locally even if filenames differ

### 🔒 Safety
- **Recycle bin instead of hard delete** — move photos to OneDrive's recycle bin first; permanently delete only after a configurable grace period (e.g. 7 days)
- **Checksum verification** — compare MD5/SHA256 of downloaded file against OneDrive's reported hash before deleting
- **Dry-run diff** — on subsequent runs, show only what's new since the last import

### 🖥️ Interface
- **Progress dashboard** — rich terminal UI with per-file progress bars, ETA, and live speed using the `rich` library
- **GUI wrapper** — simple Tkinter or PyQt window for non-technical users (folder picker, date range selector, start button)
- **Desktop notifications** — system notification when the full import completes

### ☁️ Platform Support
- **Google Photos → local** — extend the same agent pattern to Google Photos using the Google Photos API
- **iCloud → local** — support iCloud photo library export
- **Multi-account** — support switching between multiple OneDrive accounts in a single run

### 📊 Reporting
- **Import summary report** — generate an HTML or CSV report after each run: total size, count by date, any failures, time taken
- **Failed-items retry queue** — a dedicated `--retry-failed` flag that re-processes only items recorded as failed in the progress file
