# iCloud Photo Organizer (Windows)

A resilient PowerShell solution to organise iCloud Photos (via iCloud for Windows) into a clean, structured, and de-duplicated archive.

## ✨ Features

- Supports multiple iCloud accounts (e.g. personal, spouse, family)
- Automatically organises media into:
  
Person / Photos or Videos / Year / Device Model, WhatsApp, or Unknown

- Extracts metadata using ExifTool (date taken, device model)
- Builds a date-based model map to infer missing metadata for videos
- Detects and skips duplicate files per user (SHA256 hashing)
- Intelligent classification:
- MP4 → WhatsApp
- Non-EXIF photos → WhatsApp
- Non-EXIF videos → Unknown Model
- Fully compatible with iCloud "Optimise Storage" (on-demand files)
- Forces download of cloud-only files with timeout handling
- Automatically reverts files back to cloud-only after processing
- Uses a temporary working directory to avoid iCloud locking issues
- Persistent state for safe restart after crashes or interruptions
- Batch processing (2000 files per run) for stability
- Optional launcher script for fully automated end-to-end processing
- Detailed logging (runs, errors, slow downloads)
- Progress tracking with ETA and per-user statistics

---

## 📂 Example Output Structure
```
D:\Photos\Phone Media
├── User1
│   ├── Photos
│   │   ├── 2026
│   │   │   ├── iPhone 17 Pro
│   │   │   └── WhatsApp
│   └── Videos
│       ├── 2026
│       │   ├── iPhone 17 Pro
│       │   └── Unknown Model
├── User2
└── User3
```

---

## ⚙️ Requirements

- Windows 10/11
- iCloud for Windows (Photos enabled with "Optimise Storage")
- PowerShell 5.1 or later
- ExifTool

Download ExifTool:
https://exiftool.org/

---

## 🛠 Setup

1. Install iCloud for Windows and enable Photos syncing
2. Ensure photos are visible under:
```
C:\Users\<User1>\Pictures\iCloud Photos\Photos
```
3. Download ExifTool and update the path in `config.ps1`:
```
$exiftool = "C:\Tools\Exiftool\exiftool.exe"
```
4. Configure your source accounts:
```
$sources = @(
   @{ Path = "C:\Users\User1\Pictures\iCloud Photos\Photos"; Person = "User1" },
   @{ Path = "C:\Users\User2\Pictures\iCloud Photos\Photos"; Person = "User2" }
)
```
5. Set your destination folder:
```
$dest = "D:\Photos\Phone Media"
```

---

## ▶️ Usage

Option 1 – Run manually
```
powershell.exe -ExecutionPolicy Bypass -File .\download-icloud-photos.ps1
```
Option 2 – Use the launcher (recommended for large libraries)
```
powershell.exe -ExecutionPolicy Bypass -File .\launcher.ps1
```
The launcher will:
- Run the main script in batches (2000 files per run)
- Automatically restart until all files are processed
- Avoid long-running PowerShell instability (e.g. CHARTV.dll crashes)

---

## 📊 What the Script Does

For each file:

1. Forces iCloud to download the full file (if cloud-only)
2. Waits for download completion (with timeout + logging)
3. Copies to a temporary working directory
4. Extracts metadata (date taken, device model)
5. Infers missing video metadata using photo data (same day)
6. Determines:
   - Photo vs Video
   - WhatsApp vs Camera vs Unknown
7. Generates SHA256 hash for duplicate detection
8. Moves file into organised structure
9. Logs processed files for safe re-runs
10. Frees disk space (returns file to iCloud cloud-only state)

---

## 🔁 Re-running / Crash Recovery

The script is fully restart-safe:
- Previously processed files are skipped (`processed.log`)
- Duplicate content is skipped (`hashes.csv`)
- Interrupted runs resume automatically
- Launcher enables continuous processing until completion

---

## 📁 Logs

Located in:
```
D:\Photos\Phone Media\Logs\
```

- `processed.log` → files already handled
- `hashes.csv` → duplicate detection
- `errors.log` → failures
- `slow.log` → files that exceeded download timeout
- `run_*.log` → full run logs

---

## ⚠️ Known Limitations

- iCloud for Windows can delay or stall downloads for some files
- Some media (e.g. WhatsApp) may not contain EXIF metadata
- HEIC support depends on installed Windows codecs
- “Free up space” shell action may occasionally fail (logged as warning)

---

## 📜 License

MIT License

---

## 👤 Author

Ron Perkins

---

## ⭐ Contributions

Feel free to fork, improve, and submit pull requests.

