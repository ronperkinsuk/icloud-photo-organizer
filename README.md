# iCloud Photo Organizer (Windows)

A PowerShell script to organise iCloud Photos downloaded via iCloud for Windows into a clean, structured archive.

## ✨ Features

- Supports multiple iCloud accounts (e.g. personal, spouse, family)
- Automatically organises media into:

    Person / Photos or Videos / Year / Device Model or WhatsApp

- Extracts metadata using ExifTool
- Detects and skips duplicate files per user (SHA256 hashing)
- Handles WhatsApp and non-EXIF images separately
- Works with iCloud "Optimise Storage" (on-demand downloads)
- Automatically reverts files back to cloud-only after processing
- Detailed logging and resumable runs
- Progress tracking with ETA

---

## 📂 Example Output Structure
```
D:\Photos\Phone Media
├── User1
│ ├── Photos
│ │ ├── 2026
│ │ │ ├── iPhone 17 Pro
│ │ │ └── WhatsApp
│ └── Videos
├── User2
└── User3
```

---

## ⚙️ Requirements

- Windows 10/11
- iCloud for Windows (Photos enabled)
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
4. Download ExifTool and update the script path:
```
$exiftool = "C:\Tools\Exiftool\exiftool.exe"
```
4. Update the `$sources` array in the script:
```
$sources = @(
   @{ Path = "C:\Users\User1\Pictures\iCloud Photos\Photos"; Person = "User1" }
)
```
5. Set your destination folder:
```
$dest = "D:\Photos\Phone Media"
```

---

## ▶️ Usage

Run the script:
```
powershell.exe -ExecutionPolicy Bypass -File .\download-icloud-photos.ps1
```

---

## 📊 What the Script Does

For each file:

1. Forces iCloud to download the full file
2. Copies it to a temporary working folder
3. Extracts metadata (date taken, device model)
4. Determines:
   - Photo vs Video
   - WhatsApp vs Camera image
5. Detects duplicates using SHA256 hash
6. Moves file to organised structure
7. Logs processed files for safe re-runs
8. Frees up space in iCloud (returns file to cloud-only)

---

## 🔁 Re-running

The script is safe to run multiple times:

- Previously processed files are skipped
- Duplicate files are ignored
- Logs ensure incremental processing

---

## 📁 Logs

Located in:
```
D:\Photos\Phone Media\Logs\
```

- `processed.log` → files already handled
- `hashes.csv` → duplicate detection
- `errors.log` → failures
- `run_*.log` → full run logs

---

## ⚠️ Known Limitations

- iCloud Windows may delay downloads for some files
- Some media (e.g. WhatsApp images) may not contain EXIF data
- HEIC support depends on Windows codecs

---

## 📜 License

MIT License

---

## 👤 Author

Ron Perkins

---

## ⭐ Contributions

Feel free to fork, improve, and submit pull requests.

