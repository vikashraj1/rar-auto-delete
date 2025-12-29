# RAR Archive Extractor with Auto-Delete

Extracts multi-part RAR archives and automatically deletes each part as it's processed, saving disk space.

## Requirements

- Windows 10/11
- WinRAR installed at `C:\Program Files\WinRAR\rar.exe`
- **⚠️ IMPORTANT: Enough storage for extracted data + 1/n of total RAR parts**

## Installation

1. Download `extract-and-delete.ps1` and `Extract RAR Parts.bat`
2. Keep both files in the same folder
3. Done!

## How to Use

**Easy Way:**
1. Double-click `Extract RAR Parts.bat`
2. Select your `.part1.rar` file
3. Type `y` to confirm after the integrity test
4. Choose extraction folder
5. Wait for completion

**Quick Way:**
- Drag `.part1.rar` onto `Extract RAR Parts.bat`
- Or run: `powershell -ExecutionPolicy Bypass -File "extract-and-delete.ps1" "path\to\file.part1.rar"`

## What It Does

1. Tests archive integrity (prevents corruption)
2. Asks for confirmation (warns about deletion)
3. Extracts files to a new folder
4. Deletes each part after it's extracted
5. Opens the folder when done

## Why Use This?

**Without script:** Need 14GB free (7GB parts + 7GB extracted)  
**With script:** Need ~8GB free (parts deleted as you go)

## Safety Features

- Tests archive before deleting anything
- Requires typing `y` to proceed (Enter = cancel)
- Only deletes parts that are successfully extracted
- Stops immediately if extraction fails

## Configuration

Edit the `$CONFIG` section in the script to change:
- WinRAR path
- Delete delay timing
- Enable/disable integrity testing

## Troubleshooting

**"WinRAR not found"** → Install WinRAR or update path in script  
**Script opens in Notepad** → Use the `.bat` file instead  
**Parts not deleting** → Increase/Decrease `DeleteDelayMs` in config 

## ⚠️ Warning

This script **permanently deletes** RAR parts during extraction. Make sure the archive is valid before proceeding.

---

**Note:** Use at your own risk. Always backup important data.
