## Download

Download the latest ZIP from the Releases page:

- Go to Releases
- Download `mp4_duplicate_checker_v0.1.0_beta.zip`
- Extract the ZIP
- Run `run_mp4_duplicate_checker_v7.bat`

# MP4 Duplicate Checker

A lightweight Windows tool for checking duplicate MP4 videos across multiple folders.

It checks two levels of duplicates:

1. Exact duplicates
   Same file size and same SHA-256 hash. These are byte-level identical files.

2. Visual duplicate candidates
   FFmpeg samples video frames, converts them into small grayscale fingerprints, and compares visual similarity. This can help find videos that look the same after re-exporting or recompression.

The tool does not rename, move, delete, or edit your video files.

## What changed in v7

- Rewrote the exact hash stage to avoid the PowerShell type mismatch error seen in v5/v6.
- Removed reliance on `Group-Object` for final hash grouping.
- Removed `Math.Min/Math.Max` pair-key generation.
- Added an `error_debug_*.txt` file in the report folder if a fatal error occurs after report-folder creation.

## What it checks

1. Exact duplicates  
   Same file size and same SHA-256 hash. This means byte-level identical files.

2. Visual duplicate candidates  
   FFmpeg samples frames from each video, converts them to 8x8 grayscale fingerprints, and compares visual similarity.

## Requirements

- Windows 10 or Windows 11
- PowerShell
- FFmpeg

Check FFmpeg:

```bat
ffmpeg -version
```

Install FFmpeg with winget:

```bat
winget install -e --id Gyan.FFmpeg
```

## How to use

1. Extract this ZIP.
2. Keep these files in the same folder:
   - `run_mp4_duplicate_checker_v7.bat`
   - `mp4_duplicate_checker_v7.ps1`
3. Double-click `run_mp4_duplicate_checker_v7.bat`.
4. Paste a folder path into the terminal.
5. After each folder, choose:
   - `A` to add another folder
   - `S` to start checking
   - `Q` to quit
6. Choose whether to scan subfolders.
7. Run exact hash checking.
8. Run optional FFmpeg visual checking.
9. Reports are saved to a new Desktop folder.

Alternative launch:
- Drag one or more folders onto `run_mp4_duplicate_checker_v7.bat`.

## Reports

The summary begins with RESULT OVERVIEW:

- Possible duplicate pairs total
- Exact duplicate pairs by hash
- Visual duplicate candidate pairs by FFmpeg frame fingerprint
- A/B file paths for each duplicate pair

## Notes

- The tool does not rename, move, delete, or edit videos.
- Exact duplicates can be treated as the same file.
- Visual duplicate candidates should be manually reviewed.
- Visual matching is practical similarity detection, not forensic verification.
