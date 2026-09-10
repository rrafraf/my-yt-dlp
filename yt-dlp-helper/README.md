# yt-dlp Helper Package

This folder now contains the helper as a small package instead of a single monolithic script.

## Files
- `yt-dlp-helper.ps1`: interactive CLI entry point
- `yt-dlp-helper-gui.ps1`: WPF workbench for helper tasks
- `yt-dlp-helper-worker.ps1`: background worker used by the GUI
- `Helper.Core.ps1`: shared non-interactive core used by both CLI and GUI
- `yt-dlp-helper.gui.config.json`: GUI logging settings
- `user_preferences.json`: local helper state
- `cache/`: playlist feed cache
- `logs/`: GUI session logs and per-job stdout/stderr logs
- `Downloads/`: default helper download root when no custom root is chosen

Shared binaries stay at the repo root so both main tools reuse the same installs:
- `..\yt-dlp.exe`
- `..\ffmpeg_yt-dlp\`

## Launch
CLI:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\yt-dlp-helper\yt-dlp-helper.ps1
```

GUI:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\yt-dlp-helper\yt-dlp-helper-gui.ps1
```

The top-level `..\yt-dlp-helper.ps1` file is only a compatibility launcher for the CLI entry point.

## GUI Workbench
The helper GUI is organized around tasks instead of prompt order:
- `Single Video`
- `Playlist URL`
- `My Playlists`
- `Environment`
- `Folder Inspector`

Long-running work runs in a background worker. The GUI keeps the current job non-blocking, streams stdout/stderr into the activity pane, supports cancel, and keeps recent job history.

## Folder Inspector
The inspector is offline-only in v1. It reads whatever local artifacts already exist in a folder and reports:
- `download_archive.txt` summary
- playlist metadata files when present
- counts for media, `.info.json`, subtitles, and text transcripts
- inferred or confirmed playlist title and id
- certainty and completion state based only on local evidence

## Logging
`yt-dlp-helper-gui.ps1` reads `yt-dlp-helper.gui.config.json`:
```json
{
  "logging": {
    "level": "INFO",
    "retentionDays": 14
  }
}
```

## GUI Design Notes
General approach used here, and reusable for other tools:
- Design around user goals, not around the script's internal prompt sequence.
- Infer safe defaults from saved state, local context, and available system information.
- Prefer explicit controls over free-text input when values can be enumerated.
- Keep the common path simple and push power-user controls into secondary surfaces.
- Make background work legible with live logs, status, history, and cancel.
- Separate reusable execution logic from the GUI so CLI, GUI, and future automation stay aligned.
