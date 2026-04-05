# yt-dlp Helper

This folder contains the dedicated files for the interactive download helper:

- `yt-dlp-helper.ps1`: the actual helper implementation
- `user_preferences.json`: local helper runtime state
- `cache/`: playlist listing cache used by the helper

Shared binaries stay at the repo root so both main tools can reuse them:

- `..\yt-dlp.exe`
- `..\ffmpeg_yt-dlp\`

The top-level `..\yt-dlp-helper.ps1` file is now only a launcher that preserves the old entry point.
