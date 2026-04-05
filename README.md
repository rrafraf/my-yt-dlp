## my-yt-dlp Helper (Windows)

A Windows-focused toolkit to:
- Automatically fetch and keep `yt-dlp.exe` up to date (nightly builds)
- Automatically download and wire up a portable FFmpeg for `yt-dlp`
- Download YouTube videos or playlists (with metadata, thumbnails, subtitles, and duplicate prevention)
- Optionally list your YouTube playlists (requires browser cookies)
- Transcribe WhatsApp voice notes: convert `.opus` в†’ `.wav` with FFmpeg and transcribe with Whisper

---

### WhatвЂ™s in this repo
- `yt-dlp-helper.ps1`: Entry point. Interactive menu for YouTube downloads and auto-setup of tools.
- `user_preferences.json`: Stores last menu choice, playlist info, current `yt-dlp` version, and your saved Firefox profile path.
- `yt-research-gui\yt-research-gui.ps1`: Launcher for the research GUI.
- `yt-research-gui/`: Dedicated folder for the GUI implementation, config, Python project, logs, and cached transcript data.
- `ffmpeg_yt-dlp/`: Folder where the portable FFmpeg is downloaded and extracted.
- `cache/`: Caches your YouTube playlists listing (`playlists_cache.json`).
- `whatsapp_transcribe.py`: Batch transcribes WhatsApp `.opus` files using FFmpeg + OpenAI Whisper.
- `Transcripts/`: Output folder for transcription results (created on demand).
- `_tbt_audios/`: Place WhatsApp `.opus` files here for transcription.

---

## Requirements
- Windows 10/11
- PowerShell 5.0+
- Internet access (downloads yt-dlp nightly release and FFmpeg builds)
- For reliable YouTube format extraction:
  - A supported JavaScript runtime for `yt-dlp` challenge solving, such as Deno or Node.js 20+. The helper auto-detects and passes an available runtime.
- For playlist listing/authenticated downloads:
  - Firefox (Nightly or regular). The helper can scan for local Firefox profiles and save your chosen profile path in `user_preferences.json`.
- For on-demand GUI Whisper transcription:
  - `uv` installed
  - Python 3.10+ (managed through `uv` in `yt-research-gui/`)
  - FFmpeg available under `ffmpeg_yt-dlp/` (installed by the helper script)
  - GUI Python dependencies installed with `uv sync` inside `yt-research-gui/`
- For WhatsApp transcription:
  - Python 3.9+ recommended
  - FFmpeg (auto-installed into `ffmpeg_yt-dlp` by the helper script, used by the Python script)
  - Python packages: `openai-whisper` (and its dependencies, e.g. Torch)

---

## Setup

### 1) Get your Firefox profile path (for cookies)
If you want to list/download your own playlists:
1. Run `yt-dlp-helper.ps1`.
2. If no Firefox profile is saved yet, or the saved one no longer exists, the helper will scan your system for Firefox profiles and list them as choices.
3. Pick one of the detected profiles, or choose the manual-entry option and paste the absolute profile folder path yourself.
4. The selected path is saved in `user_preferences.json` for future runs.

If you need to verify the path manually first, open Firefox and go to `about:profiles`, then copy the absolute folder path for the profile you use for YouTube logins.

### 2) Allow running the script (if needed)
If your execution policy blocks local scripts, run PowerShell as your user and execute:
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

---

## Using the YouTube Helper

1. Open PowerShell in the project folder:
```powershell
cd "C:\Users\<you>\Documents\GitHub\my-yt-dlp"
```

2. Run the helper:
```powershell
./yt-dlp-helper.ps1
```

3. On first run, it will:
   - Check/download the latest yt-dlp nightly build into the repo directory (`yt-dlp.exe`)
   - Download a portable FFmpeg zip, extract it to `ffmpeg_yt-dlp/`, and pass `--ffmpeg-location` to `yt-dlp`
   - Load/save preferences in `user_preferences.json`

4. Choose a menu option:
   - 1: Download Single Video (best quality + metadata)
   - 2: Download Playlist by URL (best quality + metadata)
   - 3: List & Download My Playlists (requires Firefox cookies)
   - 4: List & Download My Playlists (Force Refresh Cache)

### Download root and per-location preferences
- On start, the helper now asks for a download root path. The choice is remembered.
- Each download root keeps its own `user_preferences.json` and `download_archive.txt` inside that path, so you can maintain separate contexts for different drives/folders.
- A global `user_preferences.json` in the project stores the last used download root, the current `yt-dlp` version, and the saved Firefox profile path.

### GUI logging configuration
`yt-research-gui\yt-research-gui.ps1` reads logging settings from `yt-research-gui\yt-research-gui.config.json`:
```json
"logging": {
  "level": "INFO",
  "retentionDays": 14
}
```
- `level`: one of `DEBUG`, `INFO`, `WARN`, `ERROR`.
- `retentionDays`: keep only log entries newer than this many days at startup (set `0` to disable trimming).

### GUI Python project
- The GUI now has a dedicated `uv` project in `yt-research-gui/`.
- The Whisper helper package lives under `yt-research-gui/src/yt_research_gui_whisper/`.
- To refresh the GUI Python environment manually:
```powershell
cd .\yt-research-gui
uv sync
```

### GUI on-demand Whisper fallback
- The Transcript tab in `yt-research-gui\yt-research-gui.ps1` now exposes a `Transcribe with Whisper` button even when YouTube subtitles/auto-subs were found.
- Whisper runs in the background and the Transcript tab shows a live activity pane plus progress text while it works.
- The GUI still auto-detects your Firefox profile path, but `Use Firefox cookies` starts unchecked by default.
- If a cookie-backed YouTube request fails with the `not available on this app` error, the GUI automatically retries that request without Firefox cookies.
- The GUI keeps extracted audio in `yt-research-gui/data/audio/<videoId>.wav`.
- The GUI keeps local transcript text in `yt-research-gui/data/transcripts/<videoId>.txt`.
- If a local Whisper transcript already exists for that video ID, the button changes to `Load Whisper Transcript` and reuses the cached file.

### Where downloads go and what gets saved
- Output root: `Downloads/`
  - Single videos: `Downloads/%(title)s [%(id)s].%(ext)s`
  - Playlists: `Downloads/<Sanitized Playlist Title>/*`
- Metadata and extras (enabled by default):
  - `--write-description`
  - `--write-info-json`
  - `--write-subs` and `--write-auto-subs` with `--sub-langs "en.*,en"`
  - `--embed-metadata`, `--embed-thumbnail`, `--embed-subs`
- Duplicate prevention via archive file: `download_archive.txt` (prevents re-downloading the same video)
- Logs for playlist downloads: `_download.log` and `_download_error.log` inside the target playlist folder

### Authentication details (cookies)
- The script extracts your Firefox profile folder name and passes it to `yt-dlp` as:
  - `--cookies-from-browser firefox:<profileName>`
- The full Firefox profile path is stored locally in `user_preferences.json`.
- On startup, the helper validates the saved path. You can keep it, choose a newly detected profile, or paste a different path manually.

### Playlist listing and cache
- Listing your playlists uses the YouTube feed URL and cookies
- Results are cached to `cache/playlists_cache.json` for 24 hours
- Use menu option 4 to refresh immediately

---

## WhatsApp Audio Transcription
The Python script converts WhatsApp `.opus` files to 16kHz mono WAV and transcribes them using OpenAI Whisper.

### Prepare
- Place WhatsApp `.opus` files in `_tbt_audios/` (example files are included)
- Ensure FFmpeg exists under `ffmpeg_yt-dlp/ffmpeg-master-latest-win64-gpl/bin/ffmpeg.exe` (the PowerShell helper can install this for you)
- Install Python dependencies:
```powershell
python -m pip install --upgrade pip
pip install -U openai-whisper
# If Whisper prompts for PyTorch, install a CPU or CUDA build as appropriate, e.g.:
pip install torch --index-url https://download.pytorch.org/whl/cpu
```

### Run
```powershell
python .\whatsapp_transcribe.py
```

### Outputs
- WAV files: `Transcripts/wav_files/<original>.wav`
- Per-audio transcript: `Transcripts/<original>.txt`
- Combined transcript: `Transcripts/_combined_transcript.txt`

Notes:
- The script temporarily adds the FFmpeg `bin` folder to `PATH` so Whisper (and FFmpeg) can run cleanly on Windows.
- Default Whisper model is `turbo` and language is set to English. Edit the top of `whatsapp_transcribe.py` to change.

---

## Customization Tips
- Change subtitle languages by editing `--sub-langs` in `yt-dlp-helper.ps1` (default: `"en.*,en"`).
- Disable/enable metadata/thumbnails/subs by removing or adding flags in the `$commonFlags` array.
- Adjust playlist page size or cache duration in `List-And-Download-My-Playlists` if desired.

---

## Troubleshooting
- Firefox profile path error: Re-run the helper and choose one of the detected Firefox profiles, or re-check `about:profiles` and paste the correct path when prompted.
- `n challenge solving failed` / `Requested format is not available`: install Deno or Node.js 20+, then re-run the helper. It will auto-detect the runtime and pass `--js-runtimes` to `yt-dlp`.
- Corporate proxy/GitHub API issues: The helper will proceed with an existing local `yt-dlp.exe` if it canвЂ™t reach GitHub; otherwise it will stop with an error.
- Script execution blocked: Use `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` for your current session.
- Whisper install issues on Windows: Ensure Visual C++ Build Tools are installed if compilation is required, or use prebuilt wheels for Torch as shown above.

---

## FAQ
- Can I use a non-Nightly Firefox? Yes. Choose the profile you actually use in the helper prompt; the script only needs the folder name to pass to `yt-dlp`.
- Where is `yt-dlp.exe` stored? In the repo root alongside the script, updated from the nightly builds.
- How do I avoid re-downloading the same videos? The helper uses `--download-archive download_archive.txt` automatically. 

