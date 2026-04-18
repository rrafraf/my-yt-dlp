## my-yt-dlp Helper (Windows)

A Windows-focused toolkit to:
- Automatically fetch and keep `yt-dlp.exe` up to date (nightly builds)
- Automatically download and wire up a portable FFmpeg for `yt-dlp`
- Download YouTube videos or playlists (with metadata, thumbnails, subtitles, and duplicate prevention)
- Optionally list your YouTube playlists (requires browser cookies)
- Transcribe WhatsApp voice notes: convert `.opus` в†’ `.wav` with FFmpeg and transcribe with Whisper

---

### WhatвЂ™s in this repo
- `yt-dlp-helper.ps1`: Compatibility launcher that preserves the old helper entry point.
- `yt-dlp-helper/`: Dedicated folder for the helper implementation, its local state, cache, and default downloads.
- `yt-research-gui\yt-research-gui.ps1`: Launcher for the research GUI.
- `yt-research-gui/`: Dedicated folder for the GUI implementation, config, Python project, logs, and cached transcript data.
- `tools/whisper_transcribe_core/`: Shared Whisper transcription core used by both the GUI and the WhatsApp transcription tool.
- `ffmpeg_yt-dlp/`: Shared portable FFmpeg used by both the helper and the GUI.
- `yt-dlp.exe`: Shared `yt-dlp` executable used by both the helper and the GUI.
- `tools\whatsapp_transcribe\whatsapp_transcribe.py`: Batch transcribes WhatsApp `.opus` files using FFmpeg + OpenAI Whisper.
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
  - Firefox (Nightly or regular). The helper can scan for local Firefox profiles and save your chosen profile path in `yt-dlp-helper\user_preferences.json`.
- For on-demand GUI Whisper transcription:
  - `uv` installed
  - Python 3.10+ (managed through `uv` in `yt-research-gui/`)
  - FFmpeg available under `ffmpeg_yt-dlp/` (installed by the helper script)
  - GUI Python dependencies installed with `uv sync` inside `yt-research-gui/`
- For GUI Gemma 4 transcript post-processing:
  - Ollama installed locally
  - Ollama running with the local API available at `http://127.0.0.1:11434`
  - A local `gemma4` model pulled into Ollama, or a different model name configured in `yt-research-gui.config.json`
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
4. The selected path is saved in `yt-dlp-helper\user_preferences.json` for future runs.

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
   - Check/download the latest shared `yt-dlp` nightly build into the repo root (`yt-dlp.exe`)
   - Download a shared portable FFmpeg zip, extract it to `ffmpeg_yt-dlp/`, and pass `--ffmpeg-location` to `yt-dlp`
   - Load/save helper preferences in `yt-dlp-helper\user_preferences.json`

4. Choose a menu option:
   - 1: Download Single Video (best quality + metadata)
   - 2: Download Playlist by URL (best quality + metadata)
   - 3: List & Download My Playlists (requires Firefox cookies)
   - 4: List & Download My Playlists (Force Refresh Cache)

### Download root and per-location preferences
- On start, the helper now asks for a download root path. The choice is remembered.
- Each download root keeps its own `user_preferences.json` and `download_archive.txt` inside that path, so you can maintain separate contexts for different drives/folders.
- The helper-global `yt-dlp-helper\user_preferences.json` stores the last used download root, the current `yt-dlp` version, and the saved Firefox profile path.

### GUI logging configuration
`yt-research-gui\yt-research-gui.ps1` reads logging settings from `yt-research-gui\yt-research-gui.config.json`:
```json
{
  "logging": {
    "level": "INFO",
    "retentionDays": 14
  },
  "ollama": {
    "model": "gemma4",
    "timeoutSeconds": 180
  }
}
```
- `level`: one of `DEBUG`, `INFO`, `WARN`, `ERROR`.
- `retentionDays`: keep only log entries newer than this many days at startup (set `0` to disable trimming).
- `ollama.model`: local Ollama model name used for transcript post-processing.
- `ollama.timeoutSeconds`: request timeout for the local Ollama API.

### GUI Python project
- The GUI now has a dedicated `uv` project in `yt-research-gui/`.
- The GUI Whisper wrapper lives under `yt-research-gui/src/yt_research_gui_whisper/`.
- The actual Whisper transcription core is shared at `tools/whisper_transcribe_core/`.
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
- The GUI keeps local Whisper timing metadata in `yt-research-gui/data/transcripts/<videoId>.timings.json` and exposes it in a `Transcript Timing` tab.
- If a local Whisper transcript already exists for that video ID, the button changes to `Load Whisper Transcript` and reuses the cached file.

### GUI Gemma 4 transcript post-processing
- The Transcript tab now includes a preset picker plus a `Run Gemma 4` button for post-processing the currently loaded transcript text.
- Built-in presets are:
  - `Clean Transcript`
  - `Dialogue Analysis`
  - `Notes and Summary`
- Gemma 4 processing is a post-transcript step only. It uses the current transcript text from YouTube subtitles or the latest Whisper result.
- Results are shown in a separate `Gemma Result` panel beside the transcript text.
- Gemma 4 runs in the background through the reusable helper in `tools\local_llm_text\`.
- Cached Gemma results are stored under `yt-research-gui/data/llm-results/<videoId>/`.
- Cache keys include the model, preset, and transcript hash, so updating the transcript causes a fresh Gemma run instead of reusing stale output.
- If a matching cached result exists, the GUI loads it immediately without calling Ollama again.

### Where downloads go and what gets saved
- Default output root: `yt-dlp-helper/Downloads/`
- If you choose a different root at launch, the same structure is created under that selected path.
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
- The full Firefox profile path is stored locally in `yt-dlp-helper\user_preferences.json`.
- On startup, the helper validates the saved path. You can keep it, choose a newly detected profile, or paste a different path manually.

### Playlist listing and cache
- Listing your playlists uses the YouTube feed URL and cookies
- Results are cached to `yt-dlp-helper/cache/playlists_cache.json` for 24 hours
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
python .\tools\whatsapp_transcribe\whatsapp_transcribe.py
```

### Outputs
- WAV files: `Transcripts/wav_files/<original>.wav`
- Per-audio transcript: `Transcripts/<original>.txt`
- Per-audio Whisper timing metadata: `Transcripts/<original>.timings.json`
- Combined transcript: `Transcripts/_combined_transcript.txt`

Notes:
- The script temporarily adds the FFmpeg `bin` folder to `PATH` so Whisper (and FFmpeg) can run cleanly on Windows.
- Default Whisper model is `turbo` and language is set to English. Edit the top of `tools\whatsapp_transcribe\whatsapp_transcribe.py` to change.

---

## Customization Tips
- Change subtitle languages by editing `--sub-langs` in `yt-dlp-helper/yt-dlp-helper.ps1` (default: `"en.*,en"`).
- Disable/enable metadata/thumbnails/subs by removing or adding flags in the `$commonFlags` array.
- Adjust playlist page size or cache duration in `List-And-Download-My-Playlists` if desired.

---

## Troubleshooting
- Firefox profile path error: Re-run the helper and choose one of the detected Firefox profiles, or re-check `about:profiles` and paste the correct path when prompted.
- `n challenge solving failed` / `Requested format is not available`: install Deno or Node.js 20+, then re-run the helper. It will auto-detect the runtime and pass `--js-runtimes` to `yt-dlp`.
- Corporate proxy/GitHub API issues: The helper will proceed with an existing local `yt-dlp.exe` if it canвЂ™t reach GitHub; otherwise it will stop with an error.
- Script execution blocked: Use `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` for your current session.
- Whisper install issues on Windows: Ensure Visual C++ Build Tools are installed if compilation is required, or use prebuilt wheels for Torch as shown above.
- Gemma 4 helper says Ollama is unavailable: Start Ollama and confirm the local API responds on `127.0.0.1:11434`, then re-run the preset from the Transcript tab.
- Gemma 4 request times out: increase `ollama.timeoutSeconds` in `yt-research-gui.config.json`, or switch the GUI to a lighter local Ollama model by changing `ollama.model`.

---

## FAQ
- Can I use a non-Nightly Firefox? Yes. Choose the profile you actually use in the helper prompt; the script only needs the folder name to pass to `yt-dlp`.
- Where is `yt-dlp.exe` stored? In the repo root as a shared tool for both the helper and the GUI, updated from the nightly builds.
- How do I avoid re-downloading the same videos? The helper uses `--download-archive download_archive.txt` automatically. 

