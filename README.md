## my-yt-dlp

Windows-focused tooling around `yt-dlp`, FFmpeg, and a small amount of Python for transcription workflows.

---

### WhatвЂ™s in this repo
- `yt-dlp-helper.ps1`: Compatibility launcher that preserves the old helper entry point.
- `yt-dlp-helper/`: Dedicated folder for the helper implementation, its local state, cache, and default downloads.
- `yt-research-gui\yt-research-gui.ps1`: Launcher for the research GUI.
- `yt-research-gui/`: Dedicated folder for the GUI implementation, config, Python project, logs, and cached transcript data.
- `tools/whisper_transcribe_core/`: Shared Whisper transcription core used by both the GUI and the WhatsApp transcription tool.
- Whisper models are stored in the user-level cache at `D:\Documents\GitHub\.cache\whisper` via `XDG_CACHE_HOME`.
- `tools/transcript_studio/`: Standalone Qt transcript exploration app for working with audio plus Whisper timing metadata.
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
- For GUI Ollama transcript post-processing:
  - Ollama installed locally
  - Ollama running with the local API available at `http://127.0.0.1:11434`
  - At least one local model pulled into Ollama (the GUI defaults to `gemma4` when available)
- For Transcript Studio:
  - `uv` installed
  - Python 3.10+ (managed through `uv` in `tools/transcript_studio/`)
  - Qt dependencies installed with `uv sync` inside `tools/transcript_studio/`
- For WhatsApp transcription:
  - Python 3.9+ recommended
  - FFmpeg (auto-installed into `ffmpeg_yt-dlp` by the helper script, used by the Python script)
  - Python packages: `openai-whisper` (and its dependencies, e.g. Torch)

## yt-dlp Helper
The helper now has a dedicated package structure under `yt-dlp-helper/`.

Shared core:
- `yt-dlp-helper/Helper.Core.ps1`

Entry points:
- CLI: `yt-dlp-helper/yt-dlp-helper.ps1`
- GUI: `yt-dlp-helper/yt-dlp-helper-gui.ps1`

### Run the helper CLI
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\yt-dlp-helper.ps1
```

### Run the helper GUI
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\yt-dlp-helper\yt-dlp-helper-gui.ps1
```

### Helper GUI workbench
The helper GUI is task-oriented and non-blocking. Current task surfaces:
- `Single Video`
- `Playlist URL`
- `My Playlists`
- `Environment`
- `Folder Inspector`

Long-running work runs in a background worker. The activity pane shows live stdout/stderr, current job state, cancel, and recent job history.

### Helper defaults and state
- Helper preferences live in `yt-dlp-helper/user_preferences.json`
- Playlist cache lives in `yt-dlp-helper/cache/playlists_cache.json`
- GUI logs and per-job logs live under `yt-dlp-helper/logs/`
- Default helper download root is `yt-dlp-helper/Downloads/`
- The selected download root is remembered and can be changed from CLI or GUI

### Download layout
- Single videos: `<download root>/Singles/%(title)s [%(id)s].%(ext)s`
- Single-video duplicate tracking: `<download root>/Singles/download_archive.txt`
- Playlist downloads: `<download root>/<Playlist Title> [<playlistId>]/`
- Playlist duplicate tracking: `<playlist folder>/download_archive.txt`
- Optional playlist sidecars: `<playlist folder>/_sidecar/`
- Converted subtitle text files: `<playlist folder>/_sidecar/text/`

Main media downloads use the current helper preset, which embeds metadata and subtitles into the media output. Optional sidecar fetches collect `.info.json` and subtitle files into `_sidecar/`.

### Folder Inspector
The helper GUI includes an offline folder inspector for existing playlist folders. It reports only what local artifacts can prove or infer, including:
- archive presence and entry count
- playlist metadata file details when present
- media, `.info.json`, subtitle, and transcript counts
- inferred or confirmed playlist title and id
- certainty states such as `Confirmed`, `Inferred`, and `Unknown`

### Helper GUI logging config
`yt-dlp-helper/yt-dlp-helper.gui.config.json` controls GUI logging:
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
- `ollama.model`: default Ollama model the GUI tries to preselect in the model dropdown.
- `ollama.timeoutSeconds`: request timeout for the local Ollama API.

### GUI Python project
- The GUI now has a dedicated `uv` project in `yt-research-gui/`.
- The GUI Whisper wrapper lives under `yt-research-gui/src/yt_research_gui_whisper/`.
- The actual Whisper transcription core is shared at `tools/whisper_transcribe_core/`.
- Whisper model files are loaded from `$env:XDG_CACHE_HOME\whisper`, configured on this machine as `D:\Documents\GitHub\.cache\whisper`.
- To refresh the GUI Python environment manually:
```powershell
cd .\yt-research-gui
uv sync
```

### Transcript Studio
- `tools/transcript_studio/` is a separate Qt app for transcript exploration once local audio and Whisper timing data exist.
- The YouTube research GUI exposes an `Open In Studio` button after a local Whisper run has produced audio and transcript artifacts.
- Transcript Studio reads a session JSON file written by the GUI and can also be launched directly against local audio/transcript/timing files.
  - It can also be launched with no arguments and then opened from `File -> Open Session`.
  - If the GUI handoff fails before the window appears, inspect `yt-research-gui/logs/transcript-studio-*.stdout.log` and `.stderr.log`.
- To install its dependencies manually:
```powershell
cd .\tools\transcript_studio
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

### GUI Ollama transcript post-processing
- The Transcript tab now includes an Ollama model dropdown, a preset picker, a `Run with Ollama` button, and a `Stop` button for canceling the active Ollama helper.
- The fetch field now accepts either a raw 11-character YouTube video ID or a full YouTube URL.
- Firefox cookie/profile controls now live under a collapsible `Browser Options` expander instead of taking permanent space at the top of the window.
- The selected Ollama prompt template is shown directly in the Transcript workspace so you can see which preset instructions are active before running the model.
- Built-in presets are:
  - `Clean Transcript`
  - `Dialogue Analysis`
  - `Notes and Summary`
- Ollama processing is a post-transcript step only. It uses the current transcript text from YouTube subtitles or the latest Whisper result.
- Results are shown in a separate `Ollama Result` panel beside the transcript text.
- The Transcript tab now also includes a live `Ollama Activity` pane that mirrors helper stdout/stderr, including streamed request progress and failure details.
- Closing the GUI while Ollama is running kills the background helper process instead of leaving it attached to the window.
- Ollama runs in the background through the reusable helper in `tools\local_llm_text\`.
- Cached Ollama results are stored under `yt-research-gui/data/llm-results/<videoId>/`.
- Cache keys include the selected model, preset, and transcript hash, so updating the transcript or switching models causes a fresh Ollama run instead of reusing stale output.
- If a matching cached result exists, the GUI loads it immediately without calling Ollama again.

### GUI video bundles
- The GUI now writes a durable per-video processing bundle under `yt-research-gui/data/video-bundles/<videoId>/`.
- Each bundle is intended to be self-contained and transferable. It keeps:
  - `manifest.json`: current artifact index, latest pointers, and tool metadata
  - `events.ndjson`: append-only action history for fetch, audio extraction, Whisper, Ollama, and Transcript Studio export/open events
  - `source/fetches/<runId>/`: metadata JSON, subtitle fetch output, selected caption file, and extracted YouTube transcript text
  - `audio/` and `audio/extractions/<runId>/`: bundled audio plus extraction logs
  - `whisper/runs/<runId>/`: transcript, timings JSON, stdout/stderr logs, and run metadata
  - `llm/runs/<runId>/`: input transcript, preset info, rendered prompt, JSON result, display text, stdout/stderr logs, and run metadata
  - `exports/transcript-studio.session.json`: the latest Transcript Studio handoff file
- Paths inside the bundle manifest and event log are relative so the bundle can be moved without breaking internal references.

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
- Whisper model files are stored under `$env:XDG_CACHE_HOME\whisper` so the WhatsApp helper, GUI, and other Whisper projects can share the same non-profile cache.
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
- Ollama helper says Ollama is unavailable: Start Ollama and confirm the local API responds on `127.0.0.1:11434`, then re-run the preset from the Transcript tab.
- Ollama request times out or model loading is too slow: increase `ollama.timeoutSeconds` in `yt-research-gui.config.json`, or switch the GUI to a lighter model from the Ollama model dropdown.

---

## FAQ
- Can I use a non-Nightly Firefox? Yes. Choose the profile you actually use in the helper prompt; the script only needs the folder name to pass to `yt-dlp`.
- Where is `yt-dlp.exe` stored? In the repo root as a shared tool for both the helper and the GUI, updated from the nightly builds.
- How do I avoid re-downloading the same videos? The helper uses `--download-archive download_archive.txt` automatically. 

## VS Code helpers
This repo includes `.vscode/launch.json` and `.vscode/tasks.json` so the main tools can be started directly from VS Code without typing the launch commands each time.
