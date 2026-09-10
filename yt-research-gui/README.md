# YT Research GUI Python Project

This folder contains the dedicated files for the YouTube research GUI:

- `yt-research-gui.ps1`: the actual GUI implementation
- `yt-research-gui.config.json`: GUI logging configuration
- `src/yt_research_gui_whisper/`: the GUI-facing Whisper wrapper package
- `..\tools\local_llm_text\`: the reusable local LLM helper used for Ollama transcript post-processing
- `..\tools\whisper_transcribe_core\`: the reusable Whisper transcription core shared with the WhatsApp tool
- `..\tools\transcript_studio\`: the standalone Qt transcript exploration app that the GUI can launch once local audio is available

The Python environment for the Whisper helper is managed with `uv`.
Whisper model files are stored in `$env:XDG_CACHE_HOME\whisper`, configured on this machine as `D:\Documents\GitHub\.cache\whisper`, so GUI runs do not use the default user-profile cache.

```powershell
cd .\yt-research-gui
uv sync
```

The top-level `..\yt-research-gui.ps1` file is only a launcher that preserves the old entry point.

The main window now uses a sidebar/workspace split so metadata and logs stop crowding the transcript artifacts. The fetch field accepts either a raw YouTube video ID or a full URL. Browser cookie/profile controls live in a collapsible `Browser Options` expander. The Transcript workspace exposes an on-demand Whisper button, keeps audio/transcript files keyed by YouTube video ID, stores Whisper timing metadata beside the cached transcript, exposes a `Transcript Timing` tab for inspecting those timings, can hand the current local audio plus transcript session off to `Transcript Studio` through an `Open In Studio` button, shows live Whisper and Ollama activity inside collapsible panels, exposes Ollama model selection plus prompt presets for transcript post-processing, shows the selected prompt template directly in the workspace, adds an explicit `Stop` button for canceling the active Ollama helper, caches Ollama results under `data/llm-results/`, writes a durable processing bundle for each video under `data/video-bundles/<videoId>/`, auto-detects your Firefox profile path while leaving the cookies checkbox off by default, and automatically retries YouTube requests without Firefox cookies when the cookie-backed request hits the `not available on this app` error.

The processing bundle is the durable record for a processed video. It contains a `manifest.json`, an `events.ndjson` action log, fetch artifacts, audio extraction logs, Whisper runs, Ollama runs, and the latest Transcript Studio session export with relative internal paths so the whole folder can be moved elsewhere.

If `Open In Studio` fails before the window appears, check `yt-research-gui/logs/transcript-studio-*.stdout.log` and `yt-research-gui/logs/transcript-studio-*.stderr.log`. The GUI now writes those launcher logs for each handoff attempt.
