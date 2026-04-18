# YT Research GUI Python Project

This folder contains the dedicated files for the YouTube research GUI:

- `yt-research-gui.ps1`: the actual GUI implementation
- `yt-research-gui.config.json`: GUI logging configuration
- `src/yt_research_gui_whisper/`: the GUI-facing Whisper wrapper package
- `..\tools\local_llm_text\`: the reusable local LLM helper used for Gemma 4 transcript post-processing
- `..\tools\whisper_transcribe_core\`: the reusable Whisper transcription core shared with the WhatsApp tool

The Python environment for the Whisper helper is managed with `uv`.

```powershell
cd .\yt-research-gui
uv sync
```

The top-level `..\yt-research-gui.ps1` file is only a launcher that preserves the old entry point.

The Transcript tab exposes an on-demand Whisper button, keeps audio/transcript files keyed by YouTube video ID, stores Whisper timing metadata beside the cached transcript, exposes a `Transcript Timing` tab for inspecting those timings, shows a live Whisper activity pane plus progress text while background transcription is running, exposes Gemma 4 prompt presets for transcript post-processing, caches Gemma results under `data/llm-results/`, auto-detects your Firefox profile path while leaving the cookies checkbox off by default, and automatically retries YouTube requests without Firefox cookies when the cookie-backed request hits the `not available on this app` error.
