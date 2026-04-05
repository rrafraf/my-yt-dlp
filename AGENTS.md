# AGENTS.md

This file defines project-specific guidance for coding agents working in this repository.

## Project Purpose
- Windows-focused tooling around `yt-dlp` for:
- interactive YouTube download workflows (`yt-dlp-helper.ps1`)
- YouTube research GUI (`yt-research-gui\yt-research-gui.ps1`)
- audio transcription helper (`whatsapp_transcribe.py`)

## Repo Boundaries
- Treat `yt-dlp-helper.ps1` and `yt-research-gui\yt-research-gui.ps1` as distinct tools.
- Do not change both in one task unless the user explicitly asks or a shared contract requires it.
- Keep changes minimal and targeted to the request.

## Config vs State
- `yt-research-gui.config.json` is checked-in app config for GUI behavior (including logging).
- `user_preferences.json` is local runtime state for helper script behavior.
- `user_preferences.json` is gitignored and should not be treated as shared project config.
- Do not commit generated runtime artifacts (for example `logs/*`) unless user explicitly asks.

## Logging Guidance
- GUI logging level/retention are controlled by `yt-research-gui.config.json`.
- Keep default logging practical for normal use (`INFO`) and avoid noisy logs unless debugging is requested.
- Retention trimming is preferred over unbounded log growth.

## Change Style
- Prefer small, reviewable diffs.
- Preserve existing script behavior unless the request is behavioral change.
- Avoid broad refactors without user approval.
- Update `README.md` when behavior/config surfaces change.

## Testing Policy (Current)
- No mandatory automated test suite is required for routine changes.
- Do not add large or brittle test scaffolding unless explicitly requested.
- If validation is needed, prioritize high-value checks around user-critical flows (especially Fetch path in GUI), not trivial UI interactions.
- Manual GUI verification by user is acceptable.

## Useful Commands
- Run helper:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\yt-dlp-helper.ps1
```
- Run GUI:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\yt-research-gui\yt-research-gui.ps1
```

## Git/Commit Expectations
- Commit only files relevant to the requested change.
- Use concise, descriptive commit messages.
- Do not amend or rewrite history unless explicitly requested.
