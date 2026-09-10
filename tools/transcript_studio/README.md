# Transcript Studio

`Transcript Studio` is a standalone Qt desktop tool for exploring transcript sessions with Whisper word timings.

It is designed to stay generic:
- any local audio file can be opened
- it can consume the existing Whisper timing JSON produced in this repo
- it exposes rendering controls so pause and confidence metadata can change how the transcript is visualized

## Current Prototype

The first prototype includes:
- timed reading view derived from word timings
- raw transcript view
- timing JSON view
- session metadata view
- transcript statistics
- sliders for pause sensitivity and confidence styling
- clickable words that seek the loaded audio

## Setup

Install the local environment with `uv`:

```powershell
cd .\tools\transcript_studio
uv sync
```

## Run

Open a saved session file:

```powershell
uv run --project . transcript-studio --session C:\path\to\session.json
```

Open the app with no input and load a session later from `File -> Open Session`:

```powershell
uv run --project . transcript-studio
```

Or pass the files directly:

```powershell
uv run --project . transcript-studio `
  --audio C:\path\to\audio.wav `
  --transcript C:\path\to\transcript.txt `
  --timings C:\path\to\transcript.timings.json `
  --title "My clip"
```

## Session Shape

The app accepts a session JSON payload with fields such as:

```json
{
  "sessionVersion": 1,
  "title": "Video title",
  "sourceKind": "youtube",
  "sourceId": "dQw4w9WgXcQ",
  "sourceUrl": "https://www.youtube.com/watch?v=dQw4w9WgXcQ",
  "audioPath": "C:\\path\\to\\audio.wav",
  "transcriptPath": "C:\\path\\to\\transcript.txt",
  "timingsPath": "C:\\path\\to\\transcript.timings.json",
  "transcriptText": "Plain transcript text",
  "timings": {
    "text": "Plain transcript text",
    "segments": []
  },
  "metadata": {}
}
```
