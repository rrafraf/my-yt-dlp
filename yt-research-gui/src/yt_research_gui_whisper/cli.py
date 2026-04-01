import argparse
import os
import sys
import time
import traceback
from pathlib import Path


def configure_stdio() -> None:
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            reconfigure(line_buffering=True)


def add_ffmpeg_to_path(ffmpeg_dir: str) -> None:
    if not ffmpeg_dir:
        return

    ffmpeg_path = Path(ffmpeg_dir)
    ffmpeg_exe = ffmpeg_path / "ffmpeg.exe"
    if not ffmpeg_exe.is_file():
        raise FileNotFoundError(f"ffmpeg.exe was not found in '{ffmpeg_dir}'")

    current_path = os.environ.get("PATH", "")
    parts = current_path.split(os.pathsep) if current_path else []
    normalized = {os.path.normcase(os.path.normpath(part)) for part in parts if part}
    ffmpeg_norm = os.path.normcase(os.path.normpath(str(ffmpeg_path)))
    if ffmpeg_norm not in normalized:
        os.environ["PATH"] = str(ffmpeg_path) + os.pathsep + current_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Transcribe an audio file with Whisper.")
    parser.add_argument("--audio", required=True, help="Path to the extracted audio file.")
    parser.add_argument("--output", required=True, help="Path to the transcript text file.")
    parser.add_argument("--ffmpeg-dir", default="", help="Directory that contains ffmpeg.exe.")
    parser.add_argument("--model", default="turbo", help="Whisper model name.")
    parser.add_argument("--language", default="en", help="Language code, or 'auto' for auto-detect.")
    return parser.parse_args()


def main() -> int:
    configure_stdio()
    args = parse_args()
    audio_path = Path(args.audio)
    output_path = Path(args.output)

    if not audio_path.is_file():
        print(f"Audio file not found: {audio_path}", file=sys.stderr)
        return 1

    try:
        add_ffmpeg_to_path(args.ffmpeg_dir)
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 1

    try:
        import whisper
    except Exception:
        print(
            "Python package 'openai-whisper' is not installed. Run 'uv sync' in the GUI folder before using Whisper transcription.",
            file=sys.stderr,
        )
        return 1

    output_path.parent.mkdir(parents=True, exist_ok=True)

    language = None if str(args.language).strip().lower() == "auto" else args.language

    print(f"Loading Whisper model '{args.model}'...")
    load_start = time.time()
    try:
        model = whisper.load_model(args.model)
    except Exception:
        traceback.print_exc()
        return 1
    print(f"Model loaded in {time.time() - load_start:.2f} seconds.")

    print(f"Transcribing '{audio_path.name}'...")
    transcribe_start = time.time()
    try:
        result = model.transcribe(
            str(audio_path),
            language=language,
            verbose=False,
            word_timestamps=False,
            fp16=False,
        )
    except Exception:
        traceback.print_exc()
        return 1

    transcript_text = str(result.get("text", "")).strip()
    output_path.write_text(transcript_text, encoding="utf-8")
    print(f"Transcript written to '{output_path}'.")
    print(f"Transcription took {time.time() - transcribe_start:.2f} seconds.")
    if transcript_text:
        print(f"Transcript length: {len(transcript_text)} characters.")
    else:
        print("Whisper completed, but produced no transcript text.")

    return 0
