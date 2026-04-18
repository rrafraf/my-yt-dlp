import argparse
import sys
import time
import traceback
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[3]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from tools.whisper_transcribe_core.core import (
    add_ffmpeg_to_path,
    configure_stdio,
    load_whisper_model,
    transcribe_audio,
    write_transcription_outputs,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Transcribe an audio file with Whisper.")
    parser.add_argument("--audio", required=True, help="Path to the extracted audio file.")
    parser.add_argument("--output", required=True, help="Path to the transcript text file.")
    parser.add_argument("--timings-output", default="", help="Optional path to a JSON file with Whisper timing metadata.")
    parser.add_argument("--ffmpeg-dir", default="", help="Directory that contains ffmpeg.exe.")
    parser.add_argument("--model", default="turbo", help="Whisper model name.")
    parser.add_argument("--language", default="en", help="Language code, or 'auto' for auto-detect.")
    return parser.parse_args()


def main() -> int:
    configure_stdio()
    args = parse_args()
    audio_path = Path(args.audio)
    output_path = Path(args.output)
    timings_output_path = Path(args.timings_output) if args.timings_output else None

    if not audio_path.is_file():
        print(f"Audio file not found: {audio_path}", file=sys.stderr)
        return 1

    print(f"Loading Whisper model '{args.model}'...")
    try:
        add_ffmpeg_to_path(args.ffmpeg_dir)
    except Exception:
        traceback.print_exc()
        return 1

    load_start = time.time()
    try:
        model = load_whisper_model(args.model)
    except Exception:
        traceback.print_exc()
        return 1
    print(f"Model loaded in {time.time() - load_start:.2f} seconds.")

    print(f"Transcribing '{audio_path.name}'...")
    transcribe_start = time.time()
    try:
        raw_result = transcribe_audio(
            model,
            audio_path,
            language=args.language,
            verbose=False,
            fp16=False,
            word_timestamps=True,
        )
        result = write_transcription_outputs(
            raw_result,
            output_path,
            timings_output_path=timings_output_path,
        )
    except Exception:
        traceback.print_exc()
        return 1

    print(f"Transcript written to '{output_path}'.")
    if timings_output_path is not None:
        print(f"Timing metadata written to '{timings_output_path}'.")
    print(f"Transcription took {time.time() - transcribe_start:.2f} seconds.")

    transcript_text = str(result.get("text", "")).strip()
    if transcript_text:
        print(f"Transcript length: {len(transcript_text)} characters.")
    else:
        print("Whisper completed, but produced no transcript text.")

    return 0
