import os
import re
import subprocess
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from tools.whisper_transcribe_core.core import (
    add_ffmpeg_to_path,
    get_whisper_model_cache_dir,
    load_whisper_model,
    transcribe_audio,
    write_transcription_outputs,
)


FFMPEG_DIR = os.path.join(os.getcwd(), r"ffmpeg_yt-dlp\ffmpeg-master-latest-win64-gpl\bin")
FFMPEG_EXE = os.path.join(FFMPEG_DIR, "ffmpeg.exe")

SCRIPT_DIR = os.getcwd()
AUDIO_DIR_NAME = "_tbt_audios"
TRANSCRIPTS_DIR_NAME = "Transcripts"
WAV_FILES_SUBDIR = "wav_files"

AUDIO_DIR = os.path.join(SCRIPT_DIR, AUDIO_DIR_NAME)
TRANSCRIPTS_DIR = os.path.join(SCRIPT_DIR, TRANSCRIPTS_DIR_NAME)
WAV_FILES_DIR = os.path.join(TRANSCRIPTS_DIR, WAV_FILES_SUBDIR)
COMBINED_TRANSCRIPT_FILE = os.path.join(TRANSCRIPTS_DIR, "_combined_transcript.txt")

DEFAULT_MODEL_NAME = "turbo"
LANGUAGE = "en"


def ensure_dir_exists(dir_path):
    if not os.path.exists(dir_path):
        print(f"  Creating directory: {dir_path}")
        os.makedirs(dir_path, exist_ok=True)
    else:
        print(f"  Directory already exists: {dir_path}")


def parse_datetime_from_filename(filename):
    match = re.search(r"(\d{4})-(\d{2})-(\d{2}) at (\d{2})\.(\d{2})\.(\d{2})", filename)
    if match:
        try:
            year, month, day, hour, minute, second = map(int, match.groups())
            return datetime(year, month, day, hour, minute, second)
        except ValueError:
            return None
    return None


def run_ffmpeg_conversion(ffmpeg_exe_path, input_path, output_path, step_name="FFmpeg conversion"):
    print(
        f"  -> Running {step_name}: {os.path.basename(ffmpeg_exe_path)} ... "
        f"'{os.path.basename(input_path)}' -> '{os.path.basename(output_path)}'"
    )
    command_list = [
        ffmpeg_exe_path,
        "-y",
        "-i",
        input_path,
        "-ac",
        "1",
        "-ar",
        "16000",
        output_path,
    ]
    try:
        start_time = time.time()
        creation_flags = 0
        if sys.platform == "win32":
            creation_flags = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(
            command_list,
            capture_output=True,
            text=True,
            check=False,
            encoding="utf-8",
            errors="replace",
            creationflags=creation_flags,
        )
        end_time = time.time()
        print(f"     FFmpeg took {end_time - start_time:.2f} seconds.")
        if result.returncode != 0:
            print(f"  Error during {step_name}:")
            error_output = result.stderr.strip() if result.stderr else result.stdout.strip()
            print(f"    Output: {error_output}")
            return False
        print(f"    Successfully created WAV: {output_path}")
        return True
    except FileNotFoundError:
        print(f"  Error: Executable not found for {step_name}. Path: {command_list[0]}")
        return False
    except Exception as exc:
        print(f"  An unexpected error occurred during {step_name}: {exc}")
        traceback.print_exc()
        return False


def select_whisper_model():
    return DEFAULT_MODEL_NAME


def main():
    print("\n--- WhatsApp Opus Transcription Script (Shared Whisper Core Mode) ---")

    print("Validating paths...")
    if not os.path.isfile(FFMPEG_EXE):
        print(f"Error: FFmpeg executable not found at {FFMPEG_EXE}")
        return
    print(f"  FFmpeg found: {FFMPEG_EXE}")
    if not os.path.isdir(AUDIO_DIR):
        print(f"Error: Audio directory not found at {AUDIO_DIR}")
        return
    print(f"  Audio directory found: {AUDIO_DIR}")

    print("Ensuring output directories exist...")
    ensure_dir_exists(TRANSCRIPTS_DIR)
    ensure_dir_exists(WAV_FILES_DIR)

    original_path = os.environ.get("PATH", "")
    path_modified = False
    try:
        add_ffmpeg_to_path(FFMPEG_DIR)
        path_modified = True
    except Exception as exc:
        print(f"Error adding FFmpeg to PATH: {exc}")
        traceback.print_exc()
        return

    try:
        chosen_model_name = select_whisper_model()
        print(f"Using Whisper model: '{chosen_model_name}'")

        model_cache_dir = get_whisper_model_cache_dir()
        print(f"\nLoading Whisper model '{chosen_model_name}' from '{model_cache_dir}'...")
        print("(This may take a while, especially on first download...)")
        load_start_time = time.time()
        try:
            model = load_whisper_model(chosen_model_name)
            load_end_time = time.time()
            print(
                f"  Whisper model '{chosen_model_name}' loaded successfully in "
                f"{load_end_time - load_start_time:.2f} seconds."
            )
        except Exception as exc:
            print(f"Error loading Whisper model '{chosen_model_name}': {exc}")
            traceback.print_exc()
            return

        print(f"\nScanning for .opus files in '{AUDIO_DIR}'...")
        opus_files = []
        for item in os.listdir(AUDIO_DIR):
            if item.lower().endswith(".opus"):
                dt = parse_datetime_from_filename(item)
                if dt:
                    opus_files.append((dt, item))
                else:
                    try:
                        mtime = datetime.fromtimestamp(os.path.getmtime(os.path.join(AUDIO_DIR, item)))
                        opus_files.append((mtime, item))
                        print(
                            f"  Warning: Could not parse datetime from filename '{item}'. Using mod time for sorting."
                        )
                    except Exception:
                        opus_files.append((datetime.min, item))
                        print(f"  Warning: Could not parse datetime for '{item}'. May be unsorted.")
        opus_files.sort()

        if not opus_files:
            print("No .opus files found to transcribe.")
            return

        print(f"Found {len(opus_files)} .opus files. Starting processing...")
        processed_transcripts = []

        for index, (_, filename) in enumerate(opus_files, start=1):
            base_name, _ = os.path.splitext(filename)
            opus_path = os.path.join(AUDIO_DIR, filename)
            wav_filename = f"{base_name}.wav"
            wav_path = os.path.join(WAV_FILES_DIR, wav_filename)
            transcript_path = os.path.join(TRANSCRIPTS_DIR, f"{base_name}.txt")
            timings_path = os.path.join(TRANSCRIPTS_DIR, f"{base_name}.timings.json")

            print(f"\n[{index}/{len(opus_files)}] Now processing: '{filename}'")

            if os.path.exists(wav_path):
                print(f"  Step 1: Found existing WAV: '{wav_path}'. Skipping FFmpeg conversion.")
            else:
                print("  Step 1: Converting Opus to WAV...")
                if not run_ffmpeg_conversion(FFMPEG_EXE, opus_path, wav_path):
                    print(f"  Skipping '{filename}' due to FFmpeg error.")
                    continue

            print("  Step 2: Transcribing WAV with Whisper...")
            transcribe_start_time = time.time()
            try:
                raw_result = transcribe_audio(
                    model,
                    wav_path,
                    language=LANGUAGE,
                    verbose=True,
                    fp16=False,
                    word_timestamps=True,
                )
                write_result = write_transcription_outputs(
                    raw_result,
                    transcript_path,
                    timings_output_path=timings_path,
                )
                transcribe_end_time = time.time()
                print(f"    Whisper transcription took {transcribe_end_time - transcribe_start_time:.2f} seconds.")
                print(f"    Successfully created transcript: {transcript_path}")
                print(f"    Timing metadata saved to: {timings_path}")
                processed_transcripts.append((base_name, transcript_path))
            except FileNotFoundError as exc:
                transcribe_end_time = time.time()
                print(
                    f"    Whisper transcription attempt took {transcribe_end_time - transcribe_start_time:.2f} "
                    "seconds before failing."
                )
                print(f"  [CRITICAL] Whisper FileNotFoundError for '{filename}' (using WAV: '{wav_path}'): {exc}")
                print("  Full traceback for FileNotFoundError:")
                traceback.print_exc()
                continue
            except Exception as exc:
                transcribe_end_time = time.time()
                print(
                    f"    Whisper transcription attempt took {transcribe_end_time - transcribe_start_time:.2f} "
                    "seconds before failing."
                )
                print(f"  Error during Whisper transcription for '{filename}' (using WAV: '{wav_path}'): {exc}")
                print("  Full traceback for other error:")
                traceback.print_exc()
                continue

        if processed_transcripts:
            print(f"\nStep 3: Combining {len(processed_transcripts)} transcripts...")
            with open(COMBINED_TRANSCRIPT_FILE, "w", encoding="utf-8") as outfile:
                for base_name, path in processed_transcripts:
                    outfile.write(f"### {base_name}\n\n")
                    try:
                        with open(path, "r", encoding="utf-8") as infile:
                            outfile.write(infile.read().strip())
                        outfile.write("\n\n")
                    except Exception as exc:
                        outfile.write(f"[Error reading transcript for {base_name}: {exc}]\n\n")
            print(f"  Combined transcript saved to: {COMBINED_TRANSCRIPT_FILE}")
        else:
            print("No transcripts were successfully processed to combine.")

    finally:
        if path_modified and "original_path" in locals():
            os.environ["PATH"] = original_path
            print("Restoring original PATH.")
        print("\n--- Transcription Complete ---")


if __name__ == "__main__":
    main()
