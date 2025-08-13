import os
import subprocess
import re
from datetime import datetime
import shutil
import sys
import traceback # Ensure traceback is imported
import time # For timing operations

import whisper

# --- Configuration ---
# Paths (please verify these are correct for your system)
FFMPEG_DIR = os.path.join(os.getcwd(), r'ffmpeg_yt-dlp\ffmpeg-master-latest-win64-gpl\bin')
FFMPEG_EXE = os.path.join(FFMPEG_DIR, 'ffmpeg.exe')

# Directories
SCRIPT_DIR = os.getcwd() # Assumes script is in the project root
AUDIO_DIR_NAME = '_tbt_audios'
TRANSCRIPTS_DIR_NAME = 'Transcripts'
WAV_FILES_SUBDIR = 'wav_files' # Subdirectory for temporary WAV files

AUDIO_DIR = os.path.join(SCRIPT_DIR, AUDIO_DIR_NAME)
TRANSCRIPTS_DIR = os.path.join(SCRIPT_DIR, TRANSCRIPTS_DIR_NAME)
WAV_FILES_DIR = os.path.join(TRANSCRIPTS_DIR, WAV_FILES_SUBDIR)
COMBINED_TRANSCRIPT_FILE = os.path.join(TRANSCRIPTS_DIR, '_combined_transcript.txt')

# Whisper settings
DEFAULT_MODEL_NAME = "turbo" # Default model
LANGUAGE = 'en'       # Language for transcription (set to None for auto-detection)

# --- Helper Functions ---
def ensure_dir_exists(dir_path):
    """Creates a directory if it doesn't exist."""
    if not os.path.exists(dir_path):
        print(f"  Creating directory: {dir_path}")
        os.makedirs(dir_path, exist_ok=True)
    else:
        print(f"  Directory already exists: {dir_path}")

def parse_datetime_from_filename(filename):
    """Extracts datetime from filenames like 'WhatsApp Audio 2025-05-25 at 15.30.16_f4a667bb.waptt.opus'."""
    match = re.search(r'(\d{4})-(\d{2})-(\d{2}) at (\d{2})\.(\d{2})\.(\d{2})', filename)
    if match:
        try:
            year, month, day, hour, minute, second = map(int, match.groups())
            return datetime(year, month, day, hour, minute, second)
        except ValueError:
            return None
    return None

def run_ffmpeg_conversion(ffmpeg_exe_path, input_path, output_path, step_name="FFmpeg conversion"):
    """Runs FFmpeg to convert audio to 16kHz mono WAV."""
    print(f"  -> Running {step_name}: {os.path.basename(ffmpeg_exe_path)} ... '{os.path.basename(input_path)}' -> '{os.path.basename(output_path)}'")
    command_list = [
        ffmpeg_exe_path,
        '-y',  # Overwrite output files without asking
        '-i', input_path,
        '-ac', '1',  # Mono
        '-ar', '16000',  # 16kHz sample rate
        output_path
    ]
    try:
        start_time = time.time()
        # Use creationflags to hide the console window for FFmpeg on Windows
        creation_flags = 0
        if sys.platform == "win32":
            creation_flags = subprocess.CREATE_NO_WINDOW
        
        result = subprocess.run(command_list, capture_output=True, text=True, check=False, encoding='utf-8', errors='replace', creationflags=creation_flags)
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
    except Exception as e:
        print(f"  An unexpected error occurred during {step_name}: {e}")
        traceback.print_exc()
        return False

def select_whisper_model(): # This function is not currently called as per user's last change
    return DEFAULT_MODEL_NAME

# --- Main Script ---
def main():
    print("\n--- WhatsApp Opus Transcription Script (Python Library Mode) ---")

    # Validate paths
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

    # Add FFmpeg to PATH temporarily so Whisper library can find it if it needs to
    original_path = os.environ.get('PATH', '')
    ffmpeg_dir_lower = FFMPEG_DIR.lower()
    original_path_lower = original_path.lower()
    ffmpeg_in_path = ffmpeg_dir_lower in original_path_lower.split(os.pathsep)
    
    path_modified = False
    if not ffmpeg_in_path:
        print(f"Temporarily adding FFmpeg directory to PATH: {FFMPEG_DIR}")
        os.environ['PATH'] = FFMPEG_DIR + os.pathsep + original_path
        path_modified = True # Track if we actually changed it
    else:
        print(f"FFmpeg directory already in PATH: {FFMPEG_DIR}")

    try:
        # Select Whisper model
        chosen_model_name = select_whisper_model() # Currently hardcoded to DEFAULT_MODEL_NAME
        print(f"Using Whisper model: '{chosen_model_name}'")

        # Load Whisper model
        print(f"\nLoading Whisper model '{chosen_model_name}'...")
        print("(This may take a while, especially on first download...)")
        load_start_time = time.time()
        try:
            model = whisper.load_model(chosen_model_name)
            load_end_time = time.time()
            print(f"  Whisper model '{chosen_model_name}' loaded successfully in {load_end_time - load_start_time:.2f} seconds.")
        except Exception as e:
            print(f"Error loading Whisper model '{chosen_model_name}': {e}")
            traceback.print_exc()
            return

        # Find and sort opus files
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
                        print(f"  Warning: Could not parse datetime from filename '{item}'. Using mod time for sorting.")
                    except:
                        opus_files.append((datetime.min, item))
                        print(f"  Warning: Could not parse datetime for '{item}'. May be unsorted.")
        opus_files.sort()

        if not opus_files:
            print("No .opus files found to transcribe.")
            return

        print(f"Found {len(opus_files)} .opus files. Starting processing...")
        processed_transcripts = []

        for i, (dt, filename) in enumerate(opus_files):
            base_name, _ = os.path.splitext(filename)
            opus_path = os.path.join(AUDIO_DIR, filename)
            wav_filename = f"{base_name}.wav"
            wav_path = os.path.join(WAV_FILES_DIR, wav_filename)
            transcript_path = os.path.join(TRANSCRIPTS_DIR, f"{base_name}.txt")

            print(f"\n[{i+1}/{len(opus_files)}] Now processing: '{filename}'")

            # 1. Convert Opus to WAV (skip if WAV already exists)
            if os.path.exists(wav_path):
                print(f"  Step 1: Found existing WAV: '{wav_path}'. Skipping FFmpeg conversion.")
            else:
                print("  Step 1: Converting Opus to WAV...")
                if not run_ffmpeg_conversion(FFMPEG_EXE, opus_path, wav_path):
                    print(f"  Skipping '{filename}' due to FFmpeg error.")
                    continue
                # run_ffmpeg_conversion now prints success internally

            # 2. Transcribe WAV using Whisper library
            print("  Step 2: Transcribing WAV with Whisper...")
            transcribe_start_time = time.time()
            try:
                result = model.transcribe(wav_path, language=LANGUAGE, verbose=True, word_timestamps=True, fp16=False)
                transcribed_text = result["text"]
                transcribe_end_time = time.time()
                print(f"    Whisper transcription took {transcribe_end_time - transcribe_start_time:.2f} seconds.")
                
                with open(transcript_path, 'w', encoding='utf-8') as f:
                    f.write(transcribed_text.strip())
                print(f"    Successfully created transcript: {transcript_path}")
                processed_transcripts.append((base_name, transcript_path))
            except FileNotFoundError as fnf_error:
                transcribe_end_time = time.time()
                print(f"    Whisper transcription attempt took {transcribe_end_time - transcribe_start_time:.2f} seconds before failing.")
                print(f"  [CRITICAL] Whisper FileNotFoundError for '{filename}' (using WAV: '{wav_path}'): {fnf_error}")
                print("  Full traceback for FileNotFoundError:")
                traceback.print_exc()
                continue
            except Exception as e:
                transcribe_end_time = time.time()
                print(f"    Whisper transcription attempt took {transcribe_end_time - transcribe_start_time:.2f} seconds before failing.")
                print(f"  Error during Whisper transcription for '{filename}' (using WAV: '{wav_path}'): {e}")
                print("  Full traceback for other error:")
                traceback.print_exc()
                continue
                
        # 3. Combine transcripts
        if processed_transcripts:
            print(f"\nStep 3: Combining {len(processed_transcripts)} transcripts...")
            with open(COMBINED_TRANSCRIPT_FILE, 'w', encoding='utf-8') as outfile:
                for base_name, path in processed_transcripts:
                    outfile.write(f"### {base_name}\n\n")
                    try:
                        with open(path, 'r', encoding='utf-8') as infile:
                            outfile.write(infile.read().strip())
                        outfile.write("\n\n")
                    except Exception as e:
                        outfile.write(f"[Error reading transcript for {base_name}: {e}]\n\n")
            print(f"  Combined transcript saved to: {COMBINED_TRANSCRIPT_FILE}")
        else:
            print("No transcripts were successfully processed to combine.")

    finally:
        # Restore original PATH if it was modified and we actually changed it
        if path_modified and 'original_path' in locals(): # ensure original_path was set
            print(f"Restoring original PATH.")
            os.environ['PATH'] = original_path
        print("\n--- Transcription Complete ---")

if __name__ == "__main__":
    main() 