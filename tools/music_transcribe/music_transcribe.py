#!/usr/bin/env python3
import argparse
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List, Optional

from colorama import Fore, Style

# Optional imports; we will check availability at runtime and degrade gracefully
try:
	import soundfile as sf  # noqa: F401
except Exception:
	pass

# Basic Pitch for general transcription
try:
	from basic_pitch.inference import predict  # noqa: F401
	from basic_pitch import ICASSP_2022_MODEL_PATH
	export_model_path = ICASSP_2022_MODEL_PATH
except Exception:
	predict = None
	export_model_path = None

# pretty_midi for MIDI IO
try:
	import pretty_midi
except Exception:
	pretty_midi = None


def run(cmd: List[str], cwd: Optional[Path] = None) -> int:
	print(Fore.CYAN + "$ " + " ".join(cmd) + Style.RESET_ALL)
	proc = subprocess.run(cmd, cwd=str(cwd) if cwd else None)
	return proc.returncode


def ensure_dir(p: Path) -> None:
	p.mkdir(parents=True, exist_ok=True)


def ffmpeg_extract_audio(input_video: Path, output_wav: Path, ffmpeg_path: Optional[Path]) -> None:
	cmd = [str(ffmpeg_path or 'ffmpeg'), '-y', '-i', str(input_video), '-vn', '-ac', '2', '-ar', '44100', str(output_wav)]
	rc = run(cmd)
	if rc != 0:
		raise RuntimeError(f"ffmpeg failed to extract audio from {input_video}")


def demucs_separate(audio_wav: Path, out_dir: Path, demucs_model: str, use_gpu: bool) -> Path:
	# Demucs outputs structure under out_dir/demucs/<model>/<track_name>/*.wav
	cmd = ['demucs', '-n', demucs_model]
	if not use_gpu:
		cmd += ['-d', 'cpu']
	cmd += ['-o', str(out_dir), str(audio_wav)]
	rc = run(cmd)
	if rc != 0:
		raise RuntimeError("Demucs separation failed")
	# Try to locate the specific output folder
	candidates = list(out_dir.rglob(audio_wav.stem))
	if candidates:
		return candidates[0]
	return out_dir


def basic_pitch_to_midi(audio_path: Path, midi_out: Path) -> None:
	if predict is None or pretty_midi is None:
		raise RuntimeError("Basic Pitch or pretty_midi not available. Ensure requirements are installed.")
	from basic_pitch import inference
	inference.predict_and_save(
		[str(audio_path)],
		output_dir=str(midi_out.parent),
		save_midi=True,
		sonify_midi=False,
		save_model_outputs=False,
		verbose=False,
		model_or_model_path=export_model_path,
	)
	generated = midi_out.parent / f"{audio_path.stem}_basic_pitch.mid"
	if generated.exists():
		shutil.move(str(generated), str(midi_out))


def combine_midis(midi_files: List[Path], combined_out: Path) -> None:
	if pretty_midi is None:
		raise RuntimeError("pretty_midi not available")
	combined = pretty_midi.PrettyMIDI()
	for mf in midi_files:
		if not mf.exists():
			continue
		m = pretty_midi.PrettyMIDI(str(mf))
		for instr in m.instruments:
			combined.instruments.append(instr)
	combined.write(str(combined_out))


def is_media_file(p: Path) -> bool:
	return p.suffix.lower() in {'.mp4', '.mkv', '.webm', '.wav', '.mp3', '.flac', '.m4a'}


def main():
	parser = argparse.ArgumentParser(description="Music transcription pipeline: extract -> separate -> transcribe")
	parser.add_argument('--input', required=True, help='Input file or directory of media')
	parser.add_argument('--out', default='music_outputs', help='Output root directory')
	parser.add_argument('--ffmpeg', default=None, help='Path to ffmpeg executable')
	parser.add_argument('--demucs-model', default='htdemucs', help='Demucs model name (e.g., htdemucs, mdx_extra_q)')
	parser.add_argument('--cpu', action='store_true', help='Force CPU for Demucs')
	parser.add_argument('--skip-separation', action='store_true', help='Skip Demucs separation (expects stems present)')
	parser.add_argument('--skip-transcribe', action='store_true', help='Skip transcription')
	args = parser.parse_args()

	inp = Path(args.input)
	out_root = Path(args.out)
	ensure_dir(out_root)

	ffmpeg_path = Path(args.ffmpeg) if args.ffmpeg else None

	media_files: List[Path] = []
	if inp.is_dir():
		for p in inp.rglob('*'):
			if p.is_file() and is_media_file(p):
				media_files.append(p)
	elif inp.is_file() and is_media_file(inp):
		media_files.append(inp)
	else:
		print(Fore.RED + f"No valid media found at {inp}" + Style.RESET_ALL)
		return 1

	for media in media_files:
		print(Fore.GREEN + f"Processing: {media}" + Style.RESET_ALL)
		work_dir = out_root / media.stem
		ensure_dir(work_dir)

		# Step 1: Extract audio
		wav_path = work_dir / f"{media.stem}.wav"
		if not wav_path.exists():
			ffmpeg_extract_audio(media, wav_path, ffmpeg_path)
		else:
			print(Fore.YELLOW + f"Audio exists: {wav_path}" + Style.RESET_ALL)

		# Step 2: Separation
		stems_dir = work_dir / 'stems'
		ensure_dir(stems_dir)
		if not args.skip_separation:
			sep_target = demucs_separate(wav_path, stems_dir, args.demucs_model, use_gpu=not args.cpu)
			print(Fore.CYAN + f"Demucs output at: {sep_target}" + Style.RESET_ALL)
		else:
			print(Fore.YELLOW + "Skipping separation as requested" + Style.RESET_ALL)

		# Heuristic: locate stems
		vocals = next(stems_dir.rglob('*vocals*.wav'), None)
		bass = next(stems_dir.rglob('*bass*.wav'), None)
		other = next(stems_dir.rglob('*other*.wav'), None)
		piano = next(stems_dir.rglob('*piano*.wav'), None)

		# Step 3: Transcription per stem (Basic Pitch baseline)
		midi_dir = work_dir / 'midi'
		ensure_dir(midi_dir)
		midi_paths: List[Path] = []

		if not args.skip_transcribe:
			for stem_name, stem_path in [('vocals', vocals), ('bass', bass), ('other', other), ('piano', piano)]:
				if stem_path and stem_path.exists():
					out_midi = midi_dir / f"{stem_name}.mid"
					try:
						basic_pitch_to_midi(stem_path, out_midi)
						midi_paths.append(out_midi)
						print(Fore.GREEN + f"Transcribed {stem_name} -> {out_midi}" + Style.RESET_ALL)
					except Exception as e:
						print(Fore.RED + f"Failed to transcribe {stem_name}: {e}" + Style.RESET_ALL)
		else:
			print(Fore.YELLOW + "Skipping transcription as requested" + Style.RESET_ALL)

		# Step 4: Combine MIDI
		if midi_paths:
			combined = work_dir / 'midi' / 'combined.mid'
			try:
				combine_midis(midi_paths, combined)
				print(Fore.CYAN + f"Combined MIDI -> {combined}" + Style.RESET_ALL)
			except Exception as e:
				print(Fore.RED + f"Failed to combine MIDI: {e}" + Style.RESET_ALL)

	print(Fore.GREEN + "All done." + Style.RESET_ALL)
	return 0


if __name__ == '__main__':
	sys.exit(main())
