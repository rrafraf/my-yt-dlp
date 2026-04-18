import json
import os
import sys
from pathlib import Path
from typing import Any


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


def load_whisper_model(model_name: str):
    try:
        import whisper
    except Exception as exc:
        raise RuntimeError(
            "Python package 'openai-whisper' is not installed. Install it before using Whisper transcription."
        ) from exc

    return whisper.load_model(model_name)


def sanitize_word(word: dict[str, Any]) -> dict[str, Any]:
    return {
        "start": word.get("start"),
        "end": word.get("end"),
        "word": str(word.get("word", "")).strip(),
        "probability": word.get("probability"),
    }


def sanitize_segment(segment: dict[str, Any]) -> dict[str, Any]:
    words = segment.get("words") or []
    return {
        "id": segment.get("id"),
        "seek": segment.get("seek"),
        "start": segment.get("start"),
        "end": segment.get("end"),
        "text": str(segment.get("text", "")).strip(),
        "tokens": segment.get("tokens") or [],
        "temperature": segment.get("temperature"),
        "avg_logprob": segment.get("avg_logprob"),
        "compression_ratio": segment.get("compression_ratio"),
        "no_speech_prob": segment.get("no_speech_prob"),
        "words": [sanitize_word(word) for word in words if isinstance(word, dict)],
    }


def build_timings_payload(result: dict[str, Any]) -> dict[str, Any]:
    transcript_text = str(result.get("text", "")).strip()
    return {
        "text": transcript_text,
        "language": result.get("language"),
        "duration": result.get("duration"),
        "segments": [
            sanitize_segment(segment)
            for segment in (result.get("segments") or [])
            if isinstance(segment, dict)
        ],
    }


def normalize_language(language: str | None) -> str | None:
    if language is None:
        return None
    if str(language).strip().lower() == "auto":
        return None
    return language


def transcribe_audio(
    model: Any,
    audio_path: str | Path,
    *,
    language: str | None = "en",
    verbose: bool = False,
    fp16: bool = False,
    word_timestamps: bool = True,
) -> dict[str, Any]:
    audio_path = Path(audio_path)
    if not audio_path.is_file():
        raise FileNotFoundError(f"Audio file not found: {audio_path}")

    return model.transcribe(
        str(audio_path),
        language=normalize_language(language),
        verbose=verbose,
        word_timestamps=word_timestamps,
        fp16=fp16,
    )


def write_transcription_outputs(
    result: dict[str, Any],
    output_path: str | Path,
    *,
    timings_output_path: str | Path | None = None,
) -> dict[str, Any]:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    transcript_text = str(result.get("text", "")).strip()
    output_path.write_text(transcript_text, encoding="utf-8")

    timings_payload = None
    if timings_output_path is not None:
        timings_output_path = Path(timings_output_path)
        timings_output_path.parent.mkdir(parents=True, exist_ok=True)
        timings_payload = build_timings_payload(result)
        timings_output_path.write_text(json.dumps(timings_payload, indent=2, ensure_ascii=False), encoding="utf-8")

    return {
        "text": transcript_text,
        "outputPath": str(output_path),
        "timingsOutputPath": str(timings_output_path) if timings_output_path is not None else "",
        "timingsPayload": timings_payload,
    }


def transcribe_audio_to_files(
    *,
    audio_path: str | Path,
    output_path: str | Path,
    ffmpeg_dir: str = "",
    model_name: str = "turbo",
    language: str | None = "en",
    timings_output_path: str | Path | None = None,
    verbose: bool = False,
    fp16: bool = False,
    word_timestamps: bool = True,
) -> dict[str, Any]:
    add_ffmpeg_to_path(ffmpeg_dir)
    model = load_whisper_model(model_name)
    result = transcribe_audio(
        model,
        audio_path,
        language=language,
        verbose=verbose,
        fp16=fp16,
        word_timestamps=word_timestamps,
    )
    return write_transcription_outputs(
        result,
        output_path,
        timings_output_path=timings_output_path,
    )
