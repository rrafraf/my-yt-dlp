from __future__ import annotations

from dataclasses import dataclass, field
import json
from pathlib import Path
from typing import Any


@dataclass
class TranscriptStudioSession:
    title: str = ""
    description: str = ""
    source_kind: str = ""
    source_id: str = ""
    source_url: str = ""
    subtitle: str = ""
    audio_path: str = ""
    transcript_path: str = ""
    timings_path: str = ""
    transcript_text: str = ""
    timings_payload: dict[str, Any] = field(default_factory=dict)
    metadata: dict[str, Any] = field(default_factory=dict)
    session_path: str = ""


def resolve_session_path(raw_path: str | None, base_dir: Path | None) -> str:
    if not raw_path:
        return ""
    candidate = Path(raw_path)
    if candidate.is_absolute() or base_dir is None:
        return str(candidate)
    return str((base_dir / candidate).resolve())


def read_text_file(path: str | Path) -> str:
    file_path = Path(path)
    try:
        return file_path.read_text(encoding="utf-8-sig")
    except FileNotFoundError as exc:
        raise FileNotFoundError(f"File not found: '{file_path}'.") from exc
    except UnicodeDecodeError as exc:
        raise ValueError(f"Could not decode '{file_path}' as UTF-8 text.") from exc
    except OSError as exc:
        raise ValueError(f"Failed to read '{file_path}': {exc}") from exc


def read_json_file(path: str | Path) -> dict[str, Any]:
    file_path = Path(path)
    try:
        payload = json.loads(read_text_file(file_path))
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"Invalid JSON in '{file_path}': {exc.msg} (line {exc.lineno}, column {exc.colno})."
        ) from exc
    if not isinstance(payload, dict):
        raise ValueError(f"Expected a JSON object in '{path}'.")
    return payload


def read_optional_text(path: str) -> str:
    if not path:
        return ""
    file_path = Path(path)
    if not file_path.is_file():
        return ""
    return read_text_file(file_path)


def read_optional_timings(path: str) -> dict[str, Any]:
    if not path:
        return {}
    file_path = Path(path)
    if not file_path.is_file():
        return {}
    payload = read_json_file(file_path)
    return payload


def build_session_from_payload(payload: dict[str, Any], *, session_path: str = "") -> TranscriptStudioSession:
    base_dir = Path(session_path).resolve().parent if session_path else None
    audio_path = resolve_session_path(str(payload.get("audioPath", "")).strip(), base_dir)
    transcript_path = resolve_session_path(str(payload.get("transcriptPath", "")).strip(), base_dir)
    timings_path = resolve_session_path(str(payload.get("timingsPath", "")).strip(), base_dir)

    transcript_text = str(payload.get("transcriptText", "")).strip()
    if not transcript_text and transcript_path:
        transcript_text = read_optional_text(transcript_path).strip()

    timings_payload: dict[str, Any] = {}
    if isinstance(payload.get("timings"), dict):
        timings_payload = payload["timings"]
    elif timings_path:
        timings_payload = read_optional_timings(timings_path)

    metadata = payload.get("metadata")
    if not isinstance(metadata, dict):
        metadata = {}

    return TranscriptStudioSession(
        title=str(payload.get("title", "")).strip(),
        description=str(payload.get("description", "")).strip(),
        source_kind=str(payload.get("sourceKind", "")).strip(),
        source_id=str(payload.get("sourceId", "")).strip(),
        source_url=str(payload.get("sourceUrl", "")).strip(),
        subtitle=str(payload.get("subtitle", "")).strip(),
        audio_path=audio_path,
        transcript_path=transcript_path,
        timings_path=timings_path,
        transcript_text=transcript_text,
        timings_payload=timings_payload,
        metadata=metadata,
        session_path=session_path,
    )


def load_session_file(path: str | Path) -> TranscriptStudioSession:
    session_path = str(Path(path).resolve())
    payload = read_json_file(session_path)
    return build_session_from_payload(payload, session_path=session_path)


def load_direct_session(
    *,
    audio: str = "",
    transcript: str = "",
    timings: str = "",
    title: str = "",
    description: str = "",
) -> TranscriptStudioSession:
    transcript_path = str(Path(transcript).resolve()) if transcript else ""
    timings_path = str(Path(timings).resolve()) if timings else ""
    audio_path = str(Path(audio).resolve()) if audio else ""

    transcript_text = read_optional_text(transcript_path).strip() if transcript_path else ""
    timings_payload = read_optional_timings(timings_path) if timings_path else {}

    return TranscriptStudioSession(
        title=title.strip(),
        description=description.strip(),
        source_kind="file",
        audio_path=audio_path,
        transcript_path=transcript_path,
        timings_path=timings_path,
        transcript_text=transcript_text,
        timings_payload=timings_payload,
    )
