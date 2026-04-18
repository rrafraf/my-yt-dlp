#!/usr/bin/env python3
import argparse
import hashlib
import json
import sys
from pathlib import Path
from typing import Any
from urllib import error, request


OLLAMA_GENERATE_URL = "http://127.0.0.1:11434/api/generate"

PRESETS: dict[str, dict[str, str]] = {
    "clean-transcript": {
        "id": "clean-transcript",
        "label": "Clean Transcript",
        "description": "Improve sentence boundaries and punctuation without summarizing.",
        "instructions": (
            "Clean the transcript for readability. Preserve meaning. Fix punctuation, casing, and sentence "
            "boundaries. Do not summarize. Do not invent missing facts. Keep uncertain words if they are still "
            "the best reading. Leave summary, key points, action items, and participants empty unless the "
            "transcript itself makes them obvious."
        ),
    },
    "dialogue-analysis": {
        "id": "dialogue-analysis",
        "label": "Dialogue Analysis",
        "description": "Clean the transcript and infer whether it is a dialogue plus likely participants.",
        "instructions": (
            "Clean the transcript for readability. Determine whether the transcript is likely a dialogue or a "
            "single-speaker monologue. Infer likely participants or roles only when the transcript supports it, "
            "and call out uncertainty in warnings. Include a concise summary if useful."
        ),
    },
    "notes-summary": {
        "id": "notes-summary",
        "label": "Notes and Summary",
        "description": "Clean the transcript and extract a concise summary, key points, and action items.",
        "instructions": (
            "Clean the transcript for readability. Then extract a concise summary, key points, and explicit or "
            "strongly implied action items. Do not add action items that are not reasonably supported by the "
            "transcript. Mention uncertainty in warnings."
        ),
    },
}


def configure_stdio() -> None:
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            reconfigure(line_buffering=True)


def normalize_source_text(text: str) -> str:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    normalized = "\n".join(line.rstrip() for line in normalized.split("\n"))
    return normalized.strip()


def hash_source_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def list_presets_payload() -> list[dict[str, str]]:
    return [
        {
            "id": preset["id"],
            "label": preset["label"],
            "description": preset["description"],
        }
        for preset in PRESETS.values()
    ]


def build_prompt(preset: dict[str, str], transcript_text: str) -> str:
    schema_description = {
        "normalizedText": "string",
        "summary": "string",
        "isDialog": "boolean",
        "participants": [{"name": "string", "role": "string", "notes": "string"}],
        "keyPoints": ["string"],
        "actionItems": ["string"],
        "warnings": ["string"],
    }
    return (
        "You are processing a speech transcript. Return exactly one JSON object and nothing else.\n"
        "Use this JSON shape:\n"
        f"{json.dumps(schema_description, ensure_ascii=True)}\n"
        "Rules:\n"
        "- Keep all fields present.\n"
        "- Use empty strings or empty arrays when a field does not apply.\n"
        "- Do not wrap the JSON in markdown fences.\n"
        "- Do not invent facts that are not supported by the transcript.\n"
        f"- Task instructions: {preset['instructions']}\n"
        "Transcript follows between the markers.\n"
        "<<<TRANSCRIPT>>>\n"
        f"{transcript_text}\n"
        "<<<END_TRANSCRIPT>>>"
    )


def call_ollama(*, model: str, prompt: str, timeout_seconds: int) -> tuple[str, dict[str, Any]]:
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "format": "json",
        "options": {
            "temperature": 0.2,
        },
    }
    body = json.dumps(payload).encode("utf-8")
    req = request.Request(
        OLLAMA_GENERATE_URL,
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with request.urlopen(req, timeout=timeout_seconds) as response:
            response_bytes = response.read()
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"Ollama returned HTTP {exc.code} for model '{model}'. Response: {detail}"
        ) from exc
    except error.URLError as exc:
        raise RuntimeError(
            "Could not reach the local Ollama API at http://127.0.0.1:11434. "
            "Start Ollama and make sure the local server is available."
        ) from exc
    except TimeoutError as exc:
        raise RuntimeError(f"Ollama request timed out after {timeout_seconds} seconds.") from exc

    response_text = response_bytes.decode("utf-8", errors="replace")
    try:
        envelope = json.loads(response_text)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Ollama returned invalid JSON: {response_text}") from exc

    raw_response = str(envelope.get("response", "")).strip()
    if not raw_response:
        raise RuntimeError(f"Ollama returned an empty response. Envelope: {response_text}")
    return raw_response, envelope


def extract_json_object(text: str) -> Any:
    stripped = text.strip()
    if stripped.startswith("```"):
        stripped = stripped.strip("`")
        if "\n" in stripped:
            stripped = stripped.split("\n", 1)[1]
        stripped = stripped.strip()
        if stripped.endswith("```"):
            stripped = stripped[:-3].strip()

    try:
        return json.loads(stripped)
    except json.JSONDecodeError:
        decoder = json.JSONDecoder()
        for index, char in enumerate(stripped):
            if char != "{":
                continue
            try:
                obj, _ = decoder.raw_decode(stripped[index:])
                return obj
            except json.JSONDecodeError:
                continue
    raise RuntimeError(f"Model output did not contain a valid JSON object: {text}")


def sanitize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def sanitize_string_list(value: Any) -> list[str]:
    if not isinstance(value, list):
        return []
    result: list[str] = []
    seen: set[str] = set()
    for item in value:
        text = sanitize_text(item)
        if not text:
            continue
        if text in seen:
            continue
        seen.add(text)
        result.append(text)
    return result


def sanitize_participants(value: Any) -> list[dict[str, str]]:
    if not isinstance(value, list):
        return []
    participants: list[dict[str, str]] = []
    for item in value:
        if not isinstance(item, dict):
            continue
        name = sanitize_text(item.get("name"))
        role = sanitize_text(item.get("role"))
        notes = sanitize_text(item.get("notes"))
        if not name and not role and not notes:
            continue
        participants.append(
            {
                "name": name,
                "role": role,
                "notes": notes,
            }
        )
    return participants


def sanitize_result(model_output: Any, source_text: str) -> dict[str, Any]:
    if not isinstance(model_output, dict):
        raise RuntimeError("Model output root must be a JSON object.")

    normalized_text = sanitize_text(model_output.get("normalizedText"))
    if not normalized_text:
        normalized_text = source_text

    return {
        "normalizedText": normalized_text,
        "summary": sanitize_text(model_output.get("summary")),
        "isDialog": bool(model_output.get("isDialog", False)),
        "participants": sanitize_participants(model_output.get("participants")),
        "keyPoints": sanitize_string_list(model_output.get("keyPoints")),
        "actionItems": sanitize_string_list(model_output.get("actionItems")),
        "warnings": sanitize_string_list(model_output.get("warnings")),
    }


def format_participants(participants: list[dict[str, str]]) -> str:
    if not participants:
        return "No clear participants inferred."
    lines = []
    for participant in participants:
        name = participant.get("name") or "Unknown participant"
        role = participant.get("role") or "Unspecified role"
        notes = participant.get("notes") or ""
        line = f"- {name} ({role})"
        if notes:
            line = f"{line}: {notes}"
        lines.append(line)
    return "\n".join(lines)


def format_string_list(items: list[str], empty_message: str) -> str:
    if not items:
        return empty_message
    return "\n".join(f"- {item}" for item in items)


def build_display_text(preset_id: str, result: dict[str, Any]) -> str:
    sections: list[str] = []
    normalized_text = sanitize_text(result.get("normalizedText"))
    if normalized_text:
        sections.append(f"Cleaned Transcript\n\n{normalized_text}")

    if preset_id == "dialogue-analysis":
        dialogue_text = "Likely a dialogue." if result.get("isDialog") else "Likely not a dialogue."
        if sanitize_text(result.get("summary")):
            dialogue_text = f"{dialogue_text}\n\nSummary: {sanitize_text(result.get('summary'))}"
        sections.append(f"Dialogue Assessment\n\n{dialogue_text}")
        sections.append(f"Participants\n\n{format_participants(result.get('participants', []))}")
    elif preset_id == "notes-summary":
        summary = sanitize_text(result.get("summary")) or "No summary extracted."
        sections.append(f"Summary\n\n{summary}")
        sections.append(
            "Key Points\n\n"
            + format_string_list(result.get("keyPoints", []), "No key points extracted.")
        )
        sections.append(
            "Action Items\n\n"
            + format_string_list(result.get("actionItems", []), "No action items extracted.")
        )

    warnings = result.get("warnings", [])
    if warnings:
        sections.append("Warnings\n\n" + format_string_list(warnings, ""))

    return "\n\n".join(section.strip() for section in sections if section.strip())


def run_command(args: argparse.Namespace) -> int:
    preset = PRESETS.get(args.preset)
    if preset is None:
        print(f"Unknown preset: {args.preset}", file=sys.stderr)
        return 1

    input_path = Path(args.input_file)
    if not input_path.is_file():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 1

    output_path = Path(args.output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    source_text = input_path.read_text(encoding="utf-8")
    normalized_source = normalize_source_text(source_text)
    if not normalized_source:
        print("Input file did not contain any transcript text.", file=sys.stderr)
        return 1

    source_hash = hash_source_text(normalized_source)
    prompt = build_prompt(preset, normalized_source)
    raw_response, envelope = call_ollama(
        model=args.model,
        prompt=prompt,
        timeout_seconds=args.timeout_seconds,
    )
    parsed_response = extract_json_object(raw_response)
    result = sanitize_result(parsed_response, normalized_source)
    display_text = build_display_text(preset["id"], result)

    output_payload = {
        "presetId": preset["id"],
        "model": args.model,
        "sourceTextHash": source_hash,
        "displayText": display_text,
        "result": result,
        "rawResponse": raw_response,
        "ollama": {
            "done": bool(envelope.get("done")),
            "totalDuration": envelope.get("total_duration"),
            "loadDuration": envelope.get("load_duration"),
            "promptEvalCount": envelope.get("prompt_eval_count"),
            "evalCount": envelope.get("eval_count"),
        },
    }
    output_path.write_text(json.dumps(output_payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Result written to '{output_path}'.")
    return 0


def list_presets_command(_: argparse.Namespace) -> int:
    print(json.dumps(list_presets_payload(), indent=2, ensure_ascii=False))
    return 0


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Reusable local LLM text processing helper.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    list_parser = subparsers.add_parser("list-presets", help="List built-in prompt presets as JSON.")
    list_parser.set_defaults(func=list_presets_command)

    run_parser = subparsers.add_parser("run", help="Process text with a local Ollama model.")
    run_parser.add_argument("--model", default="gemma4", help="Ollama model name.")
    run_parser.add_argument("--preset", required=True, choices=sorted(PRESETS.keys()), help="Preset identifier.")
    run_parser.add_argument("--input-file", required=True, help="Path to a UTF-8 text input file.")
    run_parser.add_argument("--output-file", required=True, help="Path to the JSON output file.")
    run_parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=180,
        help="Request timeout for the Ollama API.",
    )
    run_parser.set_defaults(func=run_command)

    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    configure_stdio()
    args = parse_args(argv)
    try:
        return int(args.func(args))
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
