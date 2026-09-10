from __future__ import annotations

import argparse
import sys

from .session import load_direct_session, load_session_file


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Open Transcript Studio for a saved session or direct audio/transcript inputs.")
    parser.add_argument("--session", help="Path to a transcript studio session JSON file.")
    parser.add_argument("--audio", help="Path to a local audio file.")
    parser.add_argument("--transcript", help="Path to a plain transcript text file.")
    parser.add_argument("--timings", help="Path to Whisper timings JSON.")
    parser.add_argument("--title", default="", help="Optional title override for direct input mode.")
    parser.add_argument("--description", default="", help="Optional description for direct input mode.")
    return parser.parse_args(argv)


def load_session_from_args(args: argparse.Namespace):
    if args.session:
        return load_session_file(args.session)

    if args.audio or args.transcript or args.timings:
        return load_direct_session(
            audio=args.audio or "",
            transcript=args.transcript or "",
            timings=args.timings or "",
            title=args.title or "",
            description=args.description or "",
        )

    return load_direct_session(
        title=(args.title or "Transcript Studio"),
        description=args.description or "",
    )


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    session = load_session_from_args(args)

    try:
        from PySide6.QtWidgets import QApplication
    except Exception as exc:  # pragma: no cover - import-time environment issue
        raise SystemExit(
            "PySide6 is not installed. Run 'uv sync' in tools/transcript_studio before launching Transcript Studio."
        ) from exc

    from .mainwindow import TranscriptStudioWindow

    app = QApplication(sys.argv if argv is None else ["transcript-studio", *argv])
    window = TranscriptStudioWindow(session)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
