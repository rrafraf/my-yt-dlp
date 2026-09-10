from __future__ import annotations

from dataclasses import dataclass
import html
import math
import statistics
from typing import Any


@dataclass(frozen=True)
class TimingWord:
    index: int
    text: str
    start: float
    end: float
    probability: float | None
    segment_index: int

    @property
    def duration(self) -> float:
        return max(0.0, self.end - self.start)


@dataclass
class RenderSettings:
    pause_sensitivity: int = 100
    pause_marker_threshold_tenths: int = 18
    low_confidence_threshold: int = 72
    confidence_emphasis: int = 55

    @property
    def pause_marker_threshold_seconds(self) -> float:
        return max(0.5, self.pause_marker_threshold_tenths / 10.0)


@dataclass(frozen=True)
class TranscriptStats:
    word_count: int
    segment_count: int
    duration_seconds: float
    average_confidence: float | None
    low_confidence_words: int
    median_word_duration: float
    median_pause: float
    long_pause_count: int
    words_per_minute: float
    speech_density: float

    def to_lines(self) -> list[str]:
        confidence = "n/a"
        if self.average_confidence is not None:
            confidence = f"{self.average_confidence * 100:.1f}%"

        return [
            f"Words: {self.word_count}",
            f"Segments: {self.segment_count}",
            f"Duration: {format_seconds(self.duration_seconds)}",
            f"Average confidence: {confidence}",
            f"Low-confidence words: {self.low_confidence_words}",
            f"Median word duration: {self.median_word_duration:.2f}s",
            f"Median pause: {self.median_pause:.2f}s",
            f"Long pauses: {self.long_pause_count}",
            f"Words per minute: {self.words_per_minute:.1f}",
            f"Speech density: {self.speech_density * 100:.1f}%",
            f"Inferred pace: {infer_pace_label(self.words_per_minute)}",
            f"Inferred confidence: {infer_confidence_label(self.average_confidence)}",
            f"Inferred pause style: {infer_pause_label(self.median_pause, self.long_pause_count)}",
        ]


def format_seconds(value: float) -> str:
    if value < 60:
        return f"{value:.1f}s"

    minutes = int(value // 60)
    seconds = value % 60
    return f"{minutes}m {seconds:.1f}s"


def infer_pace_label(words_per_minute: float) -> str:
    if words_per_minute <= 0:
        return "unknown"
    if words_per_minute < 95:
        return "slow"
    if words_per_minute < 145:
        return "measured"
    if words_per_minute < 185:
        return "quick"
    return "very quick"


def infer_confidence_label(average_confidence: float | None) -> str:
    if average_confidence is None:
        return "unknown"
    if average_confidence >= 0.9:
        return "strong"
    if average_confidence >= 0.78:
        return "usable"
    if average_confidence >= 0.65:
        return "mixed"
    return "fragile"


def infer_pause_label(median_pause: float, long_pause_count: int) -> str:
    if median_pause < 0.12 and long_pause_count <= 2:
        return "tight"
    if median_pause < 0.28 and long_pause_count <= 5:
        return "steady"
    if median_pause < 0.55:
        return "reflective"
    return "fragmented"


def safe_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def median_or_default(values: list[float], default: float) -> float:
    numeric = [value for value in values if isinstance(value, (int, float)) and math.isfinite(value)]
    if not numeric:
        return default
    return float(statistics.median(numeric))


def flatten_timing_words(timings_payload: dict[str, Any] | None) -> list[TimingWord]:
    if not isinstance(timings_payload, dict):
        return []

    words: list[TimingWord] = []
    segments = timings_payload.get("segments") or []
    index = 0
    for segment_index, segment in enumerate(segments):
        if not isinstance(segment, dict):
            continue
        segment_words = segment.get("words") or []
        for word in segment_words:
            if not isinstance(word, dict):
                continue
            text = str(word.get("word", "")).strip()
            start = safe_float(word.get("start"))
            end = safe_float(word.get("end"))
            if not text or start is None or end is None:
                continue
            words.append(
                TimingWord(
                    index=index,
                    text=text,
                    start=start,
                    end=end,
                    probability=safe_float(word.get("probability")),
                    segment_index=segment_index,
                )
            )
            index += 1
    return words


def estimate_duration_seconds(words: list[TimingWord], timings_payload: dict[str, Any] | None) -> float:
    if isinstance(timings_payload, dict):
        payload_duration = safe_float(timings_payload.get("duration"))
        if payload_duration is not None and payload_duration > 0:
            return payload_duration
    if not words:
        return 0.0
    return max(0.0, words[-1].end - words[0].start)


def get_pause_after(words: list[TimingWord], index: int) -> float:
    if index < 0 or index >= len(words) - 1:
        return 0.0
    return max(0.0, words[index + 1].start - words[index].end)


def get_local_gap_baseline(words: list[TimingWord], index: int, radius: int = 10) -> float:
    if not words:
        return 0.12

    start = max(0, index - radius)
    end = min(len(words) - 1, index + radius)
    gaps = [
        get_pause_after(words, position)
        for position in range(start, end)
        if get_pause_after(words, position) > 0.02
    ]
    durations = [word.duration for word in words[start : end + 1] if word.duration > 0.02]
    duration_baseline = median_or_default(durations, 0.22)
    gap_baseline = median_or_default(gaps, max(0.08, duration_baseline * 0.45))
    return max(0.08, gap_baseline)


def classify_pause(words: list[TimingWord], index: int, settings: RenderSettings) -> tuple[str, float]:
    gap = get_pause_after(words, index)
    if gap <= 0.02:
        return "inline", gap

    local_gap_baseline = get_local_gap_baseline(words, index)
    sensitivity_multiplier = max(0.35, settings.pause_sensitivity / 100.0)
    score = (gap / local_gap_baseline) * sensitivity_multiplier

    if gap >= settings.pause_marker_threshold_seconds or score >= 8.5:
        return "marker", gap
    if score >= 5.0:
        return "paragraph", gap
    if score >= 2.3:
        return "line", gap
    if score >= 1.4:
        return "soft", gap
    return "inline", gap


def build_confidence_style(word: TimingWord, settings: RenderSettings, is_selected: bool) -> str:
    styles = [
        "text-decoration:none",
        "padding:1px 2px",
        "border-radius:3px",
        "color:#14304a",
    ]

    if is_selected:
        styles.extend(
            [
                "background-color:#ffe39d",
                "color:#111111",
                "font-weight:600",
            ]
        )

    probability = word.probability
    threshold = settings.low_confidence_threshold / 100.0
    emphasis = max(0.0, min(1.0, settings.confidence_emphasis / 100.0))
    if probability is None:
        styles.append("border-bottom:1px dotted #9aa7b5")
        return ";".join(styles)

    if probability < threshold:
        severity = min(1.0, max(0.0, (threshold - probability) / max(0.05, threshold)))
        alpha = 0.18 + (0.35 * severity * emphasis)
        styles.extend(
            [
                "color:#8a4b08",
                f"background-color:rgba(255, 221, 163, {alpha:.3f})",
                f"border-bottom:1px solid rgba(194, 109, 17, {0.45 + (0.35 * severity):.3f})",
            ]
        )
    elif probability < min(0.98, threshold + 0.12):
        styles.append("color:#40566d")

    return ";".join(styles)


def build_timed_reading_html(
    words: list[TimingWord],
    settings: RenderSettings,
    *,
    selected_index: int | None = None,
    fallback_text: str = "",
) -> str:
    if not words:
        escaped = html.escape(fallback_text.strip() or "No word-level timing data loaded.")
        return (
            "<html><body style='font-family:Segoe UI;font-size:15px;line-height:1.65;"
            "color:#1f2933;background-color:#fbfbfc;padding:18px;'>"
            f"<p>{escaped}</p></body></html>"
        )

    parts = [
        "<html><body style='font-family:Segoe UI;font-size:15px;line-height:1.7;"
        "color:#1f2933;background-color:#fbfbfc;padding:18px;'>"
    ]

    for index, word in enumerate(words):
        style = build_confidence_style(word, settings, is_selected=(selected_index == index))
        label = html.escape(word.text)
        parts.append(f"<a href='word:{index}' style='{style}'>{label}</a>")

        if index >= len(words) - 1:
            continue

        pause_kind, gap = classify_pause(words, index, settings)
        if pause_kind == "inline":
            parts.append(" ")
        elif pause_kind == "soft":
            parts.append("<span style='color:#9aa7b5;'>  /  </span>")
        elif pause_kind == "line":
            parts.append("<br/>")
        elif pause_kind == "paragraph":
            parts.append("<br/><br/>")
        else:
            pause_label = html.escape(f"[{gap:.1f}s pause]")
            parts.append(
                "<br/><span style='color:#6b7280;font-size:12px;background-color:#eef2f7;"
                "padding:2px 6px;border-radius:10px;'>"
                f"{pause_label}</span><br/>"
            )

    parts.append("</body></html>")
    return "".join(parts)


def compute_stats(words: list[TimingWord], timings_payload: dict[str, Any] | None, settings: RenderSettings) -> TranscriptStats:
    segment_count = 0
    if isinstance(timings_payload, dict):
        segment_count = len([segment for segment in (timings_payload.get("segments") or []) if isinstance(segment, dict)])

    duration_seconds = estimate_duration_seconds(words, timings_payload)
    positive_probabilities = [word.probability for word in words if word.probability is not None]
    average_confidence = None
    if positive_probabilities:
        average_confidence = sum(positive_probabilities) / len(positive_probabilities)

    low_conf_threshold = settings.low_confidence_threshold / 100.0
    low_confidence_words = len([word for word in words if word.probability is not None and word.probability < low_conf_threshold])
    word_durations = [word.duration for word in words if word.duration > 0]
    pauses = [get_pause_after(words, index) for index in range(len(words) - 1)]
    positive_pauses = [pause for pause in pauses if pause > 0.02]
    median_word_duration = median_or_default(word_durations, 0.0)
    median_pause = median_or_default(positive_pauses, 0.0)
    long_pause_count = len([pause for pause in positive_pauses if pause >= settings.pause_marker_threshold_seconds])
    words_per_minute = 0.0
    if duration_seconds > 0:
        words_per_minute = len(words) / (duration_seconds / 60.0)
    speech_density = 0.0
    if duration_seconds > 0:
        speech_density = min(1.0, sum(word_durations) / duration_seconds)

    return TranscriptStats(
        word_count=len(words),
        segment_count=segment_count,
        duration_seconds=duration_seconds,
        average_confidence=average_confidence,
        low_confidence_words=low_confidence_words,
        median_word_duration=median_word_duration,
        median_pause=median_pause,
        long_pause_count=long_pause_count,
        words_per_minute=words_per_minute,
        speech_density=speech_density,
    )
