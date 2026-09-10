from __future__ import annotations

import json
from pathlib import Path

from PySide6.QtCore import QSignalBlocker, Qt, QUrl
from PySide6.QtGui import QAction
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer
from PySide6.QtWidgets import (
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSlider,
    QSplitter,
    QTabWidget,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from .analysis import (
    RenderSettings,
    TimingWord,
    build_timed_reading_html,
    compute_stats,
    flatten_timing_words,
    format_seconds,
)
from .session import TranscriptStudioSession, load_session_file


class TranscriptStudioWindow(QMainWindow):
    def __init__(self, session: TranscriptStudioSession) -> None:
        super().__init__()
        self._session = session
        self._words: list[TimingWord] = []
        self._settings = RenderSettings()
        self._selected_word_index: int | None = None

        self._player = QMediaPlayer(self)
        self._audio_output = QAudioOutput(self)
        self._player.setAudioOutput(self._audio_output)
        self._player.positionChanged.connect(self._on_player_position_changed)
        self._player.durationChanged.connect(self._on_player_duration_changed)

        self._build_ui()
        self._apply_session(session)

    def _build_ui(self) -> None:
        self.setWindowTitle("Transcript Studio")
        self.resize(1540, 960)

        central = QWidget(self)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(8, 8, 8, 8)

        splitter = QSplitter(Qt.Orientation.Horizontal, central)
        layout.addWidget(splitter)
        self.setCentralWidget(central)

        sidebar = QWidget(splitter)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(8)

        self._session_summary = QPlainTextEdit()
        self._session_summary.setReadOnly(True)
        self._session_summary.setMinimumHeight(150)
        session_group = QGroupBox("Session")
        session_layout = QVBoxLayout(session_group)
        session_layout.addWidget(self._session_summary)
        sidebar_layout.addWidget(session_group)

        controls_group = QGroupBox("Rendering Controls")
        controls_layout = QFormLayout(controls_group)
        self._pause_sensitivity_slider, self._pause_sensitivity_value = self._create_slider(
            minimum=40,
            maximum=180,
            value=self._settings.pause_sensitivity,
            suffix="%",
            callback=self._on_settings_changed,
        )
        self._pause_marker_slider, self._pause_marker_value = self._create_slider(
            minimum=8,
            maximum=40,
            value=self._settings.pause_marker_threshold_tenths,
            suffix="s",
            callback=self._on_settings_changed,
            value_formatter=lambda value: f"{value / 10.0:.1f}s",
        )
        self._low_confidence_slider, self._low_confidence_value = self._create_slider(
            minimum=50,
            maximum=95,
            value=self._settings.low_confidence_threshold,
            suffix="%",
            callback=self._on_settings_changed,
        )
        self._confidence_emphasis_slider, self._confidence_emphasis_value = self._create_slider(
            minimum=0,
            maximum=100,
            value=self._settings.confidence_emphasis,
            suffix="%",
            callback=self._on_settings_changed,
        )
        controls_layout.addRow("Pause sensitivity", self._wrap_slider(self._pause_sensitivity_slider, self._pause_sensitivity_value))
        controls_layout.addRow("Pause marker", self._wrap_slider(self._pause_marker_slider, self._pause_marker_value))
        controls_layout.addRow("Low-confidence threshold", self._wrap_slider(self._low_confidence_slider, self._low_confidence_value))
        controls_layout.addRow("Confidence emphasis", self._wrap_slider(self._confidence_emphasis_slider, self._confidence_emphasis_value))
        sidebar_layout.addWidget(controls_group)

        audio_group = QGroupBox("Audio Alignment")
        audio_layout = QVBoxLayout(audio_group)
        audio_buttons = QHBoxLayout()
        self._play_pause_button = QPushButton("Play")
        self._stop_button = QPushButton("Stop")
        self._audio_status = QLabel("No audio loaded.")
        self._audio_status.setWordWrap(True)
        self._play_pause_button.clicked.connect(self._toggle_playback)
        self._stop_button.clicked.connect(self._stop_playback)
        audio_buttons.addWidget(self._play_pause_button)
        audio_buttons.addWidget(self._stop_button)
        audio_buttons.addStretch(1)
        audio_layout.addLayout(audio_buttons)
        self._position_slider = QSlider(Qt.Orientation.Horizontal)
        self._position_slider.setRange(0, 0)
        self._position_slider.sliderReleased.connect(self._seek_to_slider_position)
        audio_layout.addWidget(self._position_slider)
        self._position_label = QLabel("0.0s / 0.0s")
        audio_layout.addWidget(self._position_label)
        audio_layout.addWidget(self._audio_status)
        sidebar_layout.addWidget(audio_group)

        self._stats_box = QPlainTextEdit()
        self._stats_box.setReadOnly(True)
        self._stats_box.setMinimumHeight(220)
        stats_group = QGroupBox("Transcript Stats")
        stats_layout = QVBoxLayout(stats_group)
        stats_layout.addWidget(self._stats_box)
        sidebar_layout.addWidget(stats_group, 1)

        workspace_tabs = QTabWidget(splitter)
        self._timed_reading_browser = QTextBrowser()
        self._timed_reading_browser.setOpenLinks(False)
        self._timed_reading_browser.setOpenExternalLinks(False)
        self._timed_reading_browser.anchorClicked.connect(self._on_word_anchor_clicked)
        workspace_tabs.addTab(self._timed_reading_browser, "Timed Reading")

        self._raw_transcript_box = QPlainTextEdit()
        self._raw_transcript_box.setReadOnly(True)
        workspace_tabs.addTab(self._raw_transcript_box, "Raw Transcript")

        self._timings_box = QPlainTextEdit()
        self._timings_box.setReadOnly(True)
        workspace_tabs.addTab(self._timings_box, "Timing JSON")

        self._metadata_box = QPlainTextEdit()
        self._metadata_box.setReadOnly(True)
        workspace_tabs.addTab(self._metadata_box, "Session Metadata")

        splitter.setSizes([390, 1100])
        self.statusBar().showMessage("Ready.")
        self._build_menu()

    def _build_menu(self) -> None:
        file_menu = self.menuBar().addMenu("&File")

        open_action = QAction("Open Session...", self)
        open_action.triggered.connect(self._open_session_dialog)
        file_menu.addAction(open_action)

        reload_action = QAction("Reload Session", self)
        reload_action.triggered.connect(self._reload_session)
        file_menu.addAction(reload_action)

        file_menu.addSeparator()

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

    def _create_slider(
        self,
        *,
        minimum: int,
        maximum: int,
        value: int,
        suffix: str,
        callback,
        value_formatter=None,
    ) -> tuple[QSlider, QLabel]:
        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(minimum, maximum)
        slider.setValue(value)
        label = QLabel("")
        label.setMinimumWidth(56)

        def update_label(current_value: int) -> None:
            if callable(value_formatter):
                label.setText(value_formatter(current_value))
            else:
                label.setText(f"{current_value}{suffix}")

        slider.valueChanged.connect(update_label)
        slider.valueChanged.connect(lambda _value: callback())
        update_label(value)
        return slider, label

    def _wrap_slider(self, slider: QSlider, label: QLabel) -> QWidget:
        wrapper = QWidget()
        layout = QHBoxLayout(wrapper)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(slider, 1)
        layout.addWidget(label)
        return wrapper

    def _apply_session(self, session: TranscriptStudioSession) -> None:
        self._session = session
        self._selected_word_index = None
        self._words = flatten_timing_words(session.timings_payload)

        transcript_text = session.transcript_text.strip()
        if not transcript_text and isinstance(session.timings_payload, dict):
            transcript_text = str(session.timings_payload.get("text", "")).strip()

        self._raw_transcript_box.setPlainText(transcript_text or "No transcript text loaded.")
        timings_text = "No timing JSON loaded."
        if session.timings_payload:
            timings_text = json.dumps(session.timings_payload, indent=2, ensure_ascii=False)
        self._timings_box.setPlainText(timings_text)
        self._metadata_box.setPlainText(json.dumps(self._build_metadata_payload(), indent=2, ensure_ascii=False))
        self._session_summary.setPlainText(self._build_session_summary_text())

        self._configure_audio_player()
        self._refresh_views()

        title = session.title.strip() or Path(session.audio_path).name or "Transcript Studio"
        self.setWindowTitle(f"Transcript Studio - {title}")
        self.statusBar().showMessage("Session loaded.")

    def _build_metadata_payload(self) -> dict:
        return {
            "sessionPath": self._session.session_path,
            "title": self._session.title,
            "description": self._session.description,
            "sourceKind": self._session.source_kind,
            "sourceId": self._session.source_id,
            "sourceUrl": self._session.source_url,
            "subtitle": self._session.subtitle,
            "audioPath": self._session.audio_path,
            "transcriptPath": self._session.transcript_path,
            "timingsPath": self._session.timings_path,
            "metadata": self._session.metadata,
        }

    def _build_session_summary_text(self) -> str:
        lines = [
            f"Title: {self._session.title or '(untitled)'}",
            f"Source kind: {self._session.source_kind or 'unknown'}",
            f"Source id: {self._session.source_id or 'n/a'}",
            f"Source url: {self._session.source_url or 'n/a'}",
            f"Audio: {self._session.audio_path or 'n/a'}",
            f"Transcript: {self._session.transcript_path or 'inline/session text'}",
            f"Timings: {self._session.timings_path or 'inline/session timings'}",
        ]
        if self._session.subtitle:
            lines.append(f"Subtitle: {self._session.subtitle}")
        if self._session.description:
            lines.extend(["", self._session.description.strip()])
        return "\n".join(lines)

    def _refresh_views(self) -> None:
        transcript_text = self._session.transcript_text.strip()
        if not transcript_text and isinstance(self._session.timings_payload, dict):
            transcript_text = str(self._session.timings_payload.get("text", "")).strip()

        html_text = build_timed_reading_html(
            self._words,
            self._settings,
            selected_index=self._selected_word_index,
            fallback_text=transcript_text,
        )
        self._timed_reading_browser.setHtml(html_text)
        stats = compute_stats(self._words, self._session.timings_payload, self._settings)
        self._stats_box.setPlainText("\n".join(stats.to_lines()))

    def _configure_audio_player(self) -> None:
        audio_path = self._session.audio_path
        has_audio = bool(audio_path and Path(audio_path).is_file())
        self._play_pause_button.setEnabled(has_audio)
        self._stop_button.setEnabled(has_audio)
        self._position_slider.setEnabled(has_audio)

        if has_audio:
            self._player.setSource(QUrl.fromLocalFile(audio_path))
            self._audio_status.setText("Click a word in Timed Reading to seek the audio.")
        else:
            self._player.setSource(QUrl())
            self._position_slider.setRange(0, 0)
            self._position_label.setText("0.0s / 0.0s")
            self._audio_status.setText("No audio file was provided in this session.")
            self._play_pause_button.setText("Play")

    def _on_settings_changed(self) -> None:
        self._settings.pause_sensitivity = self._pause_sensitivity_slider.value()
        self._settings.pause_marker_threshold_tenths = self._pause_marker_slider.value()
        self._settings.low_confidence_threshold = self._low_confidence_slider.value()
        self._settings.confidence_emphasis = self._confidence_emphasis_slider.value()
        self._refresh_views()

    def _on_word_anchor_clicked(self, url: QUrl) -> None:
        href = url.toString()
        if not href.startswith("word:"):
            return
        try:
            word_index = int(href.split(":", 1)[1])
        except ValueError:
            return
        if word_index < 0 or word_index >= len(self._words):
            return

        self._selected_word_index = word_index
        self._refresh_views()
        self._seek_to_word(word_index)

    def _seek_to_word(self, word_index: int) -> None:
        if word_index < 0 or word_index >= len(self._words):
            return
        word = self._words[word_index]
        position_ms = max(0, int(word.start * 1000))
        if self._player.source().isEmpty():
            self.statusBar().showMessage(f"Selected word {word_index + 1} at {word.start:.2f}s.")
            return

        self._player.setPosition(position_ms)
        self.statusBar().showMessage(
            f"Jumped to word {word_index + 1} at {word.start:.2f}s: {word.text}"
        )

    def _toggle_playback(self) -> None:
        if self._player.source().isEmpty():
            return
        if self._player.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
            self._player.pause()
            self._play_pause_button.setText("Play")
        else:
            self._player.play()
            self._play_pause_button.setText("Pause")

    def _stop_playback(self) -> None:
        if self._player.source().isEmpty():
            return
        self._player.stop()
        self._play_pause_button.setText("Play")
        self.statusBar().showMessage("Audio stopped.")

    def _seek_to_slider_position(self) -> None:
        if self._player.source().isEmpty():
            return
        self._player.setPosition(int(self._position_slider.value()))

    def _on_player_position_changed(self, position_ms: int) -> None:
        with QSignalBlocker(self._position_slider):
            self._position_slider.setValue(position_ms)
        duration_seconds = max(0.0, self._player.duration() / 1000.0)
        position_seconds = max(0.0, position_ms / 1000.0)
        self._position_label.setText(
            f"{format_seconds(position_seconds)} / {format_seconds(duration_seconds)}"
        )
        if self._player.playbackState() != QMediaPlayer.PlaybackState.PlayingState:
            self._play_pause_button.setText("Play")

    def _on_player_duration_changed(self, duration_ms: int) -> None:
        duration_ms = max(0, duration_ms)
        with QSignalBlocker(self._position_slider):
            self._position_slider.setRange(0, duration_ms)
        duration_seconds = max(0.0, duration_ms / 1000.0)
        position_seconds = max(0.0, self._player.position() / 1000.0)
        self._position_label.setText(
            f"{format_seconds(position_seconds)} / {format_seconds(duration_seconds)}"
        )

    def _open_session_dialog(self) -> None:
        start_dir = str(Path(self._session.session_path).parent) if self._session.session_path else ""
        session_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Transcript Session",
            start_dir,
            "Session Files (*.json);;All Files (*.*)",
        )
        if not session_path:
            return

        try:
            session = load_session_file(session_path)
        except Exception as exc:
            QMessageBox.critical(self, "Open Session Failed", str(exc))
            return

        self._apply_session(session)

    def _reload_session(self) -> None:
        if not self._session.session_path:
            QMessageBox.information(self, "Reload Session", "This window was not opened from a session file.")
            return
        try:
            session = load_session_file(self._session.session_path)
        except Exception as exc:
            QMessageBox.critical(self, "Reload Failed", str(exc))
            return

        self._apply_session(session)
