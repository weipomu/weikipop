# src/gui/popup.py
import base64
import json
import logging
import os
import threading
import time
from datetime import datetime, timezone
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

from PyQt6.QtCore import QTimer, QPoint, QSize, Qt, pyqtSignal, QEvent
from PyQt6.QtGui import QColor, QCursor, QFont, QFontMetrics, QFontInfo, QTextDocument
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel, QFrame, QApplication, QScrollArea

from src.config.config import config, IS_MACOS
from src.dictionary.lookup import DictionaryEntry, KanjiEntry
from src.dictionary.anki_client import AnkiClient
from src.gui.magpie_manager import magpie_manager
import re as _re  # hoisted — used in Anki duplicate/presence checks
from src.utils.window_info import get_active_window_title

if IS_MACOS:
    try:
        import Quartz
    except ImportError:
        Quartz = None

logger = logging.getLogger(__name__)

MINE_BAR_HEIGHT = 30  # fixed pixel height reserved for the mine status bar


class Popup(QWidget):
    # Signals are always delivered on the main thread (AutoConnection) — safe to emit from threads
    anki_presence_updated  = pyqtSignal(str, bool)  # (word, is_present) — word prevents race condition
    status_message_signal  = pyqtSignal(str)    # show a brief status message

    def __init__(self, shared_state, input_loop):
        super().__init__()
        self._latest_data    = None
        self._latest_context: Optional[Dict[str, Any]] = None
        self._last_latest_data    = None
        self._last_latest_context = None
        self._data_lock = threading.Lock()
        self._previous_active_window_on_mac = None

        self.anki_shortcut_was_pressed = False
        self.copy_shortcut_was_pressed = False
        self._anki_presence_status = None   # None=unknown, True=in Anki, False=new
        self._last_presence_word   = None   # reset on hide → always re-check on new show
        self._last_mouse_pos       = None   # throttle move_to calls
        self._last_html            = None   # skip redundant setText calls
        self._last_size            = None   # skip redundant resize calls
        self._cached_popup_size    = None   # cleared by reapply_settings
        # How many more 16ms move-timer ticks to keep forcing scroll to top.
        # Set whenever content changes; counts down to 0 so Qt layout settling
        # can never leave blank space at the top of the popup.
        self._scroll_reset_frames  = 0

        # Lazy rendering — only the first batch of entry groups is rendered on
        # initial display; remaining groups load as the user scrolls down.
        self._lazy_pending_groups   = []    # groups not yet rendered
        self._lazy_rendered_parts   = []    # accumulated HTML chunks
        self._lazy_next_group_index = 0     # absolute group index for <hr> placement
        self._dismissed_by_click   = False

        self.shared_state = shared_state
        self.input_loop   = input_loop

        self.is_visible = False
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.process_latest_data_loop)
        self.timer.start(60)

        # Separate fast timer — only for cursor tracking and show/hide.
        # Keeps popup movement at ~60fps independently of content logic.
        self._move_timer = QTimer(self)
        self._move_timer.timeout.connect(self._move_timer_tick)
        self._move_timer.start(16)

        # Off-screen probe label for height calculations (never shown)
        self.probe_label = QLabel()
        self.probe_label.setWordWrap(True)
        self.probe_label.setTextFormat(Qt.TextFormat.RichText)
        self.probe_label.hide()  # must be hidden — it has no parent so Qt would show it as a top-level window

        self.is_calibrated         = False
        self.header_chars_per_line = 50
        self.def_chars_per_line    = 50

        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)  # never steal focus
        self.setStyleSheet("background: transparent;")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.frame = QFrame()
        self._apply_frame_stylesheet()
        main_layout.addWidget(self.frame)

        self.content_layout = QVBoxLayout(self.frame)
        self.content_layout.setContentsMargins(10, 10, 10, 10)
        self.content_layout.setSpacing(4)

        # Main dictionary content (scrollable)
        self.display_label = QLabel()
        self.display_label.setWordWrap(True)
        self.display_label.setTextFormat(Qt.TextFormat.RichText)
        self.display_label.setTextInteractionFlags(Qt.TextInteractionFlag.LinksAccessibleByMouse)
        # AlignTop: pin text to top of the label so it never floats to the
        # vertical centre when the label is stretched to fill the popup height.
        self.display_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.content_scroll = QScrollArea()
        self.content_scroll.setWidgetResizable(True)
        self.content_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.content_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.content_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        # AlignTop: content always starts at the top of the viewport instead of
        # being vertically centred when it is shorter than the popup height.
        self.content_scroll.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.content_scroll.setWidget(self.display_label)
        self.content_layout.addWidget(self.content_scroll)
        self.content_scroll.verticalScrollBar().valueChanged.connect(self._on_scroll_lazy_load)

        # Brief status message (e.g. "Mined!" / "Already in Anki") — hidden by default
        self.status_label = QLabel()
        self.status_label.setWordWrap(False)
        self.status_label.setTextFormat(Qt.TextFormat.PlainText)
        self.status_label.setStyleSheet("color: #f0c674; font-size: 11px;")
        self.status_label.hide()
        self.content_layout.addWidget(self.status_label)

        # Mine bar — always shown when popup is visible
        # Green ⊕ = new word (click to mine), Grey ✓ = already mined
        self.mine_bar = QLabel()
        self.mine_bar.setTextFormat(Qt.TextFormat.RichText)
        self.mine_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.mine_bar.setFixedHeight(MINE_BAR_HEIGHT)
        self.mine_bar.linkActivated.connect(self._on_mine_clicked)
        self.mine_bar.setStyleSheet("margin-top: 2px;")
        self._set_mine_bar_new()          # default: green ⊕ Mine
        self.content_layout.addWidget(self.mine_bar)

        # Connect signals
        self.anki_presence_updated.connect(self._on_anki_presence_updated)
        self.status_message_signal.connect(self._show_status_message)

        app = QApplication.instance()
        if app:
            app.installEventFilter(self)

        self.hide()

    # ------------------------------------------------------------------ #
    #  Mine bar helpers                                                     #
    # ------------------------------------------------------------------ #

    def _set_mine_bar_new(self):
        """Green ⊕ — word not yet in Anki. Click to mine."""
        fs = max(13, config.font_size_definitions)
        self.mine_bar.setText(
            f'<div style="text-align:center;width:100%;">'
            f'<a href="mine" style="color:#4CAF50;text-decoration:none;font-size:{fs}px;"'
            f' title="Click or press Alt+A to mine">⊕ Mine</a>'
            f'</div>'
        )

    def _set_mine_bar_mined(self):
        """Grey ✓ — word already in Anki."""
        fs = max(13, config.font_size_definitions)
        self.mine_bar.setText(
            f'<div style="text-align:center;width:100%;">'
            f'<span style="color:#888888;font-size:{fs}px;">✓ Already mined</span>'
            f'</div>'
        )

    def _on_mine_clicked(self, _url):
        self.add_to_anki()

    def _on_anki_presence_updated(self, word: str, is_present: bool):
        # Only update the bar if this result is still for the current word
        # (race condition guard: async result may arrive after cursor moved)
        if word == self._last_presence_word:
            self._anki_presence_status = is_present
            if is_present:
                self._set_mine_bar_mined()
            else:
                self._set_mine_bar_new()

    # ------------------------------------------------------------------ #
    #  Frame stylesheet                                                     #
    # ------------------------------------------------------------------ #

    def _apply_frame_stylesheet(self):
        bg_color = QColor(config.color_background)
        r, g, b = bg_color.red(), bg_color.green(), bg_color.blue()
        a = config.background_opacity
        self.probe_label.setFont(QFont(config.font_family))
        self.frame.setStyleSheet(f"""
            QFrame {{
                background-color: rgba({r}, {g}, {b}, {a});
                color: {config.color_foreground};
                border-radius: 8px;
                border: 1px solid #555;
            }}
            QLabel {{
                background-color: transparent;
                border: none;
                font-family: "{config.font_family}";
            }}
            hr {{
                border: none;
                height: 1px;
            }}
        """)

    # ------------------------------------------------------------------ #
    #  Font calibration                                                     #
    # ------------------------------------------------------------------ #

    def _calibrate_empirically(self):
        logger.debug("--- Calibrating Font Metrics Empirically (One-Time) ---")
        actual_font = self.display_label.font()
        font_info   = QFontInfo(actual_font)
        logger.debug(f"[FONT] Resolved font: '{font_info.family()}' {font_info.pointSize()}pt")

        margins = self.content_layout.contentsMargins()
        border_width = 1
        horizontal_padding = margins.left() + margins.right() + (border_width * 2)

        screen = QApplication.primaryScreen()
        if screen is None:
            logger.warning("No primary screen found; skipping calibration.")
            return
        self.max_content_width = (int(screen.geometry().width() * 0.4)) - horizontal_padding

        header_font = QFont(config.font_family)
        header_font.setPixelSize(config.font_size_header)
        self.header_chars_per_line = self._find_chars_for_width(QFontMetrics(header_font))

        def_font = QFont(config.font_family)
        def_font.setPixelSize(config.font_size_definitions)
        self.def_chars_per_line = self._find_chars_for_width(QFontMetrics(def_font))

        logger.debug(f"[CALIBRATE] max_content_width={self.max_content_width}px  "
                     f"header={self.header_chars_per_line}ch  def={self.def_chars_per_line}ch")
        self.is_calibrated = True

    def _find_chars_for_width(self, metrics: QFontMetrics) -> int:
        low, high, best = 1, 500, 1
        while low <= high:
            mid = (low + high) // 2
            if metrics.horizontalAdvance('x' * mid) <= self.max_content_width:
                best = mid
                low  = mid + 1
            else:
                high = mid - 1
        return max(best, 1)

    # ------------------------------------------------------------------ #
    #  Data access                                                          #
    # ------------------------------------------------------------------ #

    def set_latest_data(self, data, context: Optional[Dict[str, Any]] = None):
        if context is None:
            context = {}
        if "document_title" not in context:
            try:
                context["document_title"] = get_active_window_title()
            except Exception:
                context["document_title"] = ""
        with self._data_lock:
            self._latest_data   = data
            self._latest_context = context

    def get_latest_data(self) -> Tuple[Any, Optional[Dict[str, Any]]]:
        with self._data_lock:
            return self._latest_data, self._latest_context

    # ------------------------------------------------------------------ #
    #  Main logic loop (runs every 60 ms — content render, presence, shortcuts) #
    # ------------------------------------------------------------------ #

    def process_latest_data_loop(self):
        if not self.is_calibrated:
            self._calibrate_empirically()
            if not self.is_calibrated:
                return  # wait until calibration succeeds

        latest_data, latest_context = self.get_latest_data()

        # Re-render only when content actually changes
        if latest_data and (latest_data    != self._last_latest_data or
                            latest_context != self._last_latest_context):
            full_html = self._calculate_content(latest_data)
            if full_html is not None:
                if full_html != self._last_html:
                    self.display_label.setText(full_html)
                    self._last_html = full_html
                    self.content_scroll.verticalScrollBar().setValue(0)
                    self._scroll_reset_frames = 8

                # Fixed consistent size — same width and height regardless of
                # how many definitions exist, so the popup never jumps around.
                new_size = self._fixed_popup_size()
                if new_size != self._last_size:
                    self.setFixedSize(new_size)
                    self._last_size = new_size
                    if self.is_visible:
                        mp = QCursor.pos()
                        self.move_to(mp.x(), mp.y())

        self._last_latest_data    = latest_data
        self._last_latest_context = latest_context

        # Hotkey state — used for both presence check and shortcuts
        _kp = getattr(self.input_loop, 'hotkey_is_pressed', False)
        _as = config.auto_scan_mode and config.auto_scan_mode_lookups_without_hotkey
        hotkey_down = latest_data and (_kp or _as)

        # Presence check — only while popup is on screen
        if self.is_visible and latest_data:
            if not isinstance(latest_data[0], KanjiEntry):
                word = (getattr(latest_data[0], "written_form", "") or
                        getattr(latest_data[0], "reading", "") or "")
                if word and word != self._last_presence_word:
                    self._last_presence_word = word
                    self._set_mine_bar_new()
                    self._check_anki_presence_async(word)

        # Shortcuts — check whenever hotkey is held, even if lock was busy
        # (mining must work even if show_popup() couldn't acquire screen_lock)
        if hotkey_down:
            anki_key = getattr(config, "add_to_anki", "Alt+A")
            anki_pressed = self.input_loop.is_key_pressed(anki_key)
            if anki_pressed and not self.anki_shortcut_was_pressed:
                self.add_to_anki()
            self.anki_shortcut_was_pressed = anki_pressed

            copy_key = getattr(config, "copy_text", "Alt+C")
            copy_pressed = self.input_loop.is_key_pressed(copy_key)
            if copy_pressed and not self.copy_shortcut_was_pressed:
                self.copy_to_clipboard()
            self.copy_shortcut_was_pressed = copy_pressed

            scroll_shortcut = getattr(config, 'scroll_popup', 'Alt+Wheel')
            if self.is_visible and self._is_scroll_shortcut_active(scroll_shortcut):
                delta = self.input_loop.get_and_reset_scroll_delta()
                if delta:
                    scrollbar = self.content_scroll.verticalScrollBar()
                    scrollbar.setValue(scrollbar.value() - (delta * 42))
            else:
                # Do not carry stale wheel deltas between frames.
                self.input_loop.get_and_reset_scroll_delta()
        else:
            self.anki_shortcut_was_pressed = False
            self.copy_shortcut_was_pressed = False
            self.input_loop.get_and_reset_scroll_delta()

    def _is_scroll_shortcut_active(self, shortcut: str) -> bool:
        shortcut = (shortcut or '').strip()
        if not shortcut:
            return False

        lower = shortcut.lower()
        if lower.endswith('+wheel'):
            key_part = shortcut[:-6].strip()
            if not key_part:
                return True

            if key_part.lower() == config.hotkey.lower():
                return getattr(self.input_loop, 'hotkey_is_pressed', False)
            return self.input_loop.is_key_pressed(key_part)

        return self.input_loop.is_key_pressed(shortcut)

    # ------------------------------------------------------------------ #
    #  Popup actions                                                        #
    # ------------------------------------------------------------------ #

    def _move_timer_tick(self):
        """Runs every 16 ms — handles show/hide and smooth cursor tracking.
        Kept cheap: only reads hotkey state and moves the window."""
        _kp = getattr(self.input_loop, 'hotkey_is_pressed', False)
        _as = config.auto_scan_mode and config.auto_scan_mode_lookups_without_hotkey
        has_data = self._latest_data is not None
        should_show = has_data and (_kp or _as)

        if not should_show:
            self._dismissed_by_click = False

        if should_show and not self._dismissed_by_click:
            self.show_popup()
            if self.is_visible:
                # Keep forcing scroll to top for a few frames after content changes
                # so Qt's layout engine can never leave blank space at the top.
                if self._scroll_reset_frames > 0:
                    self.content_scroll.verticalScrollBar().setValue(0)
                    self._scroll_reset_frames -= 1

                mouse_pos = QCursor.pos()
                mp = (mouse_pos.x(), mouse_pos.y())
                lp = self._last_mouse_pos
                if lp is None or abs(mp[0] - lp[0]) > 1 or abs(mp[1] - lp[1]) > 1:
                    self._last_mouse_pos = mp
                    self.move_to(mp[0], mp[1])
        else:
            self._scroll_reset_frames = 0
            self.hide_popup()

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.MouseButtonPress and self.is_visible:
            global_pos = event.globalPosition().toPoint() if hasattr(event, 'globalPosition') else QCursor.pos()
            if not self.frameGeometry().contains(global_pos):
                self._dismissed_by_click = True
                self.set_latest_data(None, {})
                self.hide_popup()
        return super().eventFilter(obj, event)

    def copy_to_clipboard(self):
        _, ctx = self.get_latest_data()
        if ctx:
            text = (ctx.get("context_text") or "").strip()
            if text:
                QApplication.clipboard().setText(text)

    def add_to_anki(self):
        entries, ctx = self.get_latest_data()
        if not entries or not ctx:
            return
        entry = entries[0]
        if isinstance(entry, KanjiEntry):
            return
        threading.Thread(
            target=self._add_to_anki_thread,
            args=(entry, ctx, entries),
            daemon=True
        ).start()

    def _add_to_anki_thread(self, entry: DictionaryEntry, ctx: Dict[str, Any], all_entries=None):
        """Runs in a background thread. Only emits signals to touch the UI."""
        anki = AnkiClient(getattr(config, "url", "http://127.0.0.1:8765"))
        if not anki.ping():
            logger.error("AnkiConnect not reachable")
            self.status_message_signal.emit("Error: Anki not reachable")
            return

        word    = getattr(entry, "written_form", "") or ""
        reading = getattr(entry, "reading",       "") or ""
        sentence = (ctx.get("context_text") or "").strip()

        meanings = []
        for sense in (getattr(entry, "senses", []) or []):
            glosses = sense.get("glosses", []) if isinstance(sense, dict) else []
            if glosses:
                meanings.append(glosses[0])
        meaning_str = "<br>".join(meanings)

        # Duplicate guard — same expression-field logic as presence check
        if getattr(config, "prevent_duplicates", True):
            deck_name   = getattr(config, "deck_name", "")
            deck_filter = f'deck:"{deck_name}" ' if deck_name else ""
            field_map   = getattr(config, "anki_field_map", {}) or {}
            expr_field  = next(
                (af for af, src in field_map.items()
                 if src in ("{expression}", "Word", "Expression")), None)
            dup_fields  = [expr_field] if expr_field else ["Front", "Word", "Expression"]
            safe_re     = _re.escape(word or reading)
            for field in dup_fields:
                try:
                    if anki.find_notes(f'{deck_filter}{field}:re:^{safe_re}$'):
                        self.status_message_signal.emit("Already in Anki")
                        _dup_word = word or reading
                        self.anki_presence_updated.emit(_dup_word, True)
                        return
                except Exception:
                    pass

        # Build field map
        senses = getattr(entry, "senses", []) or []

        # Full glossary — every sense, all glosses, joined
        full_glossary = "<br>".join(
            ", ".join(s.get("glosses", [])) for s in senses
        )
        # Plain-text glossary (no HTML tags)
        full_glossary_plain = "\n".join(
            ", ".join(s.get("glosses", [])) for s in senses
        )
        # First gloss of the highest-priority (first) sense
        first_dict_glossary = ""
        for s in senses:
            glosses = s.get("glosses", [])
            if glosses:
                first_dict_glossary = glosses[0]
                break
        # Part-of-speech: collect unique POS tags across all senses
        pos_set: list = []
        for s in senses:
            for p in (s.get("pos", []) or []):
                if p not in pos_set:
                    pos_set.append(p)
        part_of_speech_str = ", ".join(pos_set)

        # Furigana variants
        furigana_plain = f"{word}[{reading}]" if (word and reading) else (word or reading)
        # HTML ruby furigana: <ruby>word<rt>reading</rt></ruby>
        furigana_html = (
            f"<ruby>{word}<rt>{reading}</rt></ruby>"
            if (word and reading) else (word or reading)
        )

        freq_val = getattr(entry, "freq", 999999)
        freq_str = "" if freq_val >= 999999 else str(freq_val)

        tags_str = " ".join(sorted(getattr(entry, "tags", set()) or []))
        conj_str = " > ".join(getattr(entry, "deconjugation_process", ()) or ())
        dict_name = getattr(entry, "dictionary_name", "") or ""

        # sentence split at cloze boundary (word position)
        cloze_prefix = ""
        cloze_suffix = ""
        if sentence and word:
            idx = sentence.find(word)
            if idx >= 0:
                cloze_prefix = sentence[:idx]
                cloze_suffix = sentence[idx + len(word):]

        # -------------------------------------------------------------------
        # Build per-dictionary glossary maps from all candidate entries.
        # Covers {single-glossary-DICT}, {single-glossary-DICT-brief},
        # {single-glossary-DICT-no-dictionary}, and frequency variants.
        # -------------------------------------------------------------------
        single_glossary_by_dict: Dict[str, str] = {}
        for e in (all_entries or [entry]):
            if not isinstance(e, DictionaryEntry):
                continue
            dname = getattr(e, "dictionary_name", "") or ""
            if not dname or dname in single_glossary_by_dict:
                continue
            e_senses = getattr(e, "senses", []) or []
            all_glosses = "<br>".join(
                ", ".join(s.get("glosses", [])) for s in e_senses
            )
            single_glossary_by_dict[dname] = all_glosses

        # -------------------------------------------------------------------
        # Static marker table — every official Yomitan term-card marker.
        # Unsupported markers are left blank so they expand to "" rather than
        # being left as raw {placeholder} text on the card.
        # -------------------------------------------------------------------
        data_sources = {
            # ── audio / media ──────────────────────────────────────────────
            "{audio}":                              "",   # injected separately via AnkiConnect
            "{clipboard-image}":                    "",   # not available in desktop OCR context
            "{clipboard-text}":                     "",   # not available
            "{picture}":                            "",   # injected separately via screenshot logic
            "{screenshot}":                         "",   # injected separately
            # ── cloze ──────────────────────────────────────────────────────
            "{cloze-body}":                         word,
            "{cloze-body-kana}":                    reading,
            "{cloze-prefix}":                       cloze_prefix,
            "{cloze-suffix}":                       cloze_suffix,
            # ── expression / reading ───────────────────────────────────────
            "{expression}":                         word,
            "{reading}":                            reading,
            "{furigana}":                           furigana_html,
            "{furigana-plain}":                     furigana_plain,
            # ── glossary ───────────────────────────────────────────────────
            "{glossary}":                           full_glossary,
            "{glossary-brief}":                     meaning_str,
            "{glossary-no-dictionary}":             full_glossary,
            "{glossary-plain}":                     full_glossary_plain,
            "{glossary-plain-no-dictionary}":       full_glossary_plain,
            "{glossary-first}":                     first_dict_glossary,
            "{glossary-first-brief}":               first_dict_glossary,
            "{glossary-first-no-dictionary}":       first_dict_glossary,
            # legacy / community aliases
            "{glossary-1st-dict}":                  first_dict_glossary,
            "{jpmn-primary-definition}":            first_dict_glossary,
            # ── part of speech / conjugation ───────────────────────────────
            "{part-of-speech}":                     part_of_speech_str,
            "{conjugation}":                        conj_str,
            "{tags}":                               tags_str,
            # ── frequency ──────────────────────────────────────────────────
            "{freq}":                               freq_str,   # community alias
            "{frequencies}":                        freq_str,
            "{frequency-harmonic-rank}":            freq_str,
            "{frequency-average-rank}":             freq_str,
            "{frequency-harmonic-occurrence}":      "",   # occurrence-based, not available
            "{frequency-average-occurrence}":       "",
            # ── pitch accent (not supported) ───────────────────────────────
            "{pitch-accents}":                      "",
            "{pitch-accent-graphs}":                "",
            "{pitch-accent-graphs-jj}":             "",
            "{pitch-accent-positions}":             "",
            "{pitch-accent-categories}":            "",
            "{phonetic-transcriptions}":            "",
            # ── dictionary meta ────────────────────────────────────────────
            "{dictionary}":                         dict_name,
            "{dictionary-alias}":                   dict_name,
            # ── sentence / context ─────────────────────────────────────────
            "{sentence}":                           sentence,
            "{sentence-furigana}":                  sentence,   # furigana generation not available
            "{sentence-furigana-plain}":            sentence,
            "{search-query}":                       sentence,
            "{popup-selection-text}":               "",
            "{selection-text}":                     "",   # animecards alias
            # ── page context ───────────────────────────────────────────────
            "{document-title}":                     ctx.get("document_title", ""),
            "{url}":                                "",   # no browser context
        }

        # Inject per-dictionary single-glossary variants for every loaded dict
        for dname, gloss in single_glossary_by_dict.items():
            data_sources[f"{{single-glossary-{dname}}}"]              = gloss
            data_sources[f"{{single-glossary-{dname}-brief}}"]        = gloss
            data_sources[f"{{single-glossary-{dname}-no-dictionary}}"] = gloss
            # Frequency variants per dict — not available, leave blank
            data_sources[f"{{single-frequency-{dname}}}"]             = ""
            data_sources[f"{{single-frequency-number-{dname}}}"]      = ""

        def _resolve_template(template: str) -> str:
            """Replace every known {placeholder} within a field template string."""
            result = template
            for placeholder, value in data_sources.items():
                result = result.replace(placeholder, value or "")
            return result

        user_map = getattr(config, "anki_field_map", {}) or {}
        if user_map:
            fields = {k: _resolve_template(v) for k, v in user_map.items() if v}
        else:
            fields = {"Front": word or reading, "Back": meaning_str}

        tags = []
        if getattr(config, "add_meikipop_tag", True):
            tags.append("weikipop")
        if getattr(config, "add_document_title_tag", True):
            title = (ctx.get("document_title") or "").strip()
            if title:
                tags.append(title.replace(" ", "_"))

        note = {
            "deckName":  getattr(config, "deck_name",  "Default"),
            "modelName": getattr(config, "model_name", "Basic"),
            "fields":    fields,
            "tags":      tags,
        }

        if getattr(config, "enable_screenshot", False) and ctx.get("screenshot"):
            try:
                from PIL import Image
                screenshot = ctx["screenshot"]
                img  = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
                fname = f"weikipop_{int(time.time())}.png"
                bio   = BytesIO()
                img.save(bio, format="PNG")
                b64 = base64.b64encode(bio.getvalue()).decode("ascii")
                anki.store_media_file(fname, b64)
                img_tag = f'<img src="{fname}">'

                # Insert into whichever field the user mapped to {picture},
                # OR any field whose name suggests it holds images.
                field_map = getattr(config, "anki_field_map", {}) or {}
                pic_anki_field = next(
                    (af for af, src in field_map.items() if src == "{picture}"), None
                )
                if pic_anki_field and pic_anki_field in note["fields"]:
                    note["fields"][pic_anki_field] = img_tag
                else:
                    # Fallback: common image field names
                    for pic_field in ("Picture", "Image", "Screenshot", "image", "picture", "screenshot"):
                        if pic_field in note["fields"]:
                            note["fields"][pic_field] = img_tag
                            break
                    else:
                        # Last resort: add as extra field so it's not lost
                        note["fields"]["Picture"] = img_tag
            except Exception as e:
                logger.error(f"Screenshot failed: {e}")

        try:
            note_id = anki.add_note(note)
            logger.info(f"Added note {note_id} to Anki")
            self._append_mining_log(entry, ctx, note, note_id)
            self.status_message_signal.emit(f"Mined: {word or reading}")
            _mined_word = word or reading
            self.anki_presence_updated.emit(_mined_word, True)
        except Exception as e:
            logger.error(f"Failed to add note: {e}")
            s = str(e)
            if any(x in s for x in ("10061", "Connection refused", "Max retries",
                                     "NewConnectionError", "Failed to establish")):
                self.status_message_signal.emit("Error: Anki not running")
            else:
                self.status_message_signal.emit("Error: could not add note")

    def _append_mining_log(self, entry: DictionaryEntry, ctx: Dict[str, Any], note: Dict[str, Any], note_id: int):
        try:
            os.makedirs('data', exist_ok=True)
            glosses = []
            for sense in getattr(entry, 'senses', []) or []:
                glosses.extend(sense.get('glosses', []))
            payload = {
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'note_id': note_id,
                'expression': getattr(entry, 'written_form', ''),
                'reading': getattr(entry, 'reading', ''),
                'dictionary': getattr(entry, 'dictionary_name', ''),
                'glosses': glosses,
                'context': (ctx.get('context_text') or '').strip(),
                'document_title': ctx.get('document_title', ''),
                'anki_note': {
                    'deckName': note.get('deckName', ''),
                    'modelName': note.get('modelName', ''),
                    'fields': note.get('fields', {}),
                },
            }
            with open('data/mining_log.jsonl', 'a', encoding='utf-8') as file:
                file.write(json.dumps(payload, ensure_ascii=False) + '\n')
        except Exception as exc:
            logger.debug('Failed to append mining log: %s', exc)

    # ------------------------------------------------------------------ #
    #  Anki presence check (background thread)                             #
    # ------------------------------------------------------------------ #

    def _check_anki_presence_async(self, word: str):
        if not getattr(config, "show_hover_status", True):
            return

        def _run():
            try:
                anki = AnkiClient(getattr(config, "url", "http://127.0.0.1:8765"))
                if not anki.ping():
                    return  # Anki not running
                if not word:
                    return

                deck_name   = getattr(config, "deck_name", "")
                deck_filter = f'deck:"{deck_name}" ' if deck_name else ""

                # Find which Anki field the user mapped to {expression}.
                # If no mapping exists, fall back to common field names.
                field_map = getattr(config, "anki_field_map", {}) or {}
                expression_field = next(
                    (anki_f for anki_f, src in field_map.items()
                     if src in ("{expression}", "Word", "Expression")),
                    None
                )
                search_fields = [expression_field] if expression_field else [
                    "Front", "Word", "Expression", "Vocab", "Kanji"
                ]

                # re:^...$  = exact field match, not substring
                safe  = _re.escape(word)
                found = False
                for field in search_fields:
                    try:
                        if anki.find_notes(f'{deck_filter}{field}:re:^{safe}$'):
                            found = True
                            break
                    except Exception:
                        pass
                self.anki_presence_updated.emit(word, found)
            except Exception:
                pass

        threading.Thread(target=_run, daemon=True).start()

    # ------------------------------------------------------------------ #
    #  Status message (called on main thread via signal)                   #
    # ------------------------------------------------------------------ #

    def _show_status_message(self, message: str, duration_ms: int = 2500):
        self.status_label.setText(message)
        self.status_label.show()
        if duration_ms > 0:
            QTimer.singleShot(duration_ms, self.status_label.hide)

    # ------------------------------------------------------------------ #
    #  Content rendering                                                    #
    # ------------------------------------------------------------------ #

    def _render_senses(self, entry, max_ratio: float, inline_only: bool = False) -> tuple:
        """Render the definitions block for one DictionaryEntry. Returns (senses_html, updated_max_ratio).
        inline_only=True skips the leading <br> so the caller controls line breaks."""
        parts_calc, parts_html = [], []
        for idx, sense in enumerate(entry.senses):
            glosses   = sense.get("glosses", [])
            pos_list  = sense.get("pos",     [])
            tags_list = sense.get("tags",    [])

            gloss_str = (", ".join(glosses) if config.show_all_glosses else (glosses[0] if glosses else ""))
            s_calc = f"({idx+1})" if config.show_all_glosses else ""
            s_html = f"<b>({idx+1})</b> " if config.show_all_glosses else ""

            if config.show_pos and pos_list:
                pos_str = f' ({", ".join(pos_list)})'
                s_calc += pos_str
                s_html += f'<span style="color:{config.color_foreground};opacity:0.7;"><i>{pos_str}</i></span> '
            if config.show_tags and tags_list:
                t_str = f' [{", ".join(tags_list)}]'
                s_calc += t_str
                s_html += (f'<span style="color:{config.color_foreground};'
                           f'font-size:{config.font_size_definitions-2}px;opacity:0.7;">{t_str}</span> ')
            s_calc += gloss_str
            s_html += gloss_str
            parts_calc.append(s_calc)
            parts_html.append(s_html)

        if config.compact_mode:
            sep = "; "
            full_def_html = sep.join(parts_html)
            max_ratio = max(max_ratio, len(sep.join(parts_calc)) / self.def_chars_per_line)
        else:
            sep = "<br>"
            full_def_html = sep.join(parts_html)
            for p in parts_calc:
                max_ratio = max(max_ratio, len(p) / self.def_chars_per_line)

        if inline_only:
            senses_html = (f'<span style="font-size:{config.font_size_definitions}px;">'
                           f'{full_def_html}</span>')
        else:
            sep_space = " " if config.compact_mode else "<br>"
            senses_html = (f'{sep_space}<span style="font-size:{config.font_size_definitions}px;">'
                           f'{full_def_html}</span>')
        return senses_html, max_ratio

    # How many word/reading groups to render on first display, and per scroll trigger.
    _INITIAL_RENDER_GROUPS = 2
    _GROUPS_PER_LOAD = 3

    def _render_groups_to_html(self, groups: list, start_index: int = 0) -> tuple:
        """Render a list of entry groups to an HTML string.
        start_index is the absolute group index of groups[0], used only to decide
        whether to prepend a <hr> separator before the first group.
        Returns (html_string, max_ratio)."""
        html_parts = []
        max_ratio  = 0.0

        for i, group in enumerate(groups):
            g_idx = start_index + i
            if g_idx > 0:
                html_parts.append('<hr style="margin-top:0;margin-bottom:0;">')

            # ── Kanji entry ──────────────────────────────────────────────
            if isinstance(group, KanjiEntry):
                defn = ', '.join(group.meanings) if (config.show_examples or config.show_components) else '[字]'
                calc = f"{group.character} {', '.join(group.readings)} {defn}"
                max_ratio = max(max_ratio, len(calc) / self.header_chars_per_line, 0.7)
                html_parts.append(self._render_kanji_entry(group))
                continue

            # ── Dictionary entry group ───────────────────────────────────
            word_key, dict_entries = group
            first_entry = dict_entries[0]

            header_calc = first_entry.written_form or ""
            if first_entry.reading:
                header_calc += f" [{first_entry.reading}]"
            max_ratio = max(max_ratio, len(header_calc) / self.header_chars_per_line)

            header_html = (
                f'<span style="color:{config.color_highlight_word};'
                f'font-size:{config.font_size_header}px;">{first_entry.written_form}</span>'
            )
            if first_entry.reading:
                header_html += (
                    f' <span style="color:{config.color_highlight_reading};'
                    f'font-size:{config.font_size_header - 2}px;">[{first_entry.reading}]</span>'
                )
            if first_entry.deconjugation_process and config.show_deconjugation:
                dc = " ← ".join(p for p in first_entry.deconjugation_process if p)
                if dc:
                    header_html += (
                        f' <span style="color:{config.color_foreground};'
                        f'font-size:{config.font_size_definitions - 2}px;opacity:0.8;">({dc})</span>'
                    )
            if config.show_frequency and first_entry.freq < 999_999:
                header_html += (
                    f' <span style="color:{config.color_foreground};'
                    f'font-size:{config.font_size_definitions - 2}px;opacity:0.6;">#{first_entry.freq}</span>'
                )

            multi_dict = len(dict_entries) > 1
            body_parts = []
            for entry in dict_entries:
                if multi_dict:
                    senses_html, max_ratio = self._render_senses(entry, max_ratio, inline_only=True)
                    dict_name = getattr(entry, 'dictionary_name', '') or 'Dictionary'
                    dict_label = (
                        f'<span style="color:{config.color_foreground};'
                        f'font-size:{config.font_size_definitions}px;opacity:0.85;">'
                        f'<b>{dict_name}:</b> </span>'
                    )
                    body_parts.append(f'{dict_label}{senses_html}')
                else:
                    senses_html, max_ratio = self._render_senses(entry, max_ratio)
                    if getattr(entry, 'dictionary_name', ''):
                        header_html += (
                            f' <span style="color:{config.color_foreground};'
                            f'font-size:{config.font_size_definitions - 2}px;opacity:0.75;">'
                            f'[{entry.dictionary_name}]</span>'
                        )
                    body_parts.append(senses_html)

            if multi_dict:
                p_header = f'<p style="margin:0;padding:0;">{header_html}</p>'
                p_dicts = ''.join(
                    f'<p style="margin:0;padding:0;margin-top:3px;">{part}</p>'
                    for part in body_parts
                )
                html_parts.append(p_header + p_dicts)
            else:
                combined_body = body_parts[0] if body_parts else ''
                html_parts.append(f"{header_html}{combined_body}")

        return "".join(html_parts), max_ratio

    def _measure_html_height(self, html: str, width: int) -> int:
        """Measure the pixel height needed to render html at the given width.
        Uses a QTextDocument (same engine as QLabel) — synchronous, accurate,
        and has zero side-effects on any visible widget or scroll position."""
        doc = QTextDocument()
        doc.setDefaultFont(QFont(config.font_family))
        doc.setHtml(html)
        doc.setTextWidth(width)
        return int(doc.size().height())

    def _fixed_popup_size(self) -> QSize:
        """Return the cached fixed size for the popup.
        Recomputed only when reapply_settings() clears _last_size — not on every
        timer tick — so dictionary_sources is never read at 60ms intervals."""
        if self._cached_popup_size is not None:
            return self._cached_popup_size
        margins  = self.content_layout.contentsMargins()
        border   = 1
        h_pad    = margins.left() + margins.right() + border * 2
        v_pad    = margins.top() + margins.bottom() + border * 2 + MINE_BAR_HEIGHT + self.content_layout.spacing()
        screen   = QApplication.primaryScreen()
        screen_w = screen.geometry().width()  if screen else 1920
        screen_h = screen.geometry().height() if screen else 1080
        w = int(screen_w * 0.30)
        if config.compact_mode:
            h = min(int(screen_h * 0.22), 220)
        else:
            h = min(int(screen_h * 0.45), 420)
            sources = getattr(config, 'dictionary_sources', []) or []
            enabled = [s for s in sources if s.get('enabled', True)]
            if len(enabled) <= 1:
                h = int(h * 0.60)
        self._cached_popup_size = QSize(w + h_pad, h + v_pad)
        return self._cached_popup_size

    def _calculate_content(self, entries) -> 'str | None':
        """Build and return the initial HTML to display.  Only the first
        _INITIAL_RENDER_GROUPS groups are rendered — the rest load as the
        user scrolls via _on_scroll_lazy_load."""
        if not self.is_calibrated or not entries:
            self._lazy_pending_groups = []
            self._lazy_rendered_parts = []
            return None

        # Build display groups: entries sharing (written_form, reading) merged.
        all_groups = []
        for entry in entries:
            if isinstance(entry, KanjiEntry):
                all_groups.append(entry)
                continue
            word_key = (entry.written_form, entry.reading)
            if all_groups and isinstance(all_groups[-1], list) and all_groups[-1][0] == word_key:
                all_groups[-1][1].append(entry)
            else:
                all_groups.append([word_key, [entry]])

        initial_groups              = all_groups[:self._INITIAL_RENDER_GROUPS]
        self._lazy_pending_groups   = all_groups[self._INITIAL_RENDER_GROUPS:]
        self._lazy_next_group_index = len(initial_groups)

        initial_html, _ = self._render_groups_to_html(initial_groups, start_index=0)
        self._lazy_rendered_parts   = [initial_html]
        return initial_html

    # ------------------------------------------------------------------ #
    #  Lazy entry loading                                                   #
    # ------------------------------------------------------------------ #

    def _on_scroll_lazy_load(self, value: int):
        """Triggered by the scrollbar — appends the next batch of entry groups
        when the user has scrolled at least 70% of the way through current content."""
        if not self._lazy_pending_groups:
            return
        sb = self.content_scroll.verticalScrollBar()
        if sb.maximum() > 0 and value >= sb.maximum() * 0.70:
            self._append_next_lazy_batch()

    def _append_next_lazy_batch(self):
        """Render the next _GROUPS_PER_LOAD pending groups and append them
        without disturbing the user's current scroll position."""
        if not self._lazy_pending_groups:
            return
        sb        = self.content_scroll.verticalScrollBar()
        saved_pos = sb.value()

        batch                     = self._lazy_pending_groups[:self._GROUPS_PER_LOAD]
        self._lazy_pending_groups = self._lazy_pending_groups[self._GROUPS_PER_LOAD:]

        batch_html, _ = self._render_groups_to_html(
            batch, start_index=self._lazy_next_group_index
        )
        self._lazy_next_group_index += len(batch)
        self._lazy_rendered_parts.append(batch_html)

        full_html       = "".join(self._lazy_rendered_parts)
        self._last_html = full_html   # keep in sync so the 60ms timer doesn't re-render
        self.display_label.setText(full_html)
        # Restore position after Qt settles the layout — content above the
        # saved position is unchanged so this keeps the view perfectly stable.
        QTimer.singleShot(0, lambda pos=saved_pos: sb.setValue(pos))

    def _render_kanji_entry(self, entry: KanjiEntry) -> str:
        c_word = config.color_highlight_word
        c_read = config.color_highlight_reading
        c_text = config.color_foreground
        fs_h   = config.font_size_header
        fs_d   = config.font_size_definitions

        readings_str = f"[{', '.join(entry.readings)}]"
        header_html  = (
            f'<span style="font-size:{fs_h}px;color:{c_word};padding-right:8px;">{entry.character}</span>'
            f'<span style="font-size:{fs_h-2}px;color:{c_read};"> {readings_str}</span>'
        )
        meanings_html = f'<span style="font-size:{fs_d}px;color:{c_text};">{", ".join(entry.meanings)}</span>'
        if not config.compact_mode:
            meanings_html = (f'<span style="font-size:{fs_d}px;color:{c_text};"> [字]</span>'
                             f'<div>{meanings_html}</div>')

        examples_html = ""
        if config.show_examples:
            parts = [
                (f"<span style='font-size:{fs_h-2}px;color:{c_word}'>{e['w']}</span> "
                 f"<span style='font-size:{fs_d}px;color:{c_read}'>[{e['r']}]</span> "
                 f"<span style='font-size:{fs_d}px;color:{c_text}'>{e['m']}</span>")
                for e in entry.examples
            ]
            if parts:
                examples_html = f'<div>{"; ".join(parts)}</div>'

        components_html = ""
        if config.show_components:
            parts = [
                (f"<span style='font-size:{fs_d}px;color:{c_word}'>{c.get('c','')}</span> "
                 f"<span style='font-size:{fs_d}px;color:{c_text}'>{c.get('m','')}</span>")
                for c in entry.components
            ]
            if parts:
                components_html = f'<div>{", ".join(parts)}</div>'

        return (f'<div style="border:1px solid {c_word};">'
                f'{header_html}{meanings_html}{examples_html}{components_html}</div>')

    # ------------------------------------------------------------------ #
    #  Popup visibility & positioning                                       #
    # ------------------------------------------------------------------ #

    def show_popup(self):
        if self.is_visible:
            return
        # Non-blocking: if screenshot is in progress, skip this tick
        # rather than freezing the main thread waiting for the lock.
        if not self.shared_state.screen_lock.acquire(blocking=False):
            return
        self._store_active_window_on_mac()
        self.show()
        if IS_MACOS:
            self.raise_()
        self.is_visible = True
        self.input_loop.suppress_scroll = True

    def hide_popup(self):
        if not self.is_visible:
            return
        self.hide()
        self.is_visible = False
        self.input_loop.suppress_scroll = False
        self._last_presence_word = None  # re-check on next show, even same word
        QTimer.singleShot(50, self._release_lock_safely)
        self._restore_focus_on_mac()

    def _release_lock_safely(self):
        logger.debug("releasing screen_lock")
        self.shared_state.screen_lock.release()

    def move_to(self, x: int, y: int):
        cursor_point = QPoint(x, y)
        screen = QApplication.screenAt(cursor_point) or QApplication.primaryScreen()
        if screen is None:
            return
        screen_geo  = screen.geometry()
        popup_size  = self.size()
        offset      = 15

        ratio = screen.devicePixelRatio()
        x, y  = magpie_manager.transform_raw_to_visual((int(x), int(y)), ratio)

        mode = config.popup_position_mode

        if mode == "visual_novel_mode":
            sh = screen_geo.height()
            cy = y - screen_geo.top()
            if cy > 2 * sh / 3:
                is_below = False
            elif cy < sh / 3:
                is_below = True
            else:
                is_below = cy < sh / 2
            final_y = (y + offset) if is_below else (y - popup_size.height() - offset)
            final_y = max(screen_geo.top(), min(final_y, screen_geo.bottom() - popup_size.height()))

            sw = screen_geo.width()
            cx = x - screen_geo.left()
            pr = x + offset
            pc = x - popup_size.width() / 2
            pl = x - popup_size.width() - offset
            if cx < sw / 2:
                t = cx / (sw / 2)
                final_x = pr * (1 - t) + pc * t
            else:
                t = (cx - sw / 2) / (sw / 2)
                final_x = pc * (1 - t) + pl * t

        elif mode == "flip_horizontally":
            pref_x  = x + offset
            final_x = pref_x if pref_x + popup_size.width() <= screen_geo.right() else x - popup_size.width() - offset
            final_y = y + offset
            final_y = max(screen_geo.top(), min(final_y, screen_geo.bottom() - popup_size.height()))

        elif mode == "flip_vertically":
            final_x = x + offset
            final_x = max(screen_geo.left(), min(final_x, screen_geo.right() - popup_size.width()))
            pref_y  = y + offset
            final_y = pref_y if pref_y + popup_size.height() <= screen_geo.bottom() else y - popup_size.height() - offset

        else:  # flip_both
            pref_x  = x + offset
            final_x = pref_x if pref_x + popup_size.width() <= screen_geo.right() else x - popup_size.width() - offset
            pref_y  = y + offset
            final_y = pref_y if pref_y + popup_size.height() <= screen_geo.bottom() else y - popup_size.height() - offset

        final_x = max(screen_geo.left(), min(final_x, screen_geo.right()  - popup_size.width()))
        final_y = max(screen_geo.top(),  min(final_y, screen_geo.bottom() - popup_size.height()))
        self.move(int(final_x), int(final_y))

    def reapply_settings(self):
        logger.debug("Popup: reapplying settings")
        self._apply_frame_stylesheet()
        self.is_calibrated = False
        self._last_size = None
        self._cached_popup_size = None   # recompute size — dict count or compact mode may have changed

    # ------------------------------------------------------------------ #
    #  macOS focus management                                               #
    # ------------------------------------------------------------------ #

    def _store_active_window_on_mac(self):
        if not IS_MACOS or not Quartz:
            return
        try:
            app = Quartz.NSWorkspace.sharedWorkspace().frontmostApplication()
            if app:
                self._previous_active_window_on_mac = app
        except Exception as e:
            logger.warning(f"store_active_window failed: {e}")
            self._previous_active_window_on_mac = None

    def _restore_focus_on_mac(self):
        if not IS_MACOS or not Quartz or not self._previous_active_window_on_mac:
            return
        try:
            self._previous_active_window_on_mac.activateWithOptions_(
                Quartz.NSApplicationActivateAllWindows
            )
        except Exception as e:
            logger.warning(f"restore_focus failed: {e}")
        finally:
            self._previous_active_window_on_mac = None
