# src/config/config.py
import configparser
import json
import logging
import sys

logger = logging.getLogger(__name__)

APP_NAME = "weikipop"
APP_VERSION = "v.1.12.6"
MAX_DICT_ENTRIES = 10
IS_LINUX = sys.platform.startswith('linux')
IS_WINDOWS = sys.platform.startswith('win')
IS_MACOS = sys.platform.startswith('darwin')

class Config:
    _instance = None

    _SCHEMA = {
        'Settings': {
            'hotkey': 'alt',
            'scan_region': '0',
            'max_lookup_length': 25,
            'glens_low_bandwidth': False,
            'ocr_provider': 'meikiocr (local)',
            'auto_scan_mode': True,
            'auto_scan_mode_lookups_without_hotkey': False,
            'auto_scan_interval_seconds': 1.0,
            'auto_scan_on_mouse_move': True,
            'magpie_compatibility': True,
            'show_keyboard_shortcuts': True,
            'dictionary_sources': '[]',
        },
        'Theme': {
            'theme_name': 'Celestial Indigo',
            'font_family': 'Segoe UI',
            'font_size_definitions': 16,
            'font_size_header': 22,
            'compact_mode': True,
            'show_all_glosses': False,
            'show_deconjugation': False,
            'show_pos': False,
            'show_tags': True,
            'show_frequency': True,
            'show_pitch_accent': False,
            'show_kanji': True,
            'show_examples': False,
            'show_components': False,
            'color_background': '#281E50',
            'color_foreground': '#EAEFF5',
            'color_highlight_word': '#D4C58A',
            'color_highlight_reading': '#B5A2D4',
            'background_opacity': 245,
            'popup_position_mode': 'visual_novel_mode'
        },
        'Anki': {
            'enabled': True,
            'url': 'http://127.0.0.1:8765',
            'deck_name': 'Default',
            'model_name': 'Basic',
            'show_hover_status': True,
            'add_meikipop_tag': True,
            'add_document_title_tag': True,
            'enable_screenshot': False,
            'prevent_duplicates': True,
            'field_map': '{"Expression": "{expression}", "ExpressionFurigana": "{furigana-plain}", "ExpressionReading": "{reading}", "ExpressionAudio": "{audio}", "MainDefinition": "{glossary-first}", "Sentence": "{cloze-prefix}<b>{cloze-body}</b>{cloze-suffix}", "PitchPosition": "{pitch-accent-positions}", "FreqSort": "{frequency-harmonic-rank}", "MiscInfo": "{document-title}"}',
            'duplicate_check_fields': 'Front,Word,Expression,Vocab,Kanji,Reading,Furigana,Writing,Term,Vocabulary',
            'sentence_truncation_delimiters': '。,！,？,（,）',
            'sentence_truncation_delimiters_remove': '「,」',
        },
        'Yomitan': {
            'enabled': False,
            'api_url': 'http://127.0.0.1:19633',
        },
        'Shortcuts': {
            'add_to_anki': 'Alt+A',
            'copy_text': 'Alt+C',
            'scroll_popup': 'Alt+Wheel',
        },
    }

    def __new__(cls):
        if not cls._instance:
            cls._instance = super().__new__(cls)
            cls._instance._load()
        return cls._instance

    def _load(self):
        parser = configparser.ConfigParser()
        parser.read('config.ini', encoding='utf-8')

        for section, settings in self._SCHEMA.items():
            for key, default in settings.items():
                if parser.has_option(section, key):
                    if isinstance(default, bool):
                        val = parser.getboolean(section, key)
                    elif isinstance(default, int):
                        val = parser.getint(section, key)
                    elif isinstance(default, float):
                        val = parser.getfloat(section, key)
                    else:
                        val = parser.get(section, key)
                else:
                    val = default
                setattr(self, key, val)

                # Preserve section-specific aliases for colliding keys.
                if section == 'Anki':
                    setattr(self, f'anki_{key}', val)
                elif section == 'Yomitan':
                    setattr(self, f'yomitan_{key}', val)

        # Parse structured settings
        self.dictionary_sources = self._parse_json(getattr(self, 'dictionary_sources', '[]'), default=[])
        self.anki_field_map = self._parse_json(getattr(self, 'field_map', '{}'), default={})
        self.anki_duplicate_check_fields = self._parse_csv(getattr(self, 'duplicate_check_fields', ''), default=[])
        self.anki_sentence_delimiters = self._parse_csv(getattr(self, 'sentence_truncation_delimiters', ''), default=['。'])
        self.anki_sentence_delimiters_remove = self._parse_csv(getattr(self, 'sentence_truncation_delimiters_remove', ''), default=[])

        # Canonical aliases to avoid section key collisions (enabled/api_url/url)
        self.anki_enabled = getattr(self, 'anki_enabled', True)
        self.anki_url = getattr(self, 'anki_url', 'http://127.0.0.1:8765')
        self.yomitan_enabled = getattr(self, 'yomitan_enabled', False)
        self.yomitan_api_url = getattr(self, 'yomitan_api_url', 'http://127.0.0.1:19633')

        # Keep legacy attributes used in existing code paths
        self.enabled = self.anki_enabled
        self.url = self.anki_url
        self.api_url = self.yomitan_api_url

        self.is_enabled = True
        logger.info("Configuration loaded.")

    @staticmethod
    def _parse_json(value: str, default):
        try:
            parsed = json.loads(value) if isinstance(value, str) else value
            return parsed if parsed is not None else default
        except Exception:
            return default

    @staticmethod
    def _parse_csv(value: str, default):
        if value is None:
            return list(default)
        if not isinstance(value, str):
            return list(default)
        parts = [p.strip().strip('\'"') for p in value.split(',')]
        parts = [p for p in parts if p]
        return parts if parts else list(default)

    def save(self):
        parser = configparser.ConfigParser()
        for section, settings in self._SCHEMA.items():
            parser.add_section(section)
            for key in settings:
                val = getattr(self, key)
                # Serialize structured values back into their string fields
                if section == 'Settings' and key == 'dictionary_sources':
                    val = json.dumps(getattr(self, 'dictionary_sources', []), ensure_ascii=False)
                if section == 'Anki' and key == 'field_map':
                    val = json.dumps(getattr(self, 'anki_field_map', {}), ensure_ascii=False)
                elif section == 'Anki' and key == 'duplicate_check_fields':
                    val = ",".join(getattr(self, 'anki_duplicate_check_fields', []))
                elif section == 'Anki' and key == 'sentence_truncation_delimiters':
                    val = ",".join(getattr(self, 'anki_sentence_delimiters', []))
                elif section == 'Anki' and key == 'sentence_truncation_delimiters_remove':
                    val = ",".join(getattr(self, 'anki_sentence_delimiters_remove', []))
                elif section == 'Anki' and key == 'enabled':
                    val = getattr(self, 'anki_enabled', getattr(self, 'enabled', True))
                elif section == 'Anki' and key == 'url':
                    val = getattr(self, 'anki_url', getattr(self, 'url', 'http://127.0.0.1:8765'))
                elif section == 'Yomitan' and key == 'enabled':
                    val = getattr(self, 'yomitan_enabled', False)
                elif section == 'Yomitan' and key == 'api_url':
                    val = getattr(self, 'yomitan_api_url', 'http://127.0.0.1:19633')
                parser.set(section, key, str(val).lower() if isinstance(val, bool) else str(val))

        with open('config.ini', 'w', encoding='utf-8') as f:
            parser.write(f)
        logger.info("Settings saved to config.ini.")


config = Config()
