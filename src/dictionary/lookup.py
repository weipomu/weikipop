# lookup.py - Optimized version
import logging
import math
import os
import pickle
import re
import shutil
import threading
import time
import uuid
from collections import OrderedDict, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from src.config.config import config, MAX_DICT_ENTRIES
from src.dictionary.customdict import Dictionary, WRITTEN_FORM_INDEX, READING_INDEX, FREQUENCY_INDEX, ENTRY_ID_INDEX, DEFAULT_FREQ
from src.dictionary.deconjugator import Deconjugator, Form
from src.dictionary.yomitan_client import YomitanClient
from src.dictionary.yomitan_importer import convert_yomitan_zip_to_payload, write_payload_pickle

KANJI_REGEX = re.compile(r'[\u4e00-\u9faf]')
JAPANESE_SEPARATORS = {
    "、", "。", "「", "」", "｛", "｝", "（", "）", "【", "】",
    "『", "』", "〈", "〉", "《", "》", "：", "・", "／",
    "…", "︙", "‥", "︰", "＋", "＝", "－", "÷", "？", "！",
    "．", "～", "―", "!", "?",
}

logger = logging.getLogger(__name__)


@dataclass
class DictionaryEntry:
    id: int
    written_form: str
    reading: str
    senses: list
    freq: int
    deconjugation_process: tuple
    priority: float = 0.0
    match_len: int = 0  # Add match_len field for Yomitan entries
    dictionary_name: str = ''
    dictionary_id: str = ''


@dataclass
class KanjiEntry:
    character: str
    meanings: List[str]
    readings: List[str]
    components: List[Dict[str, str]]
    examples: List[Dict[str, str]]


def _throttled(iterable, work_ms: float = 2.0, sleep_ms: float = 8.0):
    """Yield items from iterable while capping CPU usage.
    Works for work_ms then sleeps for sleep_ms — roughly 20% CPU regardless
    of machine speed, keeping the cursor smooth during dictionary loading."""
    work_s  = work_ms  / 1000.0
    sleep_s = sleep_ms / 1000.0
    t = time.monotonic()
    for item in iterable:
        yield item
        if time.monotonic() - t >= work_s:
            time.sleep(sleep_s)
            t = time.monotonic()


class Lookup(threading.Thread):
    def __init__(self, shared_state, popup_window):
        super().__init__(daemon=True, name="Lookup")
        self.shared_state = shared_state
        self.popup_window = popup_window
        self.last_hit_result = None
        self._dict_lock = threading.RLock()

        self.user_dictionary_dir = Path('user_dictionaries')
        self.user_dictionary_dir.mkdir(exist_ok=True)

        # entry_id -> source metadata
        self.entry_sources: Dict[int, Dict[str, Any]] = {}
        self.primary_kanji_entries: Dict[str, Dict[str, Any]] = {}

        # Cache loaded Dictionary objects by (path, mtime).  This avoids
        # re-reading and re-unpickling unchanged files (e.g. the 65 MB main
        # dict) on every import or settings-save that touches dictionaries.
        self._dict_file_cache: Dict[str, tuple] = {}  # path -> (mtime, Dictionary)

        self.dictionary = Dictionary()
        self.lookup_cache: OrderedDict = OrderedDict()
        self.CACHE_SIZE = 500
        self._load_configured_dictionaries()
        self.deconjugator = Deconjugator(self.dictionary.deconjugator_rules)

        # Lazy initialization of Yomitan client - only when needed
        self._yomitan_client: Optional[YomitanClient] = None
        self._yomitan_enabled = getattr(config, "yomitan_enabled", False)
        self._yomitan_available = None  # None = untested, True/False = cached result

    @property
    def yomitan_client(self):
        """Lazy property - only create Yomitan client if actually needed"""
        if not self._yomitan_enabled:
            return None
        if self._yomitan_client is None:
            try:
                self._yomitan_client = YomitanClient(getattr(config, "yomitan_api_url", "http://127.0.0.1:19633"))
            except Exception:
                self._yomitan_enabled = False  # Disable permanently if creation fails
                return None
        return self._yomitan_client

    def clear_cache(self):
        with self._dict_lock:
            self.lookup_cache = OrderedDict()

    def _default_dictionary_sources(self) -> List[Dict[str, Any]]:
        return [{
            'id': 'builtin-main',
            'name': 'Main Dictionary',
            'path': 'dictionary.pkl',
            'enabled': True,
            'priority': 0,
            'kind': 'pickle',
            'builtin': True,
        }]

    def get_dictionary_sources(self) -> List[Dict[str, Any]]:
        sources = getattr(config, 'dictionary_sources', []) or []
        if not sources:
            sources = self._default_dictionary_sources()
        return sorted(sources, key=lambda x: int(x.get('priority', 0)))

    def set_dictionary_sources(self, sources: List[Dict[str, Any]], progress_cb=None):
        normalized = []
        for idx, source in enumerate(sources):
            normalized.append({
                'id': source.get('id') or str(uuid.uuid4()),
                'name': (source.get('name') or 'Dictionary').strip(),
                'path': source.get('path') or '',
                'enabled': bool(source.get('enabled', True)),
                'priority': idx,
                'kind': source.get('kind') or 'pickle',
                'builtin': bool(source.get('builtin', False)),
            })

        has_builtin = any(s.get('builtin') for s in normalized)
        if not has_builtin:
            normalized.insert(0, self._default_dictionary_sources()[0])
            for idx, source in enumerate(normalized):
                source['priority'] = idx

        config.dictionary_sources = normalized
        config.save()
        self._load_configured_dictionaries(progress_cb=progress_cb)

    def delete_dictionary_source(self, source_id: str) -> Tuple[bool, str]:
        """Delete a dictionary source entry and its file when appropriate.

        This is intentionally separate from enable/disable toggles.
        """
        if not source_id:
            return False, 'Missing dictionary id.'

        sources = self.get_dictionary_sources()
        target = next((s for s in sources if s.get('id') == source_id), None)
        if not target:
            return False, 'Dictionary not found.'
        if target.get('builtin'):
            return False, 'Built-in dictionary cannot be deleted.'

        path = target.get('path', '')
        if path:
            try:
                p = Path(path)
                # Only remove imported dictionary files inside managed directory.
                if p.exists() and self.user_dictionary_dir.resolve() in p.resolve().parents:
                    p.unlink()
            except Exception as exc:
                return False, f'Failed to delete dictionary file: {exc}'

        remaining = [s for s in sources if s.get('id') != source_id]
        self.set_dictionary_sources(remaining)
        return True, ''

    def import_dictionary_files(self, file_paths: List[str], progress_cb=None) -> Dict[str, Any]:
        report = {'imported': [], 'failed': [], 'skipped': []}
        if not file_paths:
            return report

        sources = self.get_dictionary_sources()
        existing_names = {s.get('name', '').lower(): s for s in sources}
        n_files = len(file_paths)

        for i, file_path in enumerate(file_paths):
            if progress_cb:
                progress_cb(i, n_files, f"Converting {Path(file_path).name}...")
            try:
                path = Path(file_path)
                if not path.exists():
                    report['failed'].append((file_path, 'File not found'))
                    continue

                suffix = path.suffix.lower()
                if suffix not in {'.zip', '.pkl'}:
                    report['failed'].append((file_path, 'Unsupported file type'))
                    continue

                if suffix == '.zip':
                    payload, suggested_name = convert_yomitan_zip_to_payload(str(path), dict_index=0)
                    safe_name = self._unique_dictionary_name(suggested_name, existing_names)
                    out_path = self.user_dictionary_dir / f'{safe_name}.pkl'
                    write_payload_pickle(payload, str(out_path))
                    source_name = safe_name
                else:
                    with open(path, 'rb') as file:
                        payload = pickle.load(file)
                    if 'entries' not in payload or 'lookup_map' not in payload:
                        report['failed'].append((file_path, 'Invalid dictionary pickle format'))
                        continue
                    source_name = self._unique_dictionary_name(path.stem, existing_names)
                    out_path = self.user_dictionary_dir / f'{source_name}.pkl'
                    shutil.copyfile(path, out_path)

                source = {
                    'id': str(uuid.uuid4()),
                    'name': source_name,
                    'path': str(out_path),
                    'enabled': True,
                    'priority': len(sources),
                    'kind': 'pickle',
                    'builtin': False,
                }
                sources.append(source)
                existing_names[source_name.lower()] = source
                report['imported'].append((file_path, source_name))
            except Exception as exc:
                report['failed'].append((file_path, str(exc)))

        # Loading phase — pass progress through so the bar updates during the
        # combined dictionary rebuild that follows the file conversion.
        def _load_progress(current, total, msg):
            if progress_cb:
                progress_cb(current, total, msg or "Loading dictionaries...")

        self.set_dictionary_sources(sources, progress_cb=_load_progress)
        return report

    @staticmethod
    def _unique_dictionary_name(base_name: str, existing_names: Dict[str, Dict[str, Any]]) -> str:
        sanitized = re.sub(r'[^\w\-\s\u3040-\u30ff\u3400-\u9fff]', '', (base_name or '').strip())
        sanitized = sanitized or 'Dictionary'
        candidate = sanitized
        counter = 2
        while candidate.lower() in existing_names:
            candidate = f'{sanitized} ({counter})'
            counter += 1
        return candidate

    def _load_configured_dictionaries(self, progress_cb=None):
        with self._dict_lock:
            sources = self.get_dictionary_sources()

            combined_entries: Dict[int, list] = {}
            combined_lookup_map: Dict[str, list] = {}
            combined_kanji_entries: Dict[str, dict] = {}
            combined_deconj_rules: list[dict] = []
            self.entry_sources = {}

            next_entry_id = 1
            enabled_sources = [s for s in sources if s.get('enabled', True)]
            if not enabled_sources:
                enabled_sources = self._default_dictionary_sources()

            n_sources = len(enabled_sources)

            for source_index, source in enumerate(sorted(enabled_sources, key=lambda x: int(x.get('priority', 0)))):
                path = source.get('path', '')
                if not path or not os.path.exists(path):
                    logger.warning("Dictionary source '%s' missing path '%s'; skipping.", source.get('name'), path)
                    continue

                if progress_cb:
                    progress_cb(source_index, n_sources, f"Loading '{source.get('name', 'Dictionary')}'...")

                # Use cached Dictionary if the file hasn't changed — avoids
                # re-reading and re-unpickling the full file (e.g. the 65 MB main
                # dict) on every import or save that touches dictionaries.
                try:
                    mtime = os.path.getmtime(path)
                except OSError:
                    mtime = 0.0
                cached = self._dict_file_cache.get(path)
                if cached and cached[0] == mtime:
                    dictionary = cached[1]
                    logger.debug("Using cached dictionary for '%s'", source.get('name'))
                else:
                    dictionary = Dictionary()
                    if not dictionary.load_dictionary(path):
                        logger.warning("Failed to load dictionary source '%s' from '%s'.", source.get('name'), path)
                        continue
                    self._dict_file_cache[path] = (mtime, dictionary)

                # Cache source metadata once — reused for every entry.
                source_meta = {
                    'dictionary_name':     source.get('name', 'Dictionary'),
                    'dictionary_id':       source.get('id', ''),
                    'dictionary_priority': source_index,
                }

                # Build id_map in one comprehension and assign sequential global IDs.
                entries_items = list(dictionary.entries.items())
                id_start = next_entry_id
                id_map = {old_id: id_start + i for i, (old_id, _) in enumerate(entries_items)}
                next_entry_id += len(entries_items)

                for i, (old_id, senses) in _throttled(enumerate(entries_items)):
                    combined_entries[id_start + i] = senses
                    self.entry_sources[id_start + i] = source_meta

                for i, (surface, map_entries) in _throttled(enumerate(dictionary.lookup_map.items())):
                    bucket = combined_lookup_map.setdefault(surface, [])
                    for map_entry in map_entries:
                        new_eid = id_map.get(map_entry[ENTRY_ID_INDEX])
                        if new_eid is None:
                            continue
                        bucket.append((
                            map_entry[WRITTEN_FORM_INDEX],
                            map_entry[READING_INDEX],
                            map_entry[FREQUENCY_INDEX],
                            new_eid,
                        ))

                if not combined_kanji_entries and dictionary.kanji_entries:
                    combined_kanji_entries = dictionary.kanji_entries

                if not combined_deconj_rules and dictionary.deconjugator_rules:
                    combined_deconj_rules = dictionary.deconjugator_rules

                if progress_cb:
                    progress_cb(source_index + 1, n_sources, "")

            self.dictionary.entries = combined_entries
            self.dictionary.lookup_map = combined_lookup_map
            self.dictionary.kanji_entries = combined_kanji_entries
            self.primary_kanji_entries = combined_kanji_entries
            self.dictionary.deconjugator_rules = combined_deconj_rules or []
            self.dictionary._is_loaded = True

            if not self.dictionary.deconjugator_rules:
                try:
                    fallback = Dictionary()
                    if fallback.load_dictionary('dictionary.pkl'):
                        self.dictionary.deconjugator_rules = fallback.deconjugator_rules
                except Exception:
                    pass

            self.deconjugator = Deconjugator(self.dictionary.deconjugator_rules)
            self.clear_cache()

    def run(self):
        logger.debug("Lookup thread started.")
        while self.shared_state.running:
            try:
                hit_result = self.shared_state.lookup_queue.get()
                if not self.shared_state.running: 
                    break
                logger.debug("Lookup: Triggered")

                current_lookup_string = self._extract_lookup_string(hit_result)
                last_lookup_string = self._extract_lookup_string(self.last_hit_result)


                self.last_hit_result = hit_result

                # skip lookup if lookup string didnt change
                if current_lookup_string == last_lookup_string:
                    continue
                

                lookup_result = self.lookup(current_lookup_string) if current_lookup_string else None
                # Pass context to popup if supported
                try:
                    self.popup_window.set_latest_data(lookup_result, hit_result if isinstance(hit_result, dict) else None)
                except TypeError:
                    self.popup_window.set_latest_data(lookup_result)
            except Exception:
                logger.exception("An unexpected error occurred in the lookup loop. Continuing...")
        logger.debug("Lookup thread stopped.")

    def _extract_lookup_string(self, hit_result: Any) -> Optional[str]:
        if not hit_result:
            return None
        if isinstance(hit_result, dict):
            return hit_result.get("lookup_string")
        if isinstance(hit_result, str):
            return hit_result
        return None

    def lookup(self, lookup_string: str) -> List:
        if not lookup_string:
            return []
        logger.info(f"Looking up: {lookup_string}")

        # Fast path: clean the text
        text = lookup_string.strip()
        text = text[:config.max_lookup_length]
        for i, ch in enumerate(text):
            if ch in JAPANESE_SEPARATORS:
                text = text[:i]
                break
        if not text:
            return []

        # Fast path: cache check (most important optimization)
        with self._dict_lock:
            if text in self.lookup_cache:
                self.lookup_cache.move_to_end(text)
                return self.lookup_cache[text]

        # Choose lookup method based on availability (cache the availability check)
        with self._dict_lock:
            results = self._fast_lookup(text)

        # Append kanji entry (cheap operation)
        if config.show_kanji and KANJI_REGEX.match(text[0]):
            kd = self.primary_kanji_entries.get(text[0])
            if kd:
                results.append(KanjiEntry(
                    character=kd['character'],
                    meanings=kd['meanings'],
                    readings=kd['readings'],
                    components=kd.get('components', []),
                    examples=kd.get('examples', []),
                ))

        # Cache results
        with self._dict_lock:
            self.lookup_cache[text] = results
            if len(self.lookup_cache) > self.CACHE_SIZE:
                self.lookup_cache.popitem(last=False)
        return results

    def _fast_lookup(self, text: str) -> List:
        """
        Optimized lookup that always uses local dictionaries and optionally
        appends Yomitan API results.
        """
        results = self._do_lookup(text)

        # Check if Yomitan is usable (cached result)
        if self._yomitan_enabled:
            if self._yomitan_available is None:
                # First time - check connection (one-time cost)
                try:
                    client = self.yomitan_client
                    self._yomitan_available = client is not None and client.check_connection()
                    if not self._yomitan_available:
                        logger.debug("Yomitan not available, falling back to local dictionary")
                except Exception:
                    self._yomitan_available = False
                    self._yomitan_enabled = False
            
            if self._yomitan_available:
                yomitan_entries = self._lookup_yomitan_optimized(text)
                for entry in yomitan_entries:
                    if hasattr(entry, 'dictionary_name'):
                        entry.dictionary_name = entry.dictionary_name or 'Yomitan API'
                    else:
                        entry.dictionary_name = 'Yomitan API'
                    if hasattr(entry, 'dictionary_id'):
                        entry.dictionary_id = entry.dictionary_id or 'yomitan-api'
                    else:
                        entry.dictionary_id = 'yomitan-api'
                results.extend(yomitan_entries)

        return results[:MAX_DICT_ENTRIES]

    def _lookup_yomitan_optimized(self, lookup_string: str) -> List[Any]:
        """
        Optimized Yomitan lookup with:
        - Early exit on perfect match
        - Minimal overhead
        - Match length tracking
        """
        if not self.yomitan_client:
            return []

        found_entries = []
        seen_keys = set()
        
        # Try exact match first (fastest path)
        exact_entries = self.yomitan_client.lookup(lookup_string) or []
        if exact_entries:
            for entry in exact_entries:
                key = (entry.written_form, entry.reading)
                if key not in seen_keys:
                    entry.match_len = len(lookup_string)
                    seen_keys.add(key)
                    found_entries.append(entry)
            # If we got an exact match, return immediately (no need to try shorter prefixes)
            if found_entries:
                return found_entries

        # No exact match - try decreasing lengths
        # Start from shorter length to avoid redundant work
        max_prefix_len = min(len(lookup_string) - 1, 20)  # Limit search depth
        for prefix_len in range(max_prefix_len, 0, -1):
            prefix = lookup_string[:prefix_len]
            entries = self.yomitan_client.lookup(prefix) or []
            if entries:
                for entry in entries:
                    key = (entry.written_form, entry.reading)
                    if key not in seen_keys:
                        entry.match_len = prefix_len
                        seen_keys.add(key)
                        found_entries.append(entry)
                
                # Stop after finding matches (Yomitan usually returns best matches first)
                if found_entries and (len(lookup_string) - prefix_len) > 3:
                    break

        return found_entries

    def _do_lookup(self, text: str) -> List[DictionaryEntry]:
        collected: Dict[int, Tuple[tuple, Form, int]] = {}
        found_primary_match = False
        first_match_len = 0
        # Per-call cache for _get_map_entries — avoids repeated hira/kata
        # conversions when the same form text appears across multiple forms.
        _map_cache: Dict[str, List] = {}

        def get_entries(form_text: str) -> List:
            hit = _map_cache.get(form_text)
            if hit is not None:
                return hit
            result = self._get_map_entries(form_text)
            _map_cache[form_text] = result
            return result

        for prefix_len in range(len(text), 0, -1):
            # Skip this prefix if we already have a longer match and this prefix
            # is too short to be meaningful — but only when the first match itself
            # was longer than 1 char, so single-char words (particles, etc.) still
            # appear when they are the only result.
            if (found_primary_match
                    and first_match_len > 1
                    and prefix_len < max(first_match_len - 2, 2)):
                break

            prefix = text[:prefix_len]

            forms = self.deconjugator.deconjugate(prefix)
            forms.add(Form(text=prefix))

            prefix_hits = []

            for form in forms:
                map_entries = get_entries(form.text)
                if not map_entries:
                    continue

                for map_entry in map_entries:
                    written = map_entry[WRITTEN_FORM_INDEX]
                    entry_id = map_entry[ENTRY_ID_INDEX]

                    if written is None and KANJI_REGEX.search(form.text):
                        logger.warning(f"Skipping malformed dictionary entry: kanji key '{form.text}'")
                        continue

                    if form.tags:
                        required_pos = form.tags[-1]
                        entry_senses = self.dictionary.entries.get(entry_id, [])
                        all_pos = {p for s in entry_senses for p in s['pos']}
                        if required_pos not in all_pos:
                            continue

                    if found_primary_match and not KANJI_REGEX.search(prefix):
                        if written and KANJI_REGEX.search(written):
                            continue

                    prefix_hits.append((map_entry, form))

            if prefix_hits:
                if not found_primary_match:
                    found_primary_match = True
                    first_match_len = prefix_len

                for map_entry, form in prefix_hits:
                    entry_id = map_entry[ENTRY_ID_INDEX]
                    if entry_id not in collected:
                        collected[entry_id] = (map_entry, form, prefix_len)

        return self._format_and_sort(list(collected.values()), text)

    def _get_map_entries(self, text: str) -> List[tuple]:
        result = self.dictionary.lookup_map.get(text, [])
        if result:
            return list(result)
        kata = self._hira_to_kata(text)
        if kata != text:
            result = self.dictionary.lookup_map.get(kata, [])
            if result:
                return list(result)
        hira = self._kata_to_hira(text)
        if hira != text:
            result = self.dictionary.lookup_map.get(hira, [])
            if result:
                return list(result)
        return []

    def _format_and_sort(
        self,
        raw: List[Tuple[tuple, Form, int]],
        original_lookup: str,
    ) -> List[DictionaryEntry]:
        merged: Dict[Tuple[str, str, str], dict] = {}

        for map_entry, form, match_len in raw:
            written = map_entry[WRITTEN_FORM_INDEX]
            reading = map_entry[READING_INDEX] or ''
            freq = map_entry[FREQUENCY_INDEX]
            entry_id = map_entry[ENTRY_ID_INDEX]
            source_meta = self.entry_sources.get(entry_id, {})
            dictionary_name = source_meta.get('dictionary_name', 'Dictionary')
            dictionary_id = source_meta.get('dictionary_id', '')
            dictionary_priority = int(source_meta.get('dictionary_priority', 9999))

            entry_senses = self.dictionary.entries.get(entry_id, [])
            priority = self._calculate_priority(written, freq, form, match_len, original_lookup)

            key = (written, reading, dictionary_id or dictionary_name)
            if key not in merged:
                merged[key] = {
                    'id': entry_id,
                    'written_form': written,
                    'reading': reading,
                    'senses': list(entry_senses),
                    'freq': freq,
                    'deconjugation_process': form.process,
                    'priority': priority,
                    'match_len': match_len,
                    'dictionary_name': dictionary_name,
                    'dictionary_id': dictionary_id,
                    'dictionary_priority': dictionary_priority,
                }
            else:
                cur = merged[key]
                if entry_id != cur['id']:
                    cur['senses'].extend(entry_senses)
                if priority > cur['priority']:
                    cur['priority'] = priority
                    cur['id'] = entry_id
                    cur['deconjugation_process'] = form.process
                if freq < cur['freq']:
                    cur['freq'] = freq
                if match_len > cur['match_len']:
                    cur['match_len'] = match_len

        # Group entries by (written_form, reading), showing all enabled dicts.
        # Users control which dictionaries appear via the enable/disable toggles
        # in Settings → Dictionaries — no artificial per-word cap needed.
        word_groups: Dict[Tuple[str, str], List[dict]] = defaultdict(list)
        for entry in merged.values():
            word_key = (entry['written_form'], entry['reading'])
            word_groups[word_key].append(entry)

        processed_groups = []
        for entries in word_groups.values():
            entries.sort(key=lambda x: x['dictionary_priority'])
            # Rank this word group by the best priority available across ALL loaded
            # dictionaries — not just the first-listed one.  This ensures that
            # main dictionary.pkl frequency data (JPDB ranks) is always used for
            # sorting 歩く vs 歩き etc. regardless of where the user placed it in
            # the dictionary order settings.
            rank_entry = max(entries, key=lambda x: (x['match_len'], x['priority']))
            processed_groups.append((-rank_entry['match_len'], -rank_entry['priority'], entries[0]['dictionary_priority'], entries))

        processed_groups.sort(key=lambda x: (x[0], x[1], x[2]))

        results = []
        for _ml, _pr, _dp, entries in processed_groups:
            for d in entries:
                results.append(DictionaryEntry(
                    id=d['id'],
                    written_form=d['written_form'],
                    reading=d['reading'],
                    senses=d['senses'],
                    freq=d['freq'],
                    deconjugation_process=d['deconjugation_process'],
                    priority=d['priority'],
                    match_len=d['match_len'],
                    dictionary_name=d['dictionary_name'],
                    dictionary_id=d['dictionary_id'],
                ))
                if len(results) >= MAX_DICT_ENTRIES:
                    return results
        return results

    def _calculate_priority(
        self,
        written_form: str,
        freq: int,
        form: Form,
        match_len: int,
        original_lookup: str,
    ) -> float:
        priority = float(match_len)

        if freq < DEFAULT_FREQ:
            priority += 10.0 * (1.0 - math.log(freq) / math.log(DEFAULT_FREQ))

        original_is_kana = not KANJI_REGEX.search(original_lookup)
        written_is_kana = not KANJI_REGEX.search(written_form) if written_form else True

        if original_is_kana:
            if written_is_kana and not form.process:
                priority += 3.0

        priority -= len(form.process)
        return priority

    def _hira_to_kata(self, text: str) -> str:
        res = []
        for c in text:
            code = ord(c)
            res.append(chr(code + 0x60) if 0x3041 <= code <= 0x3096 else c)
        return ''.join(res)

    def _kata_to_hira(self, text: str) -> str:
        res = []
        for c in text:
            code = ord(c)
            if 0x30A1 <= code <= 0x30F6:
                res.append(chr(code - 0x60))
            elif code == 0x30FD:
                res.append('\u309D')
            elif code == 0x30FE:
                res.append('\u309E')
            else:
                res.append(c)
        return ''.join(res)
