import json
import pickle
import re
import time
import zipfile
from html import escape
from collections import defaultdict
from pathlib import Path
from typing import Optional, Tuple

from src.dictionary.structured_content import handle_structured_content

DEFAULT_FREQ = 999_999
ID_NAMESPACE = 10_000_000


def _throttled(iterable, work_ms: float = 2.0, sleep_ms: float = 8.0):
    """Yield items while capping CPU to ~20% — work for 2ms then sleep 8ms."""
    work_s  = work_ms  / 1000.0
    sleep_s = sleep_ms / 1000.0
    t = time.monotonic()
    for item in iterable:
        yield item
        if time.monotonic() - t >= work_s:
            time.sleep(sleep_s)
            t = time.monotonic()


def _extract_text(node) -> str:
    if node is None:
        return ''
    if isinstance(node, str):
        return node
    if isinstance(node, list):
        return ''.join(_extract_text(child) for child in node)
    if isinstance(node, dict):
        tag = node.get('tag', '')
        content = node.get('content')
        if tag == 'ruby':
            if isinstance(content, list):
                out = []
                for child in content:
                    if isinstance(child, dict) and child.get('tag') == 'rt':
                        continue
                    out.append(_extract_text(child))
                return ''.join(out)
            return _extract_text(content)
        if tag == 'rt':
            return ''
        text = _extract_text(content)
        if tag in {'div', 'li', 'tr', 'br', 'p'}:
            return text + ' '
        return text
    return ''


def _extract_glosses(definitions: list) -> list[str]:
    glosses: list[str] = []
    for definition in definitions:
        if isinstance(definition, str):
            text = definition.strip()
            if text:
                glosses.append(escape(text))
        elif isinstance(definition, dict):
            def_type = definition.get('type')
            if def_type == 'text':
                text = (definition.get('text') or '').strip()
                if text:
                    glosses.append(escape(text))
            elif def_type == 'structured-content':
                rendered = handle_structured_content(definition)
                for html_fragment in rendered:
                    fragment = (html_fragment or '').strip()
                    if fragment:
                        glosses.append(fragment)
    return glosses


def _parse_freq_value(freq_data) -> Optional[int]:
    if isinstance(freq_data, (int, float)):
        return int(freq_data)
    if isinstance(freq_data, str):
        try:
            return int(freq_data)
        except ValueError:
            return None
    if isinstance(freq_data, dict):
        if 'value' in freq_data:
            return int(freq_data['value'])
        inner = freq_data.get('frequency')
        if inner is not None:
            return _parse_freq_value(inner)
    return None


def _load_freq_map_from_zip(zf: zipfile.ZipFile) -> dict:
    freq: dict[tuple[str, str], int] = {}
    for name in sorted(zf.namelist()):
        if not re.match(r'term_meta_bank_\d+\.json', Path(name).name):
            continue
        with zf.open(name) as file:
            rows = json.load(file)
        for row in rows:
            if len(row) < 3 or row[1] != 'freq':
                continue
            term = row[0]
            raw = row[2]
            reading = ''
            if isinstance(raw, dict) and 'reading' in raw:
                reading = raw['reading']
                rank = _parse_freq_value(raw.get('frequency'))
            else:
                rank = _parse_freq_value(raw)
            if rank is None:
                continue
            key = (term, reading)
            if key not in freq or rank < freq[key]:
                freq[key] = rank
    return freq


def _has_kanji(text: str) -> bool:
    return any(0x4E00 <= ord(c) <= 0x9FFF for c in text)


def convert_yomitan_zip_to_payload(zip_path: str, dict_index: int = 0) -> Tuple[dict, str]:
    with zipfile.ZipFile(zip_path, 'r') as zf:
        dict_title = Path(zip_path).stem
        if 'index.json' in zf.namelist():
            with zf.open('index.json') as file:
                index_meta = json.load(file)
            dict_title = index_meta.get('title') or dict_title

        freq_map = _load_freq_map_from_zip(zf)

        rows = []
        for name in sorted(zf.namelist()):
            if re.match(r'term_bank_\d+\.json', Path(name).name):
                with zf.open(name) as file:
                    rows.extend(json.load(file))

    seq_groups = defaultdict(list)
    standalone_seq = -1
    for i, row in _throttled(enumerate(rows)):
        if len(row) < 6:
            continue
        seq = row[6] if len(row) > 6 else 0
        if seq == 0:
            seq_groups[standalone_seq].append(row)
            standalone_seq -= 1
        else:
            seq_groups[seq].append(row)

    entries = {}
    lookup_map = defaultdict(list)
    id_base = dict_index * ID_NAMESPACE

    def freq_for(term: str, reading: str) -> int:
        return freq_map.get((term, reading), freq_map.get((term, ''), DEFAULT_FREQ))

    for j, (seq, group_rows) in _throttled(enumerate(seq_groups.items())):
        if seq < 0:
            entry_id = id_base + (ID_NAMESPACE + seq)
        else:
            entry_id = id_base + seq

        first_row = group_rows[0]
        canonical_term = first_row[0]
        canonical_reading = first_row[1]

        senses = []
        for row in group_rows:
            def_tags_str = row[2] if len(row) > 2 else ''
            rules_str = row[3] if len(row) > 3 else ''
            definitions = row[5] if len(row) > 5 else []
            term_tags_str = row[7] if len(row) > 7 else ''

            glosses = _extract_glosses(definitions)
            if not glosses:
                continue

            all_tag_strings = (def_tags_str + ' ' + term_tags_str).split()
            tags = [t for t in all_tag_strings if t]
            pos = [r for r in rules_str.split() if r]

            senses.append({
                'glosses': glosses,
                'pos': pos,
                'tags': tags,
                'source': dict_title,
            })

        if not senses:
            continue

        entries[entry_id] = senses

        seen_terms = set()
        seen_readings = set()

        for row in group_rows:
            term = row[0]
            reading = row[1]

            if _has_kanji(term) and term not in seen_terms:
                seen_terms.add(term)
                display_reading = reading if reading else canonical_reading
                freq = freq_for(term, display_reading)
                lookup_map[term].append((canonical_term, display_reading, freq, entry_id))

            kana_surface = reading if reading else term
            if kana_surface not in seen_readings:
                seen_readings.add(kana_surface)
                if reading:
                    freq = freq_for(kana_surface, reading)
                    lookup_map[kana_surface].append((canonical_term, reading, freq, entry_id))
                else:
                    freq = freq_for(term, '')
                    lookup_map[term].append((term, None, freq, entry_id))

    payload = {
        'entries': entries,
        'lookup_map': dict(lookup_map),
        'kanji_entries': {},
        'deconjugator_rules': [],
    }
    return payload, dict_title


def write_payload_pickle(payload: dict, output_path: str) -> None:
    with open(output_path, 'wb') as file:
        pickle.dump(payload, file, protocol=pickle.HIGHEST_PROTOCOL)
