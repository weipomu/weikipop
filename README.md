# weikipop

weikipop is a desktop Japanese OCR lookup tool built on top of the original Meikipop project line.

It continuously or manually scans a screen region, performs OCR, and shows dictionary lookups in a popup. It also supports adding cards to Anki via AnkiConnect.

<img width="1695" height="941" alt="image" src="https://github.com/user-attachments/assets/a6105d75-5556-4ea0-8eae-0a394fa52e3c" />
<img width="369" height="439" alt="image" src="https://github.com/user-attachments/assets/7cbc8a7e-f1b1-4822-b8c8-cbba22de7740" />

## Features

- Fast screen-region OCR with multiple OCR backends
- Dictionary lookup with deconjugation and kana/kanji handling, scrollable
- Multi-dictionary import support (`.zip` Yomitan and `.pkl`)
- Dictionary enable/disable + priority ordering from settings
- Optional Yomitan API integration
- AnkiConnect export with configurable field mapping
- Local mining log (`data/mining_log.jsonl`) for SRS workflows
- Global shortcuts and tray-based settings
- Cross-platform support (Windows/Linux/macOS)

## Installation

### Option A: prebuilt binaries

1. Open the latest release on GitHub.
2. Download the package for your OS.
3. Launch the executable.

### Option B: run from source

Prerequisites:

- Python 3.10+
- `pip`

Setup:

1. Clone the repository.
2. Install dependencies:
	- `pip install -r requirements.txt`
3. Run the app:
	- `python -m src.main`

## Configuration

Configuration is stored in `config.ini` at runtime and auto-created when settings are saved.

For an initial baseline, copy `config.example.ini` to `config.ini` and adjust values as needed.

If you use the **Google Lens (remote)** provider, set:

- `WEIKIPOP_GLENS_API_KEY=<your_api_key>`

## Usage

- Start the app.
- Use the configured hotkey over Japanese text to trigger lookup.
- Right-click the tray icon to:
  - select OCR provider
  - choose scan mode/region
  - open settings
- In Settings → Dictionaries:
  - import dictionaries
  - reorder dictionary priority
  - enable/disable dictionaries
- Use `Scroll Popup` shortcut (default `Alt+Wheel`) to scroll long lookup popups.

## Development workflow

- Main application entrypoint: `src/main.py`
- OCR providers: `src/ocr/providers/`
- Dictionary pipeline: `src/dictionary/`
- UI: `src/gui/`

Recommended checks before opening a PR:

- `python -m compileall src`
- Run the app and verify tray + popup flow

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for contribution standards.

## License

This project is licensed under GPL-3.0. See [LICENSE](LICENSE).

## Credits

- [rtr46](https://github.com/rtr46) for the original Meikipop project, didn't properly fork when making this project and cannot change it now, really sorry!
- [zurcGH](https://github.com/zurcGH) for Meikipop-Anki lineage
- [kqq](https://github.com/user-attachments/assets/8d1ebb4e-daeb-4bae-93cc-9ae4a78751df) for being my first initial tester!




