# Screenshot to Word

A lightweight Windows desktop utility that detects clipboard images and appends them to the active Microsoft Word desktop document.

> **Status:** early development (`0.1.0`). Test with a copy of your document first. This release is Windows-only and requires desktop Microsoft Word.

## Features in this starter release

- Start/stop clipboard image monitoring from a small GUI.
- Detect clipboard changes using the Windows clipboard sequence counter.
- Append clipboard images to the end of the active Word document.
- Keep images embedded in the document rather than linked to temporary files.
- Basic duplicate protection for repeated clipboard updates.

## Requirements

- Windows 10 or Windows 11
- Desktop Microsoft Word installed and running
- Python 3.10+ for development

## Run from source

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
python -m screenshot_to_word
```

Open the target Word document before starting monitoring. In Snipping Tool, ensure captures are copied to the clipboard. Take a screenshot; the app should append it to the active Word document. Save the document normally in Word.

## Build a Windows executable locally

```powershell
python -m pip install pyinstaller
pyinstaller --noconfirm --clean --windowed --name ScreenshotToWord --paths src --collect-all PySide6 src/screenshot_to_word/__main__.py
```

The executable is created under `dist\ScreenshotToWord\` (one-folder build). The Inno Setup script in `installer/` can package that folder into an installer.

## Project layout

```text
ScreenshotToWord/
├── src/screenshot_to_word/
│   ├── __init__.py
│   ├── __main__.py
│   ├── app.py
│   ├── clipboard_monitor.py
│   ├── main_window.py
│   └── word_writer.py
├── installer/ScreenshotToWord.iss
├── .github/workflows/build-windows.yml
├── requirements.txt
├── pyproject.toml
├── LICENSE
└── README.md
```

## Known limitations

- Windows + desktop Word only; Word in a browser and macOS/Linux are not supported.
- The app processes clipboard images, not screenshots saved only to disk.
- It writes to the **active document** in Word when each screenshot is processed. Keep the intended document active and avoid switching to another document while monitoring.
- It does not automatically save the document.
- This initial release has no floating heading bar, configurable image sizing, installer signing, or automatic update system.
- Clipboard behavior varies across Snipping Tool settings and Windows builds; test on your machine before relying on it for important work.

## Privacy

The app processes clipboard images locally and does not upload screenshots to a server. Avoid capturing sensitive information you do not intend to put in the document.

## License

MIT. See [LICENSE](LICENSE).
