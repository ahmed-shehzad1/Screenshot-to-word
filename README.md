# Screenshot to Word

A lightweight Windows desktop utility for appending clipboard screenshots to the active Microsoft Word desktop document.

## Features

- **Automatic insertion:** start monitoring and each new clipboard screenshot is appended to Word.
- **Manual insertion:** click **Insert Screenshot Now** or use the floating quick bar.
- **Floating heading bar:** keep a small always-on-top bar visible, type a task heading, and insert it as Word's Heading 2 style.
- **Start, stop, pause, and resume:** control automatic monitoring separately from manual insertion.
- **Duplicate protection:** rapid repeated clipboard updates for the same image are ignored.
- **Settings:** configure clipboard polling interval, maximum image width, and error notifications.
- **Local processing:** screenshots are handled on your computer and are not uploaded to a server.
- **Windows installer workflow:** packaging scripts are included for a future public release.

## Requirements

- Windows 10 or Windows 11
- Microsoft Word desktop application installed and running
- Python 3.10+ for development (Python 3.11 is used by the GitHub Actions build)

## Run from source

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
python -m screenshot_to_word
```

Open the target Word document first. Start automatic insertion to monitor new clipboard images, or use **Insert Screenshot Now** for manual insertion. The floating heading bar can insert a task heading into the active document. Save your document normally in Word.

## Build a Windows executable locally

```powershell
python -m pip install pyinstaller
pyinstaller --noconfirm --clean --windowed --name ScreenshotToWord --paths src --collect-all PySide6 src/screenshot_to_word/__main__.py
```

The executable folder is created under `dist\ScreenshotToWord\`. To create a setup installer, install Inno Setup 6 and compile `installer/ScreenshotToWord.iss` after building the executable folder.

## Important notes

- Windows and desktop Microsoft Word only; Word in a browser, macOS, and Linux are not supported.
- The app inserts the image currently available on the Windows clipboard. Make sure Snipping Tool copies captures to the clipboard.
- The app writes to the **active document** in Word when a screenshot or heading is inserted. Keep the intended document active.
- The app does not automatically save your Word document.
- Test on a copy of an important document before relying on it for coursework.

## Privacy

The app processes clipboard images locally and does not upload screenshots to a server. Be mindful of sensitive content visible in screenshots.

## License

MIT. See [LICENSE](LICENSE).
