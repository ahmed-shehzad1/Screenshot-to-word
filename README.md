# Screenshot to Word

### Turn screenshots into organized Word documentation — automatically.

Screenshot to Word is a Windows desktop application that makes documenting practical work, programming tasks, tutorials, and university lab assignments easier.

Capture a screenshot using Windows Snipping Tool, and Screenshot to Word can automatically insert it into your currently open Microsoft Word document. No more repeatedly switching between windows and manually copying and pasting images.

<p align="center">
  <a href="https://github.com/ahmed-shehzad1/Screenshot-to-word/releases/latest">
    <img src="https://img.shields.io/badge/Download-Windows%20Installer-2ea44f?style=for-the-badge&logo=windows&logoColor=white" alt="Download for Windows">
  </a>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011-blue?style=flat-square">
  <img src="https://img.shields.io/badge/Version-0.2.1-informational?style=flat-square">
  <img src="https://img.shields.io/badge/License-MIT-yellow?style=flat-square">
  <img src="https://img.shields.io/badge/Python-PySide6-3776AB?style=flat-square&logo=python&logoColor=white">
</p>

---

## ✨ Features

| Feature | Description |
|---|---|
| **Automatic Screenshot Insertion** | Automatically detects new clipboard screenshots and appends them to your active Word document. |
| **Manual Screenshot Insertion** | Insert screenshots whenever you want using the manual insertion control. |
| **Floating Heading Bar** | Add headings to your Word document through a small floating toolbar without repeatedly switching windows. |
| **Always-on-Top Toolbar** | Keep the floating heading bar accessible while working in other applications. |
| **Monitoring Controls** | Start, stop, pause, and resume automatic screenshot monitoring. |
| **Duplicate Protection** | Prevents rapid repeated clipboard updates for the same screenshot from being inserted multiple times. |
| **Customizable Settings** | Configure clipboard polling, maximum image width, and error notifications. |
| **Local Processing** | Screenshots are processed locally on your computer, not uploaded to a remote server. |

## 📥 Download and Install

You don't need Python, VS Code, or any development environment to use Screenshot to Word.

**1. Download**

Visit the latest GitHub Release:

[Download Screenshot to Word for Windows](https://github.com/ahmed-shehzad1/Screenshot-to-word/releases/latest)

**2. Install**

- Download `ScreenshotToWord-Setup-0.2.1.exe`.
- Run the installer.
- Follow the installation instructions.

**3. Start documenting**

- Open Microsoft Word and your target document.
- Launch Screenshot to Word.
- Start automatic screenshot monitoring.
- Capture screenshots using Windows Snipping Tool.
- Your screenshots will be inserted into the active Word document.

You can also use manual insertion or the floating heading bar whenever needed.

## 🖥️ System Requirements

- Windows 10 or Windows 11 (64-bit).
- Microsoft Word desktop application installed.
- Windows clipboard access for screenshot detection.

**Note:** Microsoft Word desktop is required. Word Online, macOS, and Linux are not currently supported.

## ⚙️ How It Works

Screenshot to Word works alongside Microsoft Word and the Windows clipboard.

1. Capture a screenshot using Windows Snipping Tool.
2. The application detects the new image on your clipboard.
3. It communicates with Microsoft Word to insert the image into the active document.
4. Continue capturing screenshots, and they are appended to your documentation.

The application also lets you insert headings to organize your work into clearly separated sections.

## 🛠️ Built With

- **Python** – Core application logic and automation.
- **PySide6 / Qt** – Desktop graphical user interface.
- **Pillow** – Image processing.
- **PyWin32** – Windows and Microsoft Word automation.
- **PyInstaller** – Packaging the application into a standalone executable.
- **Inno Setup** – Creating the Windows installer.

## 👨‍💻 Run from Source

If you want to explore the code or contribute to development:

**Prerequisites**
- Python 3.13 (64-bit recommended for the included build script).
- Windows 10 or Windows 11.
- Microsoft Word desktop.

**Setup**

```powershell
# Clone the repository
git clone https://github.com/ahmed-shehzad1/Screenshot-to-word.git

# Enter the project directory
cd Screenshot-to-word

# Create a virtual environment
py -m venv .venv

# Activate the environment
.\.venv\Scripts\Activate.ps1

# Install the application
python -m pip install --upgrade pip
python -m pip install -e .

# Run the application
python -m screenshot_to_word
```

## 📦 Build the Windows Installer

The repository includes the scripts and configuration needed to build the Windows installer.

**Requirements**
- Python 3.13 (64-bit).
- Inno Setup 6 or 7.

**Build**

1. Clone or download the repository.
2. Install the required dependencies.
3. Run `build.bat` from the project directory.
4. Wait for the build process to finish.
5. Find the generated installer inside the `release` folder.

The build bundles the Python runtime, application dependencies, graphical icon, and required visual assets. End users do not need a separate Python installation.

## 🔒 Privacy and Important Notes

- Clipboard screenshots are processed locally.
- Screenshots are not uploaded to an external server by the application.
- The application inserts content into the currently active Word document.
- Microsoft Word does not automatically save your document after insertion. Save your work normally.
- Make sure Snipping Tool copies captures to the clipboard.
- Test the application with a copy of an important document before using it for critical work.

## 🤝 Contributing

Contributions, bug reports, and feature suggestions are welcome!

If you find an issue or have an idea for improving Screenshot to Word:

1. Open an [Issue](https://github.com/ahmed-shehzad1/Screenshot-to-word/issues).
2. Describe the problem or proposed improvement.
3. If you would like to contribute code, feel free to submit a Pull Request.

## ⭐ Support the Project

If Screenshot to Word saves you time or makes your documentation workflow easier, consider:

- Giving this repository a ⭐ Star.
- Following me on GitHub for future projects.
- Sharing the application with other students and developers.
- Reporting bugs and suggesting improvements.

Every bit of support helps the project grow!

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## 👨‍💻 About the Developer

Created by **[Ahmed Shahzad](https://github.com/ahmed-shehzad1)**.

I'm a student at Air University Islamabad who enjoys building software and practical tools that can make people's everyday work easier.

Thanks for checking out Screenshot to Word!

If you find this project useful, don't forget to star the repository and follow me on GitHub.

---

<p align="center">
  <b>Made with ❤️ to make documentation easier.</b>
</p>