> 📄 **Looking for the full Product Requirements Document?** Read the [Product Requirements Document (PRD) on our Wiki](https://github.com/VarunParikh/virtual-bookkeeper-ai/wiki).

# Virtual Bookkeeper AI 🗂️✨

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-macOS-lightgrey.svg)](https://apple.com)
[![Model](https://img.shields.io/badge/Model-llama3.2%3A3b-indigo.svg)](https://ollama.com)
[![PRD](https://img.shields.io/badge/Specs-PRD.md-emerald.svg)](PRD.md)

An offline, privacy-first desktop application that organizes messy directories (such as `~/Downloads` and `~/Desktop`) into structured taxonomies and dynamic, functional subfolders using on-device LLMs via Ollama.

Zero cloud API calls. Zero manual environment configuration. 100% private.

---

## 💾 Download Standalone App

Download the latest pre-built `.dmg` installer for macOS:

👉 **[Download VirtualBookkeeper-Installer.dmg](https://github.com/VarunParikh/virtual-bookkeeper-ai/releases/latest/download/VirtualBookkeeper-Installer.dmg)** *(v1.0.0)*

Drag and drop `VirtualBookkeeper.app` into your `/Applications` folder, launch it, and the app takes care of everything else automatically.

---

## 🚀 Key Features

* **100% Offline & Private:** No sensitive financial reports, tax filings, or contracts are ever sent to remote APIs. All inference runs locally via Ollama (`llama3.2:3b`).
* **Zero-Config Runtime Provisioning:** On first launch, the app automatically checks if Ollama is running. If missing or inactive, it starts the background daemon and auto-downloads the required model weights with a live progress bar.
* **Intelligent Two-Tier Triage Engine:**
  1. **Tier 1 (Filename-First):** Evaluates filenames for clear semantic meaning to skip unnecessary document parsing and minimize resource usage.
  2. **Tier 2 (Deep Content Inspection):** Performs deep text parsing across PDFs (`pypdf`), Word documents (`docx`), Excel sheets (`openpyxl`), PowerPoint presentations (`python-pptx`), Markdown, and text files when filenames are vague.
  3. **Optional Online Fallback:** If the filename is generic and local extraction returns no text, users may enable online lookup. The app sends only the basename (not its directory or file contents) to DuckDuckGo, uses returned public snippets as unverified context for local classification, and leaves the proposed move in the review stage. The option is off by default.


* **Custom Taxonomies & Dynamic Subfolders:** Define your high-level parent groups (e.g., *Work & Projects*, *Financial & Tax*) and let the local model synthesize contextual subfolders (*Invoices*, *Engineering Resumes*, *Q3 Reviews*).
* **Two-Step Review Before Execution:** Inspect all proposed file movements, category matches, and reasoning in a staging table before committing any changes.
* **Safe, Non-Destructive Operations:**
  * **Collision Prevention:** Automatically appends numeric suffixes (`file_1.pdf`) to prevent silent file overwriting.
  * **One-Click Undo:** Manifest-backed session logging (`.custom_organizer_history.json`) lets you reverse any organization run and prune empty folders.

---

## 📖 Product Specifications

For in-depth product requirements, target user personas, technical KPIs, and future roadmap items, see the [Product Requirements Document (PRD.md)](PRD.md).

---

## 🛠️ Local Development & Manual Installation

If you prefer running from source rather than the DMG installer:

### Prerequisites
- macOS 12.0+ (Apple Silicon recommended)
- Python 3.10+

### Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/VarunParikh/virtual-bookkeeper-ai.git
   cd virtual-bookkeeper-ai
   ```

2. **Set up a virtual environment:**
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application:**
   ```bash
   python organizer_custom_groups.py
   ```

---

## 📦 Building the macOS `.app` and `.dmg`

To package the project into a standalone macOS binary and generate a distributable disk image:

```bash
# 1. Compile bundle via PyInstaller
pyinstaller --noconfirm --onedir --windowed \
  --name "VirtualBookkeeper" \
  --collect-all pypdf \
  --collect-all docx \
  --hidden-import tkinter \
  organizer_custom_groups.py

# 2. Package into DMG with Applications shortcut
mkdir -p dmg_staging
cp -R dist/VirtualBookkeeper.app dmg_staging/
ln -s /Applications dmg_staging/Applications

hdiutil create -volname "VirtualBookkeeper" \
  -srcfolder dmg_staging \
  -ov -format UDZO \
  VirtualBookkeeper-Installer.dmg

rm -rf dmg_staging
```

---

## 📄 License

This project is licensed under the terms of the [MIT License](LICENSE).
