cat << 'EOF' > README.md
# Virtual Bookkeeper AI 🗂️✨

An offline, privacy-first desktop application that organizes messy folders (like `~/Downloads`) into structured directories and functional subfolders using local LLMs (via Ollama).

Built with Python, Tkinter, and Apple Silicon-optimized inference engines.

---

## Key Features

- **100% Local & Private:** No cloud APIs. All inference runs locally via Ollama (`llama3.2:3b`).
- **Zero-Setup Runtime Manager:** Automatically detects if Ollama is running, starts the daemon if inactive, and auto-provisions the required model weights on first launch.
- **Two-Tier Triage Pipeline:**
  1. **Filename-First:** Evaluates filenames for clear semantic meaning to skip unnecessary document parsing.
  2. **Deep Content Extraction:** Parses PDFs (`pypdf`), Word docs (`docx`), Excel spreadsheets (`openpyxl`), PowerPoint decks (`python-pptx`), Markdown, and text files when filenames are ambiguous.
- **Custom Parent Groups & Dynamic Subfolders:** Define your own high-level taxonomy (e.g., *Work & Projects*, *Financial & Tax*) and let the model create contextual subfolders (*Invoices*, *Engineering Resumes*, *Q3 Reviews*).
- **Non-Destructive & Reversible:**
  - Automated collision handling (`filename_1.ext`) prevents silent file overwrites.
  - Persistent on-disk manifest (`.custom_organizer_history.json`) enables one-click full rollback anytime.

---

## Getting Started

### Prerequisites

- macOS (Apple Silicon M-series recommended)
- Python 3.10+

### Installation & Execution

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/YOUR_USERNAME/virtual-bookkeeper-ai.git](https://github.com/YOUR_USERNAME/virtual-bookkeeper-ai.git)
   cd virtual-bookkeeper-ai