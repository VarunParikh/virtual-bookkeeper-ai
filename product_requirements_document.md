# Product Requirements Document (PRD)

**Product Name:** Virtual Bookkeeper AI  
**Document Version:** 1.0  
**Target Platform:** macOS (Apple Silicon & Intel)  
**Status:** Approved / Production  

---

## 1. Executive Summary & Problem Statement

### 1.1 Problem Statement
Users routinely accumulate hundreds of unsorted documents in folders like `~/Downloads` and `~/Desktop`—invoices, bank statements, tax IDs, resumes, contracts, and presentations.

Manual sorting is tedious, and existing solutions present two significant trade-offs:
* **Cloud Privacy Risks:** Cloud-based categorizers upload sensitive personal and financial documents to external third-party servers.
* **Rigid Static Rule Engines:** Traditional regex/extension-based rules fail when filenames are ambiguous (e.g., `IMG_0023.pdf`, `Invoice_Final.pdf`), lacking semantic document comprehension.

### 1.2 Proposed Solution
**Virtual Bookkeeper AI** is a zero-configuration, privacy-first macOS desktop application that analyzes document context locally using an on-device Large Language Model (via Ollama). It maps files into user-defined parent groups and dynamically creates contextual subfolders (e.g., `Financial & Tax/GST Invoices`, `Resumes & Careers/Engineering`) with zero data leaving the host machine.

---

## 2. Target Audience & Personas

| Persona | Description | Core Pain Point |
| :--- | :--- | :--- |
| **Independent Professional / Freelancer** | Self-employed individuals, agency founders, consultants. | Overwhelmed by client receipts, invoices, custom design briefs, and vendor statements scattered in Downloads. |
| **Corporate Knowledge Worker** | Product managers, engineers, corporate executives. | Manages high volumes of research reports, executive resumes, contract drafts, and presentation decks across multiple projects. |
| **Privacy-Conscious User** | Individuals managing tax documents, medical reports, and legal agreements. | Refuses to upload personal tax identification numbers, banking records, and legal agreements to public cloud tools. |

---

## 3. Product Goals & Success Metrics (KPIs)

### 3.1 Primary Goals
* **True Zero-Config Local AI:** Abstract the complexity of local model execution; the user never needs to touch Terminal, install Homebrew, or manually download model weights.
* **High Classification Accuracy with Fast Triage:** Strike an optimal balance between execution speed and semantic categorization quality.
* **Zero Data Loss Guarantee:** Strictly non-destructive file operations with deterministic collision resolution and atomic rollback.

### 3.2 Success Metrics (KPIs)
* **Time-to-First-Run (TTFR):** $\le 3\text{ minutes}$ from DMG mount to successful first folder classification (including background runtime setup).
* **Triage Latency:** $\le 1.2\text{ seconds}$ average classification latency per file on Apple Silicon hardware (M1/M2/M3/M4).
* **Classification Acceptance Rate:** $\ge 90\%$ of AI-suggested folder and subfolder assignments approved without user rollback.
* **Undo Reliability:** $100\%$ success rate on session reversal without orphan files or path collisions.

---

## 4. Key Assumptions & Constraints

* **Platform Scope:** macOS 12.0+ (Universal / Apple Silicon optimized).
* **Hardware Profile:** Minimum 8 GB Unified Memory (optimal on 16 GB+ for concurrent model execution).
* **Network Independence:** Once the lightweight runtime and initial model weights (`llama3.2:3b`, ~2.0 GB) are provisioned, the tool runs completely offline without internet connectivity.

---

## 5. Detailed Feature Specifications & Functional Requirements

### 5.1 Automated Runtime & Lifecycle Manager (Zero-Config)
* **Service Health Check:** On application boot, ping `localhost:11434/api/tags`.
* **State Resolution:**
  * *Service Active & Model Available:* Unlock UI immediately.
  * *Service Inactive but Binary Exists:* Launch daemon in the background (`open -a Ollama --background` or `ollama serve` detached).
  * *Binary Missing:* Automatically fetch Darwin universal distribution, unpack into `~/.local/bin/ollama`, set executable permissions (`0o755`), and launch daemon.
  * *Model Missing:* Stream download of target model (`llama3.2:3b`) with real-time percentage progress displayed in the primary GUI progress bar.

### 5.2 Two-Tier Document Triage Engine
* **Tier 1 (Semantic Filename Check):** Analyze file stem. If the filename possesses distinctive descriptive semantic value (e.g., `Q3_Product_Roadmap_Review.pdf`), bypass document parsing to conserve I/O cycles and memory.
* **Tier 2 (Deep Content Inspection):** If the filename is generic (e.g., `Scan_001.pdf`, `doc_12.docx`, `download.pdf`), extract text excerpts (first 2,000–2,500 characters) across supported MIME types:
  * PDF (`pypdf` with encrypted document bypass)
  * Word (`python-docx`)
  * Excel (`openpyxl` sheet text extraction)
  * PowerPoint (`python-pptx` text run aggregation)
  * Plaintext / CSV / Markdown (`utf-8` stream parsing)

### 5.3 Custom Taxonomy & Dynamic Subfolder Generation
* **Parent Group Management:** Users can add and remove top-level parent categories via the UI.
* **LLM Subfolder Synthesis:** Prompt local model using strict JSON schema output:
  * Force assignment to one of the active Parent Groups.
  * Synthesize a concise, contextual subfolder (1–3 words) based on functional domain (e.g., *Tax Invoices*, *Design Assets*, *Engineering*).
  * Generate a single-sentence operational justification for auditing.

### 5.4 Two-Step Staging & Non-Destructive Execution
* **Step 1 (Analyze & Plan):** Populate an interactive Treeview previewing:
  * File Name
  * Triage Method (Filename Priority vs. Content Inspected)
  * Target Hierarchy (`Parent Group / Subfolder`)
  * AI Reasoning
* **Step 2 (Commit & Organize):** Batch file relocations with:
  * Collision Resolution: Append numeric suffixes (`_1`, `_2`) if identical filenames exist in target folder. Never overwrite.
  * Session Logging: Write atomic operation logs (`session_id`, `timestamp`, `original_path`, `moved_path`) to `.custom_organizer_history.json`.

### 5.5 One-Click Undo & Directory Pruning
* Load previous active session from the manifest.
* Restore all moved files back to their exact source paths.
* Automatically prune empty target folders left behind by reversed sessions.

---

## 6. Non-Functional Requirements (NFRs)

* **Data Privacy:** Zero network calls to external analytical services or remote model inference endpoints. All processing remains on the host device.
* **Fault Tolerance:** If a corrupted document fails to parse, gracefully degrade to Tier 1 triage instead of halting the queue.
* **UI Responsiveness:** All model queries, document parsing, and file I/O operations must execute on asynchronous worker threads to keep the Tkinter event loop at 60 FPS.

---

## 7. Roadmap & Future Scope (v1.1+)

* **v1.1 (Multi-Modal Vision Support):** Integrate a lightweight vision model (e.g., `llama3.2-vision` or `moondream2`) to classify scanned image receipts (`.png`, `.jpeg`) and handwritten notes via OCR.
* **v1.2 (Background Folder Watcher):** Optional daemon menu bar app (`FSEvents` listener) that silently organizes `~/Downloads` on a scheduled interval.
* **v1.3 (Custom Rule Overrides):** Allow users to pin hard regex rules (e.g., `*slip*.pdf` $\to$ `Banking`) that bypass LLM inference entirely for fixed recurring files.