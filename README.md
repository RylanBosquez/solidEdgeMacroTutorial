# solidEdgeMacroTutorial

A step-by-step tutorial for creating and running Solid Edge macros using Python and COM automation. This guide covers environment setup, script execution, and integration with Solid Edge's COM API.

---

## 🔧 Installation

### 1. Create and activate a virtual environment

Using Conda:

```bash
conda create -n solidEdgeEnv python=3.11
conda activate solidEdgeEnv
```

Using venv:

```bash
python -m venv solidEdgeEnv
solidEdgeEnv\Scripts\activate.bat  # Windows
```


### 2. Install dependencies

```bash
pip install -r requirements.txt
```


### 3. Clone the repository

```bash
git clone https://github.com/RylanBosquez/solidEdgeMacroTutorial.git
cd solidEdgeMacroTutorial
```

---

## 📁 Folder Structure

```
solidEdgeMacroTutorial/
├── assets/                         # Exported files (PDF, DXF, DWG, JT)
├── scripts/                        # Python automation scripts
│   ├── AssemblyDocumentAutomation/ # Scripts for assemblies (insert, traverse, constrain)
│   ├── BOMAndMetadata/             # Scripts for BOM export and metadata extraction
│   ├── DraftDocumentAutomation/    # Scripts for drawing views, dimensions, title block
│   ├── PartDocumentAutomation/     # Scripts for part creation and feature access
│   ├── PrintingAndExport/          # Scripts for printing and sheet/image export
│   ├── Utilities/                  # Scripts for general automation tasks
├── requirements.txt                # Python dependencies
└── README.md                       # This file
```


---

## 🚀 Usage

1. Open Solid Edge and load a document.
2. Run any script from the `scripts/` folder to automate tasks like:
   - Exporting to PDF, DXF, DWG, JT
   - Reading document properties
   - Creating parts or drawings programmatically

Example:

```bash
cd scripts/Utilities
python solidEdgeVersion.py
```

---

## 📚 Features Covered

- Launch and connect to Solid Edge
- Open existing documents
- Create new part documents
- Export drawings to PDF, DXF, DWG, JT
- Read standard and custom document properties
- Extract parts list (BOM) from Draft documents
- Create drawing views (e.g., front, side, isometric)
- Extract dimensions from Draft documents
- Read and modify title block text
- Export sheet thumbnails as images


---

## 🧩 Requirements

- Solid Edge installed
- Windows OS (COM automation is Windows-only)
- Python 3.11
- `comtypes` package

---

## 📌 Notes

- Most scripts assume Solid Edge is already running.
- Only Draft documents support DXF/DWG/JT export.
- All exports are saved to the `assets/` folder.

---

