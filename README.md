# XLSX Anonymizer

A tool for automatic anonymization and deanonymization of data in Excel files (`.xlsx`) using a local LLM to identify entities.

## 🚀 Features

- **Automatic Analysis**: Detecting persons, companies, places, and sensitive data using an LLM.
- **Full Control**: Generating a list of identified entities for manual verification before anonymization.
- **Reversibility**: Ability to restore original data thanks to token mapping.
- **Privacy**: Local data processing (requires LM Studio or a similar local server).

## 🛠 Requirements

- Python 3.x
- Libraries: `openpyxl`, `requests`
- A running local LLM server (e.g., **LM Studio**) with an OpenAI-compatible API.

Installation:
```bash
pip install openpyxl requests
```

## 📋 Workflow

The process consists of three main stages:

### 1. Analysis (`analyze.py`)
Scans Excel files in the directory and sends text fragments to a local LLM to identify sensitive data.
- **Output**: `entities.xlsx` file.
- **Action**: Open `entities.xlsx` and set the **Anonymize** column to `TAK` (YES) for elements that should be hidden. You can also modify the proposed pseudonyms.

```bash
python analyze.py your_file.xlsx
```

### 2. Anonymization (`anonymize.py`)
Replaces values in the Excel file based on the settings in `entities.xlsx`.
- **Output**: `your_file_anon.xlsx` and a mapping file `anon_map.json`.

```bash
python anonymize.py your_file.xlsx
```

### 3. Deanonymization (`deanonymize.py`)
Restores original data to the anonymized file using the `anon_map.json` map.
- **Output**: `your_file_restored.xlsx`.

```bash
python deanonymize.py your_file_anon.xlsx
```

## ⚙️ LLM Configuration

You can adjust the connection parameters in `analyze.py`:
- `LM_URL`: Server address (default: `http://localhost:1234/v1/chat/completions`).
- `MODEL`: The name of the model used in LM Studio.

## ⚠️ Notes
- The tool operates on a "find-and-replace" basis, ensuring speed and preserving the Excel file structure.
- The `anon_map.json` file is critical for data recovery – keep it in a secure location.
