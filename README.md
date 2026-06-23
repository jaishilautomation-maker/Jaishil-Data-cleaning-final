# Jaishil Data Cleaning & Studio

A unified repository containing a **Flask backend** for data cleaning and processing, and a **React/Vite frontend** for uploading, tracking, and visualising data operations.

## Repository Structure

```
├── jaishil-data-studio/     # React Frontend (Vite + TypeScript + TanStack)
├── cleaner.py               # Core cleaning logic for Sales, Purchase, JV, HO Exp, Creditors
├── main.py                  # Core cleaning logic for Stock Summary and Debtors
├── server.py                # Flask Web Server
├── .gitignore               # Root git ignore definitions
└── README.md                # This documentation
```

---

## 1. Flask Backend Setup (Python)

The backend exposes a Flask API server on port `5000` to clean uploaded Excel spreadsheets using Python libraries (`pandas` and `openpyxl`).

### Prerequisites
- Python 3.8 or higher installed.

### Setup and Running
1. **Navigate to the root directory** (where `server.py` is located).
2. **Create a virtual environment**:
   ```bash
   python -m venv venv
   ```
3. **Activate the virtual environment**:
   - **Windows (Command Prompt / PowerShell)**:
     ```powershell
     .\venv\Scripts\activate
     ```
   - **macOS / Linux**:
     ```bash
     source venv/bin/activate
     ```
4. **Install required packages**:
   ```bash
   pip install flask flask-cors pandas openpyxl
   ```
5. **Run the Flask server**:
   ```bash
   python server.py
   ```
   The backend will start and run on `http://127.0.0.1:5000`.

---

## 2. React Frontend Setup (Vite + TypeScript)

The frontend is a TanStack Start web app providing a clean interface to upload, configure, and process files using the data cleaning server.

### Prerequisites
- Node.js (v18+) or Bun installed.

### Setup and Running
1. **Navigate to the frontend folder**:
   ```bash
   cd jaishil-data-studio
   ```
2. **Install dependencies**:
   - Using **npm**:
     ```bash
     npm install
     ```
   - Using **Bun** (highly recommended as the project has `bun.lock`):
     ```bash
     bun install
     ```
3. **Run the development server**:
   - Using **npm**:
     ```bash
     npm run dev
     ```
   - Using **Bun**:
     ```bash
     bun dev
     ```
   The frontend app will launch at `http://localhost:3000` or `http://localhost:5173` (depending on configuration).

---

## Key Backend Features (Data Cleaning Logic)
The tool supports 7 profile cleaning targets:
- **Sales**: Cleans product entries, standardizes brand and SKU formatting, parses financial year, and extracts unit/quantity variables.
- **Purchase**: Corrects supplier formatting, removes non-product rows, and converts measurements (e.g., tons to kilograms).
- **JV (Journal Voucher)**: Flags and filters administrative expense entries while keeping relevant transaction rows.
- **HO Expense (Head Office Expense)**: Cross-checks column totals against grand totals to log discrepancies.
- **Sundry Creditors**: Cleans redundant company names from creditor sheets.
- **Stock Summary**: Dynamically removes intermediate subtotal rows, merges header sections, and cleans non-product rows.
- **Debtors**: Automatically removes unnecessary inter-branch rows.
