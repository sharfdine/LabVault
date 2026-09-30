# 🧪 PharmaLab OS — Automated Lab & Semester Inventory System

An interactive, pharmacy-grade laboratory operations and real-time inventory management platform built with Python & Streamlit, engineered for **spacious, multi-semester data retention** with crash-safe architecture.

---

## 🌟 Key Modules

### 1. 📅 Lab Schedule & Pre-Planning (Integrated Timetable)
- **Direct PDF Integration**: Automatically seeded from [`Lab_Timetable_135B_138B_FA26.pdf`](file:///c:/antigravity/LAB_AUTOMATION/Lab_Timetable_135B_138B_FA26.pdf) for both **Lab 135-B** and **Lab 138-B**.
- **Semester & Week-by-Week Organization**: Categorize sessions across all **18 academic weeks** of the semester.
- **Pre-Lab Planning**: Log planned experiment titles, required apparatus, and planned chemicals beforehand.
- **Workflow Continuity**: Click **"🚀 Send to Requisition Logger"** on any scheduled slot to pre-fill the Experiment Logger with zero repetitive typing.

### 2. 🧪 Experiment Logging & Pharmacy-Style Chemical Deduction (POS Requisition)
- **Point-of-Sale / Pharmacy Style Dispensing Cart**:
  - Live search across 1,340+ items by **Chemical Name**, **Item ID**, or **FR Code**.
  - Real-time stock display with storage location, expiry date, and unit (`gms`, `mL`, `Nos`).
  - Active stock validation to prevent accidental over-dispensing.
- **Simulated Requisition Slip (Bill of Materials)**:
  - Add multiple chemicals and apparatus to a single experiment slip.
  - Live projection showing current stock and expected balance after deduction.
- **ACID Transaction Deduction**:
  - Automatically deducts exact quantities from the central inventory in real time.
  - Generates a unique Experiment ID (e.g. `EXP-20260930-110147`) linked to Academic Semester & Week.
  - Records student batch, instructor, apparatus used, observations, and findings.

### 3. 📦 Central Chemical & Labware Inventory
- **Real-Time Stock Room**: Loaded directly from [`Total chemicals and lab ware list.xlsx`](file:///c:/antigravity/LAB_AUTOMATION/Total%20chemicals%20and%20lab%20ware%20list.xlsx).
- **Multi-criteria Filtering**: Filter by Lab Room (`135-B`, `138-B`, `M001`, `M003`), Group (`Chemical`, `Hardware / Labware`), or Stock Level (`Available`, `Low Stock`, `Out of Stock`).
- **Stock Inward & Restocking**: Pharmacy stock receipt modal to restock items or adjust balances with mandatory audit references.
- **One-Click Excel Export**: Download updated inventory sheets with all current balances anytime.

### 4. 📝 Categorized Notes & Correspondences
- Organize tasks into **Follow-ups**, **Correspondences**, **Safety & Compliance**, **Equipment Maintenance**, and **Procurement**.
- Priority markers (**High**, **Medium**, **Low**) and status toggles (**Pending**, **In Progress**, **Completed**).

### 5. 🧮 Lab Productivity Tools & Formulation Calculators
- **Molarity to Mass Calculator**: Computes required grams given desired Molarity ($M$), Volume ($mL$), and Molecular Weight.
- **Dilution Calculator ($C_1V_1 = C_2V_2$)**: Calculates required aliquot volume from stock solutions.
- **Percentage Solutions (% w/v and % v/v)**: Recipe helper for common laboratory reagents.
- **Reorder & Out-of-Stock Alert Radar**: Instant visibility on chemicals reaching safety thresholds.

### 6. 📜 Dispense Ledger & Experiment Audit Trail
- Complete chronological ledger of every single stock deduction and restock.
- Logs exact user, timestamp, quantity change, resulting balance, and experiment reference.

### 7. 💾 Spacious Semester Storage & Archive Vault (New!)
- **Unlimited Semester Retention**:
  - SQLite configured with **Write-Ahead Logging (WAL mode)**, B-tree indexes, and memory caching. Allows millions of records with sub-millisecond query performance.
  - Architectural capacity: Up to **281 Terabytes** per database file.
- **📸 Zero-Downtime Snapshots & Backups**:
  - Create atomic point-in-time database snapshots (`conn.backup()`) with zero downtime.
  - 1-click snapshot restore or download.
- **📎 File Attachments Repository**:
  - Store lab protocol PDFs, MSDS documents, spectra charts, equipment manuals, and experiment photos.
  - Configured for up to **2 GB upload size per file**.
- **📦 Academic Master Semester Archive**:
  - 1-click generator for comprehensive multi-sheet Excel archives (`Inventory_Master`, `Timetable_Preplan`, `Experiment_Logs`, `Stock_Transactions`, `Notes_Tasks`) formatted for semester-end audits and accreditation.

---

## 🚀 How to Run the Dashboard

The dashboard is running locally via Streamlit:

```powershell
# Open terminal in this folder
cd c:\antigravity\LAB_AUTOMATION

# Launch dashboard
python -m streamlit run app.py
```

Access the app in your browser at:
👉 **`http://localhost:8501`**

---

## 🗄️ Architecture & Data Files

- [`app.py`](file:///c:/antigravity/LAB_AUTOMATION/app.py): Streamlit dashboard with custom light-themed CSS and reactive UI components.
- [`database.py`](file:///c:/antigravity/LAB_AUTOMATION/database.py): SQLite database manager (`lab_inventory.db`) with WAL mode, attachments, and backup engine.
- [`data_storage/`](file:///c:/antigravity/LAB_AUTOMATION/data_storage/): Dedicated persistent storage folder containing:
  - `attachments/`: Uploaded lab manuals, PDFs, images, and reports.
  - `backups/`: Automated semester database snapshots.
  - `archives/`: Multi-sheet semester Excel workbooks.
- [`.streamlit/config.toml`](file:///c:/antigravity/LAB_AUTOMATION/.streamlit/config.toml): Configured for 2GB file uploads and light UI theme.
- [`Lab_Timetable_135B_138B_FA26.pdf`](file:///c:/antigravity/LAB_AUTOMATION/Lab_Timetable_135B_138B_FA26.pdf): Source timetable.
- [`Total chemicals and lab ware list.xlsx`](file:///c:/antigravity/LAB_AUTOMATION/Total%20chemicals%20and%20lab%20ware%20list.xlsx): Master inventory reference.
