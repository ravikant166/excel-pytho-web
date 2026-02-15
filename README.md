# Excel-to-HTML Pro Porter

A robust Python-based engine that converts multi-sheet Excel workbooks into an interactive, searchable HTML dashboard. Featuring automated PDF/CSV exports, dynamic totals, and a corporate theme.

## 🚀 Features
- **Multi-Tab Support:** Each Excel sheet is converted into a distinct, navigable tab.
- **Dynamic Filtering:** Search any column; rows update instantly.
- **Auto-Calculations:** The footer automatically calculates the **Sum** for numbers or **Count** for text/dates.
- **Instant Export:** Generate filtered CSVs or professional PDFs directly from the browser.
- **Date Protection:** Automatically formats dates to `DD-MM-YYYY`.
- **Branded UI:** Styled with a professional `#880055` (Plum) theme.


## ⚠️ Note for Large Files
Because your data is massive, the system now creates **two files**:
1. `Report.html` (The Viewer)
2. `Report_data.json` (The Data)

**Both files must stay in the same folder** for the dashboard to work.

### 🚀 Performance Benefits
- **Pagination:** Only 25 rows are rendered at a time, keeping the browser fast.
- **Search:** Instant global search across thousands of rows.
- **Built-in Exports:** The PDF/CSV/Excel export buttons are now handled by the DataTables library, which is much faster for large datasets.

## 🛠️ Setup
1. **Install Python 3.x**
2. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt