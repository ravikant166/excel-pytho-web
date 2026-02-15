import pandas as pd
import os
import sys
from datetime import datetime

def run_update(input_file, output_file):
    try:
        if not os.path.exists(input_file):
            print(f"Error: Input file '{input_file}' not found.")
            return

        xl = pd.ExcelFile(input_file)
        theme = "#880055"
        timestamp = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        report_title = os.path.splitext(os.path.basename(output_file))[0]
        
        # HTML Structure
        html = [f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{report_title}</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.25/jspdf.plugin.autotable.min.js"></script>
    <style>
        body {{ font-family: 'Segoe UI', sans-serif; background: #f4f4f4; margin: 0; color: #333; }}
        .top-nav {{ background: white; padding: 10px 20px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); border-bottom: 1px solid #ddd; }}
        .title-area h2 {{ margin: 0; color: {theme}; font-size: 18px; }}
        .title-area span {{ font-size: 11px; color: #777; }}
        .btn-group {{ display: flex; gap: 8px; }}
        .action-btn {{ padding: 6px 14px; cursor: pointer; border: 1px solid {theme}; background: white; color: {theme}; border-radius: 4px; font-size: 11px; font-weight: bold; }}
        .tabs {{ display: flex; gap: 5px; padding: 10px 20px 0 20px; background: #f8f8f8; overflow-x: auto; border-bottom: 2px solid {theme}; }}
        .tab-btn {{ padding: 10px 20px; cursor: pointer; border: 1px solid #ccc; border-bottom: none; background: #eee; border-radius: 6px 6px 0 0; font-size: 13px; font-weight: bold; color: #555; }}
        .tab-btn.active {{ background: {theme}; color: white; border-color: {theme}; }}
        .content {{ display: none; padding: 20px; }}
        .content.active {{ display: block; }}
        .tbl-wrap {{ background: white; padding: 15px; border-radius: 0 0 8px 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); overflow-x: auto; }}
        table {{ border-collapse: collapse; width: 100%; font-size: 13px; min-width: 600px; }}
        th {{ background: {theme}; color: white; padding: 10px; border: 1px solid #ddd; text-align: left; }}
        td {{ padding: 8px; border: 1px solid #eee; }}
        tr:nth-child(even) {{ background: #fafafa; }}
        tfoot {{ background: #fce4ec; font-weight: bold; color: {theme}; border-top: 2px solid {theme}; }}
        .filter-box {{ width: 100%; box-sizing: border-box; padding: 5px; margin-top: 8px; border: 1px solid rgba(255,255,255,0.3); border-radius: 3px; font-size: 11px; color: #333; }}
    </style>
    <script>
        function showTab(idx) {{
            document.querySelectorAll('.content').forEach(c => c.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.getElementById('tab-' + idx).classList.add('active');
            document.getElementById('btn-' + idx).classList.add('active');
        }}

        function doFilter(el) {{
            const table = el.closest('table');
            const rows = table.querySelectorAll('tbody tr');
            const filters = Array.from(table.querySelectorAll('.filter-box'));
            rows.forEach(row => {{
                let match = true;
                filters.forEach((inp, i) => {{
                    let val = inp.value.toLowerCase();
                    let text = row.cells[i].textContent.toLowerCase();
                    if (val && !text.includes(val)) match = false;
                }});
                row.style.display = match ? '' : 'none';
            }});
            updateTotals(table);
        }}

        function updateTotals(table) {{
            const rows = Array.from(table.querySelectorAll('tbody tr')).filter(r => r.style.display !== 'none');
            const fCells = table.querySelector('tfoot tr').cells;
            for (let i = 0; i < fCells.length; i++) {{
                let sum = 0, isNum = true, hasData = false;
                rows.forEach(r => {{
                    let raw = r.cells[i].textContent.trim();
                    if (raw !== "") {{
                        hasData = true;
                        let isDate = /^\\d{{2}}-\\d{{2}}-\\d{{4}}$/.test(raw);
                        let n = parseFloat(raw.replace(/,/g, ''));
                        if (!isDate && !isNaN(n)) sum += n; else isNum = false;
                    }}
                }});
                fCells[i].innerHTML = (isNum && hasData) ? 'Sum:<br>' + sum.toLocaleString() : 'Count:<br>' + rows.length;
            }}
        }}

        function exportCSV() {{
            const activeTab = document.querySelector('.content.active');
            const table = activeTab.querySelector('table');
            let csvLines = [];
            
            // Extract Headers
            const headers = Array.from(table.querySelectorAll('thead th')).map(th => th.childNodes[0].textContent.trim());
            csvLines.push(headers.join(","));
            
            // Extract Visible Rows
            table.querySelectorAll('tbody tr').forEach(row => {{
                if (row.style.display !== 'none') {{
                    const values = Array.from(row.cells).map(cell => {{
                        let data = cell.textContent.trim().replace(/"/g, '""');
                        return '"' + data + '"';
                    }});
                    csvLines.push(values.join(","));
                }}
            }});
            
            // Join with standard newline and create blob
            const csvString = csvLines.join('\\r\\n');
            const blob = new Blob([csvString], {{ type: 'text/csv;charset=utf-8;' }});
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = "{report_title}.csv";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }}

        function exportPDF() {{
            const {{ jsPDF }} = window.jspdf;
            const doc = new jsPDF('l', 'mm', 'a4');
            const activeTab = document.querySelector('.content.active');
            const sheetName = document.querySelector('.tab-btn.active').textContent;
            const table = activeTab.querySelector('table');
            const headers = [Array.from(table.querySelectorAll('thead th')).map(th => th.childNodes[0].textContent.trim())];
            const body = [];
            table.querySelectorAll('tbody tr').forEach(row => {{
                if (row.style.display !== 'none') {{
                    body.push(Array.from(row.cells).map(cell => cell.textContent.trim()));
                }}
            }});
            const footer = [Array.from(table.querySelector('tfoot tr').cells).map(cell => cell.innerText.split('\\n').join(' '))];
            doc.text("Report: {report_title}", 14, 15);
            doc.autoTable({{ head: headers, body: body, foot: footer, startY: 25, headStyles: {{ fillColor: '{theme}' }} }});
            doc.save("{report_title}.pdf");
        }}

        function clearFilters() {{
            document.querySelectorAll('.filter-box').forEach(inp => inp.value = '');
            document.querySelectorAll('table').forEach(tbl => {{
                tbl.querySelectorAll('tbody tr').forEach(r => r.style.display = '');
                updateTotals(tbl);
            }});
        }}
        window.onload = () => document.querySelectorAll('table').forEach(updateTotals);
    </script>
</head>
<body>
    <div class="top-nav">
        <div class="title-area"><h2>{report_title}</h2><span>Last Updated: {timestamp}</span></div>
        <div class="btn-group">
            <button class="action-btn" onclick="clearFilters()">Clear Filters</button>
            <button class="action-btn" onclick="exportCSV()">Export CSV</button>
            <button class="action-btn" onclick="exportPDF()">Export PDF</button>
        </div>
    </div>
    <div class="tabs">"""]

        for i, name in enumerate(xl.sheet_names):
            active = "active" if i == 0 else ""
            html.append(f'<button id="btn-{i}" class="tab-btn {active}" onclick="showTab({i})">{name}</button>')
        html.append('</div>')

        for i, name in enumerate(xl.sheet_names):
            active = "active" if i == 0 else ""
            html.append(f'<div id="tab-{i}" class="content {active}">')
            df = xl.parse(name)
            for col in df.select_dtypes(include=['datetime']).columns:
                df[col] = df[col].dt.strftime('%d-%m-%Y')
            html.append('<div class="tbl-wrap"><table><thead><tr>')
            for col in df.columns:
                html.append(f'<th>{col}<br><input class="filter-box" onkeyup="doFilter(this)" placeholder="..."></th>')
            html.append('</tr></thead><tbody>')
            for _, row in df.iterrows():
                if row.isnull().all(): continue
                html.append('<tr>' + "".join([f'<td>{("" if pd.isna(v) else str(v))}</td>' for v in row]) + '</tr>')
            html.append('</tbody><tfoot><tr>' + "".join(['<td></td>' for _ in df.columns]) + '</tr></tfoot></table></div></div>')

        html.append('</body></html>')
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("".join(html))
        print(f"[{timestamp}] CSV & PDF Fix Applied. Update successful.")

    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Required: python excel_engine.py <input_path> <output_path>")
    else:
        run_update(sys.argv[1], sys.argv[2])