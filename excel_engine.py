import pandas as pd
import os
import sys
import json
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
        
        all_data = {}
        sheet_configs = []

        for name in xl.sheet_names:
            df = xl.parse(name)
            # Standardize date formatting
            for col in df.select_dtypes(include=['datetime']).columns:
                df[col] = df[col].dt.strftime('%d-%m-%Y')
            
            all_data[name] = df.fillna("").to_dict(orient='records')
            sheet_configs.append({
                "name": name,
                "columns": [{"title": col, "data": col} for col in df.columns]
            })

        # Save data as a JavaScript Variable (NOT JSON)
        # This bypasses browser CORS security blocks
        js_filename = report_title + "_data.js"
        js_path = os.path.join(os.path.dirname(output_file), js_filename)
        with open(js_path, 'w', encoding='utf-8') as f:
            f.write(f"var masterData = {json.dumps(all_data)};")

        # HTML Structure using the local .js file
        html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{report_title}</title>
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css">
    
    <style>
        body {{ font-family: 'Segoe UI', sans-serif; background: #f4f4f4; margin: 0; padding: 20px; }}
        .header {{ background: white; padding: 15px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 20px; }}
        h2 {{ color: {theme}; margin: 0; }}
        .tabs {{ display: flex; gap: 5px; margin-bottom: -1px; overflow-x: auto; }}
        .tab-btn {{ padding: 10px 20px; cursor: pointer; border: 1px solid #ccc; background: #ddd; border-radius: 5px 5px 0 0; font-weight: bold; font-size: 13px; }}
        .tab-btn.active {{ background: {theme}; color: white; border-color: {theme}; }}
        .container {{ background: white; padding: 20px; border-radius: 0 5px 5px 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
        table.dataTable thead th {{ background-color: {theme} !important; color: white !important; font-size: 13px; }}
        .dataTables_wrapper {{ font-size: 13px; }}
    </style>
</head>
<body>
    <div class="header">
        <div><h2>{report_title}</h2><small>Offline Mode Enabled | Generated: {timestamp}</small></div>
    </div>

    <div class="tabs" id="tabBar"></div>
    <div class="container">
        <table id="mainTable" class="display nowrap" style="width:100%"></table>
    </div>

    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/dataTables.buttons.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/pdfmake.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/vfs_fonts.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.html5.min.js"></script>

    <script src="{js_filename}"></script>

    <script>
        let table;
        const sheetConfigs = {json.dumps(sheet_configs)};

        // masterData is already loaded from the external .js file
        function initTabs() {{
            const tabBar = document.getElementById('tabBar');
            sheetConfigs.forEach((sheet, idx) => {{
                const btn = document.createElement('button');
                btn.className = 'tab-btn' + (idx === 0 ? ' active' : '');
                btn.innerText = sheet.name;
                btn.onclick = (e) => {{
                    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                    e.target.classList.add('active');
                    loadSheet(sheet.name);
                }};
                tabBar.appendChild(btn);
            }});
        }}

        function loadSheet(sheetName) {{
            const config = sheetConfigs.find(s => s.name === sheetName);
            if (table) {{
                table.destroy();
                document.getElementById('mainTable').innerHTML = "";
            }}
            
            table = $('#mainTable').DataTable({{
                data: masterData[sheetName],
                columns: config.columns,
                pageLength: 25,
                scrollX: true,
                dom: 'Bfrtip',
                buttons: [
                    'copy', 'csv', 'excel', 
                    {{
                        extend: 'pdfHtml5',
                        orientation: 'landscape',
                        pageSize: 'A4',
                        title: '{report_title} - ' + sheetName
                    }}
                ]
            }});
        }}

        // Launch
        initTabs();
        loadSheet(sheetConfigs[0].name);
    </script>
</body>
</html>
"""

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"[{timestamp}] Success! Double-click {output_file} to view.")

    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Required: python excel_engine.py <input_path> <output_path>")
    else:
        run_update(sys.argv[1], sys.argv[2])