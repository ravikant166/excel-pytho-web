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
            # fillna("") prevents the 'Unknown Parameter' error
            df = xl.parse(name).fillna("")
            
            # Standardize date formatting
            for col in df.select_dtypes(include=['datetime']).columns:
                df[col] = df[col].dt.strftime('%d-%m-%Y')
            
            all_data[name] = df.to_dict(orient='records')
            
            # Map columns for DataTables
            sheet_configs.append({
                "name": name,
                "columns": [{"title": col, "data": col, "defaultContent": ""} for col in df.columns]
            })

        # Save data as a JavaScript Variable to bypass CORS/Local security
        js_filename = report_title + "_data.js"
        js_path = os.path.join(os.path.dirname(output_file), js_filename)
        with open(js_path, 'w', encoding='utf-8') as f:
            f.write(f"var masterData = {json.dumps(all_data)};")

        # HTML Structure
        html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{report_title}</title>
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css">
    
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, sans-serif; background: #f4f4f4; margin: 0; padding: 20px; }}
        
        /* Header & Top Nav */
        .top-nav {{ 
            background: white; padding: 15px 20px; 
            display: flex; justify-content: space-between; align-items: center; 
            box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 20px; border-radius: 8px;
        }}
        .title-area h2 {{ margin: 0; color: {theme}; font-size: 18px; }}
        .title-area span {{ font-size: 11px; color: #777; }}
        
        /* DataTables Buttons in Header */
        .dt-buttons {{ margin-bottom: 0 !important; }}
        .dt-button {{ 
            background: white !important; border: 1px solid {theme} !important; 
            color: {theme} !important; border-radius: 4px !important; font-size: 11px !important; 
            font-weight: bold !important; transition: 0.2s !important;
        }}
        .dt-button:hover {{ background: {theme} !important; color: white !important; }}

        /* Tabs */
        .tabs {{ display: flex; gap: 5px; margin-bottom: -1px; overflow-x: auto; }}
        .tab-btn {{ 
            padding: 10px 20px; cursor: pointer; border: 1px solid #ccc; 
            background: #ddd; border-radius: 6px 6px 0 0; font-size: 13px; font-weight: bold; 
        }}
        .tab-btn.active {{ background: {theme}; color: white; border-color: {theme}; }}

        /* Container */
        .container {{ background: white; padding: 20px; border-radius: 0 8px 8px 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }}
        
        /* Table Styling */
        table.dataTable thead th {{ background-color: {theme} !important; color: white !important; font-size: 12px; }}
        .filter-input {{ width: 100%; box-sizing: border-box; padding: 4px; font-size: 11px; border: 1px solid #ccc; border-radius: 3px; }}
        .dataTables_wrapper {{ font-size: 12px; }}
    </style>
</head>
<body>
    <div class="top-nav">
        <div class="title-area">
            <h2>{report_title}</h2>
            <span>Report Generated: {timestamp}</span>
        </div>
        <div id="exportBtns"></div> </div>

    <div class="tabs" id="tabBar"></div>
    <div class="container">
        <table id="mainTable" class="display nowrap cell-border" style="width:100%">
            <thead><tr id="filterRow"></tr></thead>
        </table>
    </div>

    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/dataTables.buttons.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/pdfmake.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/vfs_fonts.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.html5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.print.min.js"></script>

    <script src="{js_filename}"></script>

    <script>
        let table;
        const sheetConfigs = {json.dumps(sheet_configs)};

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
                document.getElementById('mainTable').innerHTML = "<thead><tr id='filterRow'></tr></thead>";
            }}

            const filterRow = document.getElementById('filterRow');
            config.columns.forEach(() => {{
                const th = document.createElement('th');
                th.innerHTML = '<input type="text" class="filter-input" placeholder="Search...">';
                filterRow.appendChild(th);
            }});
            
            table = $('#mainTable').DataTable({{
                data: masterData[sheetName],
                columns: config.columns,
                pageLength: 20,
                scrollX: true,
                dom: 'Bfrtip',
                buttons: [
                    'copy', 'csv', 'excel',
                    {{
                        extend: 'pdfHtml5',
                        text: 'PDF',
                        orientation: 'landscape',
                        pageSize: 'A3', // Larger size for many columns
                        exportOptions: {{ columns: ':visible' }},
                        customize: function (doc) {{
                            doc.defaultStyle.fontSize = 7;
                            doc.styles.tableHeader.fontSize = 8;
                            doc.styles.tableHeader.fillColor = '{theme}';
                            doc.content[1].table.widths = Array(doc.content[1].table.body[0].length + 1).join('*').split('');
                        }}
                    }}
                ],
                initComplete: function () {{
                    // Move buttons to header
                    this.buttons().container().appendTo('#exportBtns');
                    
                    // Apply column search
                    this.api().columns().every(function () {{
                        var that = this;
                        $('input', this.header()).on('keyup change clear', function () {{
                            if (that.search() !== this.value) {{
                                that.search(this.value).draw();
                            }}
                        }});
                        // Prevent sorting when clicking input
                        $('input', this.header()).on('click', function(e) {{ e.stopPropagation(); }});
                    }});
                }}
            }});
        }}

        initTabs();
        loadSheet(sheetConfigs[0].name);
    </script>
</body>
</html>
"""

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"[{timestamp}] Full Upgrade Success! Individual column filters active.")

    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Required: python excel_engine.py <input_path> <output_path>")
    else:
        run_update(sys.argv[1], sys.argv[2])