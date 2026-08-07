import http.server
import socketserver
import json
import openpyxl
import io
import os
import ctypes
import datetime

PORT = 8080

def safe_copy_file(src, dst):
    """ Safely copy a file even if locked by Excel (EXCEL.EXE) on Windows """
    try:
        kernel32 = ctypes.windll.kernel32
        GENERIC_READ = 0x80000000
        FILE_SHARE_READ = 1
        FILE_SHARE_WRITE = 2
        FILE_SHARE_DELETE = 4
        OPEN_EXISTING = 3
        
        handle = kernel32.CreateFileW(
            src, GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            None, OPEN_EXISTING, 0, None
        )
        if handle == -1 or handle == 0 or handle == 0xFFFFFFFF:
            return False
        
        buf_size = 65536
        buf = ctypes.create_string_buffer(buf_size)
        bytes_read = ctypes.c_ulong(0)
        
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        with open(dst, 'wb') as f_out:
            while True:
                res = kernel32.ReadFile(handle, buf, buf_size, ctypes.byref(bytes_read), None)
                if not res or bytes_read.value == 0:
                    break
                f_out.write(buf.raw[:bytes_read.value])
                
        kernel32.CloseHandle(handle)
        return True
    except Exception as e:
        print("Safe copy error:", e)
        return False

def ensure_master_templates():
    f1 = 'DT-3 Daily energy Consumption  Updated.xlsx'
    d1 = 'master_templates/DT3_master.xlsx'
    if os.path.exists(f1):
        safe_copy_file(f1, d1)

    f2 = 'AHU_Saved_Hours_and_Energy_Cost_Report.xlsx'
    d2 = 'master_templates/AHU_master.xlsx'
    if os.path.exists(f2):
        safe_copy_file(f2, d2)

def Math_max_0(val):
    try:
        v = float(val)
        return v if v > 0 else 0.0
    except:
        return 0.0

def scan_dt3_sheet_rows(ws):
    """ Dynamically locate exact reading and consumption row indices for any month sheet """
    meter_rows = {}
    cons_rows = {}
    
    # Readings section: rows 1 to 42
    for r in range(1, 42):
        c_val = str(ws.cell(r, 3).value or '').strip()
        if '540430038840' in c_val: meter_rows['eb_1'] = r
        elif '540430038845' in c_val: meter_rows['eb_2'] = r
        elif '540430038844' in c_val: meter_rows['eb_3'] = r
        elif '540430038821' in c_val: meter_rows['eb_4'] = r
        elif '540430038849' in c_val: meter_rows['eb_5'] = r
        elif '540430038696' in c_val: meter_rows['eb_6'] = r
        elif '540430038843' in c_val: meter_rows['eb_7'] = r
        elif '540430038848' in c_val: meter_rows['eb_8'] = r
        elif c_val == 'AHU1': meter_rows['ahu_1'] = r
        elif c_val == 'AHU2': meter_rows['ahu_2'] = r
        elif c_val == 'AHU3': meter_rows['ahu_3'] = r
        elif c_val == 'AHU4': meter_rows['ahu_4'] = r
        elif 'AHU1 - BTU' in c_val: meter_rows['btu_1'] = r
        elif 'AHU2 - BTU' in c_val: meter_rows['btu_2'] = r
        elif 'AHU3 - BTU' in c_val: meter_rows['btu_3'] = r
        elif 'AHU4 - BTU' in c_val: meter_rows['btu_4'] = r
        
    # Consumptions section: rows 40 to max_row
    for r in range(40, ws.max_row + 1):
        c_val = str(ws.cell(r, 3).value or '').strip()
        b_val = str(ws.cell(r, 2).value or '').strip()
        lbl = c_val or b_val
        if '540430038840' in lbl: cons_rows['eb_1'] = r
        elif '540430038845' in lbl: cons_rows['eb_2'] = r
        elif '540430038844' in lbl: cons_rows['eb_3'] = r
        elif '540430038821' in lbl: cons_rows['eb_4'] = r
        elif '540430038849' in lbl: cons_rows['eb_5'] = r
        elif '540430038696' in lbl: cons_rows['eb_6'] = r
        elif '540430038843' in lbl: cons_rows['eb_7'] = r
        elif '540430038848' in lbl: cons_rows['eb_8'] = r
        elif lbl == 'AHU1': cons_rows['ahu_1'] = r
        elif lbl == 'AHU2': cons_rows['ahu_2'] = r
        elif lbl == 'AHU3': cons_rows['ahu_3'] = r
        elif lbl == 'AHU4': cons_rows['ahu_4'] = r
        elif 'AHU1 - BTU' in lbl: cons_rows['btu_1'] = r
        elif 'AHU2 - BTU' in lbl: cons_rows['btu_2'] = r
        elif 'AHU3 - BTU' in lbl: cons_rows['btu_3'] = r
        elif 'AHU4 - BTU' in lbl: cons_rows['btu_4'] = r
        elif 'Total Power - EB Unit' in lbl: cons_rows['tot_eb'] = r
        elif 'Total Power- DG' in lbl: cons_rows['tot_dg'] = r
        elif 'Total Power Cumulative' in lbl: cons_rows['tot_cum'] = r
        elif 'Total AHU Power Consumption' in lbl: cons_rows['tot_ahu'] = r
        elif 'Total AHU DG consumption' in lbl: cons_rows['tot_ahu_dg'] = r
        elif 'Total BTU Consumption' in lbl: cons_rows['tot_btu'] = r

    return meter_rows, cons_rows

class EnergyTrackerHandler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self):
        if self.path == '/api/export/kwh':
            self.handle_export_kwh()
        elif self.path == '/api/export/ahu':
            self.handle_export_ahu()
        elif self.path == '/api/save':
            self.handle_save_data()
        else:
            self.send_error(404, "Endpoint not found")

    def handle_save_data(self):
        content_length = int(self.headers.get('Content-Length', 0))
        post_data = self.rfile.read(content_length)
        try:
            payload = json.loads(post_data.decode('utf-8'))
            with open('user_data.json', 'w', encoding='utf-8') as f:
                json.dump(payload, f, indent=2)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "success"}).encode('utf-8'))
        except Exception as e:
            self.send_error(500, f"Error saving data: {str(e)}")

    def handle_export_kwh(self):
        content_length = int(self.headers.get('Content-Length', 0))
        post_data = self.rfile.read(content_length)
        ensure_master_templates()

        template_file = 'master_templates/DT3_master.xlsx'
        if not os.path.exists(template_file):
            template_file = 'DT-3 Daily energy Consumption  Updated.xlsx'

        try:
            tracker_data = json.loads(post_data.decode('utf-8')) if content_length > 0 else {}
            wb = openpyxl.load_workbook(template_file)
            
            wb.calculation.calcMode = 'auto'
            wb.calculation.fullCalcOnLoad = True

            kwh_daily = tracker_data.get('kwh_daily', {})

            # Map all date columns across month sheets
            date_to_sheet_col = {}
            sheet_row_maps = {}

            for sheetname in wb.sheetnames:
                if sheetname in ['Dashboard', 'summary']:
                    continue
                sheet = wb[sheetname]
                if hasattr(sheet, 'sheet_view') and sheet.sheet_view:
                    sheet.sheet_view.showGridLines = True

                m_rows, c_rows = scan_dt3_sheet_rows(sheet)
                sheet_row_maps[sheetname] = (m_rows, c_rows)

                for r in range(2, 6):
                    for c in range(5, sheet.max_column + 1):
                        val = sheet.cell(r, c).value
                        if isinstance(val, (datetime.datetime, datetime.date)):
                            d_str = val.strftime('%Y-%m-%d')
                            date_to_sheet_col[d_str] = (sheetname, c)

            # Update readings and consumptions for each date
            for d_str, day_entry in kwh_daily.items():
                if d_str in date_to_sheet_col:
                    sname, c_idx = date_to_sheet_col[d_str]
                    sheet = wb[sname]
                    m_rows, c_rows = sheet_row_maps[sname]

                    # 1. EB Meters
                    eb_list = day_entry.get('eb', [])
                    day_eb_sum = 0
                    day_dg_sum = 0

                    for eb_item in eb_list:
                        m_id = eb_item.get('id')
                        m_key = f'eb_{m_id}'
                        if m_key in m_rows:
                            eb_row = m_rows[m_key]
                            dg_row = eb_row + 1
                            eb_cons_row = c_rows.get(m_key)
                            dg_cons_row = eb_cons_row + 1 if eb_cons_row else None

                            curr_eb = float(eb_item.get('reading', 0))
                            curr_dg = float(eb_item.get('dg_reading', 0))

                            prev_col = c_idx - 1
                            prev_eb = float(sheet.cell(eb_row, prev_col).value or 0)
                            prev_dg = float(sheet.cell(dg_row, prev_col).value or 0)

                            sheet.cell(eb_row, c_idx, curr_eb)
                            sheet.cell(dg_row, c_idx, curr_dg)

                            eb_cons = Math_max_0(curr_eb - prev_eb)
                            dg_cons = Math_max_0(curr_dg - prev_dg)

                            if eb_cons_row:
                                sheet.cell(eb_cons_row, c_idx, eb_cons)
                            if dg_cons_row:
                                sheet.cell(dg_cons_row, c_idx, dg_cons)

                            day_eb_sum += eb_cons
                            day_dg_sum += dg_cons

                    if 'tot_eb' in c_rows: sheet.cell(c_rows['tot_eb'], c_idx, day_eb_sum)
                    if 'tot_dg' in c_rows: sheet.cell(c_rows['tot_dg'], c_idx, day_dg_sum)
                    if 'tot_cum' in c_rows: sheet.cell(c_rows['tot_cum'], c_idx, day_eb_sum + day_dg_sum)

                    # 2. AHU Meters
                    ahu_list = day_entry.get('ahu', [])
                    day_ahu_sum = 0
                    day_ahu_dg_sum = 0

                    for ahu_item in ahu_list:
                        m_id = ahu_item.get('id')
                        m_key = f'ahu_{m_id}'
                        if m_key in m_rows:
                            ahu_row = m_rows[m_key]
                            dg_row = ahu_row + 1
                            ahu_cons_row = c_rows.get(m_key)
                            dg_cons_row = ahu_cons_row + 1 if ahu_cons_row else None

                            curr_ahu = float(ahu_item.get('reading', 0))
                            curr_dg = float(ahu_item.get('dg_reading', 0))

                            prev_col = c_idx - 1
                            prev_ahu = float(sheet.cell(ahu_row, prev_col).value or 0)
                            prev_dg = float(sheet.cell(dg_row, prev_col).value or 0)

                            sheet.cell(ahu_row, c_idx, curr_ahu)
                            sheet.cell(dg_row, c_idx, curr_dg)

                            ahu_cons = Math_max_0(curr_ahu - prev_ahu)
                            dg_cons = Math_max_0(curr_dg - prev_dg)

                            if ahu_cons_row:
                                sheet.cell(ahu_cons_row, c_idx, ahu_cons)
                            if dg_cons_row:
                                sheet.cell(dg_cons_row, c_idx, dg_cons)

                            day_ahu_sum += ahu_cons
                            day_ahu_dg_sum += dg_cons

                    if 'tot_ahu' in c_rows: sheet.cell(c_rows['tot_ahu'], c_idx, day_ahu_sum)
                    if 'tot_ahu_dg' in c_rows: sheet.cell(c_rows['tot_ahu_dg'], c_idx, day_ahu_dg_sum)

                    # 3. BTU Meters
                    btu_list = day_entry.get('btu', [])
                    day_btu_sum = 0

                    for btu_item in btu_list:
                        m_id = btu_item.get('id')
                        m_key = f'btu_{m_id}'
                        if m_key in m_rows:
                            btu_row = m_rows[m_key]
                            btu_cons_row = c_rows.get(m_key)

                            curr_btu = float(btu_item.get('reading', 0))
                            prev_col = c_idx - 1
                            prev_btu = float(sheet.cell(btu_row, prev_col).value or 0)

                            sheet.cell(btu_row, c_idx, curr_btu)

                            btu_cons = Math_max_0(curr_btu - prev_btu)
                            if btu_cons_row:
                                sheet.cell(btu_cons_row, c_idx, btu_cons)

                            day_btu_sum += btu_cons

                    if 'tot_btu' in c_rows: sheet.cell(c_rows['tot_btu'], c_idx, day_btu_sum)

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            file_data = output.read()

        except Exception as e:
            print("Fallback to raw master file due to:", e)
            with open(template_file, 'rb') as f_raw:
                file_data = f_raw.read()

        self.send_response(200)
        self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.send_header('Content-Disposition', 'attachment; filename="DT-3 Daily energy Consumption  Updated.xlsx"')
        self.send_header('Content-Length', str(len(file_data)))
        self.end_headers()
        self.wfile.write(file_data)

    def handle_export_ahu(self):
        content_length = int(self.headers.get('Content-Length', 0))
        post_data = self.rfile.read(content_length)
        ensure_master_templates()

        template_file = 'master_templates/AHU_master.xlsx'
        if not os.path.exists(template_file):
            template_file = 'AHU_Saved_Hours_and_Energy_Cost_Report.xlsx'

        try:
            tracker_data = json.loads(post_data.decode('utf-8')) if content_length > 0 else {}
            wb = openpyxl.load_workbook(template_file)
            
            wb.calculation.calcMode = 'auto'
            wb.calculation.fullCalcOnLoad = True

            ahu_saving = tracker_data.get('ahu_saving', {})

            for sheetname in wb.sheetnames:
                sheet = wb[sheetname]
                if hasattr(sheet, 'sheet_view') and sheet.sheet_view:
                    sheet.sheet_view.showGridLines = True

                if sheetname in ['INDEX & CONTROL PANEL', 'Executive Dashboard']:
                    continue

                date_row_map = {}
                for r in range(7, sheet.max_row + 1):
                    val = sheet.cell(r, 1).value
                    if isinstance(val, (datetime.datetime, datetime.date)):
                        d_str = val.strftime('%Y-%m-%d')
                        date_row_map[d_str] = r

                for d_str, r_idx in date_row_map.items():
                    if d_str in ahu_saving:
                        day_entry = ahu_saving[d_str]
                        ahus = day_entry.get('ahus', [])

                        col_offsets = [3, 13, 23, 33] # C, M, W, AG
                        for i, offset in enumerate(col_offsets):
                            if i < len(ahus):
                                a_item = ahus[i]
                                if a_item.get('on_time'):
                                    sheet.cell(r_idx, offset, str(a_item['on_time']))
                                if a_item.get('off_time'):
                                    sheet.cell(r_idx, offset + 1, str(a_item['off_time']))
                                if a_item.get('kwh_cons') is not None:
                                    sheet.cell(r_idx, offset + 5, float(a_item['kwh_cons']))

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            file_data = output.read()

        except Exception as e:
            print("Fallback to raw master file due to:", e)
            with open(template_file, 'rb') as f_raw:
                file_data = f_raw.read()

        self.send_response(200)
        self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.send_header('Content-Disposition', 'attachment; filename="AHU_Saved_Hours_and_Energy_Cost_Report.xlsx"')
        self.send_header('Content-Length', str(len(file_data)))
        self.end_headers()
        self.wfile.write(file_data)

if __name__ == '__main__':
    ensure_master_templates()
    handler = EnergyTrackerHandler
    with socketserver.TCPServer(("", PORT), handler) as httpd:
        print(f"Energy Tracker Server running on port {PORT} with dynamic row-scanning engine...")
        httpd.serve_forever()
