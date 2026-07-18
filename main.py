import flet as ft
import pandas as pd
import os
import re
import unicodedata
from datetime import datetime
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from collections import defaultdict
import traceback

def chuan_hoa(text):
    if pd.isna(text) or text is None:
        return ""
    text = str(text).replace('\n',' ').replace('\r',' ')
    text = unicodedata.normalize('NFD', text)
    text = ''.join(c for c in text if unicodedata.category(c)!= 'Mn')
    text = text.lower().strip()
    text = re.sub(r'\s+', ' ', text)
    return text

def parse_number(s):
    if pd.isna(s):
        return 0
    if isinstance(s, (int, float)):
        return s
    s = str(s).strip().replace(',', '').replace(' ', '')
    try:
        return float(s)
    except:
        return 0

def write_sheet(ws, df_day):
    ws.delete_rows(1, ws.max_row)
    thin = Side(style='thin')
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    fh = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    fill_xam = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    fill_trang = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    center = Alignment(horizontal='center', vertical='center')
    right_center = Alignment(horizontal='right', vertical='center')
    center_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)

    hs = ['Ngày', 'Họ Tên', 'CCCD', 'Địa Chỉ', 'Loại Vàng', 'TL Vàng', 'Giá Vàng', 'Thành Tiền']
    for i, h in enumerate(hs):
        c = ws.cell(row=1, column=i+1, value=h)
        c.font = Font(bold=True)
        c.alignment = center_wrap
        c.border = bd
        c.fill = fh
    ws.row_dimensions[1].height = 26

    cur = 2
    groups = []
    cg = []
    last_key = None
    for _, rec in df_day.iterrows():
        key = (rec['Ngay'], str(rec['Ten']).strip(), str(rec['CCCD']).strip())
        if key != last_key:
            if cg: groups.append(cg)
            cg = [rec]
            last_key = key
        else:
            cg.append(rec)
    if cg: groups.append(cg)

    for g in groups:
        n = len(g)
        st = cur
        tot = 0
        for j in range(n):
            rec = g[j]
            rr = cur + j
            if j == 0:
                ws.cell(row=rr, column=1, value=rec['Ngay']).number_format = 'DD/MM/YYYY'
                ws.cell(row=rr, column=2, value=rec['Ten'])
                ws.cell(row=rr, column=3, value=str(rec['CCCD']).split('.')[0] if pd.notna(rec['CCCD']) else "")
                ws.cell(row=rr, column=4, value=rec['DiaChi'])
            ws.cell(row=rr, column=5, value=rec['LoaiVang'])
            ws.cell(row=rr, column=6, value=rec['TLVang']).number_format = '0.00'
            ws.cell(row=rr, column=7, value=rec['GiaVang']).number_format = '#,##0'
            ws.cell(row=rr, column=8, value=rec['ThanhTien']).number_format = '#,##0'
            for cc in range(1, 9):
                cell = ws.cell(row=rr, column=cc)
                cell.border = bd
                cell.alignment = right_center if cc in [6, 7, 8] else center
            tot += rec['ThanhTien']
        if n > 1:
            for cc in [1, 2, 3, 4]:
                ws.merge_cells(start_row=st, start_column=cc, end_row=st+n-1, end_column=cc)
                cell = ws.cell(row=st, column=cc)
                cell.alignment = center
        rt = cur + n
        for cc in range(1, 8):
            a = ws.cell(row=rt, column=cc, value="")
            a.border = bd
            a.fill = fill_trang
        a = ws.cell(row=rt, column=8, value=tot)
        a.font = Font(bold=True, color="FF0000")
        a.fill = fill_xam
        a.border = bd
        a.number_format = '#,##0'
        a.alignment = right_center
        cur = rt + 1

    cur += 1
    dt0 = df_day.iloc[0]['Ngay']
    ws.merge_cells(start_row=cur, start_column=4, end_row=cur, end_column=8)
    c = ws.cell(row=cur, column=4, value=f"TỔNG HỢP CHI TIẾT HÀNG BÁN NGÀY {dt0.day:02d}/{dt0.month:02d}/{dt0.year}")
    c.font = Font(bold=True)
    c.fill = fh
    c.alignment = center
    c.border = bd
    for cc in [5, 6, 7, 8]:
        ws.cell(row=cur, column=cc).border = bd
    ws.row_dimensions[cur].height = 26
    cur += 1

    titles = ["STT", "Loại Vàng", "TL Vàng", "Giá Vàng", "Thành Tiền"]
    for i, t in enumerate(titles):
        a = ws.cell(row=cur, column=4+i, value=t)
        a.font = Font(bold=True)
        a.fill = fh
        a.alignment = center
        a.border = bd
    ws.row_dimensions[cur].height = 26
    cur += 1

    agg = defaultdict(lambda: [0, 0, 0])
    for _, rec in df_day.iterrows():
        loai = str(rec['LoaiVang']).strip()
        agg[loai][0] += rec['TLVang']
        agg[loai][1] += rec['ThanhTien']
        if agg[loai][2] == 0:
            agg[loai][2] = rec['GiaVang']

    stt = 1
    tong_tt = 0
    for loai, (tl_sum, tt_sum, gia_mau) in agg.items():
        ws.cell(row=cur, column=4, value=stt).border = bd
        ws.cell(row=cur, column=4).alignment = center
        ws.cell(row=cur, column=5, value=loai).border = bd
        ws.cell(row=cur, column=5).alignment = center
        ws.cell(row=cur, column=6, value=round(tl_sum, 2)).border = bd
        ws.cell(row=cur, column=6).number_format = '0.00'
        ws.cell(row=cur, column=6).alignment = right_center
        ws.cell(row=cur, column=7, value=gia_mau).border = bd
        ws.cell(row=cur, column=7).number_format = '#,##0'
        ws.cell(row=cur, column=7).alignment = right_center
        ws.cell(row=cur, column=8, value=tt_sum).border = bd
        ws.cell(row=cur, column=8).number_format = '#,##0'
        ws.cell(row=cur, column=8).alignment = right_center
        tong_tt += tt_sum
        cur += 1
        stt += 1

    ws.cell(row=cur, column=4, value="").border = Border(top=thin, bottom=thin, left=thin, right=None)
    ws.cell(row=cur, column=5, value="").border = Border(top=thin, bottom=thin, left=None, right=None)
    ws.cell(row=cur, column=6, value="").border = Border(top=thin, bottom=thin, left=None, right=None)
    ws.cell(row=cur, column=7, value="TỔNG").border = Border(top=thin, bottom=thin, left=None, right=None)
    ws.cell(row=cur, column=7).font = Font(bold=True)
    ws.cell(row=cur, column=7).alignment = center
    ws.cell(row=cur, column=7).fill = fill_xam
    a = ws.cell(row=cur, column=8, value=tong_tt)
    a.font = Font(bold=True, color="FF0000")
    a.fill = fill_xam
    a.border = bd
    a.number_format = '#,##0'
    a.alignment = right_center

    ws.column_dimensions['A'].width = 11
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 13
    ws.column_dimensions['H'].width = 15

# ===================================================================
# GIAO DIEN FIX CHUAN FLET 0.28 - 100% KHONG LOI TAB
# ===================================================================
def main(page: ft.Page):
    page.title = "Tool Tach Sheet Safe"
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 15
    page.theme_mode = ft.ThemeMode.LIGHT

    target_dir = os.getcwd()

    title_text = ft.Text(
        "TOOL TÁCH SHEET EXCEL (BẢN CHUẨN 0.28)", 
        size=20,
        weight=ft.FontWeight.BOLD
    )
    
    info_text = ft.Text(f"Thư mục: {target_dir}", color=ft.Colors.BLUE, size=12)

    file_dropdown = ft.Dropdown(
        label="Chọn file Excel",
        options=[],
        disabled=True,
        width=400
    )
    
    log_box = ft.TextField(
        label="Nhật ký", 
        multiline=True, 
        min_lines=8, 
        max_lines=12, 
        read_only=True,
        value="",
        width=400
    )
    
    progress_bar = ft.ProgressBar(width=400, visible=False)

    def write_log(message):
        log_box.value += f"{message}\n"
        page.update()

    def quet_tim_file_excel(e=None):
        log_box.value = ""
        write_log("🔍 Đang quét file .xlsx...")
        try:
            cac_file = [f for f in os.listdir(target_dir) if f.endswith('.xlsx') and not f.startswith('OKKK')]
            if cac_file:
                cac_file.sort(key=lambda x: os.path.getmtime(os.path.join(target_dir, x)), reverse=True)
                file_dropdown.options = [ft.dropdown.Option(f) for f in cac_file]
                file_dropdown.value = cac_file[0]
                file_dropdown.disabled = False
                btn_run.disabled = False
                write_log(f"-> Tìm thấy {len(cac_file)} file. Đã chọn: {cac_file[0]}")
            else:
                file_dropdown.options = []
                file_dropdown.value = None
                file_dropdown.disabled = True
                btn_run.disabled = True
                write_log("⚠️ Chưa có file Excel nào!")
        except Exception as ex:
            write_log(f"❌ Lỗi: {ex}\n{traceback.format_exc()}")
        page.update()

    def bat_dau_xu_ly(e):
        selected_file = file_dropdown.value
        if not selected_file:
            return
            
        file_path = os.path.join(target_dir, selected_file)
        progress_bar.visible = True
        btn_run.disabled = True
        page.update()
        
        try:
            write_log(f"\n[1] Đang xử lý: {selected_file}")
            df_all = pd.read_excel(file_path, header=None, dtype=str)
            
            header_row = None
            for i in range(min(20, len(df_all))):
                row_text = ' '.join([str(x) for x in df_all.iloc[i, :8] if pd.notna(x)])
                row_norm = chuan_hoa(row_text)
                if any(k in row_norm for k in ['ngay', 'ten', 'cccd', 'loai vang', 'ky hieu']):
                    header_row = i
                    break
            
            if header_row is None:
                for i in range(min(20, len(df_all))):
                    non_empty = sum([1 for x in df_all.iloc[i, :8] if pd.notna(x) and str(x).strip()!=''])
                    if non_empty >= 5:
                        header_row = i
                        break
            
            if header_row is None:
                header_row = 0
            
            df = df_all.iloc[header_row+1:, :8].copy()
            df.columns = ['Ngay', 'Ten', 'CCCD', 'DiaChi', 'LoaiVang', 'TLVang', 'GiaVang', 'ThanhTien']
            df = df.dropna(subset=['Ten', 'LoaiVang'], how='all')
            df = df[df['Ten'].astype(str).str.strip() != '']
            
            df['Ngay'] = pd.to_datetime(df['Ngay'], dayfirst=True, errors='coerce')
            if df['Ngay'].isna().sum() > len(df) * 0.5:
                df['Ngay'] = pd.to_datetime(df['Ngay'], errors='coerce')
            
            df = df[df['Ngay'].notna()].copy()
            df = df.sort_values('Ngay').reset_index(drop=True)
            
            for col in ['TLVang', 'GiaVang', 'ThanhTien']:
                df[col] = df[col].apply(parse_number)

            df['Ngay_date'] = df['Ngay'].dt.date
            thang, nam = df.iloc[0]['Ngay'].month, df.iloc[0]['Ngay'].year
            
            output_name = f"OKKK-T{thang}-{nam}.xlsx"
            output_path = os.path.join(target_dir, output_name)
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as w:
                df.to_excel(w, sheet_name=f"T{thang}.{nam}", index=False)
                write_sheet(w.sheets[f"T{thang}.{nam}"], df)
                for ng, dfg in df.groupby('Ngay_date'):
                    ten = ng.strftime('%d.%m')
                    dfg.to_excel(w, sheet_name=ten, index=False)
                    write_sheet(w.sheets[ten], dfg)
            
            write_log("\n============================")
            write_log("🎉 TÁCH SHEET THÀNH CÔNG! 🎉")
            write_log("============================")
            write_log(f"File: {output_name}")
            quet_tim_file_excel()
            
        except Exception as ex:
            write_log(f"\n❌ LỖI: {ex}\n{traceback.format_exc()}")
        finally:
            progress_bar.visible = False
            btn_run.disabled = False
            page.update()

    btn_refresh = ft.ElevatedButton("Quét lại file", on_click=quet_tim_file_excel, icon=ft.Icons.REFRESH)
    btn_run = ft.ElevatedButton("Bắt đầu Tách Sheet", disabled=True, on_click=bat_dau_xu_ly, bgcolor=ft.Colors.BLUE_700, color="white", icon=ft.Icons.CUT)

    page.add(
        ft.Column([
            title_text,
            info_text,
            ft.Divider(),
            btn_refresh,
            file_dropdown,
            btn_run,
            progress_bar,
            log_box
        ], spacing=10, scroll=ft.ScrollMode.AUTO)
    )
    
    quet_tim_file_excel()

ft.app(target=main)
