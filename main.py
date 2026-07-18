# ===================================================================
# GIAO DIỆN ĐỒ HỌA ỨNG DỤNG (FLET UI APP) - ĐÃ SỬA LỖI VIẾT HOA COLORS
# ===================================================================
def main(page: ft.Page):
    page.title = "Tool Tách Sheet - Bản 8 Cột Full Ghi Chú"
    page.scroll = ft.ScrollMode.AUTO
    page.theme_mode = ft.ThemeMode.DARK # Giao diện tối hiện đại
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.padding = 20

    # Các thành phần hiển thị
    title_text = ft.Text(
        "BẮT ĐẦU CHẠY TOOL TÁCH SHEET", 
        style=ft.TextThemeStyle.HEADLINE_MEDIUM, 
        weight=ft.FontWeight.BOLD, 
        color=ft.Colors.BLUE_200
    )
    sub_title = ft.Text("BẢN 8 CỘT FULL GHI CHÚ", style=ft.TextThemeStyle.TITLE_MEDIUM, color=ft.Colors.GREY_400)
    
    selected_file_text = ft.Text("Chưa chọn file Excel nào", italic=True, color=ft.Colors.AMBER_300)
    log_box = ft.TextField(
        label="Nhật ký hệ thống (Logs)", 
        multiline=True, 
        min_lines=10, 
        max_lines=15, 
        read_only=True,
        text_size=13,
        value=""
    )
    
    progress_bar = ft.ProgressBar(width=400, color="blue", visible=False)

    def write_log(message):
        log_box.value += f"{message}\n"
        page.update()

    # Hàm thực hiện xử lý tách sheet khi đã có file
    def bat_dau_xu_ly(file_path):
        progress_bar.visible = True
        log_box.value = "" # Xóa log cũ
        page.update()
        
        try:
            write_log(f"[1] Đọc file: {os.path.basename(file_path)}")
            write_log(f"    Đường dẫn: {file_path}")
            
            df_all = pd.read_excel(file_path, header=None, dtype=str)
            write_log(f" -> OK: File có {df_all.shape[0]} dòng, {df_all.shape[1]} cột")
            
            header_row = None
            for i in range(min(20, len(df_all))):
                row_text = ' '.join([str(x) for x in df_all.iloc[i, :8] if pd.notna(x)])
                row_norm = chuan_hoa(row_text)
                if any(keyword in row_norm for keyword in ['ngay', 'ten', 'cccd', 'loai vang', 'ky hieu']):
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
                write_log(" -> Không tìm thấy header rõ ràng, sử dụng dòng đầu tiên")
            else:
                write_log(f" -> Tìm thấy header ở dòng {header_row + 1}")
            
            df = df_all.iloc[header_row+1:, :8].copy()
            df.columns = ['Ngay', 'Ten', 'CCCD', 'DiaChi', 'LoaiVang', 'TLVang', 'GiaVang', 'ThanhTien']
            
            df = df.dropna(subset=['Ten', 'LoaiVang'], how='all')
            df = df[df['Ten'].astype(str).str.strip() != '']
            
            df['Ngay'] = pd.to_datetime(df['Ngay'], dayfirst=True, errors='coerce')
            if df['Ngay'].isna().sum() > len(df) * 0.5:
                df['Ngay'] = pd.to_datetime(df['Ngay'], errors='coerce')
            
            df = df[df['Ngay'].notna()].copy()
            df = df.sort_values('Ngay').reset_index(drop=True)
            
            write_log(f"[2] Lọc xong: {len(df)} dòng hợp lệ")
            if len(df) > 0:
                write_log(f" -> Ngày đầu: {df.iloc[0]['Ngay'].strftime('%d/%m/%Y')}")
                write_log(f" -> Ngày cuối: {df.iloc[-1]['Ngay'].strftime('%d/%m/%Y')}")
            else:
                write_log("!!! Không có dữ liệu ngày hợp lệ để xử lý!")
                progress_bar.visible = False
                page.update()
                return

            for col in ['TLVang', 'GiaVang', 'ThanhTien']:
                df[col] = df[col].apply(parse_number)

            df['Ngay_date'] = df['Ngay'].dt.date
            thang, nam = df.iloc[0]['Ngay'].month, df.iloc[0]['Ngay'].year
            
            output_name = f"OKKK-T{thang}-{nam}.xlsx"
            output_dir = os.path.dirname(file_path)
            output_path = os.path.join(output_dir, output_name)
            
            write_log(f"\n[5] Ghi file kết quả: {output_path}")
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as w:
                df.to_excel(w, sheet_name=f"T{thang}.{nam}", index=False)
                write_sheet(w.sheets[f"T{thang}.{nam}"], df)
                write_log(f"[6] Đã viết sheet tổng: T{thang}.{nam}")
                
                for ng, dfg in df.groupby('Ngay_date'):
                    ten = ng.strftime('%d.%m')
                    dfg.to_excel(w, sheet_name=ten, index=False)
                    write_sheet(w.sheets[ten], dfg)
                    write_log(f"[7] Đã viết sheet ngày: {ten}")
            
            write_log("\n====================================")
            write_log("🎉 THÀNH CÔNG RỒI B ƠI! 🎉")
            write_log("====================================")
            write_log(f"File kết quả lưu tại thư mục cũ với tên:\n -> {output_name}")
            
        except Exception as e:
            write_log(f"\n❌ !!! LỖI HỆ THỐNG: {e}")
            write_log(traceback.format_exc())
        finally:
            progress_bar.visible = False
            page.update()

    # Quản lý sự kiện chọn file từ FilePicker
    def pick_files_result(e: ft.FilePickerResultEvent):
        if e.files:
            file_path = e.files[0].path
            selected_file_text.value = f"Đã chọn file: {e.files[0].name}"
            btn_run.disabled = False
            btn_run.data = file_path # Lưu đường dẫn file vào nút bấm để tí gọi
            write_log(f"Đã nạp file: {file_path}")
        else:
            selected_file_text.value = "Hủy chọn file."
            btn_run.disabled = True
        page.update()

    file_picker = ft.FilePicker(on_result=pick_files_result)
    page.overlay.append(file_picker)

    # Nút bấm chính trên giao diện
    btn_select = ft.ElevatedButton(
        "Chọn file Excel (.xlsx)",
        icon=ft.icons.UPLOAD_FILE,
        on_click=lambda _: file_picker.pick_files(allowed_extensions=["xlsx"]),
        style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8))
    )
    
    btn_run = ft.ElevatedButton(
        "Bắt đầu Tách Sheet ngay",
        icon=ft.icons.PLAY_ARROW,
        disabled=True,
        color=ft.Colors.WHITE,
        bgcolor=ft.Colors.GREEN_700,
        on_click=lambda e: bat_dau_xu_ly(e.control.data),
        style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8))
    )

    # Đưa các thành phần lên màn hình app
    page.add(
        ft.Column(
            [
                title_text,
                sub_title,
                ft.Divider(height=20, color="grey"),
                btn_select,
                selected_file_text,
                ft.VerticalDivider(height=10),
                btn_run,
                progress_bar,
                ft.VerticalDivider(height=10),
                log_box
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=15
        )
    )

ft.app(target=main)
