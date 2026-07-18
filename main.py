write_log(f"File ket qua luu tai cung thu muc: {output_name}")
            
        except Exception as e:
            write_log(f"\n LOI HE THONG: {e}")
            write_log(traceback.format_exc())
        finally:
            progress_bar.visible = False
            page.update()

    # Hàm tự quét file Excel xung quanh ứng dụng
    def quet_tim_file_excel(e=None):
        try:
            # Quét ở thư mục hiện tại của app
            thu_muc_quet = os.getcwd()
            write_log(f"Dang tim file .xlsx tai: {thu_muc_quet}")
            
            cac_file = [f for f in os.listdir(thu_muc_quet) if f.endswith('.xlsx') and not f.startswith('OKKK')]
            
            if cac_file:
                # Lấy file đầu tiên tìm được
                file_target = cac_file[0]
                current_file_path[0] = os.path.join(thu_muc_quet, file_target)
                selected_file_text.value = f"Tim thay file: {file_target}"
                btn_run.disabled = False
            else:
                # Nếu không thấy, thử tìm thêm ở thư mục cha hoặc Download (nếu có quyền)
                selected_file_text.value = "Khong tim thay file .xlsx nao! Hay dat file Excel vao thu muc app."
                btn_run.disabled = True
        except Exception as ex:
            selected_file_text.value = f"Loi khi quet file: {ex}"
        page.update()

    btn_refresh = ft.ElevatedButton(
        "Quet lai thu muc (Refresh)",
        on_click=quet_tim_file_excel
    )
    
    btn_run = ft.ElevatedButton(
        "Bat dau Tach Sheet ngay",
        disabled=True,
        on_click=lambda _: bat_dau_xu_ly(current_file_path[0])
    )

    page.add(
        title_text,
        selected_file_text,
        btn_refresh,
        btn_run,
        progress_bar,
        log_box
    )
    
    # Chạy quét file ngay khi mở ứng dụng
    quet_tim_file_excel()

ft.app(target=main)
