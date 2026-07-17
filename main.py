import pandas as pd # pandas: thư viện xử lý bảng tính. Coi nó như Excel trong Python. Dữ liệu sẽ nằm trong DataFrame (bảng)
import os # Thư viện làm việc với file/thư mục: kiểm tra đường dẫn, liệt kê file, nối đường dẫn...
import sys # Thư viện hệ thống: dùng sys.exit() để thoát chương trình khi lỗi
import re # Regex: xử lý chuỗi, ví dụ thay "nhiều   khoảng   trắng" thành "một khoảng trắng"
import unicodedata # FIX LỖI TIẾNG VIỆT QUAN TRỌNG NHẤT. 'hóa' và 'hoá' nhìn giống nhau 100% nhưng mã Unicode khác nhau -> so sánh == sẽ sai
from datetime import datetime # Xử lý ngày giờ, dùng để format tên sheet
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment # Bộ đồ nghề làm đẹp Excel: Font=chữ, Border=viền, Fill=tô màu, Alignment=căn lề
from collections import defaultdict # Từ điển tự động: khi truy cập key chưa có, nó tự tạo giá trị mặc định, đỡ phải if else

print("="*50)
print("BẮT ĐẦU CHẠY TOOL TÁCH SHEET - BẢN 8 CỘT FULL GHI CHÚ")
print("="*50)

# ===================================================================
# HÀM 0: CHUẨN HÓA TIẾNG VIỆT - HÀM QUAN TRỌNG NHẤT ĐỂ FIX LỖI TÌM HEADER
# Vấn đề: 'hóa đơn' (ó + a) và 'hoá đơn' (o + á) nhìn y hệt nhau nhưng mã Unicode khác nhau
# Nếu không chuẩn hóa, code tìm chữ 'hoa don' sẽ không bao giờ thấy 'hoá đơn'
# Giải pháp: Biến hết về không dấu, chữ thường, 1 khoảng trắng
# ===================================================================
def chuan_hoa(text):
    # pd.isna() là hàm của pandas để check ô trống (NaN). Khác None của Python thường
    # Ví dụ ô Excel trống -> pd.isna sẽ True
    if pd.isna(text) or text is None:
        return "" # Trả về chuỗi rỗng để các bước sau không bị lỗi
    # Bỏ ký tự xuống dòng trong tiêu đề. Nhiều file ghi 'Ký hiệu\nhóa đơn' -> thành 'Ký hiệu hóa đơn'
    text = str(text).replace('\n',' ').replace('\r',' ')
    # NFD = Normalization Form Decomposed: tách chữ 'á' thành 2 phần: 'a' + dấu sắc riêng biệt
    text = unicodedata.normalize('NFD', text)
    # unicodedata.category(c) == 'Mn' nghĩa là ký tự này là dấu (Mark, nonspacing)
    # Dòng này giữ lại chữ cái, xóa hết dấu. 'a' + '´' -> chỉ giữ 'a'
    text = ''.join(c for c in text if unicodedata.category(c)!= 'Mn')
    # lower() đưa về chữ thường để 'Loại Vàng' == 'loai vang', strip() xóa space thừa 2 đầu
    text = text.lower().strip()
    # \s+ nghĩa là 1 hoặc nhiều khoảng trắng/tab. Thay tất cả thành 1 dấu cách duy nhất
    # 'loai    vang' -> 'loai vang'
    text = re.sub(r'\s+', ' ', text)
    return text

# ===================================================================
# HÀM 1: TÌM FILE EXCEL TRONG NHIỀU THƯ MỤC
# Tại sao cần? Trên điện thoại Android, đường dẫn hay thay đổi, bạn không nhớ để file ở đâu
# Thay vì bắt bạn gõ tay, code sẽ tự đi quét các chỗ hay để nhất
# ===================================================================
def find_excel_files():
    """Tìm file Excel trong nhiều thư mục khác nhau"""
    excel_files = [] # Danh sách rỗng để chứa kết quả
    
    # Danh sách các thư mục phổ biến trên Android. Bạn có thể thêm/bớt ở đây
    search_paths = [
        "/storage/emulated/0/AAA HỌC PYTHON/",  # Thư mục bạn hay để
        "/storage/emulated/0/Download/", # Thư mục tải về
        "/storage/emulated/0/Documents/", # Tài liệu
        "/storage/emulated/0/", # Bộ nhớ trong gốc
        "/sdcard/", # Tên khác của bộ nhớ trong trên một số máy
        # '__file__' là biến đặc biệt chỉ tồn tại khi chạy file .py. Nó chứa đường dẫn file .py đang chạy
        # os.path.abspath(__file__) lấy đường dẫn tuyệt đối, dirname lấy thư mục chứa nó
        os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else "",
        os.getcwd(),  # getcwd = get current working directory: thư mục mà Python đang đứng
    ]
    
    print("\n[0] Đang tìm file Excel trong các thư mục:")
    for path in search_paths: # Duyệt từng thư mục trong danh sách
        # Nếu path rỗng hoặc thư mục không tồn tại trên máy này thì bỏ qua
        if not path or not os.path.exists(path):
            continue
        print(f"    - {path}")
        try: # try...except để nếu không có quyền đọc thư mục thì không bị crash
            for f in os.listdir(path): # os.listdir: liệt kê tất cả file và thư mục con bên trong path
                # Điều kiện lọc file:
                # 1. f.endswith('.xlsx'): chỉ lấy file Excel mới
                # 2. "OKKK" not in f: bỏ qua file kết quả cũ mà tool đã tạo ra trước đó, tránh quét vòng lặp
                # 3. "~$" not in f: bỏ file tạm của Excel khi đang mở (Excel tạo file ~$ten_file.xlsx)
                if f.endswith('.xlsx') and "OKKK" not in f and "~$" not in f:
                    full_path = os.path.join(path, f) # Nối thư mục + tên file. Ví dụ: "/Download/" + "data.xlsx" -> "/Download/data.xlsx"
                    excel_files.append(full_path)
                    print(f"        Tìm thấy: {f}")
        except Exception as e:
            # Không có quyền đọc thư mục thì pass = bỏ qua, đi tiếp thư mục khác
            pass
    
    # Nếu quét hết mà vẫn không thấy file nào
    if not excel_files:
        print("\n[0] KHÔNG TÌM THẤY FILE EXCEL!")
        print("Bạn có thể:")
        print("1. Nhập đường dẫn thư mục chứa file Excel")
        print("2. Hoặc Enter để thoát")
        
        custom_path = input("\nNhập đường dẫn thư mục (ví dụ: /storage/emulated/0/AAA HỌC PYTHON/): ").strip()
        
        if custom_path and os.path.exists(custom_path): # Kiểm tra đường dẫn bạn nhập có tồn tại không
            print(f"\nĐang tìm trong: {custom_path}")
            try:
                for f in os.listdir(custom_path):
                    if f.endswith('.xlsx') and "OKKK" not in f and "~$" not in f:
                        full_path = os.path.join(custom_path, f)
                        excel_files.append(full_path)
                        print(f"    Tìm thấy: {f}")
            except:
                pass
    
    return excel_files # Trả về danh sách đường dẫn đầy đủ, ví dụ: ['/Download/a.xlsx', '/Documents/b.xlsx']

# ===================================================================
# HÀM 2: CHỌN FILE KHI CÓ NHIỀU FILE
# Nếu tìm thấy 5 file .xlsx thì phải hỏi bạn muốn xử lý file nào
# ===================================================================
def select_file(excel_files):
    """Chọn file từ danh sách"""
    if not excel_files: # Danh sách rỗng
        return None
    
    if len(excel_files) == 1: # Chỉ có 1 file thì chọn luôn, khỏi hỏi cho nhanh
        # os.path.basename: chỉ lấy tên file, bỏ phần đường dẫn. '/a/b/c.xlsx' -> 'c.xlsx'
        print(f"\n[0] Chọn file: {os.path.basename(excel_files[0])}")
        return excel_files[0]
    
    print(f"\n[0] Tìm thấy {len(excel_files)} file Excel:")
    # enumerate(excel_files) vừa cho số thứ tự i (bắt đầu từ 0), vừa cho giá trị f (đường dẫn)
    for i, f in enumerate(excel_files):
        print(f"    {i+1}. {os.path.basename(f)}") # i+1 để hiển thị cho người dùng bắt đầu từ 1 cho dễ hiểu
        print(f"       ({os.path.dirname(f)})") # dirname: lấy đường dẫn thư mục
    
    print(f"\nNhập số thứ tự file muốn xử lý (1-{len(excel_files)}) hoặc Enter để chọn file đầu tiên:")
    try:
        choice = input().strip() # Đọc lựa chọn của bạn từ bàn phím
        # isdigit(): kiểm tra chuỗi có phải toàn số không. "2" -> True, "abc" -> False
        if choice and choice.isdigit() and 1 <= int(choice) <= len(excel_files):
            return excel_files[int(choice)-1] # -1 vì người dùng gõ 1 nhưng list bắt đầu từ 0
        else:
            return excel_files[0] # Nhấn Enter hoặc gõ bậy thì lấy file đầu tiên cho đỡ lỗi
    except:
        return excel_files[0]

# ===================================================================
# CHƯƠNG TRÌNH CHÍNH - Đặt trong try...except để không bị tắt đột ngột
# Nếu có lỗi thì in ra màn hình cho bạn đọc, thay vì cửa sổ đen biến mất ngay
# ===================================================================
try:
    # --- BƯỚC 0: TÌM VÀ CHỌN FILE ---
    excel_files = find_excel_files() # Gọi hàm tìm file ở trên
    
    if not excel_files: # Không tìm thấy file nào
        print("\n" + "="*50)
        print("KHÔNG TÌM THẤY FILE EXCEL!")
        print("="*50)
        print("\nCách khắc phục:")
        print("1. Mở Solid Explorer, tìm thư mục chứa file Excel")
        print("2. Copy đường dẫn thư mục (ví dụ: /storage/emulated/0/AAA HỌC PYTHON/)")
        print("3. Chạy lại file .py và nhập đường dẫn đó")
        input("\nNhấn Enter để thoát...")
        sys.exit() # Lệnh thoát chương trình hoàn toàn
    
    INPUT_FILE = select_file(excel_files) # Gọi hàm chọn file
    
    if not INPUT_FILE:
        input("\nNhấn Enter để thoát...")
        sys.exit()
    
    print(f"\n[1] Đọc file: {os.path.basename(INPUT_FILE)}")
    print(f"    Đường dẫn: {INPUT_FILE}")
    
    # --- BƯỚC 1: ĐỌC FILE THÔ, CHƯA ĐẶT HEADER ---
    # header=None: bảo pandas ĐỪNG tự lấy dòng 1 làm tiêu đề, cứ đọc hết thành dữ liệu thô
    # dtype=str: ép đọc hết về dạng chuỗi, để số CCCD '012345' không bị biến thành 12345 mất số 0
    df_all = pd.read_excel(INPUT_FILE, header=None, dtype=str)
    print(f" -> OK: File có {df_all.shape[0]} dòng, {df_all.shape[1]} cột") # shape[0]=số dòng, shape[1]=số cột
    
    # --- BƯỚC 2: TÌM DÒNG TIÊU ĐỀ THẬT SỰ (FIX LỖI HÓA ĐƠN Ở ĐÂY) ---
    header_row = None # Biến lưu vị trí dòng header, ban đầu chưa biết nên để None
    # Chỉ quét 20 dòng đầu cho nhanh, vì header luôn nằm trên đầu, không ai để header ở dòng 100
    for i in range(min(20, len(df_all))):
        # df_all.iloc[i, :8]: lấy dòng i, 8 cột đầu. iloc là truy cập theo vị trí số [dòng, cột]
        # Lọc pd.notna(x) để bỏ ô trống, rồi nối tất cả thành 1 chuỗi dài để tìm từ khóa
        row_text = ' '.join([str(x) for x in df_all.iloc[i, :8] if pd.notna(x)])
        row_norm = chuan_hoa(row_text) # Chuẩn hóa không dấu: 'Ngày, Tên, CCCD, Loại Vàng' -> 'ngay ten cccd loai vang'
        # Nếu dòng này chứa các từ khóa đặc trưng của header thì đây chính là header
        # any(...): chỉ cần 1 từ khóa xuất hiện là True
        if any(keyword in row_norm for keyword in ['ngay', 'ten', 'cccd', 'loai vang', 'ky hieu']):
            header_row = i
            break # Tìm thấy rồi thì break = dừng vòng lặp ngay, không quét tiếp
    
    if header_row is None:
        # Phương án dự phòng: nếu không tìm thấy bằng từ khóa, thì tìm dòng có nhiều ô có dữ liệu nhất (>=5 ô)
        for i in range(min(20, len(df_all))):
            non_empty = sum([1 for x in df_all.iloc[i, :8] if pd.notna(x) and str(x).strip()!='']) # Đếm số ô không rỗng
            if non_empty >= 5: # Dòng nào có >=5 ô có dữ liệu thì khả năng cao là header
                header_row = i
                break
    
    if header_row is None:
        header_row = 0 # Vẫn không thấy thì mặc định lấy dòng đầu tiên làm header cho đỡ lỗi
        print(" -> Không tìm thấy header rõ ràng, sử dụng dòng đầu tiên")
    else:
        print(f" -> Tìm thấy header ở dòng {header_row + 1} (tính từ 1 cho dễ hiểu)")
    
    # --- BƯỚC 3: CẮT LẤY DỮ LIỆU THẬT VÀ ĐẶT TÊN CỘT NGẮN GỌN ---
    # iloc[header_row+1:, :8]: lấy từ dòng SAU header trở xuống hết file, và chỉ lấy 8 cột đầu
    # .copy(): tạo bản sao để tránh lỗi SettingWithCopyWarning của pandas
    df = df_all.iloc[header_row+1:, :8].copy()
    # Đặt lại tên cột ngắn gọn, không dấu để gõ code cho nhanh. Đây là tên NỘI BỘ, khác với tên hiển thị ra Excel
    # Sau này code sẽ dùng df['Ngay'] thay vì df['Ngày lập hóa đơn...']
    df.columns = ['Ngay', 'Ten', 'CCCD', 'DiaChi', 'LoaiVang', 'TLVang', 'GiaVang', 'ThanhTien']
    
    # --- BƯỚC 4: LỌC DỮ LIỆU RÁC ---
    # dropna(subset=['Ten','LoaiVang'], how='all'): Bỏ dòng mà CẢ 2 cột Ten và LoaiVang đều trống
    # how='all' = chỉ bỏ khi TẤT CẢ cột trong subset đều trống. Nếu 1 trong 2 còn thì vẫn giữ
    df = df.dropna(subset=['Ten', 'LoaiVang'], how='all')
    # astype(str): ép về chuỗi, .str.strip(): xóa space, != '': chỉ giữ dòng có Tên thật sự
    df = df[df['Ten'].astype(str).str.strip() != '']
    
    # --- BƯỚC 5: XỬ LÝ NGÀY THÁNG ---
    # pd.to_datetime: ép cột Ngày về dạng datetime thật sự của Python để sắp xếp và groupby được
    # dayfirst=True: ưu tiên ngày trước tháng sau (kiểu VN 25/12/2024). errors='coerce': nếu lỗi parse thì biến thành NaT (rỗng) chứ không crash
    df['Ngay'] = pd.to_datetime(df['Ngay'], dayfirst=True, errors='coerce')
    
    # Nếu lỗi quá nhiều (>50% dòng bị NaT), thử ép kiểu lại không có dayfirst (phòng file dạng Mỹ MM/DD/YYYY)
    if df['Ngay'].isna().sum() > len(df) * 0.5:
        df['Ngay'] = pd.to_datetime(df['Ngay'], errors='coerce')
    
    df = df[df['Ngay'].notna()].copy() # notna() = not NaN: chỉ giữ lại dòng có Ngày hợp lệ
    df = df.sort_values('Ngay').reset_index(drop=True) # sort_values: sắp xếp tăng dần theo Ngày, reset_index(drop=True): đánh lại số thứ tự 0,1,2... cho đẹp
    
    print(f"[2] Lọc xong: {len(df)} dòng hợp lệ")
    if len(df) > 0:
        # strftime('%d/%m/%Y'): format datetime thành chuỗi ngày/tháng/năm
        print(f" -> Ngày đầu: {df.iloc[0]['Ngay'].strftime('%d/%m/%Y')}")
        print(f" -> Ngày cuối: {df.iloc[-1]['Ngay'].strftime('%d/%m/%Y')}")
    
    # --- BƯỚC 6: XỬ LÝ SỐ - FIX LỖI DẤU PHẨY, CHỮ ---
    def parse_number(s):
        """Hàm nhỏ chuyển chuỗi tiền tệ '1,200,000' hoặc '1.200.000 đ' về số 1200000.0"""
        if pd.isna(s): # Ô trống
            return 0
        if isinstance(s, (int, float)): # isinstance: kiểm tra kiểu dữ liệu. Nếu đã là số rồi thì trả về luôn
            return s
        # Xóa dấu phẩy ngăn cách nghìn và khoảng trắng. '1,200,000' -> '1200000'
        s = str(s).strip().replace(',', '').replace(' ', '')
        try:
            return float(s) # Ép về float. Dùng float thay vì int vì có cột TL Vàng có số lẻ (1.5 chỉ)
        except:
            return 0 # Nếu vẫn lỗi (ví dụ chữ 'abc') thì trả về 0 cho an toàn

    for col in ['TLVang', 'GiaVang', 'ThanhTien']: # Duyệt 3 cột số
        df[col] = df[col].apply(parse_number) # apply: áp dụng hàm parse_number cho TỪNG ô trong cột

    # ========= ĐỊNH NGHĨA STYLE DÙNG CHUNG CHO ĐẸP - Sửa màu ở đây =========
    thin = Side(style='thin') # Kiểu viền mảnh. Có thể đổi thành 'thick' (dày), 'dashed' (nét đứt)
    bd = Border(left=thin,right=thin,top=thin,bottom=thin) # Tạo khung viền 4 phía
    fh = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid") # Màu xanh nhạt cho header. Mã màu HEX
    fill_xam = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid") # Màu xám cho dòng tổng
    fill_trang = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid") # Màu trắng
    center = Alignment(horizontal='center', vertical='center') # Căn giữa ngang + dọc
    right_center = Alignment(horizontal='right', vertical='center') # Căn phải, giữa dòng (dùng cho cột tiền cho thẳng hàng)
    center_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True) # wrap_text=True: tự động xuống dòng nếu chữ dài quá ô

    # ===================================================================
    # HÀM 3: VẼ LẠI 1 SHEET CHO ĐẸP - Hàm quan trọng nhất, dài nhất
    # ws: worksheet (sheet Excel cần vẽ), df_day: DataFrame chỉ chứa dữ liệu của 1 ngày
    # ===================================================================
    def write_sheet(ws, df_day):
        """Vẽ lại sheet cho đẹp, gộp hóa đơn, tính tổng"""
        ws.delete_rows(1, ws.max_row) # Xóa hết dữ liệu cũ mà pandas vừa ghi ra bằng to_excel, để mình vẽ lại theo ý mình cho đẹp
        # max_row là số dòng lớn nhất hiện có trong sheet
        
        # Tên header hiển thị ra Excel cho người dùng thấy, khác với tên nội bộ ngắn gọn ở trên
        hs = ['Ngày', 'Họ Tên', 'CCCD', 'Địa Chỉ', 'Loại Vàng', 'TL Vàng', 'Giá Vàng', 'Thành Tiền']
        for i,h in enumerate(hs): # enumerate: vừa lấy vị trí i (0-7), vừa lấy giá trị h ('Ngày'...)
            c = ws.cell(row=1, column=i+1, value=h) # Ghi vào dòng 1, cột i+1. Excel tính cột từ 1, không phải 0
            c.font = Font(bold=True) # In đậm
            c.alignment = center_wrap
            c.border = bd # Kẻ khung
            c.fill = fh # Tô màu xanh header
        ws.row_dimensions[1].height = 26 # Cho dòng header cao 26 để không bị chật. row_dimensions[1] là chiều cao dòng 1

        cur = 2 # cur: current row - dòng hiện tại đang ghi, bắt đầu từ dòng 2 (sau header)
        groups=[] # groups: danh sách các hóa đơn, mỗi hóa đơn là 1 list các món. Ví dụ: [[món1, món2], [món3]]
        cg=[] # cg: current group - hóa đơn hiện tại đang gom dở
        last_key = None # Khóa của hóa đơn trước để so sánh xem có phải cùng hóa đơn không

        for _, rec in df_day.iterrows(): # iterrows(): duyệt từng dòng dữ liệu. rec là 1 dòng dạng Series (giống dict)
            # Tạo khóa để nhận diện 1 hóa đơn: nếu giống nhau Ngày + Tên + CCCD thì là cùng 1 hóa đơn mua nhiều món
            # Dùng cho file 8 cột (không có Ký hiệu, Số HĐ nên phải dùng combo này)
            key = (rec['Ngay'], str(rec['Ten']).strip(), str(rec['CCCD']).strip())
            if key!= last_key: # Gặp hóa đơn mới (khác ngày hoặc khác khách)
                if cg: groups.append(cg) # Nếu cg đang có món cũ thì lưu hóa đơn cũ vào danh sách groups
                cg=[rec] # Bắt đầu hóa đơn mới với món đầu tiên là rec hiện tại
                last_key = key # Cập nhật khóa mới
            else: 
                cg.append(rec) # Cùng hóa đơn thì thêm món vào hóa đơn hiện tại
        if cg: groups.append(cg) # Vòng lặp kết thúc, còn sót hóa đơn cuối cùng chưa kịp append thì thêm vào
        
        # --- VẼ TỪNG HÓA ĐƠN RA EXCEL ---
        for g in groups: # Duyệt từng hóa đơn (group)
            n = len(g) # n: số món trong hóa đơn này. Ví dụ hóa đơn mua 3 chỉ vàng thì n=3
            st = cur # st: start - dòng bắt đầu của hóa đơn này
            tot = 0 # tot: total - tổng tiền của hóa đơn này
            
            for j in range(n): # Duyệt từng món trong hóa đơn, j từ 0 đến n-1
                rec = g[j] # Lấy thông tin món thứ j
                rr = cur + j # rr: real row - dòng thật sự đang ghi trên Excel
                
                if j == 0: # Chỉ ghi thông tin chung (Ngày, Tên, CCCD, Địa chỉ) ở dòng đầu tiên của hóa đơn cho gọn, các dòng sau để trống rồi gộp lại
                    ws.cell(row=rr, column=1, value=rec['Ngay']).number_format = 'DD/MM/YYYY' # Định dạng ngày hiển thị
                    ws.cell(row=rr, column=2, value=rec['Ten'])
                    # CCCD hay bị đọc thành số '123456.0', nên split('.')[0] để cắt bỏ '.0'
                    ws.cell(row=rr, column=3, value=str(rec['CCCD']).split('.')[0] if pd.notna(rec['CCCD']) else "")
                    ws.cell(row=rr, column=4, value=rec['DiaChi'])
                
                # Thông tin riêng của từng món thì ghi hết cho từng dòng
                ws.cell(row=rr, column=5, value=rec['LoaiVang'])
                ws.cell(row=rr, column=6, value=rec['TLVang']).number_format = '0.00' # 2 số thập phân cho trọng lượng
                ws.cell(row=rr, column=7, value=rec['GiaVang']).number_format = '#,##0' # '#,##0' = số có dấu phẩy ngăn cách nghìn: 1,200,000
                ws.cell(row=rr, column=8, value=rec['ThanhTien']).number_format = '#,##0'
                
                for cc in range(1, 9): # Kẻ khung và căn lề cho 8 cột (1-8)
                    cell = ws.cell(row=rr, column=cc)
                    cell.border = bd
                    # Cột 6,7,8 là tiền và trọng lượng thì căn phải cho thẳng hàng số, còn lại căn giữa cho đẹp
                    cell.alignment = right_center if cc in [6, 7, 8] else center
                
                tot += rec['ThanhTien'] # Cộng dồn tổng tiền hóa đơn
            
            if n > 1: # Nếu hóa đơn có nhiều hơn 1 món thì gộp ô cột 1-4 lại cho đẹp, nhìn như 1 hóa đơn duy nhất
                for cc in [1, 2, 3, 4]: # Gộp 4 cột thông tin chung
                    # merge_cells: gộp từ dòng st đến st+n-1. Ví dụ st=2, n=3 -> gộp dòng 2-4
                    ws.merge_cells(start_row=st, start_column=cc, end_row=st+n-1, end_column=cc)
                    # Sau khi gộp phải căn lại lề cho ô đầu tiên (ô gốc), nếu không chữ sẽ lệch
                    cell = ws.cell(row=st, column=cc)
                    cell.alignment = center
            
            rt = cur + n # rt: row total - dòng tổng cộng của hóa đơn này, nằm ngay sau dòng món cuối
            for cc in range(1, 8): # 7 ô đầu của dòng tổng để trống nhưng vẫn kẻ khung và tô trắng cho đều bảng
                a = ws.cell(row=rt, column=cc, value="")
                a.border = bd
                a.fill = fill_trang
            a = ws.cell(row=rt, column=8, value=tot) # Ô cuối cùng ghi tổng tiền hóa đơn
            a.font = Font(bold=True, color="FF0000") # Chữ đỏ, in đậm cho nổi bật. FF0000 là mã màu đỏ
            a.fill = fill_xam # Nền xám
            a.border = bd
            a.number_format = '#,##0'
            a.alignment = right_center
            cur = rt + 1 # Nhảy xuống dòng tiếp theo để ghi hóa đơn kế tiếp. cur luôn là dòng trống tiếp theo
        
        # --- VẼ BẢNG TỔNG HỢP CUỐI NGÀY (BÁO CÁO) ---
        cur += 1 # Cách 1 dòng trắng cho thoáng mắt, dễ nhìn
        dt0 = df_day.iloc[0]['Ngay'] # Lấy ngày của nhóm này để ghi tiêu đề. iloc[0] là dòng đầu tiên
        ws.merge_cells(start_row=cur, start_column=4, end_row=cur, end_column=8) # Gộp 5 cột (4-8) làm tiêu đề tổng hợp
        # :02d nghĩa là format số luôn 2 chữ số, thêm số 0 đằng trước nếu cần. 1 -> 01, 12 -> 12
        c = ws.cell(row=cur, column=4, value=f"TỔNG HỢP CHI TIẾT HÀNG BÁN NGÀY {dt0.day:02d}/{dt0.month:02d}/{dt0.year}")
        c.font = Font(bold=True)
        c.fill = fh # Nền xanh giống header
        c.alignment = center
        c.border = bd
        for cc in [5, 6, 7, 8]: # Kẻ viền cho các ô bị gộp (Excel yêu cầu phải kẻ cho cả ô bị gộp thì mới hiện viền đủ)
            ws.cell(row=cur, column=cc).border = bd
        ws.row_dimensions[cur].height = 26 # Cho cao lên
        cur += 1
        
        titles = ["STT", "Loại Vàng", "TL Vàng", "Giá Vàng", "Thành Tiền"] # Tiêu đề bảng tổng hợp
        for i, t in enumerate(titles):
            a = ws.cell(row=cur, column=4+i, value=t) # Ghi từ cột 4 trở đi (D,E,F,G,H)
            a.font = Font(bold=True)
            a.fill = fh
            a.alignment = center
            a.border = bd
        ws.row_dimensions[cur].height = 26
        cur += 1
        
        # Dùng defaultdict để gom nhóm theo Loại Vàng và tính tổng số lượng + thành tiền
        # Cấu trúc: { 'Nhẫn 18K': [tổng_TL, tổng_TT, giá_mẫu], ... }
        # lambda: [0,0,0] là hàm tạo giá trị mặc định ban đầu khi gặp key mới
        agg = defaultdict(lambda: [0, 0, 0])
        for _, rec in df_day.iterrows():
            loai = str(rec['LoaiVang']).strip() # Lấy tên loại vàng, xóa khoảng trắng thừa
            agg[loai][0] += rec['TLVang'] # Cộng dồn trọng lượng vào phần tử 0
            agg[loai][1] += rec['ThanhTien'] # Cộng dồn thành tiền vào phần tử 1
            if agg[loai][2] == 0: # Lấy giá đầu tiên làm mẫu để hiển thị, nếu chưa có giá thì gán
                agg[loai][2] = rec['GiaVang']
        
        stt = 1 # Số thứ tự cho bảng tổng hợp
        tong_tt = 0 # Tổng thành tiền cuối ngày
        # .items(): duyệt từ điển, loai là key, (tl_sum, tt_sum, gia_mau) là value được bung ra
        for loai, (tl_sum, tt_sum, gia_mau) in agg.items():
            ws.cell(row=cur, column=4, value=stt).border = bd
            ws.cell(row=cur, column=4).alignment = center
            ws.cell(row=cur, column=5, value=loai).border = bd
            ws.cell(row=cur, column=5).alignment = center # Tên vàng căn giữa
            ws.cell(row=cur, column=6, value=round(tl_sum, 2)).border = bd # round(...,2): làm tròn 2 số lẻ
            ws.cell(row=cur, column=6).number_format = '0.00'
            ws.cell(row=cur, column=6).alignment = right_center
            ws.cell(row=cur, column=7, value=gia_mau).border = bd
            ws.cell(row=cur, column=7).number_format = '#,##0'
            ws.cell(row=cur, column=7).alignment = right_center
            ws.cell(row=cur, column=8, value=tt_sum).border = bd
            ws.cell(row=cur, column=8).number_format = '#,##0'
            ws.cell(row=cur, column=8).alignment = right_center
            tong_tt += tt_sum # Cộng vào tổng cuối ngày
            cur += 1
            stt += 1
        
        # Dòng tổng cộng cuối cùng của bảng tổng hợp
        # Border đặc biệt để không bị đè viền khi in
        ws.cell(row=cur, column=4, value="").border = Border(top=thin, bottom=thin, left=thin, right=None)
        ws.cell(row=cur, column=5, value="").border = Border(top=thin, bottom=thin, left=None, right=None)
        ws.cell(row=cur, column=6, value="").border = Border(top=thin, bottom=thin, left=None, right=None)
        ws.cell(row=cur, column=7, value="").border = Border(top=thin, bottom=thin, left=None, right=None)
        a = ws.cell(row=cur, column=8, value=tong_tt)
        a.font = Font(bold=True, color="FF0000") # Tổng cuối ngày màu đỏ cho nổi
        a.fill = fill_xam
        a.border = bd
        a.number_format = '#,##0'
        a.alignment = right_center
        
        # Chỉnh độ rộng cột cho vừa mắt, không bị cắt chữ. Đơn vị là ký tự
        ws.column_dimensions['A'].width = 11 # Ngày
        ws.column_dimensions['B'].width = 24 # Tên - cho rộng vì tên người dài
        ws.column_dimensions['C'].width = 15 # CCCD
        ws.column_dimensions['D'].width = 30 # Địa chỉ - dài nhất
        ws.column_dimensions['E'].width = 20 # Loại vàng
        ws.column_dimensions['F'].width = 10 # TL
        ws.column_dimensions['G'].width = 13 # Giá
        ws.column_dimensions['H'].width = 15 # Thành tiền
    
    # ========= BƯỚC 7: XUẤT FILE =========
    df['Ngay_date'] = df['Ngay'].dt.date # Tạo cột mới chỉ có ngày (không có giờ phút) để groupby theo ngày cho chính xác. .dt là accessor cho datetime
    thang, nam = df.iloc[0]['Ngay'].month, df.iloc[0]['Ngay'].year # Lấy tháng năm của dòng đầu tiên để đặt tên file
    
    output_name = f"OKKK-T{thang}-{nam}.xlsx" # Tên file kết quả, thêm OKKK để dễ phân biệt và tránh bị quét lại ở lần chạy sau
    output_dir = os.path.dirname(INPUT_FILE) # Lấy thư mục chứa file gốc để lưu file kết quả cùng chỗ cho dễ tìm
    output_path = os.path.join(output_dir, output_name) # Nối thành đường dẫn đầy đủ
    
    print(f"\n[5] Ghi file: {output_path}")
    
    # Mở 1 file Excel mới để ghi nhiều sheet, dùng engine openpyxl để giữ được style đẹp (màu, viền)
    # with ... as: tự động đóng file khi ghi xong, dù có lỗi cũng đóng, an toàn hơn open/close thủ công
    with pd.ExcelWriter(output_path, engine='openpyxl') as w:
        # Ghi sheet tổng trước (chứa tất cả các ngày)
        # index=False: đừng ghi cột số thứ tự 0,1,2 của pandas ra Excel, chỉ ghi dữ liệu thôi
        df.to_excel(w, sheet_name=f"T{thang}.{nam}", index=False)
        w.sheets[f"T{thang}.{nam}"] # w.sheets là dict chứa các sheet vừa tạo, lấy sheet tổng ra
        write_sheet(w.sheets[f"T{thang}.{nam}"], df) # Gọi hàm vẽ lại cho đẹp (ghi đè lên dữ liệu thô vừa ghi)
        print(f"[6] Đã viết sheet tổng: T{thang}.{nam}")
        
        # Chia theo từng ngày và ghi mỗi ngày 1 sheet riêng
        # groupby('Ngay_date'): tự động chia DataFrame thành các nhóm theo ngày. Mỗi nhóm là 1 ngày
        for ng, dfg in df.groupby('Ngay_date'): # ng là ngày (2024-05-01), dfg là DataFrame con của ngày đó
            ten = ng.strftime('%d.%m') # Đặt tên sheet là 01.05, 02.05... cho ngắn gọn, dễ nhìn. strftime format ngày
            dfg.to_excel(w, sheet_name=ten, index=False)
            write_sheet(w.sheets[ten], dfg) # Vẽ lại cho đẹp
            print(f"[7] Đã viết sheet ngày: {ten}")
    
    print(f"\n" + "="*50)
    print(f"THÀNH CÔNG!")
    print(f"="*50)
    print(f"File: {output_path}")
    print(f"Size: {os.path.getsize(output_path)/1024:.1f} KB") # getsize lấy dung lượng file (byte), /1024 để đổi ra KB, :.1f làm tròn 1 số lẻ

except Exception as e: # Nếu có lỗi bất kỳ ở các bước trên thì nhảy vào đây
    print(f"\n!!! LỖI: {e}")
    import traceback
    traceback.print_exc() # In chi tiết lỗi ra để biết lỗi dòng nào, hàm nào, file nào. Rất quan trọng để debug

input("\nNhấn Enter để thoát...") # Dừng màn hình lại cho bạn kịp đọc kết quả, không tắt ngay khi chạy bằng double-click trên Windows/Android
