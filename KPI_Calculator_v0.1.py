import tkinter as tk
from tkinter import filedialog, messagebox
import os
import openpyxl
import pandas as pd
import polars as pl
import numpy as np 
from pandas.tseries.offsets import CustomBusinessDay 
from datetime import datetime
import matplotlib.pyplot as plt 
import matplotlib.pyplot as plt1 
import numpy as np 
import tkinter as tk
from tkinter import messagebox, ttk, scrolledtext

#ghi chú
#df_CBBT_MienNam = KPI_DF
#df_HS_MienNam = BTTH_MienNam_DF
#df_HS_DaTinhKPI = HS_DaTinhKPI_DF
##Combine_DF = HS_CanKiemTra_DF

def Processing_PBI_WEB_data():
    #Chạy hàm chọn HR_info
    display_message("--Chọn file Thông tin nhân sự - HR_Miền Nam--","")
    KPI_DF, CB_KoKPI_DF, KPI_TungCB_DF = Hr_info()
    display_message("--Chọn Dữ liệu báo cáo lấy từ PBI WEB và Dữ liệu HS đã tính KPI--","")
    BTTH_DF, HS_DaTinhKPI_DS = select_excel_files()
    BTTH_MienNam_DF = BTTH_DF.unique(subset='SO_TO_TRINH')
    #Chuyển BTTH_DF lọc danh sách miền nam
    BTTH_MienNam_DF = BTTH_MienNam_DF.filter(pl.col('CAN_BO_BT').is_in(KPI_DF['Họ và tên']))
    # Tạo biến
    So_Hoso = 'SO_HS'
   # Step 1: Lọc các giá trị trùng lặp (duplicate) trong 'SO_HS'
    duplicate_so_hoso = BTTH_MienNam_DF[So_Hoso].is_in(BTTH_MienNam_DF[So_Hoso].unique().to_list())

    # Step 2: Lọc ra các dòng trùng lặp
    df_duplicate_So_Hoso_SameMonth = BTTH_MienNam_DF.filter(
        pl.col("SO_HS").is_in(BTTH_MienNam_DF["SO_HS"].value_counts().filter(pl.col("count") > 1)["SO_HS"])
    )
    # TStep 3: Thêm cột "Lý do" với giá trị "Trùng trong tháng"
    df_duplicate_So_Hoso_SameMonth = df_duplicate_So_Hoso_SameMonth.with_columns(
        pl.lit("Trùng trong tháng").alias("Lý do")
    )

    # Step 4: Filter for duplicates from the previous month
    df_duplicate_So_Hoso_PreviousMonth = BTTH_MienNam_DF.filter(pl.col('SO_HS').is_in(HS_DaTinhKPI_DS['SO_HS']))

    # Step 5: Merge to get the third column value from df_HS_DaTinhKPI
    df_duplicate_So_Hoso_PreviousMonth = df_duplicate_So_Hoso_PreviousMonth.join(
        HS_DaTinhKPI_DS[['SO_HS', 'NOTE']],
        on='SO_HS',
        how='left'
    )
    # Thay thế giá trị cột "Lý do" bằng giá trị từ cột 
    df_duplicate_So_Hoso_PreviousMonth = df_duplicate_So_Hoso_PreviousMonth.with_columns(
    pl.col('NOTE').alias('Lý do') 
)

    # Xóa cột nếu không cần thiết
    df_duplicate_So_Hoso_PreviousMonth = df_duplicate_So_Hoso_PreviousMonth.drop('NOTE')


##------------------------------------Xử lý hồ sơ cần kiểm tra trùng, đóng, sai quy định.--------------------------------------
            ##Giá trị cần kiểm tra
    check_Dong = 'đóng'
    check_TuChoi = 'từ chối'
    check_MaHS_regex = r'vp\..*\.hs.*'
    check_HanMuc = 'hết hạn mức'
    check_saiNDBH = 'sai người được bảo hiểm'
    check_saiNDBH1 = 'sai n.*bh'
    check_SaiSoThe = 'sai số thẻ'
    check_SaiSoHD = 'sai số HD'
    check_SaiSoHopDong = 'sai số hợp đồng'
    check_SaiHD = 'sai hd'
    check_SaiHopDong = 'sai hợp đồng'
    check_ngoaiHieuLuc = 'ngoài hiệu lực'
    check_ngoaiHLBH = 'ngoài hlbh'
    check_ngoaiTGBG = 'ngoài thời gian bảo hiểm'        
    check_khongGQAPP = 'không.*quy định.*giải quyết.*app'
    kekhai = ['kê khai', 'không thuộc đối tượng bảo hiểm', 'kiểm tra csyt', 'không thực hiện giám định']
    kekhai_regex = '|'.join(kekhai)    
    # Filter the DataFrame
    df_DongHSdoTrung = BTTH_MienNam_DF.filter(
        (pl.col("TIEN_BT_GQ") == 0) & 
        (
            (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_TuChoi.lower())) | 
            (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_Dong.lower()))
        ) & 
        (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_MaHS_regex.lower())) & 
        (~pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_HanMuc.lower())) & 
        (~pl.col("GIAI_QUYET").str.to_lowercase().str.contains(kekhai_regex.lower()))
    )

    # Hồ sơ bồi thường ít hơn 10k
    df_HSBoiThuongNghiNgo = BTTH_MienNam_DF.filter((pl.col("TIEN_BT_GQ") < 10000) & (pl.col("TIEN_BT_GQ") > 0))
    # Hồ sơ đóng do sai số hợp đồng, sai số thẻ, sai NĐBH   
    df_HSSaiHopDong = BTTH_MienNam_DF.filter(
        (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_SaiSoThe.lower())) |
        (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_SaiSoHD.lower())) |
        (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_SaiSoHopDong.lower())) |
        (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_SaiHD.lower())) |
        (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_SaiHopDong.lower())) |
        (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_saiNDBH.lower())) |
        (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_saiNDBH1.lower()))  
    )
    # Hồ sơ đóng do ngoài hiệu lực, ngoài HLBH, ngoài thời gian bảo hiểm
    df_HSngoaiHLBH = BTTH_MienNam_DF.filter(
        (pl.col("TIEN_BT_GQ") == 0) & (
            (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_ngoaiHieuLuc.lower())) |
            (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_ngoaiHLBH.lower())) |
            (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_ngoaiTGBG.lower()))
        )
    )
    # Hồ sơ đóng do hồ sơ không giải quyết qua app nhưng không yêu cầu bản cứng như quy định
    df_khongGQAPP = BTTH_MienNam_DF.filter(
        (pl.col("TIEN_BT_GQ") == 0) & 
        (pl.col("GIAI_QUYET").str.to_lowercase().str.contains(check_khongGQAPP.lower()))
    )
    # THÊM Cột lý do cho các DF trên
    df_DongHSdoTrung = df_DongHSdoTrung.with_columns(pl.lit("Đóng do trùng").alias("Lý do"))
    df_HSBoiThuongNghiNgo = df_HSBoiThuongNghiNgo.with_columns(pl.lit("Bồi thường < 10.000đ").alias("Lý do"))
    df_HSSaiHopDong = df_HSSaiHopDong.with_columns(pl.lit("Đóng do sai số thẻ/ sai hợp đồng/ sai NĐBH").alias("Lý do"))
    df_HSngoaiHLBH = df_HSngoaiHLBH.with_columns(pl.lit("Đóng do ngoài HLBH").alias("Lý do"))
    df_khongGQAPP = df_khongGQAPP.with_columns(pl.lit("Đóng do hồ sơ không giải quyết qua app nhưng không yêu cầu bản cứng như quy định").alias("Lý do"))


    HS_CanKiemTra_DF = pl.concat([
        df_DongHSdoTrung,  # Không cần chuyển nếu đã là Polars DataFrame
        df_HSBoiThuongNghiNgo,
        df_HSSaiHopDong,
        df_HSngoaiHLBH,
        df_khongGQAPP,
        df_duplicate_So_Hoso_SameMonth,
        df_duplicate_So_Hoso_PreviousMonth
    ])
    
  
    # Tạo DF D99MemThieuHASHTAG
    hashtagcung = 'bản cứng|hs cứng'
    df_D99MemThieuHASHTAG = filter_df_HS_MEM_thieuHashtag(BTTH_MienNam_DF, "vp.d99", "hsmem", hashtagcung)

    #Tạo DF HS mềm 
    hsMem = ['hsmem', 'hscung']
    filter_condition = pl.col("HAU_QUA").str.to_lowercase().str.contains(hsMem[0].lower())  # Start with the first condition

    # Add the remaining conditions using a loop
    for x in hsMem[1:]:
        filter_condition |= pl.col("HAU_QUA").str.to_lowercase().str.contains(x.lower())

    # Filter the DataFrame based on the constructed condition
    df_HsMem = BTTH_MienNam_DF.filter(filter_condition)

    #Create DF ALL HS MIỀN NAM CHUẨN TEMPLATE
    # Step 1: Create a new DataFrame with selected columns
    df_MienNam_Template_chuan = pl.DataFrame({
        'STT': pl.Series(BTTH_MienNam_DF['STT'].to_list()),  # Dùng to_list() để chuyển thành danh sách
        'Ngày yêu cầu': pl.Series(BTTH_MienNam_DF['NGAY_YC'].to_list()),
        'Ngày lập HS': pl.Series(BTTH_MienNam_DF['NGAY_LAP_HS'].to_list()),
        'Ngày bổ sung HS': pl.Series(BTTH_MienNam_DF['NGAY_BO_SUNG_HS'].to_list()),
        'Ngày duyệt': pl.Series(BTTH_MienNam_DF['APPROVE_DATE'].to_list()),
        'Ngày lập tờ trình': pl.Series(BTTH_MienNam_DF['NGAY_LAP_BT'].to_list()),
        'Số GYCTT': pl.Series(BTTH_MienNam_DF['SO_HS'].to_list()),
        'Số hồ sơ tờ trình bồi thường': pl.Series(BTTH_MienNam_DF['SO_TO_TRINH'].to_list()),
        'Cán bộ nhập GYCTT': pl.Series(BTTH_MienNam_DF['CAN_BO_NHAP_HS'].to_list()),
        'Cán bộ giải quyết bồi thường': pl.Series(BTTH_MienNam_DF['CAN_BO_BT'].to_list()),
        'Ngày lập tờ trình BT': pl.Series(BTTH_MienNam_DF['NGAY_LAP_BT'].to_list()),
        'Ngày gửi TBBT': pl.Series(BTTH_MienNam_DF['NGAY_THONG_BAO_BT'].to_list()),
        'Ngày giải quyết': pl.Series(BTTH_MienNam_DF['STT'].to_list()),  # Kiểm tra lại cột này
        'Phòng ban': pl.Series(BTTH_MienNam_DF['PHONG_BAN'].to_list()),
        'Số tiền bồi thường': pl.Series(BTTH_MienNam_DF['TIEN_BT_GQ'].to_list()),
        'Nguyên nhân tổn thất': pl.Series(BTTH_MienNam_DF['NGUYEN_NHAN_RUI_RO'].to_list()),
        'Giải quyết': pl.Series(BTTH_MienNam_DF['GIAI_QUYET'].to_list()),
        'CB_DUYET': pl.Series(BTTH_MienNam_DF['NGUOI_DUYET'].to_list()),
        'Giờ lập': pl.Series(BTTH_MienNam_DF['TG_LAP_TT'].dt.strftime('%H:%M').to_list()),  # Chắc chắn TG_LAP_TT là datetime
        'Hậu quả': pl.Series(BTTH_MienNam_DF['HAU_QUA'].to_list())
    })
    date_columns = ['Ngày yêu cầu', 'Ngày lập HS', 'Ngày bổ sung HS', 'Ngày duyệt', 'Ngày lập tờ trình BT', 'Ngày gửi TBBT']

    for col in date_columns:
        df_MienNam_Template_chuan = df_MienNam_Template_chuan.with_columns(
            pl.col(col).cast(pl.Datetime).dt.date()
            # pl.col(col).cast(pl.Datetime).dt.date().dt.strftime("%d/%m/%Y")  -> CHuyển về text
        )

   
    # Xuất file Excel DataFrame HS_CanKiemTra_DF và df_D99MemThieuHASHTAG
    file_path = select_file_path("XXX_CanKiemTra.xlsx")  # Gọi hàm chọn nơi lưu file, bạn có thể truyền tên file mặc định
    if file_path:  # Kiểm tra xem người dùng có chọn file không
        dataframes = [HS_CanKiemTra_DF.to_pandas(), df_D99MemThieuHASHTAG.to_pandas()]
        sheet_names = ["HS_Cho_Kiem_Tra", "Thieu_Hashtag"]
        export_to_excel(dataframes, sheet_names, file_path)  # Gọi hàm xuất dữ liệu

    # Xuất file Excel DataFrame HS_Mem và BTTH_MienNam
    file_path1 = select_file_path("BTTH_MienNam.xlsx")  # Gọi hàm chọn nơi lưu file với tên mặc định
    if file_path1:  # Kiểm tra xem người dùng có chọn file không
        dataframes1 = [df_MienNam_Template_chuan.to_pandas(), df_HsMem.to_pandas()]
        sheet_names1 = ["BTTH_MienNam", "HS_Mem"]
        export_to_excel(dataframes1, sheet_names1, file_path1)  # Gọi hàm xuất dữ liệu

    display_message("Dữ liệu đã được xử lý và xuất ra file Excel.", "Vui lòng kiểm tra file Excel đã lưu.")  # Hiển thị thông báo

    

def calculate_btth_data():
    BTTD_DF = BTTD_Folder()
    df_MienNam_Template_chuan, df_File_D99_MEM, df_HS_Khong_tinh_KPI, df_File_D99_BSCT, df_HS_KHCN = select_files_BTTH()
    


def kpi_report():
    # Example KPI report
    report = "KPI Report: All metrics are green."
    display_message("KPI Report has been generated.", report)

def plot():
    # Example plot message
    data = [1, 2, 3, 4]
    plt.plot(data)
    plt.title("Sample Plot")
    plt.xlabel("X-axis")
    plt.ylabel("Y-axis")
    plt.show()
    display_message("Plot has been created.", "Plot data: [1, 2, 3, 4]")

def display_message(message, data):
    output_area.insert(tk.END, message + "\n\n")  # Display message
    
    if isinstance(data, pl.DataFrame):
        output_area.insert(tk.END, data.to_string() + "\n")  # Display DataFrame contents
    else:
        output_area.insert(tk.END, str(data) + "\n")  # Display other data types

    output_area.insert(tk.END, "\n")  # Add a new line for better readability


def BTTD_Folder():
    # Tạo một cửa sổ ẩn để mở hộp thoại chọn thư mục
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ chính

    # Mở hộp thoại chọn thư mục
    folder_path = filedialog.askdirectory(title="Chọn thư mục chứa dữ liệu BTTĐ")

    if not folder_path:
        print("Không có thư mục nào được chọn.")
        return None

    # Danh sách để lưu dữ liệu
    data = []

    # Duyệt qua tất cả các tệp trong thư mục
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            # Bạn có thể tùy chỉnh cách đọc dữ liệu từ tệp ở đây
            # Ví dụ: Đọc tệp excel và thêm vào danh sách
            if filename.endswith('.xlsx'):
                df = pl.read_excel(file_path)
                data.append(df)

    # Kết hợp tất cả các DataFrame thành một DataFrame duy nhất
    if data:
        BTTD_DF = pl.concat(data)
        display_message("Dữ liệu đã được tải thành công.", BTTD_DF.shape)
        return BTTD_DF
    else:
        print("Không có tệp dữ liệu nào được tìm thấy trong thư mục.")
        return None

def filter_df_HS_MEM_thieuHashtag(df, loai_HS_D99, hashtagmem, hashtagcung):
    min_tien_bt_gq = 10000001

    # Filter the DataFrame using Polars
    df_D99MemThieuHASHTAG = df.filter(
        (pl.col("TIEN_BT_GQ") < min_tien_bt_gq) &
        (pl.col("SO_HS").str.to_lowercase().str.contains(loai_HS_D99.lower())) &
        (~pl.col("HAU_QUA").str.to_lowercase().str.contains(hashtagmem.lower())) &
        (~pl.col("HAU_QUA").str.to_lowercase().str.contains(hashtagcung.lower()))
    )
    
    return df_D99MemThieuHASHTAG

def Hr_info():
    # Create a hidden Tkinter window
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Open a file dialog to select an Excel file
    file_path = filedialog.askopenfilename(title="Chọn file HR_Info", filetypes=[("Excel files", "*.xlsx;*.xls")])

    if not file_path:
        print("No file selected.")
        return None, None, None

    # Read the specified sheets into DataFrames
    KPI_DF = pl.read_excel(file_path, sheet_name="KPI")
    CB_KoKPI_DF = pl.read_excel(file_path, sheet_name="CB_Khong_Tinh_KPI")
    KPI_TungCB_DF = pl.read_excel(file_path, sheet_name="KPI_Tung_CB")

    return KPI_DF, CB_KoKPI_DF, KPI_TungCB_DF

def select_excel_files():
    # Tạo một cửa sổ Tkinter ẩn
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ chính

    # Chọn tệp báo cáo gốc
    report_file_path = filedialog.askopenfilename(title="Chọn file báo cáo gốc export từ PBI", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not report_file_path:
        print("Không có tệp báo cáo nào được chọn.")
        return None, None

    # Chọn tệp hồ sơ đã tính KPI
    kpi_file_path = filedialog.askopenfilename(title="Chọn file hồ sơ đã tính KPI", filetypes=[("Excel files", "*.xlsm")])
    if not kpi_file_path:
        print("Không có tệp hồ sơ nào được chọn.")
        return None, None

    # Đọc các tệp Excel vào DataFrame
    BTTH_DF = pl.read_excel(report_file_path)
    HS_DaTinhKPI_DF = pl.read_excel(kpi_file_path)

    return BTTH_DF, HS_DaTinhKPI_DF


def select_files_BTTH():
    # Tạo một cửa sổ Tkinter ẩn
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ chính

    # Chọn file 1
    display_message("--Chọn file BTTH_MienNam--","")
    file1_path = filedialog.askopenfilename(title="Chọn file 1 (có sheet BTTH_MienNam)", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file1_path:
        print("Không có file 1 nào được chọn.")
        return None, None, None, None, None
    # Đọc sheet BTTH_MienNam từ file 1
    df_MienNam_Template_chuan = pl.read_excel(file1_path, sheet_name='BTTH_MienNam')

    # Chọn file 2
    display_message("--Chọn file Các loại HS KHÁC--","")
    file2_path = filedialog.askopenfilename(title="Chọn file 2 (có các sheet D99_MEM, HS_KhongTinhKPI, D99_BSCT, KHCN)", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file2_path:
        print("Không có file 2 nào được chọn.")
        return None, None, None, None, None
    # Đọc các sheet từ file 2
    df_File_D99_MEM = pl.read_excel(file2_path, sheet_name='D99_MEM')
    df_HS_Khong_tinh_KPI = pl.read_excel(file2_path, sheet_name='HS_KhongTinhKPI')
    df_File_D99_BSCT = pl.read_excel(file2_path, sheet_name='D99_BSCT')
    df_HS_KHCN = pl.read_excel(file2_path, sheet_name='KHCN')
    
    display_message("Bảng BTTH_Miền Nam", df_MienNam_Template_chuan.shape)
    display_message("Bảng D99_MEM", df_File_D99_MEM.shape)
    display_message("Bảng HS_KhongTinhKPI", df_HS_Khong_tinh_KPI.shape)
    display_message("Bảng D99_BSCT", df_File_D99_BSCT.shape)
    display_message("Bảng KHCN", df_HS_KHCN.shape)

    # Hiển thị thông báo
    display_message("Chọn file thành công", "")

    return df_MienNam_Template_chuan, df_File_D99_MEM, df_HS_Khong_tinh_KPI, df_File_D99_BSCT, df_HS_KHCN
     # Hiển thị thông báo

def select_file_path(default_name="output.xlsx"):
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ chính của tkinter
    file_path = filedialog.asksaveasfilename(title="Chọn nơi lưu file", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=default_name)
    return file_path

# Hàm xuất DataFrame vào Excel, sử dụng đường dẫn file đã chọn
def export_to_excel(dataframes, sheet_names, file_path=None):
    # Nếu không truyền tham số file_path, gọi hàm chọn nơi lưu
    if not file_path:
        file_path = select_file_path()  # Lấy đường dẫn nơi lưu file
        
    if file_path:  # Kiểm tra nếu người dùng chọn file
        with pd.ExcelWriter(file_path) as writer:
            for df, sheet_name in zip(dataframes, sheet_names):
                df.to_excel(writer, index=False, sheet_name=sheet_name)
        print(f"File đã được lưu tại: {file_path}")
    else:
        print("Không có file được lưu.")





# Create main window
root = tk.Tk()
root.title("BTTH Application")
root.geometry("1200x600")
root.configure(bg='#f0f0f0')

# Create a frame for buttons
button_frame = tk.Frame(root, bg='#f0f0f0')
button_frame.pack(side=tk.LEFT, padx=20, pady=20)

# Create a frame for output area
output_frame = tk.Frame(root, bg='#f0f0f0')
output_frame.pack(side=tk.RIGHT, padx=20, pady=20)

# Create title
title_label = tk.Label(root, text="BTTH Application", font=("Helvetica", 24), bg='#f0f0f0')
title_label.pack(pady=20)

# Create buttons with improved design
button_style = {
    "width": 24,
    "height": 2,
    "font": ("Helvetica", 12),
    "bg": "#007BFF",
    "fg": "white",
    "activebackground": "#0056b3",
    "activeforeground": "white"
}

button1 = tk.Button(button_frame, text="Processing PBI_WEB Data", command=Processing_PBI_WEB_data, **button_style)
button1.pack(pady=10)

button2 = tk.Button(button_frame, text="Calculate BTTH Data", command=calculate_btth_data, **button_style)
button2.pack(pady=10)

button3 = tk.Button(button_frame, text="KPI Report", command=kpi_report, **button_style)
button3.pack(pady=10)

button4 = tk.Button(button_frame, text="Plot", command=plot, **button_style)
button4.pack(pady=10)

# Create a scrolled text area for output
output_area = scrolledtext.ScrolledText(output_frame, width=120, height=30, font=("Helvetica", 12))
output_area.insert(tk.END, "--- WELOME TO KPI APPLICATION ---\n\n")
output_area.pack()

# Run the main loop
try:
    root.mainloop()
    
except Exception as e:
    print(f"An error occurred: {e}")
finally:
    root.destroy()  # Đảm bảo rằng ứng dụng được đóng