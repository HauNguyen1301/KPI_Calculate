import tkinter as tk
from tkinter import Tk, filedialog, messagebox
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
#GLOBAL Vrriables
global_BTTH_DF = None
global_HS_DaTinhKPI_DF = None
global_KPI_DF = None
global_CB_KoKPI_DF = None
global_KPI_TungCB_DF = None
global_BTTD_DF = None
global_HS_MienNam_DF = None
global_df_MienNam_Template_chuan = None
global_df_File_D99_MEM, global_df_HS_Khong_tinh_KPI, global_df_File_D99_BSCT, global_df_HS_KHCN = None, None, None, None
global_File_ALL_DLBT_ChuanTemp = None


#ghi chú
#df_CBBT_MienNam = KPI_DF
#df_HS_MienNam = BTTH_MienNam_DF
#df_HS_DaTinhKPI = HS_DaTinhKPI_DF
##Combine_DF = HS_CanKiemTra_DF

def Processing_PBI_WEB_data():
    KPI_DF = global_KPI_DF
    CB_KoKPI_DF = global_CB_KoKPI_DF
    BTTH_DF = global_BTTH_DF
    HS_DaTinhKPI_DF = global_HS_DaTinhKPI_DF
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
    df_duplicate_So_Hoso_PreviousMonth = BTTH_MienNam_DF.filter(pl.col('SO_HS').is_in(HS_DaTinhKPI_DF['SO_HS']))

    # Step 5: Merge to get the third column value from df_HS_DaTinhKPI
    df_duplicate_So_Hoso_PreviousMonth = df_duplicate_So_Hoso_PreviousMonth.join(
        HS_DaTinhKPI_DF[['SO_HS', 'NOTE']],
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

    for x in hsMem[1:]:
        filter_condition |= pl.col("HAU_QUA").str.to_lowercase().str.contains(x.lower())
    filter_condition &= pl.col("SO_HS").str.contains("D99")
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

    display_message("DỮ LIỆU ĐÃ ĐƯỢC XỬ LÝ VÀ XUẤT RA FILE EXCEL.\n VUI LÒNG KIỂM TRA FILE EXCEL ĐÃ LƯU.","")  # Hiển thị thông báo

    

def calculate_btth_data():
    df_HS_Khong_tinh_KPI = global_df_HS_Khong_tinh_KPI
    File_D99_BSCT = global_df_File_D99_BSCT
    df_HS_KHCN = global_df_HS_KHCN
    df_File_D99_MEM = global_df_File_D99_MEM
    df_BTTD = global_BTTD_DF
    df_MienNam_Template_chuan = global_df_MienNam_Template_chuan
    df_MienNam_Template_chuan_pandas = df_MienNam_Template_chuan.to_pandas()

    holidays = pd.to_datetime(['2024-09-03', '2024-09-02', '2024-05-01', '2024-04-30', 
                            '2024-04-29', '2024-04-18', '2024-02-14', '2024-02-13', 
                            '2024-02-12', '2024-02-11', '2024-02-10', '2024-02-09', 
                            '2024-02-08', '2024-01-01', '2023-09-04', '2023-09-03', 
                            '2023-09-02', '2023-09-01', '2023-05-03', '2023-05-02', 
                            '2023-05-01', '2023-04-30', '2023-04-29', '2023-01-26', 
                            '2023-01-25', '2023-01-24', '2023-01-23', '2023-01-22', 
                            '2023-01-21', '2023-01-20', '2023-01-02', '2023-01-01', 
                            '2023-01-02' ])

    # Create a custom business day calendar excluding weekends and holidays
    custom_bday = CustomBusinessDay(holidays=holidays)

    # Convert date columns to datetime if they're not already
    date_columns = ['Ngày yêu cầu', 'Ngày duyệt', 'Ngày bổ sung HS', 'Ngày gửi TBBT']
    for col in date_columns:
        df_MienNam_Template_chuan_pandas[col] = pd.to_datetime(df_MienNam_Template_chuan_pandas[col], errors='coerce')

    # Function to calculate network days
    def calculate_network_days(start_date, end_date):
        if pd.isnull(start_date) or pd.isnull(end_date):
            return None
        return len(pd.date_range(start=start_date, end=end_date, freq=custom_bday)) - 1

    # Calculate new columns
    df_MienNam_Template_chuan_pandas['1_NgàyYêuCầu_NgàyDuyệt'] = df_MienNam_Template_chuan_pandas.apply(
        lambda row: calculate_network_days(row['Ngày yêu cầu'], row['Ngày duyệt']), axis=1)

    df_MienNam_Template_chuan_pandas['2_NgàyYêuCầu_NgàyGửiTBBT'] = df_MienNam_Template_chuan_pandas.apply(
        lambda row: calculate_network_days(row['Ngày yêu cầu'], row['Ngày gửi TBBT']), axis=1)

    df_MienNam_Template_chuan_pandas['3_NgàyBổSungHS_NgàyDuyệt'] = df_MienNam_Template_chuan_pandas.apply(
        lambda row: calculate_network_days(row['Ngày bổ sung HS'], row['Ngày duyệt']), axis=1)

    df_MienNam_Template_chuan_pandas['4_NgàyBổSungHS_NgàyGửiTBBT'] = df_MienNam_Template_chuan_pandas.apply(
        lambda row: calculate_network_days(row['Ngày bổ sung HS'], row['Ngày gửi TBBT']), axis=1)

    # Function to replace negative values with 9999
    def replace_negative(value):
        return 9999 if value is not None and value < 0 else value

    # Apply the replace_negative function to the specified columns
    columns_to_check = [
        '1_NgàyYêuCầu_NgàyDuyệt',
        '2_NgàyYêuCầu_NgàyGửiTBBT',
        '3_NgàyBổSungHS_NgàyDuyệt',
        '4_NgàyBổSungHS_NgàyGửiTBBT'
    ]

    for col in columns_to_check:
        df_MienNam_Template_chuan_pandas[col] = df_MienNam_Template_chuan_pandas[col].apply(replace_negative)

    # Function to determine status
    def determine_status(row):
        if pd.isnull(row['3_NgàyBổSungHS_NgàyDuyệt']):
            if (row['1_NgàyYêuCầu_NgàyDuyệt'] is not None and row['1_NgàyYêuCầu_NgàyDuyệt'] < 8) or \
            (row['2_NgàyYêuCầu_NgàyGửiTBBT'] is not None and row['2_NgàyYêuCầu_NgàyGửiTBBT'] < 9):
                return "Đúng hạn"
            else:
                return "Trễ hạn"
        else:
            if (row['3_NgàyBổSungHS_NgàyDuyệt'] is not None and row['3_NgàyBổSungHS_NgàyDuyệt'] < 6) or \
            (row['4_NgàyBổSungHS_NgàyGửiTBBT'] is not None and row['4_NgàyBổSungHS_NgàyGửiTBBT'] < 7):
                return "Đúng hạn"
            else:
                return "Trễ hạn"

    def calculate_working_days(row):
        if pd.isnull(row['Ngày bổ sung HS']):
            if pd.isnull(row['Ngày gửi TBBT']):
                return row['1_NgàyYêuCầu_NgàyDuyệt']
            else:
                return row['2_NgàyYêuCầu_NgàyGửiTBBT']
        else:
            if pd.isnull(row['Ngày gửi TBBT']):
                return row['3_NgàyBổSungHS_NgàyDuyệt']
            else:
                return row['4_NgàyBổSungHS_NgàyGửiTBBT']


    # Add 5_Status column
    df_MienNam_Template_chuan_pandas['5_Status'] = df_MienNam_Template_chuan_pandas.apply(determine_status, axis=1)
    df_MienNam_Template_chuan_pandas['Số_Ngày_Làm_việc'] = df_MienNam_Template_chuan_pandas.apply(calculate_working_days, axis=1)

    df_MienNam_Template_chuan = pl.DataFrame(df_MienNam_Template_chuan_pandas)
    # Thêm cột 'Loại HS'
    df_MienNam_Template_chuan = df_MienNam_Template_chuan.with_columns(
        pl.when(pl.col("Nguyên nhân tổn thất") == "Ngoại trú").then(1)
        .when(pl.col("Nguyên nhân tổn thất") == "Nội trú").then(2)
        .when(pl.col("Nguyên nhân tổn thất") == "Sinh mạng").then(3)
        .otherwise(None).alias("Loại HS")
    )
    
    
    so_hsmem_set = set(df_File_D99_MEM['SO_HS'].to_list())

    # Thêm cột 'HS mềm'
    df_MienNam_Template_chuan = df_MienNam_Template_chuan.with_columns([
        pl.when(pl.col("Số GYCTT").is_in(so_hsmem_set))
        .then(pl.lit("HSMEM"))
        .otherwise(pl.lit(None))
        .alias("HS mềm")
    ])

    # Thêm cột 'QLHA'
    df_MienNam_Template_chuan = df_MienNam_Template_chuan.with_columns(
        pl.when(pl.col("Hậu quả").str.contains("#QLHA", literal=True))
        .then(pl.lit("QLHA"))
        .otherwise(pl.lit(None))
        .alias("QLHA")
    )

    # Thêm cột 'MaLoaiHS'
    df_MienNam_Template_chuan = df_MienNam_Template_chuan.with_columns(
        pl.col("Số GYCTT").str.slice(3, 3).alias("MaLoaiHS")
    )

        # Thêm cột 'CoBSCT'
    so_hs_BSCT_set = set(File_D99_BSCT["Giấy YCTT"].to_list())
    df_MienNam_Template_chuan = df_MienNam_Template_chuan.with_columns(
        pl.when(pl.col("Số GYCTT").is_in(so_hs_BSCT_set))
        .then(pl.lit("Có BSCT"))
        .otherwise(pl.lit(None))
        .alias("CoBSCT")
    )

    # Thêm cột 'BTTD'
    so_hs_BTTD_set = set(df_BTTD["Số TTBT"].to_list())
    df_MienNam_Template_chuan = df_MienNam_Template_chuan.with_columns(
        pl.when(pl.col("Số hồ sơ tờ trình bồi thường").is_in(so_hs_BTTD_set))
        .then(pl.lit("BTTĐ"))
        .otherwise(pl.lit(None))
        .alias("BTTD")
    )

    # Tạo điều kiện và thêm cột 'HS IP bị giảm tỉ lệ quy đổi'
    condition = (
        (pl.col("Số tiền bồi thường") == 0)
        & (pl.col("Loại HS") == 2)
        & (
            pl.col("Giải quyết").str.to_lowercase().str.contains("hết hạn mức")
            | pl.col("Giải quyết").str.to_lowercase().str.contains("không bổ sung")
            | pl.col("Giải quyết").str.to_lowercase().str.contains("hạn bổ sung")
        )
    )
    df_MienNam_Template_chuan = df_MienNam_Template_chuan.with_columns(
        pl.when(condition).then(pl.lit("X")).otherwise(pl.lit(None)).alias("HS IP bị giảm tỉ lệ quy đổi")
    )

    # Thêm cột 'Hồ sơ không tính KPI'
    df_MienNam_Template_chuan = df_MienNam_Template_chuan.with_columns(
        pl.when(
            pl.col("Số hồ sơ tờ trình bồi thường").is_in(
                df_HS_Khong_tinh_KPI["Số tờ trình bồi thường"].to_list()
            )
        )
        .then(pl.lit("X"))
        .otherwise(pl.lit(None))
        .alias("Hồ sơ không tính KPI")
    )
    
    # Xóa cột không cần thiết
    df_MienNam_Template_chuan = df_MienNam_Template_chuan.drop(
        ["1_NgàyYêuCầu_NgàyDuyệt", "2_NgàyYêuCầu_NgàyGửiTBBT", "3_NgàyBổSungHS_NgàyDuyệt", "4_NgàyBổSungHS_NgàyGửiTBBT"]
    )    

    # Tạo điều kiện và thêm cột 'HS IP bị giảm tỉ lệ quy đổi'
    # Lọc bỏ các hàng có 'Hồ sơ không tính KPI' là 'X'
    df_MienNam_Template_chuan = df_MienNam_Template_chuan.filter(
        pl.col("Hồ sơ không tính KPI").is_null()
    )
    #print(df_MienNam_Template_chuan)
    # Export the DataFrame to an Excel file
    df_MienNam_Template_chuan_pandas = df_MienNam_Template_chuan.to_pandas()
    file_path = select_file_path("BC_BTTH_Processed.xlsx")  # Gọi hàm chọn nơi lưu file, bạn có thể truyền tên file mặc định
    if file_path:  # Kiểm tra xem người d�ng có chọn file không
        dataframes = [df_MienNam_Template_chuan_pandas,df_HS_KHCN.to_pandas(),df_HS_Khong_tinh_KPI.to_pandas()]
        sheet_names = ["BTTH_MienNam_Calculated","HS_KHCN","HS_Khong_tinh_KPI"]
        export_to_excel(dataframes, sheet_names, file_path)  # Gọi hàm xuất dữ liệu
        display_message("DỮ LIỆU ĐÃ ĐƯỢC TÍNH TOÁN VÀ XUẤT RA FILE EXCEL BC_BTTH_PROCESSED.XLSX.", "VUI LÒNG KIỂM TRA FILE EXCEL ĐÃ LƯU.")  # Hiển thị thông báo
    #print(df_MienNam_Template_chuan)


def kpi_report():
    # Example KPI report
    df_HS_KHCN = global_df_HS_KHCN
    df_HS_Khong_tinh_KPI = global_df_HS_Khong_tinh_KPI
    File_ALL_DLBT_ChuanTemp = global_File_ALL_DLBT_ChuanTemp
    KPI_DF = global_KPI_DF
    KPI_TungCB_DF = global_KPI_TungCB_DF
    
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT NGOẠI TRÚ---------------------------------------------------------
    counts = File_ALL_DLBT_ChuanTemp.group_by("Cán bộ giải quyết bồi thường").agg([
        # Ngoại trú - giữ nguyên
        ((pl.col("MaLoaiHS").is_in(["D31", "D98", "D33", "D15"])) & 
        (pl.col("Loại HS") == 1) & 
        (pl.col("QLHA").is_null()) & 
        (pl.col("BTTD").is_null()) & 
        (pl.col("Hồ sơ không tính KPI").is_null())
        ).sum().alias("Ngoại trú"),
        
        # Nội trú - giữ nguyên
        ((pl.col("MaLoaiHS").is_in(["D31", "D98", "D33", "D15"])) & 
        (pl.col("Loại HS") == 2) & 
        (pl.col("QLHA").is_null()) & 
        (pl.col("Hồ sơ không tính KPI").is_null())
        ).sum().alias("Nội trú"),

        # QLHA ngoai tru - từ code đính kèm
        ((pl.col("MaLoaiHS").is_in(["D31", "D33", "D15"])) &
        (pl.col("Loại HS") == 1) &
        (pl.col("QLHA") == "QLHA") &
        (pl.col("Hồ sơ không tính KPI").is_null())
        ).sum().alias("QLHA Ngoại trú"),

        # QLHA noi tru - từ code đính kèm
        ((pl.col("MaLoaiHS").is_in(["D31", "D33", "D15"])) &
        (pl.col("Loại HS") == 2) &
        (pl.col("QLHA") == "QLHA") &
        (pl.col("Hồ sơ không tính KPI").is_null())
        ).sum().alias("QLHA Nội trú"),

        # Tu vong - từ code đính kèm 
        ((pl.col("Loại HS") == 3) &
        (pl.col("Hồ sơ không tính KPI").is_null())
        ).sum().alias("Tử vong"),

        # D99 mềm BS OP - từ code đính kèm
        ((pl.col("Hồ sơ không tính KPI").is_null()) &
        (pl.col("Loại HS") == 1) &
        (pl.col("MaLoaiHS") == "D99") &
        (pl.col("HS mềm") == "HSMEM") &
        (pl.col("CoBSCT") == "Có BSCT")
        ).sum().alias("D99 BSCT OP"),

        # D99 mềm bước 2 OP - từ code đính kèm
        ((pl.col("Hồ sơ không tính KPI").is_null()) &
        (pl.col("Loại HS") == 1) &
        (pl.col("MaLoaiHS") == "D99") &
        (pl.col("HS mềm") == "HSMEM") &
        (pl.col("BTTD").is_null())
        ).sum().alias("D99 mềm B2 OP"),

        # D99 cứng OP - từ code đính kèm
        ((pl.col("Hồ sơ không tính KPI").is_null()) &
        (pl.col("Loại HS") == 1) &
        (pl.col("MaLoaiHS") == "D99") &
        (pl.col("HS mềm").is_null()) &
        (pl.col("BTTD").is_null())
        ).sum().alias("D99 cứng OP"),

        # D99 mềm BS IP - từ code đính kèm
        ((pl.col("Hồ sơ không tính KPI").is_null()) &
        (pl.col("Loại HS") == 2) &
        (pl.col("MaLoaiHS") == "D99") &
        (pl.col("HS mềm") == "HSMEM") &
        (pl.col("CoBSCT") == "Có BSCT")
        ).sum().alias("D99 BSCT IP"),

        # D99 mềm bước 2 IP - từ code đính kèm
        ((pl.col("Hồ sơ không tính KPI").is_null()) &
        (pl.col("Loại HS") == 2) &
        (pl.col("MaLoaiHS") == "D99") &
        (pl.col("HS mềm") == "HSMEM")
        ).sum().alias("D99 mềm B2 IP"),

        # D99 cứng IP - từ code đính kèm
        ((pl.col("Hồ sơ không tính KPI").is_null()) &
        (pl.col("Loại HS") == 2) &
        (pl.col("MaLoaiHS") == "D99") &
        (pl.col("HS mềm").is_null())
        ).sum().alias("D99 cứng IP"),

        # BTTD D31 - từ code đính kèm
        ((pl.col("Hồ sơ không tính KPI").is_null()) &
        (pl.col("Loại HS") == 1) &
        (pl.col("MaLoaiHS").is_in(["D31", "D98", "D33", "D15"])) &
        (pl.col("BTTD") == "BTTĐ") &
        (pl.col("QLHA").is_null())
        ).sum().alias("BTTD D31"),

        # BTTD D99 - từ code đính kèm
        ((pl.col("Hồ sơ không tính KPI").is_null()) &
        (pl.col("Loại HS") == 1) &
        (pl.col("MaLoaiHS") == "D99") &
        (pl.col("BTTD") == "BTTĐ")
        ).sum().alias("BTTD D99"),

        # Không tính KPI - từ code đính kèm
        (pl.col("Hồ sơ không tính KPI") == "X").sum().alias("HS_Khong_Tinh_KPI"),

        # Giảm tỉ lệ quy đổi - từ code đính kèm
        ((pl.col("Hồ sơ không tính KPI").is_null()) &
        (pl.col("HS IP bị giảm tỉ lệ quy đổi") == "X") &
        (pl.col("Loại HS") == 2) &
        (pl.col("Số tiền bồi thường") == 0)
        ).sum().alias("HS Giảm tỉ lệ quy đổi"),

        # Tổng số hồ sơ - từ code đính kèm
        pl.col("Số hồ sơ tờ trình bồi thường")
        .filter(pl.col("Hồ sơ không tính KPI").is_null())
        .count().alias("Số lượng hồ sơ BVCare chưa quy đổi"),

        # Số hồ sơ đúng hạn - từ code đính kèm
        pl.col("Số hồ sơ tờ trình bồi thường")
        .filter((pl.col("Hồ sơ không tính KPI").is_null()) &
                (pl.col("5_Status") == "Đúng hạn"))
        .count().alias("Số hồ sơ đúng hạn")
    ])
    # Tính toán riêng cho df_HS_KHCN
    counts_khcn = df_HS_KHCN.group_by("Cán bộ BT").agg([
        # KHCN
        (pl.col("Nghiệp vụ").str.to_lowercase().is_in(["khn", "ats", "pai", "ytk"]))
        .sum().alias("KHCN"),
        
        # PA
        (pl.col("Nghiệp vụ").is_in(["PA"]))
        .sum().alias("PA"),
        
        # Du lịch
        (pl.col("Nghiệp vụ").str.to_lowercase().is_in(["fle", "dqt", "ydl"]))
        .sum().alias("Du lịch"),
        
        # Du lịch ECEP  
        (pl.col("Nghiệp vụ").str.to_lowercase().is_in(["ecep"]))
        .sum().alias("Du lịch ECEP"),
        
        # Kcare
        (pl.col("Nghiệp vụ").str.to_lowercase().is_in(["kcare"]))
        .sum().alias("Kcare")
    ])
    # Ghép nối với KPI_DF - giữ nguyên
    # Chuẩn hóa tên cột trước khi join
    counts = counts.rename({"Cán bộ giải quyết bồi thường": "Họ và tên"})
    counts_khcn = counts_khcn.rename({"Cán bộ BT": "Họ và tên"}) 

    # Strip whitespace
    KPI_DF = KPI_DF.with_columns(pl.col("Họ và tên"))
    counts = counts.with_columns(pl.col("Họ và tên"))
    counts_khcn = counts_khcn.with_columns(pl.col("Họ và tên"))

    # Join với xử lý null
    KPI_DF = KPI_DF.join(
        counts,
        on="Họ và tên",
        how="left"
    ).join(
        counts_khcn, 
        on="Họ và tên",
        how="left"
    ).fill_null(0)
    KPI_DF = KPI_DF.with_columns(
        pl.lit(0).alias("IJ Ngoại trú"),
        pl.lit(0).alias("IJ Nội trú"),
        pl.lit(0).alias("Ngoại giao"), 
        pl.lit(0).alias("CI37"),
        pl.lit(0).alias("D99 B1 OP"),  
        pl.lit(0).alias("D99 B1 IP"),
        pl.lit(0).alias("HS Hỗ trợ")
    )
    # Thêm cột mới "%Thời Gian GQ HS_QT063"
    KPI_DF = KPI_DF.with_columns([
        (pl.col("Số hồ sơ đúng hạn") / pl.col("Số lượng hồ sơ BVCare chưa quy đổi")).alias("%Thời Gian GQ HS_QT063")
    ])

    # Sắp xếp lại cột
    KPI_DF = KPI_DF.select([
        "STT",
        "Họ và tên", 
        "Mã nhân viên",
        "Tháng",
        "Năm",
        "Ngoại trú",
        "Nội trú", 
        "QLHA Ngoại trú",
        "QLHA Nội trú",
        "IJ Ngoại trú",
        "IJ Nội trú",
        "KHCN",
        "PA",
        "Ngoại giao",
        "Du lịch",
        "Tử vong",
        "Kcare",
        "CI37",
        "D99 B1 OP",
        "D99 BSCT OP", 
        "D99 mềm B2 OP",
        "D99 cứng OP",
        "D99 B1 IP",
        "D99 BSCT IP",
        "D99 mềm B2 IP", 
        "D99 cứng IP",
        "HS Hỗ trợ",
        "BTTD D31",
        "BTTD D99",
        "Du lịch ECEP",
        "HS Giảm tỉ lệ quy đổi",
        "Số lượng hồ sơ BVCare chưa quy đổi",
        "Số hồ sơ đúng hạn",
        "%Thời Gian GQ HS_QT063"
    ])
    if KPI_TungCB_DF is None:
        return
    # Tính toán riêng cho KPI_TungCB_DF
    # Thêm cột KPI
    KPI_DF = KPI_DF.join(
        KPI_TungCB_DF.select(["Họ và tên", "KPI"]),
        on="Họ và tên",
        how="left"
    ).fill_null(0)
    # Thêm cột Tổng sau quy đổi
    KPI_DF = KPI_DF.with_columns(
        (
            pl.col("Ngoại trú") * 1 +
            pl.col("Nội trú") * 2.4 +
            pl.col("QLHA Ngoại trú") * 1.1 +
            pl.col("QLHA Nội trú") * 2.5 +
            pl.col("IJ Ngoại trú") * 2 +
            pl.col("IJ Nội trú") * 4 +
            pl.col("KHCN") * 2 +
            pl.col("PA") * 3 +
            pl.col("Ngoại giao") * 6 +
            pl.col("Du lịch") * 12 +
            pl.col("Tử vong") * 15 +
            pl.col("Kcare") * 45 +
            pl.col("CI37") * 60 +
            pl.col("D99 B1 OP") * 0.5 +
            pl.col("D99 BSCT OP") * 0.5 +
            pl.col("D99 mềm B2 OP") * 1 +
            pl.col("D99 cứng OP") * 1.5 +
            pl.col("D99 B1 IP") * 0.5 +
            pl.col("D99 BSCT IP") * 0.5 +
            pl.col("D99 mềm B2 IP") * 2.4 +
            pl.col("D99 cứng IP") * 2.9 +
            pl.col("BTTD D31") * 0.85 +
            pl.col("BTTD D99") * 0.85 +
            pl.col("Du lịch ECEP") * 1 +
            (pl.col("HS Giảm tỉ lệ quy đổi") * (-1.4))
        ).alias("Tổng sau quy đổi")
    )
    # Thêm cột Tỉ lệ GQ đúng hạn
    KPI_DF = KPI_DF.with_columns([
        pl.when(pl.col("%Thời Gian GQ HS_QT063") / 0.95 < 1)
        .then(pl.col("%Thời Gian GQ HS_QT063") / 0.95 * 0.2)
        .otherwise(0.2)
        .alias("TiLeGQDungHan")
    ])
    # Tạo cột Rate
    KPI_DF = KPI_DF.with_columns([
        (
            (pl.col("Tổng sau quy đổi") / pl.col("KPI")) * 0.7 + 0.1 + pl.col("TiLeGQDungHan")
        ).alias("Rate")
    ])
    # tạo cột Tỉ lệ % theo KPI_QT063
    KPI_DF = KPI_DF.with_columns([
        pl.when(pl.col("KPI").is_null())
        .then(0)
        .otherwise(pl.col("Rate"))
        .alias("Tỉ lệ % theo KPI_QT063")
    ])
    # Tính toán tạo cột Phần vượt
    KPI_DF = KPI_DF.with_columns([
        pl.when(pl.col("Tỉ lệ % theo KPI_QT063") > 1.2)
        .then(pl.col("Tỉ lệ % theo KPI_QT063") - 1.2)
        .otherwise(pl.col("Tỉ lệ % theo KPI_QT063"))
        .alias("PhanVuot")
    ])
    # Tính toán 'PhanTangThem'
    KPI_DF = KPI_DF.with_columns([
        (pl.col("PhanVuot") * pl.col("KPI") / 0.7 / pl.col("KPI")).alias("PhanTangThem")
    ])
    # Tính toán 'KPI sau vượt 120%_QT063'
    KPI_DF = KPI_DF.with_columns([
        pl.when(pl.col("Tỉ lệ % theo KPI_QT063") > 1.2)
        .then(1.2 + pl.col("PhanTangThem"))
        .otherwise(pl.col("Tỉ lệ % theo KPI_QT063"))
        .alias("KPI sau vượt 120%_QT063")
    ])
    #DROP các cột không cần thiết
    KPI_DF = KPI_DF.drop(["TiLeGQDungHan", "Rate", "PhanVuot", "PhanTangThem"])
    #Filter các Tổng sau quy đổi >= 10
    KPI_DF = KPI_DF.filter(pl.col("Tổng sau quy đổi") >= 10)
    KPI_DF_ForCBBT = KPI_DF.drop(["Tổng sau quy đổi", "Tỉ lệ % theo KPI_QT063", "KPI sau vượt 120%_QT063", "KPI"])


    # Export the DataFrame to an Excel file
    def export_to_excel(df, default_filename="KPI_Report.xlsx"):
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_filename
        )
        if file_path:
            df.write_excel(file_path)
        root.destroy()
    # Sử dụng hàm
    export_to_excel(KPI_DF, "KPI_Report.xlsx")
    export_to_excel(KPI_DF_ForCBBT, "KPI_Report_ForCBBT.xlsx")
    display_message("DỮ LIỆU ĐÃ ĐƯỢC TÍNH TOÁN VÀ XUẤT RA FILE EXCEL KPI_REPORT.XLSX.", "VUI LÒNG KIỂM TRA FILE EXCEL ĐÃ LƯU.")



def plot():
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    def select_excel_file():
    # Display initial message
        display_message("--> Chọn file KPI_Report để vẽ biểu đồ", "")
        
        # Open file dialog
        file_path = filedialog.askopenfilename(
            title="Chọn file KPI_Report", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if file_path:
            # Display success message after selection
            display_message("--> KPI_Report Loaded!", "")
            return file_path
        else:
            display_message("!! Không có file được chọn", "")
            return None

    file_path = filedialog.askopenfilename(title="Chọn file KPI", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        KPI_DF_ForPlot = pd.read_excel(file_path, sheet_name='Sheet1')

        X_axis = KPI_DF_ForPlot['Họ và tên']
        Column_y_axis = KPI_DF_ForPlot['Tổng sau quy đổi']
        Line_y_axis = np.round(KPI_DF_ForPlot['KPI sau vượt 120%_QT063'] * 100)

        # Filter the Line_y_axis values that are greater than 0
        filtered_Line_y_axis = Line_y_axis[Line_y_axis > 0.5]
        filtered_X_axis = X_axis[Line_y_axis > 0.5]

        # Create the figure and axes
        fig, ax1 = plt.subplots(figsize=(15, 6))
        ax2 = ax1.twinx()

        # Plot the clustered column chart
        bars = ax1.bar(filtered_X_axis, Column_y_axis[Line_y_axis > 0.5], label='Số lượng HS sau quy đổi')

        # Plot the line chart with the filtered data
        ax2.plot(filtered_X_axis, filtered_Line_y_axis, color='red',  label='Tỉ lệ sau 120%')

        # Set the labels and title
        ax1.set_xlabel('Họ và tên')
        ax1.set_ylabel('Số lượng hồ sơ')
        ax2.set_ylabel('Tỉ lệ %')

        # Rotate the x-axis labels by 45 degrees
        ax1.set_xticklabels(filtered_X_axis, rotation=45, ha='right')
        fig.tight_layout()
        fig.subplots_adjust(bottom=0.2)  # Increase bottom margin

        # Add data labels to the clustered column chart
        for bar in bars:
            height = bar.get_height()
            ax1.annotate(str(round(height, 2)), xy=(bar.get_x() + bar.get_width() / 2, height), 
                        xytext=(0, 3), textcoords="offset points", 
                        ha='center', va='bottom', fontsize=8)

        # Add data labels to the line chart
        for x, y in zip(filtered_X_axis, filtered_Line_y_axis):
            ax2.annotate(str(round(y, 2)) + '%', xy=(x, y), 
                        xytext=(0, 3), textcoords="offset points", 
                        ha='center', va='bottom', fontsize=8, color='red',
                        bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="red", alpha=0.8))

        # Combine the legends
        lines, labels = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax2.legend(lines + lines2, labels + labels2)

        # Create a Toplevel window
        plot_window = tk.Toplevel(root)
        plot_window.title("KPI Plot")

        # Embed the plot in the Tkinter window
        canvas = FigureCanvasTkAgg(fig, master=plot_window)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

        # Add a close button
        close_button = tk.Button(plot_window, text="Close", command=plot_window.destroy)
        close_button.pack()

        # Show the plot
        plot_window.mainloop()
        
        pass
    

   
def display_message(message, data):
    output_area.insert(tk.END, message + "\n")  # Display message
    
    if isinstance(data, pl.DataFrame):
        output_area.insert(tk.END, data.to_string() + "\n")  # Display DataFrame contents
    else:
        output_area.insert(tk.END, str(data) + "\n")  # Display other data types

    output_area.insert(tk.END, "\n")  # Add a new line for better readability
    output_area.see(tk.END)  # Tự động cuộn xuống dòng mới nhất


def BTTD_Folder():
    # Tạo một cửa sổ ẩn để mở hộp thoại chọn thư mục
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ chính

    # Mở hộp thoại chọn thư mục
    folder_path = filedialog.askdirectory(title="Chọn thư mục chứa dữ liệu BTTĐ")

    if not folder_path:
        display_message("!! Không có thư mục nào được chọn.","")
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
        display_message("Dữ liệu BTTĐ được tải thành công.", BTTD_DF.shape)
        return BTTD_DF
    else:
        display_message("!! Không có tệp dữ liệu nào được tìm thấy trong thư mục.","")
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
        messagebox.showwarning("Cảnh báo", "Vui lòng thêm thông tin HR trước khi thực hiện!")
        display_message("!! Chưa có thông tin HR\n--> Vui lòng thao tác lại\n", "")
        return None, None, None

    # Read the specified sheets into DataFrames
    KPI_DF = pl.read_excel(file_path, sheet_name="KPI")
    CB_KoKPI_DF = pl.read_excel(file_path, sheet_name="CB_Khong_Tinh_KPI")
    KPI_TungCB_DF = pl.read_excel(file_path, sheet_name="KPI_Tung_CB")

    return KPI_DF, CB_KoKPI_DF, KPI_TungCB_DF

def select_excel_PBI_file():
    # Tạo một cửa sổ Tkinter ẩn
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ chính

    # Chọn tệp báo cáo gốc
    report_file_path = filedialog.askopenfilename(title="Chọn file báo cáo gốc export từ PBI", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not report_file_path:
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn dữ liệu để thực hiện!")
        display_message("!! Chưa có thông tin Báo cáo gốc từ PBI\n--> Vui lòng thao tác lại\n", "")
        return None, None

    # Chọn tệp hồ sơ đã tính KPI
    kpi_file_path = filedialog.askopenfilename(title="Chọn file hồ sơ đã tính KPI", filetypes=[("Excel files", "*.xlsm")])
    if not kpi_file_path:
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn file HS_ĐÃ_TÍNH_KPI thực hiện!")
        display_message("!! Chưa có thông tin hồ sơ đã tính KPI\n--> Vui lòng thao tác lại\n", "")
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
    display_message("--> Chọn file BTTH_MienNam","")
    file1_path = filedialog.askopenfilename(title="Chọn file 1 (có sheet BTTH_MienNam)", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file1_path:
        display_message("!! Không có file BTTH_MienNam nào được chọn.", "")
        return None, None, None, None, None
    # Đọc sheet BTTH_MienNam từ file 1
    df_MienNam_Template_chuan = pl.read_excel(file1_path, sheet_name='BTTH_MienNam')

    # Chọn file 2
    display_message("--Chọn file Các loại HS_KHÁC--","")
    file2_path = filedialog.askopenfilename(title="Chọn file HS_KHÁC (có các sheet D99_MEM, HS_KhongTinhKPI, D99_BSCT, KHCN)", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file2_path:
        display_message("Không có file HS_KHÁC nào được chọn.", "")
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

def select_file_path(default_name="output.xlsx"): ##XUẤT FILE EXCEL
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
        display_message(f"FILE ĐÃ ĐƯỢC LƯU TẠI: {os.path.basename(file_path).upper()}", "")
    else:
        display_message("!! Không có file được lưu.","")



def add_file_HR():
    global global_KPI_DF, global_CB_KoKPI_DF, global_KPI_TungCB_DF  # Khai báo biến toàn cục
    # Tạo một cửa sổ Tkinter ẩn
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ chính
    display_message("--> Chọn file Thông tin nhân sự - HR_Miền Nam","")
    # Chọn file HR_Info
    global_KPI_DF, global_CB_KoKPI_DF, global_KPI_TungCB_DF = Hr_info()
    if global_KPI_DF is None or global_CB_KoKPI_DF is None or global_KPI_TungCB_DF is None:
        return
    else:
        label0.config(text="0. HR_File Loaded!")
        output_area.insert(tk.END, "--> FILE HR_INFO: SHEET KPI, CB_KHONG_TINH_KPI, KPI_TUNG_CB LOADED!\n\n")
    update_button_states()

def add_file1():
    # Add your existing file adding logic here
    global global_BTTH_DF, global_HS_DaTinhKPI_DF# Khai báo biến toàn cục
    display_message("--> Chọn Dữ liệu báo cáo lấy từ PBI WEB & Dữ liệu HS đã tính KPI","")
    global_BTTH_DF, global_HS_DaTinhKPI_DF = select_excel_PBI_file()
    if (global_BTTH_DF is None or global_HS_DaTinhKPI_DF is None or 
        len(global_BTTH_DF) == 0 or len(global_HS_DaTinhKPI_DF) == 0):
        messagebox.showwarning("Cảnh báo", "Vui lòng thêm dữ liệu trước khi thực hiện")
        display_message("!! Không có dữ liệu nào được chọn.", "")
        return
    else:
        label1.config(text="1. PBI_WEB, HS_Đã_tính_KPI Loaded!")
        output_area.insert(tk.END, "--> FILE PBI_WEB, FILE HS ĐÃ TÍNH KPI LOADED!\n\n")
        if global_KPI_DF is None:
            output_area.insert(tk.END, "!! Vui lòng add thêm HR_Info\n")
    update_button_states()   

def add_file2():
    # Add your existing file adding logic here
    # Khởi tạo một cửa sổ Tkinter ẩn
    global global_BTTD_DF, global_HS_MienNam_DF
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ chính
    display_message("--> Chọn folder BTTĐ","")
    folder_path = filedialog.askdirectory(title="Chọn thư mục chứa dữ liệu BTTĐ")
    if not folder_path:
        display_message("!! Không có thư mục nào được chọn.", "")
    else:
        # Danh sách để lưu dữ liệu
        data = []
        # Duyệt qua tất cả các tệp trong thư mục
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                # Đọc tệp excel và thêm vào danh sách
                if filename.endswith('.xlsx'):
                    df = pl.read_excel(file_path)
                    data.append(df)

        # Kết hợp tất cả các DataFrame thành một DataFrame duy nhất
        if data:
            global_BTTD_DF = pl.concat(data)
            display_message("Dữ liệu đã được tải thành công.", global_BTTD_DF.shape)
        else:
            display_message("!! Không có tệp dữ liệu nào được tìm thấy trong thư mục BTTĐ.", "")

    # Chọn file BTTH_MienNam
    global global_df_MienNam_Template_chuan
    display_message("--> Chọn file BTTH_MienNam","")
    file1_path = filedialog.askopenfilename(title="Chọn file 1 (có sheet BTTH_MienNam)", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file1_path:
        display_message("!! Không có file BTTH_MienNam được chọn.","")
        return None, None, None, None, None
    # Đọc sheet BTTH_MienNam từ file 1
    global_df_MienNam_Template_chuan = pl.read_excel(file1_path, sheet_name='BTTH_MienNam')

    # Chọn file HS_Khác
    global global_df_File_D99_MEM, global_df_HS_Khong_tinh_KPI, global_df_File_D99_BSCT, global_df_HS_KHCN
    display_message("--> Chọn file Các loại HS khác","")
    file2_path = filedialog.askopenfilename(title="Chọn file HS_KHÁC (có các sheet D99_MEM, HS_KhongTinhKPI, D99_BSCT, KHCN)", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file2_path:
        display_message("!! Không có file HS_Khác được chọn.")
        return None, None, None, None, None
    # Đọc các sheet từ file 2
    global_df_File_D99_MEM = pl.read_excel(file2_path, sheet_name='D99_MEM')
    global_df_HS_Khong_tinh_KPI = pl.read_excel(file2_path, sheet_name='HS_KhongTinhKPI')
    global_df_File_D99_BSCT = pl.read_excel(file2_path, sheet_name='D99_BSCT')
    global_df_HS_KHCN = pl.read_excel(file2_path, sheet_name='KHCN')
    

    if global_BTTD_DF.is_empty() or global_df_MienNam_Template_chuan.is_empty() or global_df_File_D99_MEM.is_empty() or global_df_HS_Khong_tinh_KPI.is_empty() or global_df_File_D99_BSCT.is_empty() or global_df_HS_KHCN.is_empty():
        messagebox.showwarning("Cảnh báo", "Vui lòng Chọn dữ liệu BTTD, BTTH_MienNam!")
        display_message("!! Chưa có thông tin dữ liệu\n--> Vui lòng thao tác lại\n", "")
        return
    else:
        label2.config(text="2. BTTĐ, HS_MienNam Loaded!")
        output_area.insert(tk.END, "--> FILE HS_MIENNAM: SHEET D99_MEM, KOTINHKPI, BSCT, KHCN LOADED!\n")
    update_button_states()

def add_file3():
    global global_KPI_TungCB_DF, global_File_ALL_DLBT_ChuanTemp, global_df_HS_KHCN, global_df_HS_Khong_tinh_KPI
    global global_KPI_DF, global_CB_KoKPI_DF, global_KPI_TungCB_DF

    if global_KPI_TungCB_DF is None:  # Kiểm tra xem global_KPI_TungCB_DF có giá trị hay không
        output_area.insert(tk.END, "!! Chưa có thông tin HR\n--> Add HR_Info trước khi chọn file BC_BTTH\n")
        global_KPI_DF, global_CB_KoKPI_DF, global_KPI_TungCB_DF = Hr_info()
        if global_KPI_DF is None or global_CB_KoKPI_DF is None or global_KPI_TungCB_DF is None:
            messagebox.showwarning("Cảnh báo", "Vui lòng thêm thông tin HR trước khi thực hiện!")
            display_message("!! Chưa có thông tin HR\n--> Vui lòng thao tác lại\n", "")
            return
        else:
            label0.config(text="0. HR_File Loaded!")
            output_area.insert(tk.END, "--> FILE HR_INFO: SHEET KPI, CB_KHONG_TINH_KPI, KPI_TUNG_CB LOADED!\n\n")
        
        # Mở hộp thoại chọn file
        root = tk.Tk()
        root.withdraw()  # Ẩn cửa sổ chính
        output_area.insert(tk.END, "--> Select BC_BTTH_Processed\n")
        file_path = filedialog.askopenfilename(title="Chọn file BC_BTTH_Processed", filetypes=[("Excel files", "*.xlsx;*.xls")])

        if file_path:  # Kiểm tra xem người dùng đã chọn file hay chưa
            # Đọc dữ liệu từ file đã chọn và thêm vào global_File_ALL_DLBT_ChuanTemp
            global_File_ALL_DLBT_ChuanTemp = pl.read_excel(file_path, sheet_name="BTTH_MienNam_Calculated")
            global_df_HS_KHCN = pl.read_excel(file_path, sheet_name="HS_KHCN")
            global_df_HS_Khong_tinh_KPI = pl.read_excel(file_path, sheet_name="HS_Khong_tinh_KPI") 
            # Đọc file Excel# Kết hợp với DataFrame hiện tại
            label3.config(text="3. BC_BTTH loaded!")
            output_area.insert(tk.END, "--> BC_BTTH LOADED!\n")
                       
    else:
        # Mở hộp thoại chọn file
        root = tk.Tk()
        root.withdraw()  # Ẩn cửa sổ chính
        output_area.insert(tk.END, "--> Select BC_BTTH_Processed\n")
        file_path = filedialog.askopenfilename(title="Chọn file BC_BTTH_Processed", filetypes=[("Excel files", "*.xlsx;*.xls")])

        if file_path:  # Kiểm tra xem người dùng đã chọn file hay chưa
            # Đọc dữ liệu từ file đã chọn và thêm vào global_File_ALL_DLBT_ChuanTemp
            global_File_ALL_DLBT_ChuanTemp = pl.read_excel(file_path, sheet_name="BTTH_MienNam_Calculated")
            global_df_HS_KHCN = pl.read_excel(file_path, sheet_name="HS_KHCN")
            global_df_HS_Khong_tinh_KPI = pl.read_excel(file_path, sheet_name="HS_Khong_tinh_KPI") 
            # Đọc file Excel# Kết hợp với DataFrame hiện tại
            label3.config(text="3. BC_BTTH loaded!")
            output_area.insert(tk.END, "--> BC_BTTH loaded!\n")

    update_button_states()        

def update_button_states():
    button1['state'] = 'normal' if (label1.cget("text") != "" and label0.cget("text") != "") else 'disabled'
    button2['state'] = 'normal' if (label2.cget("text") != "" and label0.cget("text") != "") else 'disabled'
    button3['state'] = 'normal' if (label3.cget("text") != "" and label0.cget("text") != "") else 'disabled'


# Create main window
root = tk.Tk()
root.title("KPI Calculator Application")
root.geometry("1200x701+0+0")
root.configure(bg='#2C3E50')

# Tạo style chung
style = {
    "bg": "#2C3E50",
    "fg": "#ECF0F1"
}

# Create a frame for buttons
button_frame = tk.Frame(root, bg='#34495E', padx=20, pady=20) #PADX
button_frame.pack(side=tk.LEFT, padx=15, pady=15)

# Tiêu đề với màu sắc nổi bật
title_label = tk.Label(root, 
                      text="KPI Calculator Application",
                      font=("Montserrat", 28, "bold"),
                      bg='#2C3E50',
                      fg='#3498DB')
title_label.pack(pady=20)

# Style cho các nút chính
button_style = {
    "width": 24,
    "height": 2,
    "font": ("Roboto", 12),
    "bg": "#3498DB",
    "fg": "white",
    "activebackground": "#2980B9",
    "activeforeground": "white",
    "border": 0,
    "cursor": "hand2",
    "relief": "flat"
}

# Tạo các nút với hiệu ứng hover
def on_enter(e):
    e.widget['background'] = '#2980B9' if e.widget['background'] == '#3498DB' else '#219A52'

def on_leave(e):
    e.widget['background'] = '#3498DB' if e.widget['background'] == '#2980B9' else '#27AE60'


# Tạo một frame cho khu vực hiển thị output
output_frame = tk.Frame(root, bg='#34495E', padx=20, pady=20)
output_frame.pack(side=tk.RIGHT, padx=20, pady=0)

# Tạo Text widget với scrollbar
output_area = tk.Text(output_frame, wrap=tk.WORD, height=20, width=0)
output_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Tạo scrollbar và kết nối với Text widget 
scrollbar = tk.Scrollbar(output_frame, command=output_area.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Cấu hình Text widget để sử dụng scrollbar
output_area.config(yscrollcommand=scrollbar.set)



# Original buttons
button1 = tk.Button(button_frame, text="1. Processing PBI_WEB Data", command=Processing_PBI_WEB_data, state='disabled', **button_style)
button1.pack(pady=10)
button1.bind("<Enter>", on_enter)
button1.bind("<Leave>", on_leave)

button2 = tk.Button(button_frame, text="2. Calculate BTTH Data", command=calculate_btth_data, state='disabled', **button_style)
button2.pack(pady=10)
button2.bind("<Enter>", on_enter)
button2.bind("<Leave>", on_leave)

button3 = tk.Button(button_frame, text="3. Calculate KPI", command=kpi_report, state='disabled', **button_style)
button3.pack(pady=10)
button3.bind("<Enter>", on_enter)
button3.bind("<Leave>", on_leave)

button4 = tk.Button(button_frame, text="4. Plot", command=plot, state='active', **button_style)
button4.pack(pady=10)
button4.bind("<Enter>", on_enter)
button4.bind("<Leave>", on_leave)

# New file buttons
file_button_style = {
    "bg": "#28a745",  # Green color for file buttons
    "activebackground": "#218838",
    "width": 20,      # Set the width of the button
    "height": 2       # Set the height of the button
}


button_add1 = tk.Button(button_frame, text="Add Source Step 1", command=add_file1,  **file_button_style)
button_add1.pack(pady=8)
button_add1.bind("<Enter>", on_enter)
button_add1.bind("<Leave>", on_leave)

button_add2 = tk.Button(button_frame, text="Add Source Step 2", command=add_file2, **file_button_style)
button_add2.pack(pady=8)
button_add2.bind("<Enter>", on_enter)
button_add2.bind("<Leave>", on_leave)

button_add3 = tk.Button(button_frame, text="Add Source Step 3", command=add_file3, **file_button_style)
button_add3.pack(pady=8)
button_add3.bind("<Enter>", on_enter)
button_add3.bind("<Leave>", on_leave)

button_add0 = tk.Button(button_frame, text="Add HR Info", command=add_file_HR, **file_button_style)
button_add0.pack(pady=8)
button_add0.bind("<Enter>", on_enter)
button_add0.bind("<Leave>", on_leave)

# Style cho labels
label_style = {
    "font": ("Roboto", 10),
    "bg": '#34495E',
    "fg": "#2ECC71"
}
# Output area với màu sắc hiện đại
output_area = scrolledtext.ScrolledText(
    output_frame,
    width=120,
    height=30,
    font=("Roboto", 12),
    bg="#ECF0F1",
    fg="#2C3E50"
)
label0 = tk.Label(button_frame, text="", **label_style)
label0.pack(pady=5,padx=5)

label1 = tk.Label(button_frame, text="", **label_style)
label1.pack(pady=5,padx=5)

label2 = tk.Label(button_frame, text="", **label_style)
label2.pack(pady=5,padx=5)

label3 = tk.Label(button_frame, text="", **label_style)
label3.pack(pady=5,padx=5)

# Create a scrolled text area for output
output_area = scrolledtext.ScrolledText(output_frame, width=120, height=30, font=("Helvetica", 12))
welcome_message = "--- WELCOME TO KPI APPLICATION Version 1.0 Final ---\n\n"
welcome_message += "   Hướng dẫn. Bạn cần thực hiện các bước sau:\n"
welcome_message += "   Add source tương ứng từng bước. *Lưu ý HR_Info cần cho bước 2 3 4\n\n"
output_area.insert(tk.END, welcome_message)
output_area.pack(fill=tk.BOTH, expand=True)

# Start the application
root.mainloop()

