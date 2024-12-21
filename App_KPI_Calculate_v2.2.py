import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd 
import numpy as np 
from pandas.tseries.offsets import CustomBusinessDay 
from datetime import datetime
import matplotlib.pyplot as plt 
import matplotlib.pyplot as plt1 
import numpy as np 
import tkinter as tk
from tkinter import ttk



import tkinter as tk
from tkinter import filedialog
import polars as pl




def select_file_CBBT_mienNam():
    path_CBBT = filedialog.askopenfilename(title="Chọn File CBBT_MienNam", filetypes=[("Excel files", "*.xlsx *.xlsm")])
    if not path_CBBT:
        result_label.config(text="Không có file CBBT được chọn. Quá trình bị hủy.")
        return None, None
    try:
        # Đọc chỉ sheet KPI từ file Excel
        CBBT_MienNam_DF = pd.read_excel(path_CBBT, sheet_name='KPI')
        return path_CBBT, CBBT_MienNam_DF
    except Exception as e:
        result_label.config(text=f"Lỗi khi đọc sheet KPI: {str(e)}")
        return None, None

def select_file_CB_KoTinhKPI():
    path_CB_koTinhKPI = filedialog.askopenfilename(title="Chọn File CB không tính KPI", filetypes=[("Excel files", "*.xlsx *.xlsm")])
    if not path_CB_koTinhKPI:
        result_label.config(text="Không có file CB không tính KPI. Quá trình bị hủy.")
        return None, None
    try:
        # Đọc chỉ sheet KPI từ file Excel
        CB_koTinhKPI_DF = pd.read_excel(path_CB_koTinhKPI, sheet_name='Sheet1')
        return path_CB_koTinhKPI, CB_koTinhKPI_DF
    except Exception as e:
        result_label.config(text=f"Lỗi khi đọc sheet KPI: {str(e)}")
        return None, None

def select_file_KPI_Tung_CBBT():
    path_KPI_CBBT = filedialog.askopenfilename(title="Chọn File KPI_Tung_CBBT", filetypes=[("Excel files", "*.xlsx *.xlsm")])
    if not path_KPI_CBBT:
        result_label.config(text="Không có file KPI_CBBT được chọn. Quá trình bị hủy.")
        return None, None
    try:
        # Đọc chỉ sheet KPI từ file Excel
        File_KPI_Tung_Canbo = pd.read_excel(path_KPI_CBBT, sheet_name='Sheet1')
        return path_KPI_CBBT, File_KPI_Tung_Canbo
    except Exception as e:
        result_label.config(text=f"Lỗi khi đọc sheet KPI từng CBBT: {str(e)}")
        return None, None        
    

def process_raw_report():
    # Chọn file báo cáo gốc
    raw_report_path = filedialog.askopenfilename(title="Chọn File Báo Cáo Gốc từ PBI", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not raw_report_path:
        result_label.config(text="Không có file báo cáo gốc được chọn. Quá trình bị hủy.")
        return

    # Chọn file CBBT
    path_CBBT, CBBT_MienNam_DF = select_file_CBBT_mienNam()
    if not path_CBBT or CBBT_MienNam_DF is None:
        return

    # Chọn file HS Đã Tính KPI
    path_HS_DaTinhKPI = filedialog.askopenfilename(title="Chọn File HS Đã Tính KPI", filetypes=[("Excel files", "*.xlsx *.xlsm")])
    if not path_HS_DaTinhKPI:
        result_label.config(text="Không có file HS Đã Tính KPI được chọn. Quá trình bị hủy.")
        return

    try:
        # Đọc các file
        df = pd.read_excel(raw_report_path)
        df = df.drop_duplicates(subset='SO_TO_TRINH')
        df_CBBT_MienNam = pd.read_excel(path_CBBT, sheet_name='KPI', usecols=['Họ và tên'])
        df_HS_DaTinhKPI = pd.read_excel(path_HS_DaTinhKPI, sheet_name='Sheet1', usecols=['SO_HS', 'NOTE'])

        # Process the data (add your processing logic here)
        # ...

        df = df.drop_duplicates(subset='SO_TO_TRINH') #Remove duplicate values from SO TO TRINH
        df_CBBT_MienNam = pd.read_excel(path_CBBT,sheet_name='KPI', usecols=['Họ và tên'], ) ## DF CBBT thuộc miền nam
        df_HS_DaTinhKPI = pd.read_excel(path_HS_DaTinhKPI, sheet_name='Sheet1', usecols=['SO_HS', 'NOTE'])  # Đọc dữ liệu đã tính KPI trước đó



        #Xử lý data
        #Lọc CBBT Miền Nam
        df_HS_MienNam = df[df['CAN_BO_BT'].isin(df_CBBT_MienNam['Họ và tên'])]

        #TÁCH HS trùng trong tháng
        So_Hoso = 'SO_HS'
        duplicate_So_Hoso_SameMonth = df_HS_MienNam.duplicated(subset=So_Hoso, keep=False)
        df_duplicate_So_Hoso_SameMonth = df_HS_MienNam.loc[duplicate_So_Hoso_SameMonth]
        df_duplicate_So_Hoso_SameMonth.insert(len(df_duplicate_So_Hoso_SameMonth.columns), "Lý do", "Trùng trong tháng này")


        # Tách HS trùng tháng cũ và thêm giá trị từ cột thứ ba của df_HS_DaTinhKPI
        df_duplicate_So_Hoso_PreviousMonth = df_HS_MienNam[df_HS_MienNam['SO_HS'].isin(df_HS_DaTinhKPI['SO_HS'])]

        # Merge để lấy giá trị cột thứ ba từ df_HS_DaTinhKPI
        df_duplicate_So_Hoso_PreviousMonth = df_duplicate_So_Hoso_PreviousMonth.merge(
            df_HS_DaTinhKPI[['SO_HS','NOTE']],
            on='SO_HS',
            how='left'
        )

        # Thay thế giá trị cột "Lý do" bằng giá trị từ cột 
        df_duplicate_So_Hoso_PreviousMonth['Lý do'] = df_duplicate_So_Hoso_PreviousMonth['NOTE']

        # Xóa cột nếu không cần thiết
        df_duplicate_So_Hoso_PreviousMonth = df_duplicate_So_Hoso_PreviousMonth.drop(columns=['NOTE'])


        ##------------------------------------Xử lý hồ sơ cần kiểm tra trùng, đóng, sai quy định.--------------------------------------
            ##Giá trị cần kiểm tra
        check_Dong = 'Đóng'
        check_TuChoi = 'Từ chối'
        check_MaHS_regex = r'VP\..*\.HS.*'
        check_HanMuc = 'hết hạn mức'
        check_saiNDBH = 'sai người được bảo hiểm'
        check_saiNDBH1 = 'sai N.*BH'
        check_SaiSoThe = 'Sai số thẻ'
        check_SaiSoHD = 'Sai số HD'
        check_SaiSoHopDong = 'Sai số hợp đồng'
        check_SaiHD = 'Sai hd'
        check_SaiHopDong = 'Sai hợp đồng'
        check_ngoaiHieuLuc = 'ngoài hiệu lực'
        check_ngoaiHLBH = 'ngoài HLBH'
        check_ngoaiTGBG = 'ngoài thời gian bảo hiểm'        
        check_khongGQAPP = 'không.*quy định.*giải quyết.*app'
        kekhai = ['kê khai', 'không thuộc đối tượng bảo hiểm', 'kiểm tra CSYT', 'không thực hiện giám định']
        kekhai_regex = '|'.join(kekhai)

        df_DongHSdoTrung = df_HS_MienNam.query(
            '(TIEN_BT_GQ == 0) & '
            '( (GIAI_QUYET.str.contains(@check_TuChoi, case=False)) | '
            '(GIAI_QUYET.str.contains(@check_Dong, case=False)) ) & '
            '(GIAI_QUYET.str.contains(@check_MaHS_regex, case=False, regex=True)) & '
            '(GIAI_QUYET.str.contains(@check_HanMuc, case=False)==False) & '
            '(GIAI_QUYET.str.contains(@kekhai_regex, case=False, regex=True)==False)'
        )

        ## Đóng trùng/từ chối --- Có số hồ sơ --- không bao gồm chữ hết hạn mức.
        df_HSBoiThuongNghiNgo = df_HS_MienNam.query('(TIEN_BT_GQ < 10000) & (TIEN_BT_GQ > 0)') 
        ## Hồ sơ bồi thường ít hơn 5000
        df_HSSaiHopDong = df_HS_MienNam.query('(GIAI_QUYET.str.contains(@check_SaiSoThe, case=False)) | (GIAI_QUYET.str.contains(@check_SaiSoHD, case=False)) | (GIAI_QUYET.str.contains(@check_SaiSoHopDong, case=False)) | (GIAI_QUYET.str.contains(@check_SaiHD, case=False)) | (GIAI_QUYET.str.contains(@check_SaiHopDong, case=False)) | (GIAI_QUYET.str.contains(@check_saiNDBH, case=False)) | (GIAI_QUYET.str.contains(@check_saiNDBH1, case=False))')
        ## Hồ sưo đóng do sai số hợp đồng, sai số thẻ, sai NĐBH
        df_HSngoaiHLBH = df_HS_MienNam.query('(TIEN_BT_GQ == 0) & (   (GIAI_QUYET.str.contains(@check_ngoaiHieuLuc, case=False)) | (GIAI_QUYET.str.contains(@check_ngoaiHLBH, case=False)) | (GIAI_QUYET.str.contains(@check_ngoaiTGBG, case=False))     )')
        ## Hồ sơ đóng do ngoài hiệu lực, đóng theo yêu cầu khách hàng
        df_khongGQAPP = df_HS_MienNam.query('(TIEN_BT_GQ == 0) & (   (GIAI_QUYET.str.contains(@check_khongGQAPP, case=False))   )')
        ##THÊM Cột lý do cho các DF trên
        df_DongHSdoTrung.insert(len(df_DongHSdoTrung.columns),"Lý do",value="Đóng do trùng")
        df_HSBoiThuongNghiNgo.insert(len(df_HSBoiThuongNghiNgo.columns),"Lý do",value="Bồi thường < 10.000đ")
        df_HSSaiHopDong.insert(len(df_HSSaiHopDong.columns),"Lý do",value="Đóng do sai số thẻ/ sai hợp đồng/ sai NĐBH")
        df_HSngoaiHLBH.insert(len(df_HSngoaiHLBH.columns),"Lý do",value="Đóng do ngoài HLBH")
        df_khongGQAPP.insert(len(df_khongGQAPP.columns),"Lý do",value="Đóng do hồ sơ không giải quyết qua app nhưng không yêu cầu bản cứng như quy định")
        ##Combine các DF trên
        Combine_DF = pd.concat([df_DongHSdoTrung, df_HSBoiThuongNghiNgo, df_HSSaiHopDong, df_HSngoaiHLBH, df_khongGQAPP, df_duplicate_So_Hoso_SameMonth, df_duplicate_So_Hoso_PreviousMonth], axis=0, ignore_index=True)
        #______________________________________________________________________________________________________________________________________



        ##-------------------------------HÀM LỌC HS MỀM----------------------------------------------------------------------------------------
        def filter_df_HS_MEM(df, loai_HS_D99, hashtagmem, hashtagcung):
            # Define variables for the query conditions
            min_tien_bt_gq = 10000001
            loai_hs_d99_condition = df['SO_HS'].str.contains(loai_HS_D99, case=False)
            hashtagmem_condition = df['HAU_QUA'].str.contains(hashtagmem, case=False) == False
            hashtagcung_condition = df['HAU_QUA'].str.contains(hashtagcung, case=False) == False

            # Define the query condition as a string expression
            query_condition = (
                "(TIEN_BT_GQ < @min_tien_bt_gq) & "
                "(SO_HS.str.contains(@loai_HS_D99, case=False)) & "
                "(HAU_QUA.str.contains(@hashtagmem, case=False)==False) & "
                "(HAU_QUA.str.contains(@hashtagcung, case=False)==False)"
            ).format(
                min_tien_bt_gq=min_tien_bt_gq,
                loai_HS_D99=loai_HS_D99,
                hashtagmem=hashtagmem,
                hashtagcung=hashtagcung
            )

            # Filter the DataFrame using the query condition
            df_D99MemThieuHASHTAG = df.query(query_condition)

            return df_D99MemThieuHASHTAG
        loai_HS_D99 = 'VP.D99'
        hashtagmem = 'hsmem'
        hashtagcung = 'bản cứng|hs cứng'
        df_D99MemThieuHASHTAG = filter_df_HS_MEM(df_HS_MienNam, loai_HS_D99, hashtagmem, hashtagcung)
        ##-----------------------------------------------------------------------------------------------------------------------
        #______________________________________________________________________________________________________________________________________

        ##############################################################################################################################
        ############################################              Cần xóa                                #############################
        ##############################################################################################################################
        ##Tao File HS QLHA
        hsQLHA = 'qlha'
        df_HsQLHA = df_HS_MienNam.query('HAU_QUA.str.contains(@hsQLHA, case=False, na=False)')
        ##-----------------------------------------------------------------------------------------------------------------------
        ##############################################################################################################################
        ##############################################################################################################################
        ##############################################################################################################################




        ##Tao File HS D99 Mềm
        hsMem = ['hsmem', 'hscung']
        df_HsMem = df_HS_MienNam.query(' | '.join([f"HAU_QUA.str.contains('{x}', case=False, na=False)" for x in hsMem]))

        #______________________________________________________________________________________________________________________________________

        
        #______________________________________________________________________________________________________________________________________


        #add Data ITC to Template KPI 
        df_MienNam_Template_chuan = pd.DataFrame()
        df_MienNam_Template_chuan.insert(0,'STT',df_HS_MienNam['STT'])
        df_MienNam_Template_chuan.insert(1,'Ngày yêu cầu',df_HS_MienNam['NGAY_YC'])
        df_MienNam_Template_chuan.insert(2,'Ngày lập HS',df_HS_MienNam['NGAY_LAP_HS'])
        df_MienNam_Template_chuan.insert(3,'Ngày bổ sung HS',df_HS_MienNam['NGAY_BO_SUNG_HS'])
        df_MienNam_Template_chuan.insert(4,'Ngày duyệt',df_HS_MienNam['APPROVE_DATE'])
        df_MienNam_Template_chuan.insert(5,'Ngày lập tờ trình',df_HS_MienNam['NGAY_LAP_BT'])
        df_MienNam_Template_chuan.insert(6,'Số GYCTT',df_HS_MienNam['SO_HS'])
        df_MienNam_Template_chuan.insert(7,'Số hồ sơ tờ trình bồi thường',df_HS_MienNam['SO_TO_TRINH'])
        df_MienNam_Template_chuan.insert(8,'Cán bộ nhập GYCTT',df_HS_MienNam['CAN_BO_NHAP_HS'])
        df_MienNam_Template_chuan.insert(9,'Cán bộ giải quyết bồi thường',df_HS_MienNam['CAN_BO_BT'])
        df_MienNam_Template_chuan.insert(10,'Ngày lập tờ trình BT',df_HS_MienNam['NGAY_LAP_BT'])
        df_MienNam_Template_chuan.insert(11,'Ngày gửi TBBT',df_HS_MienNam['NGAY_THONG_BAO_BT'])
        df_MienNam_Template_chuan.insert(12,'Ngày giải quyết',df_HS_MienNam['STT'])
        df_MienNam_Template_chuan.insert(13,'Phòng ban',df_HS_MienNam['PHONG_BAN'])
        df_MienNam_Template_chuan.insert(14,'Số tiền bồi thường',df_HS_MienNam['TIEN_BT_GQ'])
        df_MienNam_Template_chuan.insert(15,'Nguyên nhân tổn thất',df_HS_MienNam['NGUYEN_NHAN_RUI_RO'])
        df_MienNam_Template_chuan.insert(16,'Giải quyết',df_HS_MienNam['GIAI_QUYET'])
        df_MienNam_Template_chuan.insert(17,'CB_DUYET',df_HS_MienNam['NGUOI_DUYET'])
        df_MienNam_Template_chuan.insert(18,'Giờ lập',df_HS_MienNam['TG_LAP_TT'].dt.strftime('%H:%M'))
        df_MienNam_Template_chuan.insert(19,'Hậu quả',df_HS_MienNam['HAU_QUA'])                          
        #______________________________________________________________________________________________________________________________________


        #Chỉnh ngày tháng xóa giờ phút giây
        df_MienNam_Template_chuan['Ngày yêu cầu'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày yêu cầu'])
        df_MienNam_Template_chuan['Ngày lập HS'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày lập HS'])
        df_MienNam_Template_chuan['Ngày bổ sung HS'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày bổ sung HS'])
        df_MienNam_Template_chuan['Ngày duyệt'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày duyệt'])
        df_MienNam_Template_chuan['Ngày lập tờ trình BT'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày lập tờ trình BT'])
        df_MienNam_Template_chuan['Ngày gửi TBBT'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày gửi TBBT'])
        for col in ['Ngày yêu cầu', 'Ngày lập HS', 'Ngày bổ sung HS', 'Ngày duyệt', 'Ngày lập tờ trình BT', 'Ngày gửi TBBT']:
            df_MienNam_Template_chuan[col] = df_MienNam_Template_chuan[col].dt.date
        #______________________________________________________________________________________________________________________________________


        ##--------------------------------TẠO FILE ALL DLBT Chuẩn template--------------------------------

        # Ask user for the output directory
        output_dir = filedialog.askdirectory(title="Select Output Directory")

        if output_dir:
            # Create full file paths
            file_path_1 = os.path.join(output_dir, '1-ALL-DLBT-ChuanTemp.xlsx')
            file_path_2 = os.path.join(output_dir, '2-HS D99 Mem.xlsx')
            file_path_3 = os.path.join(output_dir, 'XXXXX_HS_CanKiemTra.xlsx')

            # Write the files
            with pd.ExcelWriter(file_path_1, engine='xlsxwriter') as writer:
                df_MienNam_Template_chuan.to_excel(writer, sheet_name='Baocao', index=False)
                # df_KPI.to_excel(writer, sheet_name='KPI', index=False)
                df_CBBT_MienNam.to_excel(writer, sheet_name='Check-CBBT', index=False)
            
            df_HsMem.to_excel(file_path_2, sheet_name='D99 MEM', index=False, header=True, columns=['SO_HS', 'SO_TO_TRINH', 'HAU_QUA'])

            # Export file cần kiểm tra
            with pd.ExcelWriter(file_path_3, engine='xlsxwriter') as writer:
                Combine_DF.to_excel(writer, sheet_name='HS trùng', index=False)
                df_D99MemThieuHASHTAG.to_excel(writer, sheet_name='D99-Hashtag', index=False)

            result_label.config(text=f"Files saved successfully in {output_dir}")
        else:
            result_label.config(text="No output directory selected. Process cancelled.")


        messagebox.showinfo("Created files successfully ( 1. ALL-DLBT-ChuanTemp.xlsx, 2. HS D99 Mem.xlsx, 3. XXXXXX_HS_CanKiemTra.xlsx )")

        
    except Exception as e:
        result_label.config(text=f"Error processing files: {str(e)}")

    



def create_btth_data():
    def import_file():
        # 1. Chọn file '1-ALL-DLBT-ChuanTemp.xlsx'
        file_path = filedialog.askopenfilename(title="Chọn file '1-ALL-DLBT-ChuanTemp.xlsx'", filetypes=[("Excel files", "*.xlsx")])
        df_MienNam_Template_chuan = pd.read_excel(file_path, sheet_name='Baocao')

        # 2. Chọn file '2-HS D99 Mem.xlsx'
        file_path = filedialog.askopenfilename(title="Chọn file '2-HS D99 Mem.xlsx'", filetypes=[("Excel files", "*.xlsx")])
        df_File_D99_MEM = pd.read_excel(file_path, sheet_name='D99 MEM', usecols=['SO_HS'])

        # 3. Chọn file '6-HS_Trung_KoTinhKPI.xlsx'
        file_path = filedialog.askopenfilename(title="Chọn file '6-HS_Trung_KoTinhKPI.xlsx'", filetypes=[("Excel files", "*.xlsx")])
        df_HS_Khong_tinh_KPI = pd.read_excel(file_path, sheet_name='Sheet1')

        # 4. Chọn file '4-D99 BSCT.xlsx'
        file_path = filedialog.askopenfilename(title="Chọn file '4-D99 BSCT.xlsx'", filetypes=[("Excel files", "*.xlsx")])
        File_D99_BSCT = pd.read_excel(file_path, sheet_name='D99 BSCT', usecols=['Giấy YCTT'])

        # 5. Chọn folder và gộp tất cả file xlsx, xls
        folder_path = filedialog.askdirectory(title="Chọn thư mục chứa các file BTTD")
        dfs = []

        for filename in os.listdir(folder_path):
            if filename.endswith('.xlsx') or filename.endswith('.xls'):
                file_path = os.path.join(folder_path, filename)
                try:
                    df = pd.read_excel(file_path)
                    dfs.append(df)
                except Exception as e:
                    print(f"Lỗi khi đọc file {filename}: {str(e)}")

        df_BTTD = pd.concat(dfs, ignore_index=True)

        return df_MienNam_Template_chuan, df_File_D99_MEM, df_HS_Khong_tinh_KPI, File_D99_BSCT, df_BTTD        
        
    df_MienNam_Template_chuan, df_File_D99_MEM, df_HS_Khong_tinh_KPI, File_D99_BSCT, df_BTTD = import_file()
    # Xử lý dữ liệu tiếp theo ở đây
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
        df_MienNam_Template_chuan[col] = pd.to_datetime(df_MienNam_Template_chuan[col], errors='coerce')

    # Function to calculate network days
    def calculate_network_days(start_date, end_date):
        if pd.isnull(start_date) or pd.isnull(end_date):
            return None
        return len(pd.date_range(start=start_date, end=end_date, freq=custom_bday)) - 1

    # Calculate new columns
    df_MienNam_Template_chuan['1_NgàyYêuCầu_NgàyDuyệt'] = df_MienNam_Template_chuan.apply(
        lambda row: calculate_network_days(row['Ngày yêu cầu'], row['Ngày duyệt']), axis=1)

    df_MienNam_Template_chuan['2_NgàyYêuCầu_NgàyGửiTBBT'] = df_MienNam_Template_chuan.apply(
        lambda row: calculate_network_days(row['Ngày yêu cầu'], row['Ngày gửi TBBT']), axis=1)

    df_MienNam_Template_chuan['3_NgàyBổSungHS_NgàyDuyệt'] = df_MienNam_Template_chuan.apply(
        lambda row: calculate_network_days(row['Ngày bổ sung HS'], row['Ngày duyệt']), axis=1)

    df_MienNam_Template_chuan['4_NgàyBổSungHS_NgàyGửiTBBT'] = df_MienNam_Template_chuan.apply(
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
        df_MienNam_Template_chuan[col] = df_MienNam_Template_chuan[col].apply(replace_negative)

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
    df_MienNam_Template_chuan['5_Status'] = df_MienNam_Template_chuan.apply(determine_status, axis=1)
    df_MienNam_Template_chuan['Số_Ngày_Làm_việc'] = df_MienNam_Template_chuan.apply(calculate_working_days, axis=1)

    def assign_loai_hs(nguyen_nhan):
        if nguyen_nhan == 'Ngoại trú':
            return 1
        elif nguyen_nhan == 'Nội trú':
            return 2
        elif nguyen_nhan == 'Sinh mạng':
            return 3
        else:
            return None  # or any default value you prefer

    df_MienNam_Template_chuan['Loại HS'] = df_MienNam_Template_chuan['Nguyên nhân tổn thất'].apply(assign_loai_hs)
    so_hsmem_set = set(df_File_D99_MEM['SO_HS'])
    df_MienNam_Template_chuan['HS mềm'] = df_MienNam_Template_chuan['Số GYCTT'].apply(lambda x: 'HSMEM' if x in so_hsmem_set else '')
    df_MienNam_Template_chuan['QLHA'] = df_MienNam_Template_chuan['Hậu quả'].apply(lambda x: 'QLHA' if '#QLHA' in str(x) else '')
    df_MienNam_Template_chuan['MaLoaiHS'] = df_MienNam_Template_chuan['Số GYCTT'].str[3:6]
    so_hs_BSCT_set = set(File_D99_BSCT['Giấy YCTT'])
    df_MienNam_Template_chuan['CoBSCT'] = df_MienNam_Template_chuan['Số GYCTT'].apply(lambda x: 'Có BSCT' if x in so_hs_BSCT_set else '')
    so_hs_BTTD_set = set(df_BTTD['Số TTBT'])
    df_MienNam_Template_chuan['BTTD'] = df_MienNam_Template_chuan['Số hồ sơ tờ trình bồi thường'].apply(lambda x: 'BTTĐ' if x in so_hs_BTTD_set else '')
    condition = (
        (df_MienNam_Template_chuan['Số tiền bồi thường'] == 0) &
        (df_MienNam_Template_chuan['Loại HS'] == 2) &
        (
            df_MienNam_Template_chuan['Giải quyết'].str.contains('hết hạn mức', case=False, na=False) |
            df_MienNam_Template_chuan['Giải quyết'].str.contains('không bổ sung', case=False, na=False) |
            df_MienNam_Template_chuan['Giải quyết'].str.contains('hạn bổ sung', case=False, na=False)
        )
    )
    df_MienNam_Template_chuan['HS IP bị giảm tỉ lệ quy đổi'] = np.where(condition, 'X', '')
    df_MienNam_Template_chuan['Hồ sơ không tính KPI'] = np.where(
        df_MienNam_Template_chuan['Số hồ sơ tờ trình bồi thường'].isin(df_HS_Khong_tinh_KPI['Số tờ trình bồi thường']),'X','')
    df_MienNam_Template_chuan.drop(['1_NgàyYêuCầu_NgàyDuyệt','2_NgàyYêuCầu_NgàyGửiTBBT','3_NgàyBổSungHS_NgàyDuyệt','4_NgàyBổSungHS_NgàyGửiTBBT'], axis=1, inplace=True)
    df_MienNam_Template_chuan = df_MienNam_Template_chuan[df_MienNam_Template_chuan['Hồ sơ không tính KPI'] != 'X']


    df_MienNam_Template_chuan['Ngày yêu cầu'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày yêu cầu'])
    df_MienNam_Template_chuan['Ngày lập HS'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày lập HS'])
    df_MienNam_Template_chuan['Ngày bổ sung HS'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày bổ sung HS'])
    df_MienNam_Template_chuan['Ngày duyệt'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày duyệt'])
    df_MienNam_Template_chuan['Ngày lập tờ trình'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày lập tờ trình'])
    df_MienNam_Template_chuan['Ngày lập tờ trình BT'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày lập tờ trình BT'])
    df_MienNam_Template_chuan['Ngày gửi TBBT'] = pd.to_datetime(df_MienNam_Template_chuan['Ngày gửi TBBT'])
    for col in ['Ngày yêu cầu', 'Ngày lập HS', 'Ngày bổ sung HS', 'Ngày duyệt', 'Ngày lập tờ trình BT', 'Ngày gửi TBBT']:
        df_MienNam_Template_chuan[col] = df_MienNam_Template_chuan[col].dt.date

    save_folder = filedialog.askdirectory(title="Chọn thư mục để lưu file Excel")

    if save_folder:
    # Create full path for the file
        file_path = os.path.join(save_folder, 'BC_BTTH_Processed(Python_App).xlsx')

        try:
            # Try to save the Excel file
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df_MienNam_Template_chuan.to_excel(writer, sheet_name='Baocao', index=False)
            
            print(f"File đã được lưu tại: {file_path}")
            messagebox.showinfo("Thành công", "Tạo file dữ liệu BTTH thành công!")
        except PermissionError:
            # Display the custom error message
            error_message = "Không thể lưu file do đang mở"
            print(error_message)
            messagebox.showerror("Lỗi", error_message)
        except Exception as e:
            # Handle other potential errors
            error_message = f"Lỗi khi lưu file: {str(e)}"
            print(error_message)
            messagebox.showerror("Lỗi", error_message)
    else:
        print("Không có thư mục nào được chọn. File không được lưu.")


# Sử dụng hàm
# df_MienNam_Template_chuan, path_File_D99_MEM, df_HS_Khong_tinh_KPI, File_D99_BSCT, df_BTTD = create_btth_data()
   

def count_kpi():
    
    def import_file():
        file_path = filedialog.askopenfilename(title="Chọn file 'BC_BTTH_Processed(Python_App).xlsx'", filetypes=[("Excel files", "*.xlsx")])
        File_ALL_DLBT_ChuanTemp = pd.read_excel(file_path, sheet_name='Baocao')
        return File_ALL_DLBT_ChuanTemp
    File_ALL_DLBT_ChuanTemp = import_file()
    def import_file_KHCN():
        file_path = filedialog.askopenfilename(title="Chọn file '5-KHCN.xlsx'", filetypes=[("Excel files", "*.xlsx")])
        df_HS_KHCN = pd.read_excel(file_path, sheet_name='KHCN')
        return df_HS_KHCN
    df_HS_KHCN =  import_file_KHCN()


    # Chọn file CBBT
    path_CBBT, CBBT_MienNam_DF = select_file_CBBT_mienNam()
    if not path_CBBT or CBBT_MienNam_DF is None:
        return

    KPI_DF = CBBT_MienNam_DF    

    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT NGOẠI TRÚ---------------------------------------------------------------------------------------
    def count_ngoai_tru(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D31', 'D98', 'D33','D15'])) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 1) &
            (File_ALL_DLBT_ChuanTemp['QLHA'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['BTTD'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull())
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT NỘI TRÚ---------------------------------------------------------------------------------------
    def count_noi_tru(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D31', 'D98', 'D33','D15'])) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 2) &
            (File_ALL_DLBT_ChuanTemp['QLHA'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull())
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT QLHA NGOẠI TRÚ---------------------------------------------------------------------------------------
    def count_QLHA_ngoai_tru(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D31', 'D33','D15'])) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 1) &
            (File_ALL_DLBT_ChuanTemp['QLHA']=='QLHA') &
        #    (File_ALL_DLBT_ChuanTemp['BTTD'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull())
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT QLHA NỘI TRÚ---------------------------------------------------------------------------------------
    def count_QLHA_noi_tru(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D31', 'D33','D15'])) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 2) &
            (File_ALL_DLBT_ChuanTemp['QLHA']=='QLHA') &
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull())
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT KHCN---------------------------------------------------------------------------------------
    def count_KHCN(cbbt_name):
        count = (
            (df_HS_KHCN['Cán bộ BT'] == cbbt_name) &
            (df_HS_KHCN['Nghiệp vụ'].isin(['KHN','ATS','PAI','YTK']))         
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT PA,WCI---------------------------------------------------------------------------------------
    def count_PA(cbbt_name):
        count = (
            (df_HS_KHCN['Cán bộ BT'] == cbbt_name) &
            (df_HS_KHCN['Nghiệp vụ'].isin(['PA']))         
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT Du lịch---------------------------------------------------------------------------------------
    def count_DuLich(cbbt_name):
        count = (
            (df_HS_KHCN['Cán bộ BT'] == cbbt_name) &
            (df_HS_KHCN['Nghiệp vụ'].str.lower().isin(['fle','dqt','ydl']))        
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT Du lịch ECEP---------------------------------------------------------------------------------------
    def count_DuLichECEP(cbbt_name):
        count = (
            (df_HS_KHCN['Cán bộ BT'] == cbbt_name) &
            (df_HS_KHCN['Nghiệp vụ'].str.lower().isin(['ecep']))        
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT Tử vong---------------------------------------------------------------------------------------
    def count_Tu_vong(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 3) &
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull())
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT KCARE---------------------------------------------------------------------------------------
    def count_Kcare(cbbt_name):
        count = (
            (df_HS_KHCN['Cán bộ BT'] == cbbt_name) &
            (df_HS_KHCN['Nghiệp vụ'].str.lower().isin(['kcare']))        
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT D99 mềm BS OP---------------------------------------------------------------------------------------
    def count_d99_bsct_op(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &        
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 1) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D99'])) &
            (File_ALL_DLBT_ChuanTemp['HS mềm'] == 'HSMEM') &
            (File_ALL_DLBT_ChuanTemp['CoBSCT'] == 'Có BSCT')
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT D99 mềm Bước 2 OP---------------------------------------------------------------------------------------
    def count_d99_Mem_Buoc2_OP(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &        
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 1) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D99'])) &
            (File_ALL_DLBT_ChuanTemp['HS mềm'] == 'HSMEM') &
            (File_ALL_DLBT_ChuanTemp['BTTD'].isnull())         
        ).sum()
        return count

    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT D99 cứng OP---------------------------------------------------------------------------------------
    def count_d99_Cung_OP(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &        
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 1) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D99'])) &
            (File_ALL_DLBT_ChuanTemp['HS mềm'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['BTTD'].isnull()) 
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT D99 mềm BS IP---------------------------------------------------------------------------------------
    def count_d99_bsct_ip(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &        
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 2) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D99'])) &
            (File_ALL_DLBT_ChuanTemp['HS mềm'] == 'HSMEM') &
            (File_ALL_DLBT_ChuanTemp['CoBSCT'] == 'Có BSCT')
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT D99 mềm Bước 2 IP---------------------------------------------------------------------------------------
    def count_d99_Mem_Buoc2_IP(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &        
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 2) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D99'])) &
            (File_ALL_DLBT_ChuanTemp['HS mềm'] == 'HSMEM')         
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT D99 Cứng IP---------------------------------------------------------------------------------------
    def count_d99_cung_IP(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &        
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 2) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D99'])) &
            (File_ALL_DLBT_ChuanTemp['HS mềm'].isnull())         
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT BTTD D31--------------------------------------------------------------------------------------
    def count_BTTD_D31(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &        
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 1) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D31','D98', 'D33','D15'])) &
            (File_ALL_DLBT_ChuanTemp['BTTD'] == 'BTTĐ') &        
            (File_ALL_DLBT_ChuanTemp['QLHA'].isnull())
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT BTTD D99--------------------------------------------------------------------------------------
    def count_BTTD_D99(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &        
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 1) &
            (File_ALL_DLBT_ChuanTemp['MaLoaiHS'].isin(['D99'])) &
            (File_ALL_DLBT_ChuanTemp['BTTD'] == 'BTTĐ')
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT Không tính KPI--------------------------------------------------------------------------------------
    def count_HS_Khong_Tinh_KPI(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &        
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI']=="X")
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT HSS Giảm tỉ lệ quy đổi IP -> OP--------------------------------------------------------------------------------------
    def count_HS_Giam_TL_Quy_Doi(cbbt_name):
        count = (
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &        
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['HS IP bị giảm tỉ lệ quy đổi'] == 'X' )&
            (File_ALL_DLBT_ChuanTemp['Loại HS'] == 2 ) & 
            (File_ALL_DLBT_ChuanTemp['Số tiền bồi thường'] == 0 )
        ).sum()
        return count
    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT Số lượng hs by CBBT--------------------------------------------------------------------------------------
    def count_tong_HS(cbbt_name):
        return File_ALL_DLBT_ChuanTemp[
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull())
        ]['Số hồ sơ tờ trình bồi thường'].count()

    ##CÁC FUNCTION COUNT HỒ SƠ -----------------COUNT Số lượng HS ĐÚNG HẠN--------------------------------------------------------------------------------------
    def count_ho_so_dung_han(cbbt_name):
        return File_ALL_DLBT_ChuanTemp[
            (File_ALL_DLBT_ChuanTemp['Cán bộ giải quyết bồi thường'] == cbbt_name) &
            (File_ALL_DLBT_ChuanTemp['Hồ sơ không tính KPI'].isnull()) &
            (File_ALL_DLBT_ChuanTemp['5_Status'] == "Đúng hạn")
        ]['Số hồ sơ tờ trình bồi thường'].count()





    #__________________________________________________________________________________________________________________________________________________

    ##COUNT - PHÂN LOẠI HỒ SƠ
    KPI_DF['Ngoại trú'] = KPI_DF['Họ và tên'].apply(count_ngoai_tru)
    KPI_DF['Nội trú'] = KPI_DF['Họ và tên'].apply(count_noi_tru)
    KPI_DF['QLHA Ngoại trú'] = KPI_DF['Họ và tên'].apply(count_QLHA_ngoai_tru)
    KPI_DF['QLHA Nội trú'] = KPI_DF['Họ và tên'].apply(count_QLHA_noi_tru)
    KPI_DF['IJ Ngoại trú'] = 0
    KPI_DF['IJ Nội trú'] = 0
    KPI_DF['KHCN'] = KPI_DF['Họ và tên'].apply(count_KHCN)
    KPI_DF['PA/WCI'] = KPI_DF['Họ và tên'].apply(count_PA)
    KPI_DF['Ngoại giao'] = 0
    KPI_DF['Du Lịch'] = KPI_DF['Họ và tên'].apply(count_DuLich)
    KPI_DF['Tử vong'] = KPI_DF['Họ và tên'].apply(count_Tu_vong)
    KPI_DF['Kcare'] = KPI_DF['Họ và tên'].apply(count_Kcare)
    KPI_DF['CI37'] = 0
    KPI_DF['D99 B1 OP'] = 0
    KPI_DF['D99 mềm BS OP'] = KPI_DF['Họ và tên'].apply(count_d99_bsct_op)
    KPI_DF['D99 mềm Bước 2 OP'] = KPI_DF['Họ và tên'].apply(count_d99_Mem_Buoc2_OP)
    KPI_DF['D99 cứng OP'] = KPI_DF['Họ và tên'].apply(count_d99_Cung_OP)
    KPI_DF['D99 B1 IP'] = 0
    KPI_DF['D99 mềm BS IP'] = KPI_DF['Họ và tên'].apply(count_d99_bsct_ip)
    KPI_DF['D99 mềm B2 IP'] = KPI_DF['Họ và tên'].apply(count_d99_Mem_Buoc2_IP)
    KPI_DF['D99 Cứng IP'] = KPI_DF['Họ và tên'].apply(count_d99_cung_IP)
    KPI_DF['HS Hỗ trợ']= 0
    KPI_DF['BTTD D31'] = KPI_DF['Họ và tên'].apply(count_BTTD_D31)
    KPI_DF['BTTD D99'] = KPI_DF['Họ và tên'].apply(count_BTTD_D99)
    KPI_DF['HS IP -> OP'] = KPI_DF['Họ và tên'].apply(count_HS_Giam_TL_Quy_Doi)
    KPI_DF['HS Không tính KPI'] = KPI_DF['Họ và tên'].apply(count_HS_Khong_Tinh_KPI)
    KPI_DF['Hồ sơ du lịch qua cổng ECEP'] = KPI_DF['Họ và tên'].apply(count_DuLichECEP)
    KPI_DF['Số lượng hồ sơ BVCare chưa quy đổi'] = KPI_DF['Họ và tên'].apply(count_tong_HS)
    KPI_DF['Số hồ sơ đúng hạn'] = KPI_DF['Họ và tên'].apply(count_ho_so_dung_han)
    KPI_DF['%Thời Gian GQ HS_QT063'] = KPI_DF['Số hồ sơ đúng hạn'] / KPI_DF['Số lượng hồ sơ BVCare chưa quy đổi']

    #______________________________________________________________________________________________________________________________________________________________


    ##------------------------------------- TÍNH TOÁN KPI / KPI từng cán bộ---------------------------------------------------------------------------------------
    path_KPI_CBBT, File_KPI_Tung_Canbo = select_file_KPI_Tung_CBBT()
    if not path_KPI_CBBT or File_KPI_Tung_Canbo is None:
        return

    KPI_DF = KPI_DF.merge(File_KPI_Tung_Canbo[['Họ và tên', 'KPI']], on='Họ và tên', how='left')

    KPI_DF['Tổng sau quy đổi'] = (KPI_DF['Ngoại trú'] *1 +KPI_DF['Nội trú'] *2.4 +
                                KPI_DF['QLHA Ngoại trú'] *1.1 +KPI_DF['QLHA Nội trú'] *2.5 +
                                KPI_DF['IJ Ngoại trú'] *2 +KPI_DF['IJ Nội trú'] *4 +
                                KPI_DF['KHCN'] *2 +KPI_DF['PA/WCI'] *3 +KPI_DF['Ngoại giao'] *6 +KPI_DF['Du Lịch'] *12 +
                                KPI_DF['Tử vong'] *15 +KPI_DF['Kcare'] *45 +KPI_DF['CI37'] *60 +
                                KPI_DF['D99 mềm BS OP'] *0.5 +KPI_DF['D99 mềm Bước 2 OP'] *1 +KPI_DF['D99 cứng OP'] *1.5 +
                                KPI_DF['D99 mềm BS IP'] *0.5 +KPI_DF['D99 mềm B2 IP'] *2.4 +KPI_DF['D99 Cứng IP'] *2.9 +
                                KPI_DF['BTTD D31'] *0.85 +KPI_DF['BTTD D99'] *0.85 + KPI_DF['Hồ sơ du lịch qua cổng ECEP'] *1 +
                                (KPI_DF['HS IP -> OP'] *(-1.4)))

    # Tính toán 'TiLeGQDungHan'
    KPI_DF['TiLeGQDungHan'] = np.where(KPI_DF['%Thời Gian GQ HS_QT063'] / 0.95 < 1,
                                    KPI_DF['%Thời Gian GQ HS_QT063'] / 0.95 * 0.2, 0.2)
    # Tính toán 'Rate'
    KPI_DF['Rate'] = (KPI_DF['Tổng sau quy đổi'] / KPI_DF['KPI']) * 0.7 + 0.1 + KPI_DF['TiLeGQDungHan']
    #
    KPI_DF['Tỉ lệ % theo KPI_QT063'] = np.where(KPI_DF['KPI'].isnull(), 0, KPI_DF['Rate'])

    # Tính toán KPI_DF['PhanVuot']
    KPI_DF['PhanVuot'] = np.where(KPI_DF['Tỉ lệ % theo KPI_QT063'] > 1.2,
                                KPI_DF['Tỉ lệ % theo KPI_QT063'] - 1.2,
                                KPI_DF['Tỉ lệ % theo KPI_QT063'])
    # Tính toán KPI_DF['PhanTangThem']
    KPI_DF['PhanTangThem'] = KPI_DF['PhanVuot'] * KPI_DF['KPI'] / 0.7 / KPI_DF['KPI']
    # Tính toán KPI_DF['KPI sau vượt 120%_QT063'] & Gán cột 'Tỉ lệ % theo KPI_QT063'
    KPI_DF['KPI sau vượt 120%_QT063'] = np.where(KPI_DF['Tỉ lệ % theo KPI_QT063'] > 1.2,
                                                1.2 + KPI_DF['PhanTangThem'],
                                                KPI_DF['Tỉ lệ % theo KPI_QT063'])
    KPI_DF = KPI_DF.drop(columns=['TiLeGQDungHan', 'Rate', 'PhanVuot', 'PhanTangThem'])

    KPI_DF = KPI_DF[KPI_DF['Tổng sau quy đổi'] >= 10]



    KPI_columns_to_export = ['Mã nhân viên','Họ và tên','Ngoại trú','Nội trú',
                            'QLHA Ngoại trú','QLHA Nội trú','IJ Ngoại trú','IJ Nội trú',
                            'KHCN','PA/WCI','Ngoại giao','Du Lịch','Tử vong','Kcare','CI37',
                            'D99 B1 OP','D99 mềm BS OP','D99 mềm Bước 2 OP','D99 cứng OP',
                            'D99 B1 IP','D99 mềm BS IP','D99 mềm B2 IP','D99 Cứng IP',
                            'HS Hỗ trợ','BTTD D31','BTTD D99','HS IP -> OP','Hồ sơ du lịch qua cổng ECEP',
                            'Số lượng hồ sơ BVCare chưa quy đổi','Số hồ sơ đúng hạn','%Thời Gian GQ HS_QT063'
                            ]
    KPI_columns_to_export2 = ['Mã nhân viên','Họ và tên','Ngoại trú','Nội trú',
                            'QLHA Ngoại trú','QLHA Nội trú','IJ Ngoại trú','IJ Nội trú',
                            'KHCN','PA/WCI','Ngoại giao','Du Lịch','Tử vong','Kcare','CI37',
                            'D99 B1 OP','D99 mềm BS OP','D99 mềm Bước 2 OP','D99 cứng OP',
                            'D99 B1 IP','D99 mềm BS IP','D99 mềm B2 IP','D99 Cứng IP',
                            'HS Hỗ trợ','BTTD D31','BTTD D99','HS IP -> OP','Hồ sơ du lịch qua cổng ECEP',
                            'Số lượng hồ sơ BVCare chưa quy đổi','Số hồ sơ đúng hạn','%Thời Gian GQ HS_QT063',
                            'Tổng sau quy đổi','Tỉ lệ % theo KPI_QT063','KPI sau vượt 120%_QT063'
                            ]    
    def export_kpi_files(KPI_DF, KPI_columns_to_export, KPI_columns_to_export2, File_ALL_DLBT_ChuanTemp):
        # Chọn thư mục để lưu cả hai file
        save_directory = filedialog.askdirectory(title="Chọn thư mục để lưu các file KPI")
        
        if not save_directory:
            return "Không có thư mục nào được chọn. Quá trình xuất file bị hủy."

        # Tạo đường dẫn đầy đủ cho hai file
        file_path_1 = os.path.join(save_directory, 'KPI_Thang_XX(byPY).xlsx')
        file_path_2 = os.path.join(save_directory, 'KPI_Thang_XX(for_PBI).xlsx')

        # Xuất file KPI_Thang_XX(byPY).xlsx
        KPI_DF_export = KPI_DF[KPI_columns_to_export]
        with pd.ExcelWriter(file_path_1, engine='xlsxwriter') as writer:
            KPI_DF_export.to_excel(writer, sheet_name='Baocao', index=False)
            File_ALL_DLBT_ChuanTemp.to_excel(writer, sheet_name='BTTH', index=False)

        # Xuất file KPI_Thang_XX(for_PBI).xlsx
        KPI_DF_export2 = KPI_DF[KPI_columns_to_export2]
        with pd.ExcelWriter(file_path_2, engine='xlsxwriter') as writer:
            KPI_DF_export2.to_excel(writer, sheet_name='KPI', index=False)
            File_ALL_DLBT_ChuanTemp.to_excel(writer, sheet_name='BTTH', index=False)

        return f"Hai file KPI đã được xuất thành công vào thư mục:\n{save_directory}"
    

    def xu_ly_NLC2():
        return
    


# Sử dụng hàm
    result = export_kpi_files(KPI_DF, KPI_columns_to_export, KPI_columns_to_export2, File_ALL_DLBT_ChuanTemp)



    
    messagebox.showinfo("Thành công", "Quá trình tính toán KPI đã hoàn tất.")




def plot_data():
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    plt.ioff()  # Turn off interactive mode

    def import_file_KPI():
        file_path = filedialog.askopenfilename(title="Chọn file 'KPI_Thang_XX(FOR_PBI).xlsx'", filetypes=[("Excel files", "*.xlsx")])
        KPI_DF = pd.read_excel(file_path, sheet_name='KPI')
        return KPI_DF

    KPI_DF = import_file_KPI()
    X_axis = KPI_DF['Họ và tên']
    Column_y_axis = KPI_DF['Tổng sau quy đổi']
    Line_y_axis = np.round(KPI_DF['KPI sau vượt 120%_QT063'] * 100)

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


    """ import matplotlib.pyplot as plt # type: ignore

    # Assuming you have the following data
    X_axis1 = KPI_DF['Họ và tên']
    Column_y_axis1 = KPI_DF['Số lượng hồ sơ BVCare chưa quy đổi']
    Line_y_axis1 = KPI_DF['Tổng sau quy đổi']

    # Filter the Line_y_axis values that are greater than 100
    filtered_Line_y_axis1 = Line_y_axis1[Line_y_axis1 > 100]
    filtered_X_axis1 = X_axis1[Line_y_axis1 > 100]
    filtered_Column_y_axis1 = Column_y_axis1[Line_y_axis1 > 100]

    # Create the figure and axes
    fig1, ax1 = plt.subplots(figsize=(15, 6))

    # Plot the clustered column chart
    bars1 = ax1.bar(filtered_X_axis1, filtered_Column_y_axis1, label='Số lượng HS chưa quy đổi')

    # Plot the line chart with the filtered data
    ax1.plot(filtered_X_axis1, filtered_Line_y_axis1, color='red', label='Tổng sau quy đổi')

    # Set the labels and title
    ax1.set_xlabel('Họ và tên')
    ax1.set_ylabel('Số lượng hồ sơ')

    # Rotate the x-axis labels by 90 degrees
    ax1.set_xticklabels(filtered_X_axis1, rotation=90)

    # Add data labels to the clustered column chart
    for bar in bars1:
        height = bar.get_height()
        ax1.annotate(str(round(height, 2)), xy=(bar.get_x() + bar.get_width() / 2, height), 
                    xytext=(0, 3), textcoords="offset points", 
                    ha='center', va='bottom', fontsize=8)
    # Add data labels to the line chart
    for x, y in zip(filtered_X_axis1, filtered_Line_y_axis1):
        ax1.annotate(str(round(y, 2)), 
                    xy=(x, y),
                    xytext=(0, 3), textcoords="offset points",
                    ha='center', va='bottom', fontsize=8, color='red',
                    bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="red", alpha=0.8))

    # Add a legend
    ax1.legend()

    # Show the chart
    plt.show() """




# Create the main window
root = tk.Tk()
root.title("Report Processing App")
root.geometry("500x400")

# Configure styles
style = ttk.Style()
style.theme_use('clam')

# Create a frame for buttons
button_frame = ttk.Frame(root, padding=20)
button_frame.pack(expand=True, fill="both")

# Create buttons
btn_process = ttk.Button(button_frame, text="1. Process Raw Report", command=process_raw_report, width=25)
btn_process.pack(pady=10)

btn_create_btth = ttk.Button(button_frame, text="2. Create BTTH Data File", command=create_btth_data, width=25)
btn_create_btth.pack(pady=10)

btn_count_kpi = ttk.Button(button_frame, text="3. Count KPI", command=count_kpi, width=25)
btn_count_kpi.pack(pady=10)

btn_plot = ttk.Button(button_frame, text="4. Plot Data", command=plot_data, width=25)
btn_plot.pack(pady=10)



# Create a label to display the result
result_label = ttk.Label(root, text="", font=("Helvetica", 12))
result_label.pack(pady=20)

# Start the application
root.mainloop()
