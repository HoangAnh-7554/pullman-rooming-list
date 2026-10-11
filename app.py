import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import numpy as np
from io import BytesIO
import unicodedata
import re

st.set_page_config(page_title="Rooming List Checker", page_icon="🏨", layout="wide")

# ==============================================================================
# BỘ LỌC LÀM SẠCH VÀ NHẬN DIỆN NGÔN NGỮ
# ==============================================================================
def normalize_name(name):
    if not isinstance(name, str) or pd.isna(name):
        return ""
    name = name.replace('Đ', 'D').replace('đ', 'd')
    name = unicodedata.normalize('NFD', name).encode('ascii', 'ignore').decode('utf-8')
    name = name.upper()
    name = re.sub(r'[^A-Z\s]', ' ', name)
    name = re.sub(r'\b(MR|MS|MRS|MISS)\b', '', name)
    name = ' '.join(name.split())
    return name

def normalize_date(date_val):
    if pd.isna(date_val) or str(date_val).strip() in ["", "NAN", "NAT"]:
        return ""
    val_str = str(date_val).strip()
    try:
        if len(val_str) >= 10 and val_str[4] == '-':
            dt = pd.to_datetime(val_str)
        else:
            dt = pd.to_datetime(val_str, dayfirst=True)
        return dt.strftime('%d-%m-%y')
    except:
        return val_str

def guess_nationality(name):
    """Thuật toán phát hiện khách nước ngoài dựa trên cấu trúc từ"""
    if not isinstance(name, str) or name.strip() == "":
        return "Chưa rõ"
    name = name.upper()
    # 1. Chứa các chữ cái không có trong bảng chữ cái tiếng Việt
    if re.search(r'[FJWZ]', name):
        return "🌍 Nước Ngoài"
    # 2. Chứa từ kết thúc bằng các phụ âm không dùng trong tiếng Việt
    words = name.split()
    for word in words:
        if len(word) > 1:
            if re.search(r'[BDFJKLRSVWXZ]$', word) or word.endswith('TH'):
                return "🌍 Nước Ngoài"
    return "Việt Nam"

# ==============================================================================
# HÀM TẠO FILE EXCEL MẪU
# ==============================================================================
def create_template_excel():
    data = {
        'So_Phong': ['', '0201', '0201', '0301'],
        'Ten_Khach': ['Joe Woolf', 'Huynh Duc Phong', 'Nguyen Tai Uyen', 'Khach Moi'],
        'Ngay_Den': ['03-10-26', '03-10-26', '03-10-26', '03-10-26'],
        'Ngay_Di': ['04-10-26', '04-10-26', '04-10-26', '05-10-26']
    }
    df_template = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Danh_Sach_Doan')
        worksheet = writer.sheets['Danh_Sach_Doan']
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['B'].width = 30
        worksheet.column_dimensions['C'].width = 20
        worksheet.column_dimensions['D'].width = 20
    return output.getvalue()

# ==============================================================================
# GIAO DIỆN CHÍNH
# ==============================================================================
st.title("Phần mềm Đối chiếu Rooming List (Tích hợp Dashboard)")

with st.info("Tải file Excel mẫu về máy để điền thông tin đoàn trước khi tải lên (Để trống 'So_Phong' nếu muốn Tool tự gán)."):
    st.download_button(
        label="📥 Tải File Excel Mẫu (Template)",
        data=create_template_excel(),
        file_name="Template_Rooming_List_Doan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("---")
col1, col2 = st.columns(2)
with col1:
    xml_file = st.file_uploader("1. Tải lên file XML (Từ hệ thống Opera)", type=['xml'])
with col2:
    excel_file = st.file_uploader("2. Tải lên file Excel (Từ Khách đoàn)", type=['xlsx', 'xls'])

if xml_file and excel_file:
    _, btn_col, _ = st.columns([1, 1, 1])
    with btn_col:
        start_check = st.button("🚀 Bắt đầu đối chiếu", type="primary", use_container_width=True)

    if start_check:
        with st.spinner("Đang xử lý dữ liệu và nhận diện Quốc tịch..."):
            try:
                # --- ĐỌC XML (OPERA) ---
                tree = ET.parse(xml_file)
                root = tree.getroot()
                opera_guests = []
                for guest in root.iter('G_CONFIRMATION_NO'):
                    room = guest.find('ROOM').text if guest.find('ROOM') is not None else ''
                    name = guest.find('FULL_NAME').text if guest.find('FULL_NAME') is not None else ''
                    arrival = guest.find('ARRIVAL').text if guest.find('ARRIVAL') is not None else ''
                    departure = guest.find('DEPARTURE').text if guest.find('DEPARTURE') is not None else ''
                    
                    opera_guests.append({
                        'So_Phong': str(room).strip(),
                        'Ten_Khach_Opera': normalize_name(name),
                        'Ngay_Den_Opera': normalize_date(arrival),
                        'Ngay_Di_Opera': normalize_date(departure)
                    })
                df_opera = pd.DataFrame(opera_guests)
                opera_all_names = df_opera['Ten_Khach_Opera'].tolist()

                # --- ĐỌC EXCEL (ĐOÀN) ---
                df_excel = pd.read_excel(excel_file, dtype=str)
                if 'Ten_Khach' in df_excel.columns:
                    df_excel['Ten_Khach_Excel'] = df_excel['Ten_Khach'].apply(normalize_name)
                else:
                    st.error("Lỗi: File Excel của bạn không có cột 'Ten_Khach'.")
                    st.stop()
                
                df_excel = df_excel.rename(columns={'Ngay_Den': 'Ngay_Den_Excel', 'Ngay_Di': 'Ngay_Di_Excel'})
                if 'Ngay_Den_Excel' in df_excel.columns:
                    df_excel['Ngay_Den_Excel'] = df_excel['Ngay_Den_Excel'].apply(normalize_date)
                if 'Ngay_Di_Excel' in df_excel.columns:
                    df_excel['Ngay_Di_Excel'] = df_excel['Ngay_Di_Excel'].apply(normalize_date)
                if 'So_Phong' not in df_excel.columns:
                    df_excel['So_Phong'] = ""

                # --- TỰ ĐỘNG GÁN SỐ PHÒNG ---
                room_mapping = df_opera.groupby('Ten_Khach_Opera')['So_Phong'].apply(list).to_dict()
                def auto_assign_room(row):
                    current_room = str(row['So_Phong']).strip().upper()
                    if current_room in ["", "NAN", "NAT", "NONE"]:
                        name = row['Ten_Khach_Excel']
                        if name in room_mapping and len(room_mapping[name]) > 0:
                            return room_mapping[name].pop(0)
                        else:
                            return "Chưa gán được"
                    return current_room

                df_excel['So_Phong'] = df_excel.apply(auto_assign_room, axis=1)

                # --- DASHBOARD TỔNG QUAN ---
                # Đếm số phòng và số khách
                tong_khach_opera = len(df_opera)
                tong_phong_opera = df_opera['So_Phong'].nunique()
                
                tong_khach_excel = len(df_excel)
                # Đếm phòng Excel (bỏ qua những khách chưa gán được phòng)
                phong_excel_valid = df_excel['So_Phong'].astype(str).str.upper().replace(['', 'NAN', 'NAT', 'NONE', 'CHƯA GÁN ĐƯỢC'], np.nan).dropna()
                tong_phong_excel = phong_excel_valid.nunique()

                st.markdown("### 📊 THỐNG KÊ TỔNG QUAN ĐOÀN")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Tổng Khách (Opera)", f"{tong_khach_opera} pax")
                m2.metric("Tổng Phòng (Opera)", f"{tong_phong_opera} phòng")
                m3.metric("Tổng Khách (Excel)", f"{tong_khach_excel} pax", delta=f"{tong_khach_excel - tong_khach_opera} khách" if tong_khach_excel != tong_khach_opera else None)
                m4.metric("Tổng Phòng (Excel)", f"{tong_phong_excel} phòng", delta=f"{tong_phong_excel - tong_phong_opera} phòng" if tong_phong_excel != tong_phong_opera else None)

                # --- ĐỐI CHIẾU ---
                df_opera = df_opera.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])
                df_excel = df_excel.sort_values(by=['So_Phong', 'Ten_Khach_Excel'])

                df_opera['Guest_Index'] = df_opera.groupby('So_Phong').cumcount()
                df_excel['Guest_Index'] = df_excel.groupby('So_Phong').cumcount()

                df_compare = pd.merge(df_opera, df_excel, on=['So_Phong', 'Guest_Index'], how='outer', indicator=True)
                
                # --- PHÂN LOẠI QUỐC TỊCH (Khách Nước Ngoài) ---
                # Quét tên khách từ Opera (ưu tiên), nếu không có thì lấy tên từ Excel
                df_compare['Quoc_Tich'] = df_compare.apply(
                    lambda row: guess_nationality(row['Ten_Khach_Opera']) if pd.notna(row['Ten_Khach_Opera']) else guess_nationality(row['Ten_Khach_Excel']), 
                    axis=1
                )

                def deep_check_status(row):
                    merge_status = row['_merge']
                    room = str(row['So_Phong'])
                    name_op = str(row['Ten_Khach_Opera'])
                    name_ex = str(row['Ten_Khach_Excel'])
                    date_in_op = str(row['Ngay_Den_Opera'])
                    date_out_op = str(row['Ngay_Di_Opera'])
                    date_in_ex = str(row['Ngay_Den_Excel'])
                    date_out_ex = str(row['Ngay_Di_Excel'])
                    
                    if room == "Chưa gán được": return "Lỗi nặng: Khách Excel không có trên hệ thống Opera"
                    if merge_status == 'left_only': return "Hủy phòng / Excel gửi sót khách này"
                    if merge_status == 'right_only':
                        if date_in_ex in ['NAN', ''] or date_out_ex in ['NAN', '']: return "Lỗi định dạng: File Excel bị trống ngày In/Out"
                        if name_ex in opera_all_names: return "Sai số phòng (Khách này đang xếp ở phòng khác trên Opera)"
                        if row['Guest_Index'] > 0: return "Phát sinh thêm người ghép (Dư pax so với Opera)"
                        else: return "Phòng mới (Có trong Excel, chưa có trên Opera)"
                            
                    if merge_status == 'both':
                        if date_in_ex in ['NAN', ''] or date_out_ex in ['NAN', '']: return "Lỗi định dạng: File Excel bị trống ngày In/Out"
                        name_mismatch = (name_op != name_ex)
                        date_mismatch = (date_in_op != date_in_ex) or (date_out_op != date_out_ex)
                        
                        if name_mismatch and date_mismatch: return "Sai lệch KÉP (Sai tên & Sai Lịch trình)"
                        elif name_mismatch: return "Sai lệch Tên khách"
                        elif date_mismatch: return "Sai lệch ngày In/Out"
                        else: return "Khớp hoàn toàn"
                    return "Cần kiểm tra thủ công"

                df_compare['Trang_Thai'] = df_compare.apply(deep_check_status, axis=1)
                
                # Làm gọn bảng và sắp xếp
                df_compare = df_compare.drop(columns=['_merge', 'Guest_Index'])
                df_compare['Is_Perfect'] = df_compare['Trang_Thai'] == 'Khớp hoàn toàn'
                df_compare = df_compare.sort_values(by=['Is_Perfect', 'So_Phong'])
                df_compare = df_compare.drop(columns=['Is_Perfect'])
                
                # Sắp xếp lại thứ tự cột cho đẹp
                cols = ['So_Phong', 'Quoc_Tich', 'Trang_Thai', 'Ten_Khach_Opera', 'Ten_Khach_Excel', 'Ngay_Den_Opera', 'Ngay_Di_Opera', 'Ngay_Den_Excel', 'Ngay_Di_Excel']
                df_compare = df_compare[cols]

                st.markdown("---")
                st.success("✅ Đã xử lý tất cả các trường hợp ngoại lệ thành công!")
                
                # Hiển thị trên web với màu sắc Highlight cho Khách nước ngoài
                def highlight_foreigner(s):
                    return ['background-color: #ffcccc' if v == '🌍 Nước Ngoài' else '' for v in s]
                
                st.dataframe(df_compare.style.apply(highlight_foreigner, subset=['Quoc_Tich']), use_container_width=True)

                # --- XUẤT FILE EXCEL ---
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_compare.to_excel(writer, index=False, sheet_name='Ket_Qua_Doi_Chieu')
                    worksheet = writer.sheets['Ket_Qua_Doi_Chieu']
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        worksheet.column_dimensions[column_letter].width = max_length + 2

                _, dwn_col, _ = st.columns([1, 1, 1])
                with dwn_col:
                    st.download_button(
                        label="📥 TẢI FILE BÁO CÁO CÁC LỖI ĐỐI CHIẾU",
                        data=output.getvalue(),
                        file_name="Bao_Cao_Rooming_List_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Đã xảy ra lỗi hệ thống: {e}")