import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import numpy as np
from io import BytesIO
import unicodedata
import re

st.set_page_config(page_title="Rooming List Checker", page_icon="🏨", layout="wide")

# ==============================================================================
# HÀM LÀM SẠCH VÀ NHẬN DIỆN
# ==============================================================================
def normalize_name(name):
    """Làm sạch tiếng Việt, sửa lỗi chữ Đ, xóa danh xưng (MR/MS)"""
    if pd.isna(name) or str(name).strip() == "": return ""
    name = str(name).replace('Đ', 'D').replace('đ', 'd')
    name = unicodedata.normalize('NFD', name).encode('ascii', 'ignore').decode('utf-8').upper()
    name = re.sub(r'[^A-Z\s]', ' ', name)
    name = re.sub(r'\b(MR|MS|MRS|MISS)\b', '', name)
    return ' '.join(name.split())

def normalize_date(date_val):
    """Chuẩn hóa ngày tháng về định dạng chuẩn DD-MM-YY"""
    if pd.isna(date_val) or str(date_val).strip() in ["", "NAN", "NAT"]: return ""
    val_str = str(date_val).strip()
    try:
        # Nhận diện định dạng hệ thống YYYY-MM-DD
        if len(val_str) >= 10 and val_str[4] == '-':
            dt = pd.to_datetime(val_str)
        else:
            dt = pd.to_datetime(val_str, dayfirst=True)
        return dt.strftime('%d-%m-%y')
    except:
        return val_str

def guess_nationality(name):
    """Phát hiện khách nước ngoài thông qua cấu trúc tên"""
    if pd.isna(name) or str(name).strip() == "": return "Chưa rõ"
    name = str(name).upper()
    if re.search(r'[FJWZ]', name): return "🌍 Nước Ngoài"
    words = name.split()
    for word in words:
        if len(word) > 1 and (re.search(r'[BDFJKLRSVWXZ]$', word) or word.endswith('TH')):
            return "🌍 Nước Ngoài"
    return "Việt Nam"

# ==============================================================================
# HÀM TẠO FILE EXCEL MẪU (TEMPLATE)
# ==============================================================================
def create_template_excel():
    data = {
        'So_Phong': ['', '0201', '0201', '0301'],
        'Ten_Khach': ['Joe Woolf', 'Huynh Duc Phong', 'Nguyen Tai Uyen', 'Khach Moi'],
        'Nhom_Phong': ['1', '2', '2', '3'],
        'Bus': ['Bus 1', 'Bus 2', 'Bus 2', 'Bus 3'],
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
        worksheet.column_dimensions['C'].width = 15
        worksheet.column_dimensions['D'].width = 15
        worksheet.column_dimensions['E'].width = 20
        worksheet.column_dimensions['F'].width = 20
    return output.getvalue()

# ==============================================================================
# GIAO DIỆN CHÍNH
# ==============================================================================
st.title("Phần mềm Đối chiếu Rooming List (Bản Master)")

with st.info("Tải file Excel mẫu về máy để điền thông tin đoàn (Hỗ trợ file Tour tự động nhận diện cột Tên, Nhóm và Bus)."):
    st.download_button(
        label="📥 TẢI FILE EXCEL MẪU TRẮNG",
        data=create_template_excel(),
        file_name="Template_Rooming_List.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("---")
col1, col2 = st.columns(2)
with col1:
    xml_file = st.file_uploader("1. Tải lên file XML (Từ hệ thống Opera)", type=['xml'])
with col2:
    excel_file = st.file_uploader("2. Tải lên file Excel (Từ Khách đoàn / TA)", type=['xlsx', 'xls'])

if xml_file and excel_file:
    _, btn_col, _ = st.columns([1, 1, 1])
    with btn_col:
        start_check = st.button("🚀 Bắt đầu phân tích toàn diện", type="primary", use_container_width=True)

    if start_check:
        with st.spinner("Đang đọc luồng dữ liệu, gán phòng theo nhóm và check xe Bus..."):
            try:
                # ---------------- BƯỚC 1: ĐỌC XML OPERA ----------------
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

                # ---------------- BƯỚC 2: ĐỌC EXCEL ĐOÀN THÔNG MINH ----------------
                excel_bytes = excel_file.read() # Đọc file vào bộ nhớ để xử lý an toàn
                df_raw = pd.read_excel(BytesIO(excel_bytes))
                
                # AI tự nhận diện xem đây có phải là file do Tour thiết kế riêng không (Có các dòng tiêu đề rác)
                is_tour_format = any('COMPANY TRIP' in str(c).upper() for c in df_raw.columns) or any('COMPANY TRIP' in str(v).upper() for v in df_raw.iloc[0].values)
                
                if is_tour_format:
                    df_excel = pd.read_excel(BytesIO(excel_bytes), skiprows=2, dtype=str)
                    # Gán tên cột linh hoạt cho file Tour
                    if 'Full Name' in df_excel.columns: df_excel['Ten_Khach'] = df_excel['Full Name']
                    if 'Shared room' in df_excel.columns: df_excel['Nhom_Phong'] = df_excel['Shared room']
                    if 'Bus' in df_excel.columns: df_excel['Bus_Excel'] = df_excel['Bus']
                else:
                    df_excel = pd.read_excel(BytesIO(excel_bytes), dtype=str)

                # Kiểm tra lại tên cột
                if 'Ten_Khach' in df_excel.columns:
                    df_excel['Ten_Khach_Excel'] = df_excel['Ten_Khach'].apply(normalize_name)
                else:
                    st.error("Lỗi: Không tìm thấy cột 'Ten_Khach' hoặc 'Full Name' trong file Excel.")
                    st.stop()
                    
                df_excel = df_excel.rename(columns={'Ngay_Den': 'Ngay_Den_Excel', 'Ngay_Di': 'Ngay_Di_Excel'})
                if 'Ngay_Den_Excel' in df_excel.columns: df_excel['Ngay_Den_Excel'] = df_excel['Ngay_Den_Excel'].apply(normalize_date)
                if 'Ngay_Di_Excel' in df_excel.columns: df_excel['Ngay_Di_Excel'] = df_excel['Ngay_Di_Excel'].apply(normalize_date)
                
                # Thêm các cột thiếu để tránh lỗi
                for col in ['So_Phong', 'Nhom_Phong', 'Bus_Excel', 'Ngay_Den_Excel', 'Ngay_Di_Excel']:
                    if col not in df_excel.columns: df_excel[col] = ""

                # ---------------- BƯỚC 3: GÁN PHÒNG THEO NHÓM (SHARED ROOM) ----------------
                if df_excel['Nhom_Phong'].str.strip().astype(bool).any():
                    # Nếu file có chia nhóm chung phòng
                    name_to_room = df_opera.groupby('Ten_Khach_Opera')['So_Phong'].first().to_dict()
                    df_excel['Mapped_Room'] = df_excel['Ten_Khach_Excel'].map(name_to_room)
                    
                    def get_group_room(group):
                        valid = group.replace(['', 'NAN', 'NAT', 'NONE'], pd.NA).dropna()
                        return valid.iloc[0] if not valid.empty else "Chưa gán được"
                    
                    group_to_room = df_excel.groupby('Nhom_Phong')['Mapped_Room'].apply(get_group_room).to_dict()
                    
                    def assign_by_group(row):
                        grp = row['Nhom_Phong']
                        if pd.notna(grp) and str(grp).strip() not in ["", "NAN", "NAT"]:
                            return group_to_room.get(grp, "Chưa gán được")
                        return "Chưa gán được"
                    
                    # Cập nhật số phòng mới chỉ cho những khách bị trống
                    df_excel['So_Phong'] = np.where(df_excel['So_Phong'].replace(['', 'NAN'], pd.NA).isna(), df_excel.apply(assign_by_group, axis=1), df_excel['So_Phong'])
                else:
                    # Gán lẻ tẻ nếu file không chia cột Nhóm
                    room_mapping = df_opera.groupby('Ten_Khach_Opera')['So_Phong'].apply(list).to_dict()
                    def auto_assign_room(row):
                        current_room = str(row['So_Phong']).strip().upper()
                        if current_room in ["", "NAN", "NAT", "NONE"]:
                            name = row['Ten_Khach_Excel']
                            if name in room_mapping and len(room_mapping[name]) > 0: 
                                return room_mapping[name].pop(0)
                            else: return "Chưa gán được"
                        return current_room
                    df_excel['So_Phong'] = df_excel.apply(auto_assign_room, axis=1)

                # ---------------- BƯỚC 4: ĐỐI CHIẾU & CHECK XE BUS ----------------
                df_opera = df_opera.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])
                df_excel = df_excel.sort_values(by=['So_Phong', 'Ten_Khach_Excel'])

                # Đánh số thứ tự trong phòng để so khớp chính xác
                df_opera['Guest_Index'] = df_opera.groupby('So_Phong').cumcount()
                df_excel['Guest_Index'] = df_excel.groupby('So_Phong').cumcount()

                df_compare = pd.merge(df_opera, df_excel, on=['So_Phong', 'Guest_Index'], how='outer', indicator=True)
                
                # Cảnh báo lệch Bus (Chỉ hiện khi cùng 1 phòng mà có > 1 Bus)
                bus_conflicts = []
                for room, group in df_compare.groupby('So_Phong'):
                    if room == "Chưa gán được" or pd.isna(room): continue
                    valid_buses = group['Bus_Excel'].astype(str).replace(['NAN', '', 'NONE', 'NAT'], pd.NA).dropna().unique()
                    if len(valid_buses) > 1: bus_conflicts.append(room)
                
                df_compare['Canh_Bao_Bus'] = ""
                df_compare.loc[df_compare['So_Phong'].isin(bus_conflicts), 'Canh_Bao_Bus'] = df_compare['Bus_Excel'].apply(
                    lambda x: f"⚠️ Khác Bus ({x})" if pd.notna(x) and str(x).strip() not in ["", "NAN"] else ""
                )

                # Nhận diện quốc tịch
                df_compare['Quoc_Tich'] = df_compare.apply(lambda row: guess_nationality(row['Ten_Khach_Opera']) if pd.notna(row['Ten_Khach_Opera']) else guess_nationality(row['Ten_Khach_Excel']), axis=1)

                # Quy tắc Trạng Thái Báo Lỗi
                def deep_check_status(row):
                    merge_status = row['_merge']
                    room = str(row['So_Phong'])
                    name_op = str(row['Ten_Khach_Opera'])
                    name_ex = str(row['Ten_Khach_Excel'])
                    date_in_op = str(row['Ngay_Den_Opera'])
                    date_out_op = str(row['Ngay_Di_Opera'])
                    date_in_ex = str(row['Ngay_Den_Excel']).strip()
                    date_out_ex = str(row['Ngay_Di_Excel']).strip()

                    if room == "Chưa gán được": return "Lỗi nặng: Khách Excel không có trên hệ thống Opera"
                    if merge_status == 'left_only': return "Hủy phòng / Excel gửi sót khách này"
                    if merge_status == 'right_only': return "Phát sinh thêm người ghép" if row['Guest_Index'] > 0 else "Phòng mới (Có trong Excel, chưa có trên Opera)"
                    
                    if merge_status == 'both':
                        # Nếu file Excel không có cột ngày, ta châm chước bỏ qua lỗi ngày
                        if date_in_ex in ['NAN', ''] and date_out_ex in ['NAN', '']:
                            date_mismatch = False
                        else:
                            date_mismatch = (date_in_op != date_in_ex) or (date_out_op != date_out_ex)
                            
                        name_mismatch = (name_op != name_ex)
                        
                        if name_mismatch and date_mismatch: return "Sai lệch KÉP (Sai tên & Lịch trình)"
                        elif name_mismatch: return "Sai lệch Tên khách"
                        elif date_mismatch: return "Sai lệch ngày In/Out"
                        return "Khớp hoàn toàn"
                    return "Cần kiểm tra thủ công"

                df_compare['Trang_Thai'] = df_compare.apply(deep_check_status, axis=1)
                
                # ---------------- BƯỚC 5: TẠO GIAO DIỆN TỔNG QUAN ----------------
                tong_khach_opera = len(df_opera)
                tong_phong_opera = df_opera['So_Phong'].nunique()
                tong_khach_excel = len(df_excel)
                tong_phong_excel = df_excel['So_Phong'].astype(str).str.upper().replace(['', 'NAN', 'NAT', 'NONE', 'CHƯA GÁN ĐƯỢC'], np.nan).dropna().nunique()

                st.markdown("### 📊 THỐNG KÊ TỔNG QUAN ĐOÀN")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Tổng Khách (Opera)", f"{tong_khach_opera} pax")
                m2.metric("Tổng Phòng (Opera)", f"{tong_phong_opera} phòng")
                m3.metric("Tổng Khách (Excel)", f"{tong_khach_excel} pax", delta=f"{tong_khach_excel - tong_khach_opera} khách" if tong_khach_excel != tong_khach_opera else None)
                m4.metric("Tổng Phòng (Excel)", f"{tong_phong_excel} phòng", delta=f"{tong_phong_excel - tong_phong_opera} phòng" if tong_phong_excel != tong_phong_opera else None)

                # Format lại Cột cho gọn
                cols = ['So_Phong', 'Quoc_Tich', 'Canh_Bao_Bus', 'Trang_Thai', 'Ten_Khach_Opera', 'Ten_Khach_Excel', 'Bus_Excel', 'Nhom_Phong', 'Ngay_Den_Opera', 'Ngay_Di_Opera', 'Ngay_Den_Excel', 'Ngay_Di_Excel']
                df_compare = df_compare[cols]

                # Tách 2 Sheet báo cáo
                df_errors = df_compare[(df_compare['Trang_Thai'] != 'Khớp hoàn toàn') | (df_compare['Canh_Bao_Bus'] != "")].sort_values(by=['Trang_Thai', 'So_Phong'])
                df_full = df_compare.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])

                st.markdown("---")
                st.success("✅ Đã xử lý xong! Bạn có thể xem kết quả trực tiếp hoặc tải file về.")
                
                def highlight_special(s):
                    return ['background-color: #ffcccc' if s['Quoc_Tich'] == '🌍 Nước Ngoài' else 'background-color: #ffe5b4' if '⚠️' in str(s['Canh_Bao_Bus']) else '' for _ in s]
                
                tab1, tab2 = st.tabs(["⚠️ CÁC PHÒNG BỊ LỖI / LỆCH BUS", "📋 DANH SÁCH TỔNG ĐOÀN"])
                
                with tab1:
                    if df_errors.empty: st.info("🎉 Tuyệt vời! Tất cả các phòng đều khớp 100% và đi chung Bus.")
                    else: st.dataframe(df_errors.style.apply(highlight_special, axis=1), use_container_width=True)
                with tab2:
                    st.dataframe(df_full.style.apply(highlight_special, axis=1), use_container_width=True)

                # Xuất ra bộ nhớ để tải
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_errors.to_excel(writer, index=False, sheet_name='Phong_Can_Luu_Y')
                    df_full.to_excel(writer, index=False, sheet_name='Danh_Sach_Tong')
                    for sheet_name in ['Phong_Can_Luu_Y', 'Danh_Sach_Tong']:
                        worksheet = writer.sheets[sheet_name]
                        for column in worksheet.columns:
                            max_length = max((len(str(cell.value)) for cell in column), default=0)
                            worksheet.column_dimensions[column[0].column_letter].width = max_length + 2

                # Nút tải file Excel kết quả
                st.markdown("<br>", unsafe_allow_html=True)
                _, dwn_col, _ = st.columns([1, 1, 1])
                with dwn_col:
                    st.download_button(
                        label="📥 TẢI BÁO CÁO KẾT QUẢ ĐỐI CHIẾU (2 SHEETS)",
                        data=output.getvalue(),
                        file_name="Bao_Cao_Rooming_List.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Đã xảy ra lỗi hệ thống: {e}")