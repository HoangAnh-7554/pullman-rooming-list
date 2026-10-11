import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import numpy as np
from io import BytesIO
import unicodedata
import re

st.set_page_config(page_title="Rooming List Checker", page_icon="🏨", layout="wide")

# ==============================================================================
# BỘ LỌC LÀM SẠCH DỮ LIỆU & NHẬN DIỆN
# ==============================================================================
def normalize_name(name):
    """Làm sạch tiếng Việt, sửa lỗi chữ Đ, xóa danh xưng (MR/MS)"""
    if pd.isna(name) or str(name).strip() == "" or str(name).strip().upper() == "NAN": 
        return ""
    # Thay thế chữ Đ tiếng Việt trước khi lọc dấu
    name = str(name).replace('Đ', 'D').replace('đ', 'd')
    name = unicodedata.normalize('NFD', name).encode('ascii', 'ignore').decode('utf-8').upper()
    # Chỉ giữ chữ cái
    name = re.sub(r'[^A-Z\s]', ' ', name)
    # Xóa danh xưng
    name = re.sub(r'\b(MR|MS|MRS|MISS)\b', '', name)
    return ' '.join(name.split())

def normalize_date(date_val):
    """Chuẩn hóa ngày tháng về định dạng chuẩn DD-MM-YY"""
    if pd.isna(date_val) or str(date_val).strip() in ["", "NAN", "NAT"]: 
        return ""
    val_str = str(date_val).strip()
    try:
        # Nếu Excel trả về định dạng YYYY-MM-DD HH:MM:SS
        if len(val_str) >= 10 and val_str[4] == '-':
            dt = pd.to_datetime(val_str)
        else:
            dt = pd.to_datetime(val_str, dayfirst=True)
        return dt.strftime('%d-%m-%y')
    except:
        return val_str

def guess_nationality(name):
    """Phát hiện khách nước ngoài thông qua cấu trúc tên"""
    if pd.isna(name) or str(name).strip() == "" or str(name).strip().upper() == "NAN": 
        return "Chưa rõ"
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
st.title("Phần mềm Đối chiếu Rooming List (Bản Chuẩn File Tour)")

with st.info("Hỗ trợ tải lên biểu mẫu trực tiếp của công ty lữ hành. Hoặc bạn có thể tải Form mẫu tại đây:"):
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
        with st.spinner("Đang lọc dòng rác, bầu chọn phòng gán khách và check xe Bus..."):
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

                # ---------------- BƯỚC 2: ĐỌC VÀ LỌC RÁC EXCEL ĐOÀN ----------------
                excel_bytes = excel_file.read()
                df_raw = pd.read_excel(BytesIO(excel_bytes))
                
                # Cảm biến nhận diện file Tour
                is_tour_format = any('COMPANY TRIP' in str(c).upper() for c in df_raw.columns) or any('COMPANY TRIP' in str(v).upper() for v in df_raw.iloc[0].values)
                
                if is_tour_format:
                    df_excel = pd.read_excel(BytesIO(excel_bytes), skiprows=2, dtype=str)
                    if 'Full Name' in df_excel.columns: df_excel['Ten_Khach'] = df_excel['Full Name']
                    if 'Shared room' in df_excel.columns: df_excel['Nhom_Phong'] = df_excel['Shared room']
                    if 'Bus' in df_excel.columns: df_excel['Bus_Excel'] = df_excel['Bus']
                else:
                    df_excel = pd.read_excel(BytesIO(excel_bytes), dtype=str)

                # KIỂM TRA TÊN VÀ XÓA DÒNG RÁC
                if 'Ten_Khach' in df_excel.columns:
                    df_excel['Ten_Khach_Excel'] = df_excel['Ten_Khach'].apply(normalize_name)
                    # Lọc bỏ các dòng bị trống tên (xóa dòng rác)
                    df_excel = df_excel[df_excel['Ten_Khach_Excel'].str.strip() != ""]
                    df_excel = df_excel[df_excel['Ten_Khach_Excel'] != "NAN"]
                else:
                    st.error("Lỗi: Không tìm thấy cột 'Ten_Khach' hoặc 'Full Name' trong file Excel.")
                    st.stop()
                    
                df_excel = df_excel.rename(columns={'Ngay_Den': 'Ngay_Den_Excel', 'Ngay_Di': 'Ngay_Di_Excel'})
                if 'Ngay_Den_Excel' in df_excel.columns: df_excel['Ngay_Den_Excel'] = df_excel['Ngay_Den_Excel'].apply(normalize_date)
                if 'Ngay_Di_Excel' in df_excel.columns: df_excel['Ngay_Di_Excel'] = df_excel['Ngay_Di_Excel'].apply(normalize_date)
                
                for col in ['So_Phong', 'Nhom_Phong', 'Bus_Excel', 'Ngay_Den_Excel', 'Ngay_Di_Excel']:
                    if col not in df_excel.columns: df_excel[col] = ""

                # ---------------- BƯỚC 3: THUẬT TOÁN "BẦU CHỌN" GÁN PHÒNG ----------------
                name_to_rooms = df_opera.groupby('Ten_Khach_Opera')['So_Phong'].apply(list).to_dict()

                if df_excel['Nhom_Phong'].str.strip().astype(bool).any():
                    group_to_room = {}
                    
                    # Xét từng nhóm (Shared room)
                    for grp, grp_df in df_excel.groupby('Nhom_Phong'):
                        if pd.isna(grp) or str(grp).strip() in ["", "NAN", "NAT", "NONE"]: continue
                        
                        # Gom tất cả các phòng mà thành viên nhóm này CÓ THỂ ở
                        possible_rooms = []
                        for name in grp_df['Ten_Khach_Excel']:
                            if name in name_to_rooms:
                                possible_rooms.extend(name_to_rooms[name])
                                
                        if possible_rooms:
                            # Lấy số phòng xuất hiện nhiều nhất trong nhóm để gán
                            best_room = pd.Series(possible_rooms).mode()[0]
                            group_to_room[grp] = best_room
                            
                            # Xóa phòng đã sử dụng để không gán trùng cho người khác
                            for name in grp_df['Ten_Khach_Excel']:
                                if name in name_to_rooms and best_room in name_to_rooms[name]:
                                    name_to_rooms[name].remove(best_room)
                                    
                    def assign_by_group(row):
                        grp = row['Nhom_Phong']
                        if pd.notna(grp) and str(grp).strip() not in ["", "NAN", "NAT", "NONE"]:
                            return group_to_room.get(grp, "Chưa gán được")
                        
                        # Gán lẻ cho những người không có nhóm
                        name = row['Ten_Khach_Excel']
                        if name in name_to_rooms and len(name_to_rooms[name]) > 0:
                            return name_to_rooms[name].pop(0)
                        return "Chưa gán được"
                    
                    # Chỉ gán tự động cho những người bị trống số phòng
                    df_excel['So_Phong'] = np.where(df_excel['So_Phong'].replace(['', 'NAN', 'NONE'], pd.NA).isna(), df_excel.apply(assign_by_group, axis=1), df_excel['So_Phong'])
                else:
                    # Gán lẻ nếu hoàn toàn không chia nhóm
                    def auto_assign_room(row):
                        current_room = str(row['So_Phong']).strip().upper()
                        if current_room in ["", "NAN", "NAT", "NONE"]:
                            name = row['Ten_Khach_Excel']
                            if name in name_to_rooms and len(name_to_rooms[name]) > 0: 
                                return name_to_rooms[name].pop(0)
                            else: return "Chưa gán được"
                        return current_room
                    df_excel['So_Phong'] = df_excel.apply(auto_assign_room, axis=1)

                # ---------------- BƯỚC 4: ĐỐI CHIẾU & KIỂM TRA XE BUS ----------------
                df_opera = df_opera.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])
                df_excel = df_excel.sort_values(by=['So_Phong', 'Ten_Khach_Excel'])

                # Đánh Index trong phòng để so khớp 1-1 chính xác
                df_opera['Guest_Index'] = df_opera.groupby('So_Phong').cumcount()
                df_excel['Guest_Index'] = df_excel.groupby('So_Phong').cumcount()

                df_compare = pd.merge(df_opera, df_excel, on=['So_Phong', 'Guest_Index'], how='outer', indicator=True)
                
                # Báo động Lệch Bus
                bus_conflicts = []
                for room, group in df_compare.groupby('So_Phong'):
                    if room == "Chưa gán được" or pd.isna(room): continue
                    valid_buses = group['Bus_Excel'].astype(str).replace(['NAN', '', 'NONE', 'NAT'], pd.NA).dropna().unique()
                    if len(valid_buses) > 1: bus_conflicts.append(room)
                
                df_compare['Canh_Bao_Bus'] = ""
                df_compare.loc[df_compare['So_Phong'].isin(bus_conflicts), 'Canh_Bao_Bus'] = df_compare['Bus_Excel'].apply(
                    lambda x: f"⚠️ Khác Bus ({x})" if pd.notna(x) and str(x).strip() not in ["", "NAN"] else ""
                )

                # Phân loại Quốc tịch
                df_compare['Quoc_Tich'] = df_compare.apply(lambda row: guess_nationality(row['Ten_Khach_Opera']) if pd.notna(row['Ten_Khach_Opera']) else guess_nationality(row['Ten_Khach_Excel']), axis=1)

                # Xác định trạng thái lỗi
                def deep_check_status(row):
                    if str(row['So_Phong']) == "Chưa gán được": return "Lỗi nặng: Khách Excel không có trên hệ thống Opera"
                    if row['_merge'] == 'left_only': return "Hủy phòng / Excel gửi sót khách này"
                    if row['_merge'] == 'right_only': return "Phát sinh thêm người ghép" if row['Guest_Index'] > 0 else "Phòng mới (Có trong Excel, chưa có trên Opera)"
                    
                    if row['_merge'] == 'both':
                        if str(row['Ngay_Den_Excel']).strip() in ['NAN', ''] and str(row['Ngay_Di_Excel']).strip() in ['NAN', '']:
                            date_mismatch = False
                        else:
                            date_mismatch = (str(row['Ngay_Den_Opera']) != str(row['Ngay_Den_Excel'])) or (str(row['Ngay_Di_Opera']) != str(row['Ngay_Di_Excel']))
                            
                        name_mismatch = (str(row['Ten_Khach_Opera']) != str(row['Ten_Khach_Excel']))
                        
                        if name_mismatch and date_mismatch: return "Sai lệch KÉP (Sai tên & Lịch trình)"
                        elif name_mismatch: return "Sai lệch Tên khách"
                        elif date_mismatch: return "Sai lệch ngày In/Out"
                        return "Khớp hoàn toàn"
                    return "Cần kiểm tra thủ công"

                df_compare['Trang_Thai'] = df_compare.apply(deep_check_status, axis=1)
                
                # ---------------- BƯỚC 5: DASHBOARD & XUẤT SHEET ----------------
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

                # Format cột xuất báo cáo
                cols = ['So_Phong', 'Quoc_Tich', 'Canh_Bao_Bus', 'Trang_Thai', 'Ten_Khach_Opera', 'Ten_Khach_Excel', 'Bus_Excel', 'Nhom_Phong', 'Ngay_Den_Opera', 'Ngay_Di_Opera', 'Ngay_Den_Excel', 'Ngay_Di_Excel']
                df_compare = df_compare[cols]

                # Chia 2 tập dữ liệu
                df_errors = df_compare[(df_compare['Trang_Thai'] != 'Khớp hoàn toàn') | (df_compare['Canh_Bao_Bus'] != "")].sort_values(by=['Trang_Thai', 'So_Phong'])
                df_full = df_compare.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])

                st.markdown("---")
                st.success("✅ Phân tích hoàn tất! Đã dọn sạch dòng rác, tối ưu việc gán phòng và phát hiện khách lệch Bus.")
                
                def highlight_special(s):
                    return ['background-color: #ffcccc' if s['Quoc_Tich'] == '🌍 Nước Ngoài' else 'background-color: #ffe5b4' if '⚠️' in str(s['Canh_Bao_Bus']) else '' for _ in s]
                
                tab1, tab2 = st.tabs(["⚠️ CÁC PHÒNG BỊ LỖI / LỆCH BUS", "📋 DANH SÁCH TỔNG ĐOÀN"])
                with tab1:
                    if df_errors.empty: st.info("🎉 Tuyệt vời! Tất cả các phòng đều khớp 100% và đi chung Bus.")
                    else: st.dataframe(df_errors.style.apply(highlight_special, axis=1), use_container_width=True)
                with tab2:
                    st.dataframe(df_full.style.apply(highlight_special, axis=1), use_container_width=True)

                # Ghi ra Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_errors.to_excel(writer, index=False, sheet_name='Phong_Can_Luu_Y')
                    df_full.to_excel(writer, index=False, sheet_name='Danh_Sach_Tong')
                    for sheet_name in ['Phong_Can_Luu_Y', 'Danh_Sach_Tong']:
                        worksheet = writer.sheets[sheet_name]
                        for column in worksheet.columns:
                            max_length = max((len(str(cell.value)) for cell in column), default=0)
                            worksheet.column_dimensions[column[0].column_letter].width = max_length + 2

                # Nút tải file
                st.markdown("<br>", unsafe_allow_html=True)
                _, dwn_col, _ = st.columns([1, 1, 1])
                with dwn_col:
                    st.download_button(
                        label="📥 TẢI BÁO CÁO (CHIA 2 SHEET CHUẨN)",
                        data=output.getvalue(),
                        file_name="Bao_Cao_Rooming_List_Hoan_Chinh.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Đã xảy ra lỗi: {e}")