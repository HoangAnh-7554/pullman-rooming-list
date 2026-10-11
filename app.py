import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import numpy as np
from io import BytesIO
import unicodedata
import re

st.set_page_config(page_title="Rooming List Checker Master", page_icon="🏨", layout="wide")

# ==============================================================================
# 1. BỘ LỌC LÀM SẠCH, CHUẨN HÓA VÀ NHẬN DIỆN THÔNG MINH
# ==============================================================================
def normalize_name(name):
    """Làm sạch tiếng Việt, sửa lỗi chữ Đ, xóa danh xưng, đảo tên"""
    if pd.isna(name) or str(name).strip() in ["", "NAN", "NAT", "NONE"]: 
        return ""
    name = str(name).replace('Đ', 'D').replace('đ', 'd')
    name = unicodedata.normalize('NFD', name).encode('ascii', 'ignore').decode('utf-8').upper()
    name = re.sub(r'[^A-Z\s]', ' ', name) # Xóa ký tự đặc biệt
    name = re.sub(r'\b(MR|MS|MRS|MISS|DR|PROF|CHILD)\b', '', name) # Xóa danh xưng
    return ' '.join(name.split())

def normalize_room(room):
    """Khắc phục lỗi mất số 0 ở đầu (Ví dụ: 201 -> 0201)"""
    r = str(room).replace('.0', '').strip().upper()
    if r.isdigit() and len(r) == 3:
        return '0' + r
    return r

def normalize_date(date_val):
    """Chuẩn hóa ngày tháng về định dạng chuẩn DD-MM-YY cho mọi trường hợp"""
    if pd.isna(date_val) or str(date_val).strip() in ["", "NAN", "NAT"]: return ""
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
    """Phát hiện khách nước ngoài thông qua cấu trúc từ vựng"""
    if pd.isna(name) or str(name).strip() == "" or str(name).strip().upper() == "NAN": return "Chưa rõ"
    name = str(name).upper()
    if re.search(r'[FJWZ]', name): return "🌍 Nước Ngoài"
    words = name.split()
    for word in words:
        if len(word) > 1 and (re.search(r'[BDFJKLRSVWXZ]$', word) or word.endswith('TH')):
            return "🌍 Nước Ngoài"
    return "Việt Nam"

def find_column(df, keywords):
    """Tự động tìm kiếm tên cột dựa trên từ khóa (Auto-detect Headers)"""
    for col in df.columns:
        c_clean = str(col).upper().replace('_', ' ').replace('-', ' ')
        for kw in keywords:
            if kw in c_clean:
                return col
    return None

# ==============================================================================
# 2. TẠO FILE EXCEL MẪU
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
        ws = writer.sheets['Danh_Sach_Doan']
        for col, width in zip(['A','B','C','D','E','F'], [15, 30, 15, 15, 20, 20]):
            ws.column_dimensions[col].width = width
    return output.getvalue()

# ==============================================================================
# 3. GIAO DIỆN CHÍNH STREAMLIT
# ==============================================================================
st.title("Phần mềm Đối chiếu Rooming List (Ultimate Version)")

with st.info("Hỗ trợ đọc thẳng file gốc của Công ty lữ hành (Không cần format lại). Hoặc dùng file Template chuẩn bên dưới:"):
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
        start_check = st.button("🚀 Bắt đầu Quét Đa Luồng", type="primary", use_container_width=True)

    if start_check:
        with st.spinner("Hệ thống đang dò tìm cột, nhận diện đảo tên và bầu chọn phòng..."):
            try:
                # ---------------- BƯỚC 1: ĐỌC XML OPERA ----------------
                tree = ET.parse(xml_file)
                root = tree.getroot()
                opera_guests = []
                for guest in root.iter('G_CONFIRMATION_NO'):
                    # Lọc bỏ khách đã CANCELLED
                    status = guest.find('RESV_STATUS').text if guest.find('RESV_STATUS') is not None else ''
                    if 'CANCEL' in str(status).upper():
                        continue
                        
                    room = guest.find('ROOM').text if guest.find('ROOM') is not None else ''
                    name = guest.find('FULL_NAME').text if guest.find('FULL_NAME') is not None else ''
                    arrival = guest.find('ARRIVAL').text if guest.find('ARRIVAL') is not None else ''
                    departure = guest.find('DEPARTURE').text if guest.find('DEPARTURE') is not None else ''
                    
                    opera_guests.append({
                        'So_Phong': normalize_room(room),
                        'Ten_Khach_Opera': normalize_name(name),
                        'Ngay_Den_Opera': normalize_date(arrival),
                        'Ngay_Di_Opera': normalize_date(departure)
                    })
                df_opera = pd.DataFrame(opera_guests)
                opera_all_names = df_opera['Ten_Khach_Opera'].tolist()

                # ---------------- BƯỚC 2: ĐỌC EXCEL ĐOÀN AUTO-DETECT ----------------
                excel_bytes = excel_file.read()
                df_raw = pd.read_excel(BytesIO(excel_bytes))
                
                # Bỏ qua các dòng rác Header của file Tour an toàn tuyệt đối với mọi kiểu dữ liệu
                is_tour_format = False
                for idx, row in df_raw.head(10).iterrows():
                    row_str = ' '.join([str(val).upper() for val in row.values])
                    if 'FULL NAME' in row_str or 'TÊN' in row_str:
                        df_excel = pd.read_excel(BytesIO(excel_bytes), skiprows=idx, dtype=str)
                        is_tour_format = True
                        break
                
                if not is_tour_format:
                    df_excel = pd.read_excel(BytesIO(excel_bytes), dtype=str)

                # TỰ ĐỘNG MAP TÊN CỘT (Thông minh)
                col_name = find_column(df_excel, ['FULL NAME', 'TEN KHACH', 'TÊN KHÁCH', 'HỌ TÊN', 'GUEST'])
                col_room = find_column(df_excel, ['SO PHONG', 'SỐ PHÒNG', 'ROOM'])
                col_group = find_column(df_excel, ['SHARED ROOM', 'NHOM PHONG', 'NHÓM', 'SHARE'])
                col_bus = find_column(df_excel, ['BUS', 'XE', 'CHUYEN XE', 'SỐ XE'])
                col_in = find_column(df_excel, ['ARRIVAL', 'NGAY DEN', 'NGÀY ĐẾN', 'CHECK IN'])
                col_out = find_column(df_excel, ['DEPARTURE', 'NGAY DI', 'NGÀY ĐI', 'CHECK OUT'])

                if col_name: df_excel['Ten_Khach_Excel'] = df_excel[col_name].apply(normalize_name)
                else:
                    st.error("Lỗi: Không thể tìm thấy cột chứa Tên Khách trong file Excel.")
                    st.stop()

                df_excel['So_Phong'] = df_excel[col_room].apply(normalize_room) if col_room else ""
                df_excel['Nhom_Phong'] = df_excel[col_group] if col_group else ""
                df_excel['Bus_Excel'] = df_excel[col_bus] if col_bus else ""
                df_excel['Ngay_Den_Excel'] = df_excel[col_in].apply(normalize_date) if col_in else ""
                df_excel['Ngay_Di_Excel'] = df_excel[col_out].apply(normalize_date) if col_out else ""

                # Xóa dòng rác
                df_excel = df_excel[df_excel['Ten_Khach_Excel'].str.strip() != ""]
                df_excel = df_excel[df_excel['Ten_Khach_Excel'] != "NAN"]

                # ---------------- BƯỚC 3: SMART MAPPING & FUZZY MATCH ----------------
                name_to_rooms = df_opera.groupby('Ten_Khach_Opera')['So_Phong'].apply(list).to_dict()
                
                fuzzy_name_map = {frozenset(k.split()): k for k in name_to_rooms.keys()}

                def get_possible_rooms(excel_name):
                    if excel_name in name_to_rooms:
                        return name_to_rooms[excel_name]
                    excel_set = frozenset(excel_name.split())
                    if excel_set in fuzzy_name_map:
                        real_op_name = fuzzy_name_map[excel_set]
                        return name_to_rooms[real_op_name]
                    return []

                def consume_room(excel_name, room_to_consume):
                    if excel_name in name_to_rooms and room_to_consume in name_to_rooms[excel_name]:
                        name_to_rooms[excel_name].remove(room_to_consume)
                        return
                    excel_set = frozenset(excel_name.split())
                    if excel_set in fuzzy_name_map:
                        real_op_name = fuzzy_name_map[excel_set]
                        if room_to_consume in name_to_rooms[real_op_name]:
                            name_to_rooms[real_op_name].remove(room_to_consume)

                # Thuật toán Bầu chọn (Nhóm)
                if df_excel['Nhom_Phong'].str.strip().astype(bool).any():
                    group_to_room = {}
                    for grp, grp_df in df_excel.groupby('Nhom_Phong'):
                        if pd.isna(grp) or str(grp).strip() in ["", "NAN", "NAT", "NONE"]: continue
                        
                        possible_rooms = []
                        for name in grp_df['Ten_Khach_Excel']:
                            possible_rooms.extend(get_possible_rooms(name))
                                
                        if possible_rooms:
                            best_room = pd.Series(possible_rooms).mode()[0]
                            group_to_room[grp] = best_room
                            for name in grp_df['Ten_Khach_Excel']:
                                consume_room(name, best_room)
                                    
                    def assign_by_group(row):
                        grp = row['Nhom_Phong']
                        if pd.notna(grp) and str(grp).strip() not in ["", "NAN", "NAT", "NONE"]:
                            return group_to_room.get(grp, "Chưa gán được")
                        rooms = get_possible_rooms(row['Ten_Khach_Excel'])
                        if rooms:
                            rm = rooms[0]
                            consume_room(row['Ten_Khach_Excel'], rm)
                            return rm
                        return "Chưa gán được"
                    
                    df_excel['So_Phong'] = np.where(df_excel['So_Phong'].replace(['', 'NAN', 'NONE'], pd.NA).isna(), df_excel.apply(assign_by_group, axis=1), df_excel['So_Phong'])
                else:
                    def auto_assign_room(row):
                        current_room = str(row['So_Phong']).strip()
                        if current_room in ["", "NAN", "NAT", "NONE"]:
                            rooms = get_possible_rooms(row['Ten_Khach_Excel'])
                            if rooms:
                                rm = rooms[0]
                                consume_room(row['Ten_Khach_Excel'], rm)
                                return rm
                            return "Chưa gán được"
                        return current_room
                    df_excel['So_Phong'] = df_excel.apply(auto_assign_room, axis=1)

                # ---------------- BƯỚC 4: ĐỐI CHIẾU & KIỂM TRA BUS ----------------
                df_opera = df_opera.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])
                df_excel = df_excel.sort_values(by=['So_Phong', 'Ten_Khach_Excel'])

                df_opera['Guest_Index'] = df_opera.groupby('So_Phong').cumcount()
                df_excel['Guest_Index'] = df_excel.groupby('So_Phong').cumcount()

                df_compare = pd.merge(df_opera, df_excel, on=['So_Phong', 'Guest_Index'], how='outer', indicator=True)
                
                bus_conflicts = []
                for room, group in df_compare.groupby('So_Phong'):
                    if room == "Chưa gán được" or pd.isna(room): continue
                    valid_buses = group['Bus_Excel'].astype(str).replace(['NAN', '', 'NONE', 'NAT'], pd.NA).dropna().unique()
                    if len(valid_buses) > 1: bus_conflicts.append(room)
                
                df_compare['Canh_Bao_Bus'] = ""
                df_compare.loc[df_compare['So_Phong'].isin(bus_conflicts), 'Canh_Bao_Bus'] = df_compare['Bus_Excel'].apply(
                    lambda x: f"⚠️ Khác Bus ({x})" if pd.notna(x) and str(x).strip() not in ["", "NAN"] else ""
                )

                df_compare['Quoc_Tich'] = df_compare.apply(lambda row: guess_nationality(row['Ten_Khach_Opera']) if pd.notna(row['Ten_Khach_Opera']) else guess_nationality(row['Ten_Khach_Excel']), axis=1)

                def deep_check_status(row):
                    if str(row['So_Phong']) == "Chưa gán được": return "Lỗi nặng: Khách Excel không có trên hệ thống Opera"
                    if row['_merge'] == 'left_only': return "Hủy phòng / Excel gửi sót khách này"
                    if row['_merge'] == 'right_only': return "Phát sinh thêm người ghép" if row['Guest_Index'] > 0 else "Phòng mới (Có trong Excel, chưa có trên Opera)"
                    
                    if row['_merge'] == 'both':
                        if str(row['Ngay_Den_Excel']).strip() in ['NAN', ''] and str(row['Ngay_Di_Excel']).strip() in ['NAN', '']:
                            date_mismatch = False
                        else:
                            date_mismatch = (str(row['Ngay_Den_Opera']) != str(row['Ngay_Den_Excel'])) or (str(row['Ngay_Di_Opera']) != str(row['Ngay_Di_Excel']))
                            
                        name_op = str(row['Ten_Khach_Opera'])
                        name_ex = str(row['Ten_Khach_Excel'])
                        name_mismatch = set(name_op.split()) != set(name_ex.split())
                        
                        if name_mismatch and date_mismatch: return "Sai lệch KÉP (Sai tên & Lịch trình)"
                        elif name_mismatch: return "Sai lệch Tên khách"
                        elif date_mismatch: return "Sai lệch ngày In/Out"
                        return "Khớp hoàn toàn"
                    return "Cần kiểm tra thủ công"

                df_compare['Trang_Thai'] = df_compare.apply(deep_check_status, axis=1)
                
                # ---------------- BƯỚC 5: DASHBOARD & XUẤT 3 SHEET ----------------
                tong_khach_opera = len(df_opera)
                tong_phong_opera = df_opera['So_Phong'].nunique()
                tong_khach_excel = len(df_excel)
                tong_phong_excel = df_excel['So_Phong'].astype(str).replace(['', 'NAN', 'NAT', 'NONE', 'CHƯA GÁN ĐƯỢC'], np.nan).dropna().nunique()

                st.markdown("### 📊 THỐNG KÊ TỔNG QUAN ĐOÀN")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Tổng Khách (Opera)", f"{tong_khach_opera} pax")
                m2.metric("Tổng Phòng (Opera)", f"{tong_phong_opera} phòng")
                m3.metric("Tổng Khách (Excel)", f"{tong_khach_excel} pax", delta=f"{tong_khach_excel - tong_khach_opera} khách" if tong_khach_excel != tong_khach_opera else None)
                m4.metric("Tổng Phòng (Excel)", f"{tong_phong_excel} phòng", delta=f"{tong_phong_excel - tong_phong_opera} phòng" if tong_phong_excel != tong_phong_opera else None)

                cols = ['So_Phong', 'Quoc_Tich', 'Canh_Bao_Bus', 'Trang_Thai', 'Ten_Khach_Opera', 'Ten_Khach_Excel', 'Bus_Excel', 'Nhom_Phong', 'Ngay_Den_Opera', 'Ngay_Di_Opera', 'Ngay_Den_Excel', 'Ngay_Di_Excel']
                df_compare = df_compare[cols]

                df_errors = df_compare[
                    (df_compare['Trang_Thai'] != 'Khớp hoàn toàn') | 
                    (df_compare['Canh_Bao_Bus'] != "") |
                    (df_compare['Quoc_Tich'] == '🌍 Nước Ngoài')
                ].sort_values(by=['Trang_Thai', 'So_Phong'])
                
                df_full = df_compare.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])

                df_valid_rooms = df_compare[df_compare['So_Phong'].astype(str).str.strip() != "Chưa gán được"].copy()
                df_valid_rooms['Bus_Excel'] = df_valid_rooms['Bus_Excel'].replace(['', 'NAN', 'NAT', 'NONE'], 'Tự túc / KHÔNG CÓ XE')
                df_valid_rooms['Ten_Hien_Thi'] = df_valid_rooms.apply(lambda r: str(r['Ten_Khach_Opera']) if pd.notna(r['Ten_Khach_Opera']) and str(r['Ten_Khach_Opera']).strip() != 'NAN' else str(r['Ten_Khach_Excel']), axis=1)
                
                df_key_cut = df_valid_rooms.groupby(['Bus_Excel', 'So_Phong']).agg(
                    So_Luong_The_Can_Cat=('Ten_Hien_Thi', 'count'),
                    Danh_Sach_Khach_Trong_Phong=('Ten_Hien_Thi', lambda x: '  |  '.join([str(i) for i in x if str(i).strip() not in ["", "NAN"]]))
                ).reset_index()
                df_key_cut = df_key_cut.rename(columns={'Bus_Excel': 'Chuyến Xe (Bus)', 'So_Phong': 'Số Phòng'})
                df_key_cut = df_key_cut.sort_values(by=['Chuyến Xe (Bus)', 'Số Phòng'])

                st.markdown("---")
                st.success("✅ Phân tích hoàn tất! Dữ liệu đã được chuẩn hóa, đảo chiều tên thành công và trích xuất Bảng cắt Keycard.")
                
                def highlight_special(s):
                    return ['background-color: #ffcccc' if s.get('Quoc_Tich') == '🌍 Nước Ngoài' else 'background-color: #ffe5b4' if '⚠️' in str(s.get('Canh_Bao_Bus', '')) else '' for _ in s]
                
                tab1, tab2, tab3 = st.tabs(["⚠️ BẢNG CẦN LƯU Ý", "📋 DANH SÁCH TỔNG ĐOÀN", "🔑 BẢNG LÀM THẺ KHÓA (CHIA BUS)"])
                with tab1:
                    if df_errors.empty: st.info("🎉 Tuyệt vời! Đoàn khớp 100%, đi chung Bus và thuần Việt.")
                    else: st.dataframe(df_errors.style.apply(highlight_special, axis=1), use_container_width=True)
                with tab2:
                    st.dataframe(df_full.style.apply(highlight_special, axis=1), use_container_width=True)
                with tab3:
                    st.dataframe(df_key_cut, use_container_width=True)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_errors.to_excel(writer, index=False, sheet_name='Phong_Can_Luu_Y')
                    df_full.to_excel(writer, index=False, sheet_name='Danh_Sach_Tong')
                    df_key_cut.to_excel(writer, index=False, sheet_name='Cat_Key_Chia_Bus')
                    
                    for sheet_name in ['Phong_Can_Luu_Y', 'Danh_Sach_Tong', 'Cat_Key_Chia_Bus']:
                        worksheet = writer.sheets[sheet_name]
                        for column in worksheet.columns:
                            max_length = max((len(str(cell.value)) for cell in column), default=0)
                            worksheet.column_dimensions[column[0].column_letter].width = max_length + 2

                st.markdown("<br>", unsafe_allow_html=True)
                _, dwn_col, _ = st.columns([1, 1, 1])
                with dwn_col:
                    st.download_button(
                        label="📥 TẢI BÁO CÁO ULTIMATE (3 SHEETS)",
                        data=output.getvalue(),
                        file_name="Bao_Cao_Rooming_List_Ultimate.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Đã xảy ra lỗi hệ thống: {e}")