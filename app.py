import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import numpy as np
from io import BytesIO
import unicodedata
import re
from collections import defaultdict
from difflib import SequenceMatcher

st.set_page_config(page_title="Rooming List Checker Master", page_icon="🏨", layout="wide")

# ==============================================================================
# 1. CẤU HÌNH NHẬN DIỆN ÂM TIẾT TIẾNG VIỆT & BỘ LỌC LÀM SẠCH DỮ LIỆU
# ==============================================================================
VN_SYLLABLE_STRICT = re.compile(
    r'^('
    r'(NGH|NG|NH|CH|GH|GI|KH|PH|TH|TR|[BCDGHKLMNPRSTVX])?'
    r'(A|AI|AO|AU|AY|AM|AN|ANG|ANH|ACH|AP|AT|AC|'
    r'E|EO|EM|EN|ENG|ENH|ECH|EP|ET|EC|'
    r'I|IA|IEU|IU|IM|IN|INH|ICH|IP|IT|IC|IEP|IET|IEC|IENG|IEN|IEM|'
    r'O|OI|OA|OAI|OAY|OAN|OANG|OANH|OACH|OAP|OAT|OAC|OE|OEM|OEN|OET|OM|ON|ONG|OP|OT|OC|OO|OONG|'
    r'U|UA|UOI|UI|UU|UOU|UE|UEN|UET|UENH|UECH|UYA|UY|UAI|UAY|UAN|UANG|UAT|UAC|UON|UONG|UOM|UOP|UOT|UOC|UM|UN|UNG|UP|UT|UC|UYEN|UYET|UYNH|UYCH|'
    r'Y|YEU|YEN|YET|YEM)'
    r'|'
    r'QU(A|AI|AO|AU|AY|AM|AN|ANG|ANH|ACH|AP|AT|AC|E|EO|EM|EN|ENG|ENH|ECH|EP|ET|I|IA|IM|IN|INH|ICH|IP|IT|IEP|IET|IEN|IEM|O|OI|OC|ON|ONG|OT|U|Y|YA|YET|YEN|YNH|YCH)'
    r')$',
    re.IGNORECASE
)

def normalize_name(name):
    """Làm sạch tiếng Việt, sửa lỗi chữ Đ, xóa danh xưng & học hàm/học vị (PGS, TS, MD, Assoc. Prof...)"""
    if pd.isna(name) or str(name).strip().upper() in ["", "NAN", "NAT", "NONE"]: 
        return ""
    name = str(name).replace('Đ', 'D').replace('đ', 'd')
    name = unicodedata.normalize('NFD', name).encode('ascii', 'ignore').decode('utf-8').upper()
    name = re.sub(r'[^A-Z\s]', ' ', name)
    name = re.sub(r'\b(MR|MS|MRS|MISS|MSTR|DR|PROF|ASSOC|ASST|PGS|GS|TS|THS|BS|CKI|CKII|MD|PHD|MBA|BSC|MSC|CHILD|BABY|INFANT)\b', '', name)
    return ' '.join(name.split())

def is_valid_guest_row(raw_name):
    """Loại bỏ các dòng tổng kết cuối bảng như 'Tổng số phòng', 'Total', 'Grand Total'"""
    if pd.isna(raw_name): return False
    s = str(raw_name).strip().upper()
    if s in ["", "NAN", "NAT", "NONE"]: return False
    norm_s = normalize_name(s)
    if norm_s == "": return False
    summary_keywords = ['TONG SO PHONG', 'TONG CONG', 'TONG SO', 'TOTAL', 'GRAND TOTAL', 'SUBTOTAL', 'TONG KHACH']
    for kw in summary_keywords:
        if kw in norm_s:
            return False
    return True

def normalize_room(room):
    """Khắc phục lỗi mất số 0 ở đầu của Excel (Ví dụ: 402 -> 0402)"""
    if pd.isna(room): return ""
    r = str(room).replace('.0', '').strip().upper()
    if r in ["", "NAN", "NAT", "NONE"]: return ""
    if r.isdigit() and len(r) == 3:
        return '0' + r
    return r

def is_real_room_column(series):
    """Phân biệt cột Số phòng thật (0201, 0402...) với cột đếm số lượng phòng (1, 1, 1...)"""
    valid_vals = [str(v).replace('.0', '').strip() for v in series.dropna() if str(v).strip().upper() not in ["", "NAN", "NAT", "NONE"]]
    if not valid_vals:
        return False
    numeric_vals = [int(v) for v in valid_vals if v.isdigit()]
    if numeric_vals and max(numeric_vals) < 100:
        return False
    return True

def is_real_group_column(series):
    """Phân biệt cột Nhóm Phòng (Shared room) với cột Chức danh/Vai trò (Chủ tọa, Báo cáo viên...)"""
    valid_vals = [str(v).replace('.0', '').strip() for v in series.dropna() if str(v).strip().upper() not in ["", "NAN", "NAT", "NONE"]]
    if not valid_vals:
        return False
    val_counts = pd.Series(valid_vals).value_counts()
    # Một phòng khách sạn không thể ghép quá 6 người; nếu > 6 người cùng 1 mã nhóm thì đó là cột Chức danh/Bộ phận
    if val_counts.max() > 6:
        return False
    return True

def normalize_date(date_val):
    """Chuẩn hóa ngày tháng về định dạng chuẩn DD-MM-YY"""
    if pd.isna(date_val) or str(date_val).strip().upper() in ["", "NAN", "NAT", "NONE"]: 
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

def standardize_room_type(rt):
    """Chuẩn hóa Loại phòng (King/Twin/Single) để so sánh chéo"""
    if pd.isna(rt) or str(rt).strip().upper() in ["", "NAN", "NAT", "NONE"]: 
        return "UNKNOWN"
    rt_str = str(rt).upper().strip()
    if 'TWIN' in rt_str or 'TWB' in rt_str or re.search(r'\bTW\b', rt_str):
        return "TWIN"
    if 'KING' in rt_str or 'KGB' in rt_str or 'KGA' in rt_str or 'DBL' in rt_str or 'DOUBLE' in rt_str or 'SINGLE' in rt_str or 'SGL' in rt_str:
        return "KING"
    return "UNKNOWN"

def extract_special_notes(text):
    """Tự động phát hiện các yêu cầu đặc biệt: Late Check-out, Early Check-in, Extra Bed, Baby Cot, VIP"""
    if pd.isna(text) or str(text).strip().upper() in ["", "NAN", "NAT", "NONE"]:
        return ""
    s_raw = ' '.join(str(text).replace('\n', ' - ').split())
    s_up = s_raw.upper()
    notes = []
    if any(k in s_up for k in ['LATE', 'CHECKOUT LATE', 'LATE CHECK OUT', 'LATE C/O', 'TRẢ MUỘN', 'TRA MUON']):
        notes.append(f"⏰ Late C/O ({s_raw})")
    elif any(k in s_up for k in ['EARLY', 'SỚM', 'EARLY C/I']):
        notes.append(f"🌅 Early C/I ({s_raw})")
    if any(k in s_up for k in ['EXTRA BED', 'EXB', 'GIƯỜNG PHỤ', 'GHÉP GIƯỜNG', 'GHEP GIUONG']):
        notes.append(f"🛏️ Extra/Ghép ({s_raw})")
    if any(k in s_up for k in ['BABY COT', 'CRIB', 'NÔI']):
        notes.append(f"👶 Baby Cot ({s_raw})")
    if 'VIP' in s_up:
        notes.append(f"⭐ VIP ({s_raw})")
    return " | ".join(notes)

def guess_nationality(name, excel_foreigner_flag=""):
    """Phát hiện khách nước ngoài qua cột đánh dấu Excel HOẶC cấu trúc âm tiết tiếng Việt"""
    if pd.notna(excel_foreigner_flag) and str(excel_foreigner_flag).strip().upper() not in ["", "NAN", "NAT", "NONE", "0", "0.0", "NO", "N", "FALSE", "VN", "VIET NAM", "VIETNAM", "VIỆT NAM"]:
        return "🌍 Nước Ngoài"
    if pd.isna(name) or str(name).strip().upper() in ["", "NAN", "NAT", "NONE"]: 
        return "Chưa rõ"
    words = str(name).upper().strip().split()
    if not words:
        return "Chưa rõ"
    for w in words:
        if not VN_SYLLABLE_STRICT.match(w):
            return "🌍 Nước Ngoài"
    return "Việt Nam"

def find_column(df, keywords):
    """Tự động tìm kiếm tên cột (Ưu tiên khớp chính xác 100% trước khi khớp một phần)"""
    for kw in keywords:
        for col in df.columns:
            c_clean = ' '.join(str(col).upper().replace('_', ' ').replace('-', ' ').replace('\n', ' ').split())
            if kw.strip() == c_clean:
                return col
    for kw in keywords:
        for col in df.columns:
            c_clean = ' '.join(str(col).upper().replace('_', ' ').replace('-', ' ').replace('\n', ' ').split())
            if kw in c_clean:
                return col
    return None

def pair_guests_in_rooms(df_opera, df_excel):
    """Ghép cặp khách thông minh bên trong từng phòng để chống lệch thứ tự khi tên bị đảo chữ"""
    paired_rows = []
    all_rooms = sorted(list(set(df_opera['So_Phong'].dropna().unique()).union(set(df_excel['So_Phong'].dropna().unique()))))
    
    for room in all_rooms:
        op_sub = df_opera[df_opera['So_Phong'] == room].to_dict('records')
        ex_sub = df_excel[df_excel['So_Phong'] == room].to_dict('records')
        
        if room == "Chưa gán được" or not op_sub or not ex_sub:
            max_len = max(len(op_sub), len(ex_sub))
            for i in range(max_len):
                row_dict = {'So_Phong': room, 'Guest_Index': i}
                if i < len(op_sub) and i < len(ex_sub):
                    row_dict.update(op_sub[i])
                    row_dict.update(ex_sub[i])
                    row_dict['_merge'] = 'both'
                elif i < len(op_sub):
                    row_dict.update(op_sub[i])
                    row_dict['_merge'] = 'left_only'
                else:
                    row_dict.update(ex_sub[i])
                    row_dict['_merge'] = 'right_only'
                paired_rows.append(row_dict)
            continue
            
        used_ex = set()
        used_op = set()
        guest_idx = 0
        
        # Bước 1: Ưu tiên ghép các khách có tập hợp từ trong tên trùng khớp 100% (kể cả đảo thứ tự họ tên)
        for i_op, op_g in enumerate(op_sub):
            op_words = set(str(op_g['Ten_Khach_Opera']).split())
            for i_ex, ex_g in enumerate(ex_sub):
                if i_ex in used_ex: continue
                ex_words = set(str(ex_g['Ten_Khach_Excel']).split())
                if op_words == ex_words:
                    used_op.add(i_op)
                    used_ex.add(i_ex)
                    r = {'So_Phong': room, 'Guest_Index': guest_idx, '_merge': 'both'}
                    r.update(op_g)
                    r.update(ex_g)
                    paired_rows.append(r)
                    guest_idx += 1
                    break
                    
        # Bước 2: Ghép các khách còn lại trong phòng theo độ tương đồng chữ cái cao nhất
        rem_op = [i for i in range(len(op_sub)) if i not in used_op]
        rem_ex = [i for i in range(len(ex_sub)) if i not in used_ex]
        
        while rem_op and rem_ex:
            best_pair = (rem_op[0], rem_ex[0])
            best_score = -1
            for i_op in rem_op:
                for i_ex in rem_ex:
                    s_op = ' '.join(sorted(str(op_sub[i_op]['Ten_Khach_Opera']).split()))
                    s_ex = ' '.join(sorted(str(ex_sub[i_ex]['Ten_Khach_Excel']).split()))
                    score = SequenceMatcher(None, s_op, s_ex).ratio()
                    if score > best_score:
                        best_score = score
                        best_pair = (i_op, i_ex)
            i_op, i_ex = best_pair
            rem_op.remove(i_op)
            rem_ex.remove(i_ex)
            r = {'So_Phong': room, 'Guest_Index': guest_idx, '_merge': 'both'}
            r.update(op_sub[i_op])
            r.update(ex_sub[i_ex])
            paired_rows.append(r)
            guest_idx += 1
            
        for i_op in rem_op:
            r = {'So_Phong': room, 'Guest_Index': guest_idx, '_merge': 'left_only'}
            r.update(op_sub[i_op])
            paired_rows.append(r)
            guest_idx += 1
            
        for i_ex in rem_ex:
            r = {'So_Phong': room, 'Guest_Index': guest_idx, '_merge': 'right_only'}
            r.update(ex_sub[i_ex])
            paired_rows.append(r)
            guest_idx += 1
            
    return pd.DataFrame(paired_rows)

# ==============================================================================
# 2. TẠO FILE EXCEL MẪU
# ==============================================================================
def create_template_excel():
    data = {
        'So_Phong': ['', '0201', '0201', '0301'],
        'Ten_Khach': ['Joe Woolf', 'Huynh Duc Phong', 'Nguyen Tai Uyen', 'Khach Moi'],
        'Loai_Phong': ['King', 'Twin', 'Twin', 'King'],
        'Nhom_Phong': ['1', '2', '2', '3'],
        'Bus': ['Bus 1', 'Bus 2', 'Bus 2', 'Bus 3'],
        'Ngay_Den': ['03-10-26', '03-10-26', '03-10-26', '03-10-26'],
        'Ngay_Di': ['04-10-26', '04-10-26', '04-10-26', '05-10-26'],
        'Ghi_Chu': ['VIP', '', '', 'Late check out 18/10']
    }
    df_template = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Danh_Sach_Doan')
        ws = writer.sheets['Danh_Sach_Doan']
        for col, width in zip(['A','B','C','D','E','F','G','H'], [15, 30, 15, 15, 15, 18, 18, 25]):
            ws.column_dimensions[col].width = width
    return output.getvalue()

# ==============================================================================
# 3. GIAO DIỆN CHÍNH STREAMLIT
# ==============================================================================
st.title("Phần mềm Đối chiếu Rooming List (Ultimate Master)")

with st.info("Hỗ trợ đọc thẳng mọi định dạng file Tour (Company Trip & Hội nghị). Tự động nhận diện phòng ghép, lọc học hàm/học vị, bắt ghi chú Late C/O và kiểm tra lệch giường."):
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
    excel_bytes = excel_file.read()
    xls_ inspect = pd.ExcelFile(BytesIO(excel_bytes))
    sheet_list = xls_inspect.sheet_names
    
    selected_sheet = sheet_list[0]
    if len(sheet_list) > 1:
        selected_sheet = st.selectbox("📑 File Excel có nhiều Sheet, vui lòng chọn Sheet chứa Rooming List:", sheet_list)

    _, btn_col, _ = st.columns([1, 1, 1])
    with btn_col:
        start_check = st.button("🚀 Bắt đầu Quét Đa Luồng", type="primary", use_container_width=True)

    if start_check:
        with st.spinner("Hệ thống đang phân tích cấu trúc file, nhận diện phòng ghép và đối chiếu dữ liệu..."):
            try:
                # ---------------- BƯỚC 1: ĐỌC XML OPERA ----------------
                tree = ET.parse(xml_file)
                root = tree.getroot()
                opera_guests = []
                for guest in root.iter('G_CONFIRMATION_NO'):
                    status = guest.find('RESV_STATUS').text if guest.find('RESV_STATUS') is not None else ''
                    if 'CANCEL' in str(status).upper(): continue
                        
                    room = guest.find('ROOM').text if guest.find('ROOM') is not None else ''
                    name = guest.find('FULL_NAME').text if guest.find('FULL_NAME') is not None else ''
                    room_cat = guest.find('ROOM_CATEGORY_LABEL').text if guest.find('ROOM_CATEGORY_LABEL') is not None else ''
                    arrival = guest.find('ARRIVAL').text if guest.find('ARRIVAL') is not None else ''
                    departure = guest.find('DEPARTURE').text if guest.find('DEPARTURE') is not None else ''
                    
                    opera_guests.append({
                        'So_Phong': normalize_room(room),
                        'Ten_Khach_Opera': normalize_name(name),
                        'Loai_Phong_Opera': str(room_cat).strip(),
                        'Ngay_Den_Opera': normalize_date(arrival),
                        'Ngay_Di_Opera': normalize_date(departure)
                    })
                df_opera = pd.DataFrame(opera_guests)

                # ---------------- BƯỚC 2: ĐỌC EXCEL ĐOÀN AUTO-DETECT ----------------
                df_raw = pd.read_excel(BytesIO(excel_bytes), sheet_name=selected_sheet, header=None)
                
                is_tour_format = False
                skip_count = 0
                for idx, row in df_raw.head(15).iterrows():
                    row_str = ' '.join([str(val).upper() for val in row.values])
                    if any(k in row_str for k in ['FULL NAME', 'TÊN KHÁCH', 'HỌ VÀ TÊN', 'HỌ TÊN', 'GUEST NAME']):
                        skip_count = idx
                        is_tour_format = True
                        break
                
                if is_tour_format:
                    df_excel = pd.read_excel(BytesIO(excel_bytes), sheet_name=selected_sheet, skiprows=skip_count, dtype=str)
                else:
                    df_excel = pd.read_excel(BytesIO(excel_bytes), sheet_name=selected_sheet, dtype=str)

                # MAP CỘT THÔNG MINH
                col_name = find_column(df_excel, ['FULL NAME', 'TEN KHACH', 'TÊN KHÁCH', 'HỌ VÀ TÊN', 'HỌ TÊN', 'GUEST NAME', 'GUEST'])
                col_room = find_column(df_excel, ['ROOM NO', 'ROOM NUMBER', 'SO PHONG', 'SỐ PHÒNG'])
                col_room_type = find_column(df_excel, ['ROOM TYPE', 'LOAI PHONG', 'LOẠI PHÒNG', 'GHI CHÚ', 'GHI CHU', 'NOTE'])
                col_note = find_column(df_excel, ['GHI CHÚ', 'GHI CHU', 'NOTE', 'REMARK', 'SPECIAL REQUEST'])
                col_group = find_column(df_excel, ['SHARED ROOM', 'NHOM PHONG', 'NHÓM PHÒNG', 'SHARE ROOM', 'SHARE', 'NHÓM', 'NHOM'])
                col_bus = find_column(df_excel, ['BUS', 'CHUYEN XE', 'CHUYẾN XE', 'SỐ XE', 'SO XE'])
                col_in = find_column(df_excel, ['CHECK IN', 'CHECKIN', 'ARRIVAL', 'NGAY DEN', 'NGÀY ĐẾN'])
                col_out = find_column(df_excel, ['CHECK OUT', 'CHECKOUT', 'DEPARTURE', 'NGAY DI', 'NGÀY ĐI'])
                col_foreign = find_column(df_excel, ['FOREIGNER', 'QUOC TICH', 'QUỐC TỊCH', 'NATIONALITY'])

                if not col_name:
                    st.error("Lỗi: Không tìm thấy cột 'Tên Khách' / 'Họ tên' trong file Excel.")
                    st.stop()

                # Lọc bỏ dòng trống và dòng 'Tổng số phòng' / 'Total' TRƯỚC khi xử lý gộp phòng
                df_excel = df_excel[df_excel[col_name].apply(is_valid_guest_row)].copy()
                df_excel['Ten_Khach_Excel'] = df_excel[col_name].apply(normalize_name)

                # Kiểm tra cột Số phòng thật hay cột đếm số lượng phòng (1, 1, 1...)
                if col_room and is_real_room_column(df_excel[col_room]):
                    df_excel['So_Phong'] = df_excel[col_room].apply(normalize_room)
                else:
                    df_excel['So_Phong'] = ""

                # Nhận diện Nhóm ở ghép (Phân biệt cột Shared Room thật với cột Chức danh/Vai trò Hội nghị)
                if col_group and is_real_group_column(df_excel[col_group]):
                    df_excel['Nhom_Phong'] = df_excel[col_group].fillna("")
                else:
                    first_col = df_excel.columns[0]
                    stt_series = df_excel[first_col].replace(['nan', 'NaN', 'None', ''], pd.NA)
                    if stt_series.isna().any() and stt_series.notna().any():
                        df_excel['Nhom_Phong'] = stt_series.ffill().fillna("")
                    else:
                        df_excel['Nhom_Phong'] = ""

                if col_room_type:
                    df_excel['Loai_Phong_Excel'] = df_excel[col_room_type].replace(['nan', 'NaN', 'None', ''], pd.NA)
                    if df_excel['Nhom_Phong'].str.strip().astype(bool).any():
                        df_excel['Loai_Phong_Excel'] = df_excel.groupby('Nhom_Phong')['Loai_Phong_Excel'].ffill()
                    df_excel['Loai_Phong_Excel'] = df_excel['Loai_Phong_Excel'].fillna("")
                else:
                    df_excel['Loai_Phong_Excel'] = ""

                # Trích xuất Ghi chú vận hành (Late C/O, Extra Bed, VIP...)
                raw_notes_series = pd.Series([""] * len(df_excel), index=df_excel.index)
                if col_note:
                    raw_notes_series = raw_notes_series + " " + df_excel[col_note].fillna("")
                if col_room_type and col_room_type != col_note:
                    raw_notes_series = raw_notes_series + " " + df_excel[col_room_type].fillna("")
                df_excel['Ghi_Chu_Van_Hanh'] = raw_notes_series.apply(extract_special_notes)

                df_excel['Foreign_Flag_Excel'] = df_excel[col_foreign].fillna("") if col_foreign else ""
                df_excel['Bus_Excel'] = df_excel[col_bus].fillna("") if col_bus else ""
                df_excel['Ngay_Den_Excel'] = df_excel[col_in].apply(normalize_date) if col_in else ""
                df_excel['Ngay_Di_Excel'] = df_excel[col_out].apply(normalize_date) if col_out else ""

                # ---------------- BƯỚC 3: SMART MAPPING & MULTI-LAYER FUZZY MATCH ----------------
                name_to_rooms = df_opera.groupby('Ten_Khach_Opera')['So_Phong'].apply(list).to_dict()
                fuzzy_name_map = defaultdict(list)
                for k in name_to_rooms.keys():
                    fuzzy_name_map[frozenset(k.split())].append(k)

                def get_possible_rooms(excel_name):
                    # Tầng 1: Khớp chính xác
                    if excel_name in name_to_rooms and name_to_rooms[excel_name]:
                        return list(name_to_rooms[excel_name])
                    # Tầng 2: Khớp đảo thứ tự Họ Tên
                    excel_set = frozenset(excel_name.split())
                    if excel_set in fuzzy_name_map:
                        rooms = []
                        for real_op_name in fuzzy_name_map[excel_set]:
                            rooms.extend(name_to_rooms.get(real_op_name, []))
                        if rooms:
                            return rooms
                    # Tầng 3: Khớp mờ chống lỗi đánh máy 1-2 ký tự (Độ giống >= 88%)
                    for op_name, r_list in name_to_rooms.items():
                        if r_list and SequenceMatcher(None, excel_name, op_name).ratio() >= 0.88:
                            return list(r_list)
                    return []

                def consume_room(excel_name, room_to_consume):
                    if excel_name in name_to_rooms and room_to_consume in name_to_rooms[excel_name]:
                        name_to_rooms[excel_name].remove(room_to_consume)
                        return
                    excel_set = frozenset(excel_name.split())
                    if excel_set in fuzzy_name_map:
                        for real_op_name in fuzzy_name_map[excel_set]:
                            if room_to_consume in name_to_rooms.get(real_op_name, []):
                                name_to_rooms[real_op_name].remove(room_to_consume)
                                return
                    for op_name, r_list in name_to_rooms.items():
                        if room_to_consume in r_list and SequenceMatcher(None, excel_name, op_name).ratio() >= 0.88:
                            r_list.remove(room_to_consume)
                            return

                if df_excel['Nhom_Phong'].str.strip().astype(bool).any():
                    group_to_room = {}
                    for grp, grp_df in df_excel.groupby('Nhom_Phong'):
                        if pd.isna(grp) or str(grp).strip().upper() in ["", "NAN", "NAT", "NONE"]: continue
                        
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
                        if pd.notna(grp) and str(grp).strip().upper() not in ["", "NAN", "NAT", "NONE"]:
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
                        current_room = str(row['So_Phong']).strip().upper()
                        if current_room in ["", "NAN", "NAT", "NONE"]:
                            rooms = get_possible_rooms(row['Ten_Khach_Excel'])
                            if rooms:
                                rm = rooms[0]
                                consume_room(row['Ten_Khach_Excel'], rm)
                                return rm
                            return "Chưa gán được"
                        return current_room
                    df_excel['So_Phong'] = df_excel.apply(auto_assign_room, axis=1)

                # ---------------- BƯỚC 4: ĐỐI CHIẾU DỮ LIỆU & KIỂM TRA LỆCH ----------------
                df_opera = df_opera.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])
                df_excel = df_excel.sort_values(by=['So_Phong', 'Ten_Khach_Excel'])

                df_compare = pair_guests_in_rooms(df_opera, df_excel)
                
                # Cảnh báo khác tuyến xe Bus trong cùng 1 phòng
                bus_conflicts = []
                for room, group in df_compare.groupby('So_Phong'):
                    if room == "Chưa gán được" or pd.isna(room): continue
                    valid_buses = group['Bus_Excel'].astype(str).replace(['NAN', '', 'NONE', 'NAT'], pd.NA).dropna().unique()
                    if len(valid_buses) > 1: bus_conflicts.append(room)
                
                df_compare['Canh_Bao_Bus'] = ""
                df_compare.loc[df_compare['So_Phong'].isin(bus_conflicts), 'Canh_Bao_Bus'] = df_compare['Bus_Excel'].apply(
                    lambda x: f"⚠️ Khác Bus ({x})" if pd.notna(x) and str(x).strip().upper() not in ["", "NAN"] else ""
                )

                # Cảnh báo lệch Loại Phòng (Giường King / Twin)
                def check_room_type(row):
                    op_rt = standardize_room_type(row.get('Loai_Phong_Opera', ''))
                    ex_rt = standardize_room_type(row.get('Loai_Phong_Excel', ''))
                    if ex_rt == "UNKNOWN" or op_rt == "UNKNOWN":
                        return ""
                    if op_rt != ex_rt:
                        ex_clean = ' '.join(str(row.get('Loai_Phong_Excel','')).split())
                        return f"⚠️ Lệch Giường (Opera: {row.get('Loai_Phong_Opera','')} | Excel: {ex_clean})"
                    return ""
                
                df_compare['Canh_Bao_Loai_Phong'] = df_compare.apply(check_room_type, axis=1)
                df_compare['Ghi_Chu_Van_Hanh'] = df_compare.get('Ghi_Chu_Van_Hanh', pd.Series([""] * len(df_compare))).fillna("")

                df_compare['Quoc_Tich'] = df_compare.apply(
                    lambda row: guess_nationality(
                        row['Ten_Khach_Opera'] if pd.notna(row.get('Ten_Khach_Opera')) and str(row.get('Ten_Khach_Opera')).strip() != "" else row.get('Ten_Khach_Excel', ''),
                        row.get('Foreign_Flag_Excel', '')
                    ), axis=1
                )

                def deep_check_status(row):
                    if str(row['So_Phong']) == "Chưa gán được": return "Lỗi nặng: Khách Excel không có trên hệ thống Opera"
                    if row['_merge'] == 'left_only': return "Hủy phòng / Excel gửi sót khách này"
                    if row['_merge'] == 'right_only': return "Phát sinh thêm người ghép" if row['Guest_Index'] > 0 else "Phòng mới (Có trong Excel, chưa có trên Opera)"
                    
                    if row['_merge'] == 'both':
                        if str(row.get('Ngay_Den_Excel', '')).strip() in ['NAN', ''] and str(row.get('Ngay_Di_Excel', '')).strip() in ['NAN', '']:
                            date_mismatch = False
                        else:
                            date_mismatch = (str(row.get('Ngay_Den_Opera', '')) != str(row.get('Ngay_Den_Excel', ''))) or (str(row.get('Ngay_Di_Opera', '')) != str(row.get('Ngay_Di_Excel', '')))
                            
                        name_op = str(row.get('Ten_Khach_Opera', ''))
                        name_ex = str(row.get('Ten_Khach_Excel', ''))
                        exact_words = set(name_op.split()) == set(name_ex.split())
                        typo_similarity = SequenceMatcher(None, ' '.join(sorted(name_op.split())), ' '.join(sorted(name_ex.split()))).ratio()
                        
                        if not exact_words and typo_similarity >= 0.88:
                            if date_mismatch: return "Sai lệch KÉP (Gõ sai chính tả tên & Lệch ngày)"
                            return "⚠️ Nghi ngờ gõ sai chính tả tên (Khớp 90%)"
                            
                        name_mismatch = not exact_words
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
                tong_phong_excel = df_excel['So_Phong'].astype(str).replace(['', 'NAN', 'NAT', 'NONE', 'CHƯA GÁN ĐƯỢC'], np.nan).dropna().nunique()

                st.markdown("### 📊 THỐNG KÊ TỔNG QUAN ĐOÀN")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Tổng Khách (Opera)", f"{tong_khach_opera} pax")
                m2.metric("Tổng Phòng (Opera)", f"{tong_phong_opera} phòng")
                m3.metric("Tổng Khách (Excel)", f"{tong_khach_excel} pax", delta=f"{tong_khach_excel - tong_khach_opera} khách" if tong_khach_excel != tong_khach_opera else None)
                m4.metric("Tổng Phòng (Đã Gán)", f"{tong_phong_excel} phòng", delta=f"{tong_phong_excel - tong_phong_opera} phòng" if tong_phong_excel != tong_phong_opera else None)

                cols = [
                    'So_Phong', 'Quoc_Tich', 'Trang_Thai', 'Canh_Bao_Bus', 'Canh_Bao_Loai_Phong', 'Ghi_Chu_Van_Hanh',
                    'Ten_Khach_Opera', 'Ten_Khach_Excel', 'Loai_Phong_Opera', 'Loai_Phong_Excel',
                    'Bus_Excel', 'Nhom_Phong', 'Ngay_Den_Opera', 'Ngay_Di_Opera', 'Ngay_Den_Excel', 'Ngay_Di_Excel'
                ]
                for c in cols:
                    if c not in df_compare.columns: df_compare[c] = ""
                df_compare = df_compare[cols]

                # Sheet 1: Phòng Cần Lưu Ý (Lỗi + Lệch Bus + Lệch Giường + Ghi chú Late C/O, Extra Bed + Khách Nước Ngoài)
                df_errors = df_compare[
                    (df_compare['Trang_Thai'] != 'Khớp hoàn toàn') | 
                    (df_compare['Canh_Bao_Bus'] != "") |
                    (df_compare['Canh_Bao_Loai_Phong'] != "") |
                    (df_compare['Ghi_Chu_Van_Hanh'] != "") |
                    (df_compare['Quoc_Tich'] == '🌍 Nước Ngoài')
                ].sort_values(by=['So_Phong', 'Trang_Thai'])
                
                # Sheet 2: Danh sách Tổng
                df_full = df_compare.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])

                # Sheet 3: Bảng Cắt Key - Chia Bus
                df_valid_rooms = df_compare[df_compare['So_Phong'].astype(str).str.strip() != "Chưa gán được"].copy()
                df_valid_rooms['Bus_Excel'] = df_valid_rooms['Bus_Excel'].replace(['', 'NAN', 'NAT', 'NONE'], 'Tự túc / KHÔNG CÓ XE')
                df_valid_rooms['Ten_Hien_Thi'] = df_valid_rooms.apply(
                    lambda r: str(r['Ten_Khach_Opera']) if pd.notna(r['Ten_Khach_Opera']) and str(r['Ten_Khach_Opera']).strip() not in ['', 'NAN'] else str(r['Ten_Khach_Excel']), axis=1
                )
                
                df_key_cut = df_valid_rooms.groupby(['Bus_Excel', 'So_Phong']).agg(
                    Loai_Phong=('Loai_Phong_Opera', 'first'),
                    So_Luong_The_Can_Cat=('Ten_Hien_Thi', 'count'),
                    Danh_Sach_Khach_Trong_Phong=('Ten_Hien_Thi', lambda x: '  |  '.join([str(i) for i in x if str(i).strip() not in ["", "NAN"]])),
                    Ghi_Chu=('Ghi_Chu_Van_Hanh', lambda x: ' | '.join(set([str(i) for i in x if str(i).strip() not in ["", "NAN"]])))
                ).reset_index()
                df_key_cut = df_key_cut.rename(columns={'Bus_Excel': 'Chuyến Xe (Bus)', 'So_Phong': 'Số Phòng'})
                df_key_cut = df_key_cut.sort_values(by=['Chuyến Xe (Bus)', 'Số Phòng'])

                st.markdown("---")
                st.success("✅ Phân tích hoàn tất! Đã xử lý tự động học hàm/học vị, gộp phòng Merge Cell, bắt ghi chú vận hành và kiểm tra chéo lịch trình.")
                
                def highlight_special(s):
                    if str(s.get('Trang_Thai', '')) != 'Khớp hoàn toàn': return ['background-color: #ffb3b3'] * len(s)
                    if '⚠️' in str(s.get('Canh_Bao_Loai_Phong', '')): return ['background-color: #ffff99'] * len(s)
                    if '⚠️' in str(s.get('Canh_Bao_Bus', '')): return ['background-color: #ffe5b4'] * len(s)
                    if str(s.get('Ghi_Chu_Van_Hanh', '')) != '': return ['background-color: #d9f2d9'] * len(s)
                    if s.get('Quoc_Tich') == '🌍 Nước Ngoài': return ['background-color: #ffcccc'] * len(s)
                    return [''] * len(s)
                
                tab1, tab2, tab3 = st.tabs(["⚠️ BẢNG CẦN LƯU Ý", "📋 DANH SÁCH TỔNG ĐOÀN", "🔑 BẢNG LÀM THẺ KHÓA (CHIA BUS)"])
                with tab1:
                    if df_errors.empty: st.info("🎉 Tuyệt vời! Đoàn khớp 100%, đi chung Bus, đúng loại giường và hoàn toàn thuần Việt.")
                    else: st.dataframe(df_errors.style.apply(highlight_special, axis=1), use_container_width=True)
                with tab2:
                    st.dataframe(df_full.style.apply(highlight_special, axis=1), use_container_width=True)
                with tab3:
                    st.info("💡 Lễ tân dùng bảng này để cắt thẻ phòng theo xấp tương ứng với từng chuyến Bus hoặc in gửi cho Tour Leader / Bellman.")
                    st.dataframe(df_key_cut, use_container_width=True)

                # ---------------- BƯỚC 6: XUẤT EXCEL ----------------
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_errors.to_excel(writer, index=False, sheet_name='Phong_Can_Luu_Y')
                    df_full.to_excel(writer, index=False, sheet_name='Danh_Sach_Tong')
                    df_key_cut.to_excel(writer, index=False, sheet_name='Cat_Key_Chia_Bus')
                    
                    for sheet_name in ['Phong_Can_Luu_Y', 'Danh_Sach_Tong', 'Cat_Key_Chia_Bus']:
                        worksheet = writer.sheets[sheet_name]
                        for column in worksheet.columns:
                            max_length = max((len(str(cell.value)) for cell in column), default=0)
                            worksheet.column_dimensions[column[0].column_letter].width = min(max_length + 3, 55)

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