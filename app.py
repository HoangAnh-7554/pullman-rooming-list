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
# 1. CẤU HÌNH NHẬN DIỆN ÂM TIẾT TIẾNG VIỆT, PHÒNG EXBED SẴN & BỘ LỌC DỮ LIỆU
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

UNASSIGNED_OPERA_LABEL = "⏳ Chưa gán phòng (Opera)"

# Danh sách 6 phòng có sẵn Extra Bed cố định tại khách sạn (518, 618, 718, 720, 818, 820)
BUILT_IN_EXBED_ROOMS = {'0518', '0618', '0718', '0720', '0818', '0820'}

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

def normalize_room(room, is_opera=False):
    """Khắc phục lỗi mất số 0 ở đầu của Excel (Ví dụ: 518 -> 0518) và đánh nhãn phòng trống trên Opera"""
    if pd.isna(room): 
        return UNASSIGNED_OPERA_LABEL if is_opera else ""
    r = str(room).replace('.0', '').strip().upper()
    if r in ["", "NAN", "NAT", "NONE"]: 
        return UNASSIGNED_OPERA_LABEL if is_opera else ""
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

def has_exbed_keyword(text):
    """Kiểm tra chuỗi có chứa từ khóa yêu cầu Extra Bed / Giường phụ / Triple không"""
    if pd.isna(text): return False
    s = str(text).upper()
    return any(k in s for k in ['EXTRA BED', 'EXBED', 'EXB', 'GIƯỜNG PHỤ', 'GIUONG PHU', 'TRIPLE', '3 NGƯỜI', '3 NGUOI'])

def is_room_3_keys(room_df):
    """
    Xác định phòng có thuộc diện cắt 3 key hay không:
    1. Có ghi chú Extra Bed / Giường phụ trong Excel (Room Type / Note) hoặc trên Opera (Products / Special Requests)
    2. Hoặc tổng số người lớn trên Opera (Adults_Opera) >= 3
    3. Hoặc phòng nằm trong 6 phòng có Exbed sẵn (0518, 0618, 0718, 0720, 0818, 0820) VÀ có từ 3 khách trở lên trong phòng
    """
    rm_code = str(room_df['So_Phong'].iloc[0]).strip()
    total_guests_in_room = len(room_df)
    total_op_adults = room_df['Adults_Opera'].sum() if 'Adults_Opera' in room_df.columns else 0
    
    exb_requested = False
    for _, r in room_df.iterrows():
        if (has_exbed_keyword(r.get('Loai_Phong_Excel', '')) or 
            has_exbed_keyword(r.get('Raw_Note_Excel', '')) or 
            has_exbed_keyword(r.get('Req_Opera', '')) or 
            has_exbed_keyword(r.get('Products_Opera', ''))):
            exb_requested = True
            break
            
    is_builtin_exb = (rm_code in BUILT_IN_EXBED_ROOMS)
    if exb_requested or total_op_adults >= 3 or (is_builtin_exb and total_guests_in_room >= 3):
        return True, is_builtin_exb, exb_requested
    return False, is_builtin_exb, exb_requested

def extract_special_notes(text, opera_vip="", opera_req="", opera_prod="", rm_code="", is_3k=False, is_builtin_exb=False, exb_requested=False):
    """Tự động phát hiện các yêu cầu đặc biệt: Late C/O, Early C/I, Extra Bed (Sẵn có vs Kê thêm), Baby Cot, VIP, Đổi phòng, Nối booking"""
    notes = []
    if pd.notna(opera_vip) and str(opera_vip).strip().upper() not in ["", "NAN", "NONE"]:
        notes.append(f"⭐ Opera VIP {str(opera_vip).strip()}")
    if pd.notna(opera_req) and str(opera_req).strip().upper() not in ["", "NAN", "NONE"]:
        notes.append(f"📝 Req: {str(opera_req).strip()}")
    if pd.notna(opera_prod) and str(opera_prod).strip().upper() not in ["", "NAN", "NONE"]:
        notes.append(f"📦 Prod: {str(opera_prod).strip()}")
        
    # Phân biệt rõ Phòng có Exbed sẵn (518, 618, 718, 720, 818, 820) và Phòng cần kê thêm Extra Bed
    if is_3k:
        if is_builtin_exb:
            notes.append(f"🛋️ Phòng {rm_code} có Exbed sẵn (Báo HKP trải ga Exbed + Cắt 3 key)")
        else:
            notes.append("🛏️ Có Extra Bed / 3 người (Báo HKP kê Exbed + Cắt 3 key)")
    elif is_builtin_exb:
        notes.append(f"🛋️ Phòng {rm_code} có Exbed sẵn (Đang ở ≤2 khách -> Cắt 2 key)")

    if pd.notna(text) and str(text).strip().upper() not in ["", "NAN", "NAT", "NONE"]:
        s_raw = ' '.join(str(text).replace('\n', ' - ').split())
        s_up = s_raw.upper()
        if any(k in s_up for k in ['LATE', 'CHECKOUT LATE', 'LATE CHECK OUT', 'LATE C/O', 'TRẢ MUỘN', 'TRA MUON']):
            notes.append(f"⏰ Late C/O ({s_raw})")
        elif any(k in s_up for k in ['EARLY', 'SỚM', 'EARLY C/I']):
            notes.append(f"🌅 Early C/I ({s_raw})")
        if any(k in s_up for k in ['GHÉP GIƯỜNG', 'GHEP GIUONG', 'HOLLYWOOD']):
            notes.append(f"🛏️ Ghép giường ({s_raw})")
        if any(k in s_up for k in ['BABY COT', 'CRIB', 'NÔI']):
            notes.append(f"👶 Baby Cot ({s_raw})")
        if any(k in s_up for k in ['ĐỔI PHÒNG', 'DOI PHONG', 'ROOM MOVE', 'CHUYỂN PHÒNG', 'CHUYEN PHONG', 'MOVE FROM', 'MOVE TO']):
            notes.append(f"🔄 Đổi phòng ({s_raw})")
        if any(k in s_up for k in ['NỐI PHÒNG', 'NOI PHONG', 'NỐI BOOKING', 'NOI BOOKING', 'CONNECT', 'EXTEND', 'STAY OVER', 'STAYOVER', 'IN-HOUSE', 'IN HOUSE', 'ĐANG Ở']):
            notes.append(f"🔗 Nối booking ({s_raw})")
        if 'VIP' in s_up:
            notes.append(f"⭐ VIP ({s_raw})")
    return " | ".join(dict.fromkeys(notes))

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
    """Ghép cặp khách thông minh bên trong từng phòng (Ưu tiên khớp cả Tên + Ngày In/Out nếu khách có nhiều chặng nối booking)"""
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
        
        for i_op, op_g in enumerate(op_sub):
            op_words = set(str(op_g['Ten_Khach_Opera']).split())
            op_dates = (str(op_g.get('Ngay_Den_Opera', '')), str(op_g.get('Ngay_Di_Opera', '')))
            for i_ex, ex_g in enumerate(ex_sub):
                if i_ex in used_ex: continue
                ex_words = set(str(ex_g['Ten_Khach_Excel']).split())
                ex_dates = (str(ex_g.get('Ngay_Den_Excel', '')), str(ex_g.get('Ngay_Di_Excel', '')))
                if op_words == ex_words and op_dates == ex_dates:
                    used_op.add(i_op)
                    used_ex.add(i_ex)
                    r = {'So_Phong': room, 'Guest_Index': guest_idx, '_merge': 'both'}
                    r.update(op_g)
                    r.update(ex_g)
                    paired_rows.append(r)
                    guest_idx += 1
                    break

        for i_op, op_g in enumerate(op_sub):
            if i_op in used_op: continue
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

def calc_key_and_booklet_rule(sub_df, full_room_df, total_buses_for_room):
    """
    Quy chuẩn cắt key khách sạn:
    - Mặc định 1 phòng = 2 key (Phòng KING gộp 1 bìa 2 key; Phòng TWIN tách riêng mỗi người 1 key / 2 bìa).
    - Trường hợp Phòng có Extra Bed HOẶC thuộc 6 phòng có Exbed sẵn (0518, 0618, 0718, 0720, 0818, 0820) ở >= 3 khách:
      Cắt 3 key / phòng!
    - Nếu phòng bị tách ra nhiều xe Bus khác nhau (total_buses_for_room > 1):
      + Phòng 2 key chuẩn: Mỗi nhánh xe cắt 1 key (1 bìa).
      + Phòng 3 key (Exbed): Chia số key theo số khách trên từng xe (nhánh 1 khách -> 1 key, nhánh 2 khách -> 2 key) để tổng = 3 key.
    """
    n_guests_in_bus = len(sub_df)
    rm_code = str(full_room_df['So_Phong'].iloc[0]).strip()
    op_rt = str(full_room_df['Loai_Phong_Opera'].iloc[0]) if 'Loai_Phong_Opera' in full_room_df.columns else ""
    ex_rt = str(full_room_df['Loai_Phong_Excel'].iloc[0]) if 'Loai_Phong_Excel' in full_room_df.columns else ""
    rt_std = standardize_room_type(op_rt) if standardize_room_type(op_rt) != "UNKNOWN" else standardize_room_type(ex_rt)
    
    is_3k, is_bi, ex_req = is_room_3_keys(full_room_df)
    
    if total_buses_for_room > 1:
        if is_3k:
            k_count = min(n_guests_in_bus, 2) if len(full_room_df) >= 3 else 1
            b_count = k_count if rt_std == "TWIN" else 1
            return k_count, b_count, f"🚌 Khác Bus (Phòng Exbed 3 key): Cắt {k_count} key ({b_count} bìa) cho {n_guests_in_bus} khách xe này"
        else:
            return 1, 1, "🚌 Khác Bus: Cắt riêng 1 key (1 bìa) cho nhánh xe này"
            
    if is_3k:
        exb_tag = f"Phòng {rm_code} có Exbed sẵn" if is_bi else "Kê Extra Bed"
        if rt_std == "TWIN":
            return 3, 3, f"🛏️ TWIN + EXBED ({exb_tag}): 3 key (Tách riêng mỗi người 1 key / 3 bìa)"
        else:
            return 3, 1, f"🛋️ KING + EXBED ({exb_tag}): 3 key / phòng (Gộp 1 bìa)"
            
    if rt_std == "TWIN":
        if n_guests_in_bus >= 2:
            return 2, 2, "🛏️ TWIN: 2 key (Tách riêng mỗi người 1 key / 2 bìa)"
        else:
            return 2, 1, "🛏️ TWIN (1 khách): 2 key (1 bìa)"
    else:
        if is_bi:
            return 2, 1, f"👑 KING (Phòng {rm_code} có Exbed sẵn nhưng ở {n_guests_in_bus} khách): 2 key (1 bìa)"
        return 2, 1, "👑 KING: 2 key / phòng (Gộp 1 bìa)"

# ==============================================================================
# 2. TẠO FILE EXCEL MẪU
# ==============================================================================
def create_template_excel():
    data = {
        'So_Phong': ['', '0201', '0201', '0618'],
        'Ten_Khach': ['Joe Woolf', 'Huynh Duc Phong', 'Nguyen Tai Uyen', 'Khach Gia Dinh'],
        'Loai_Phong': ['King', 'Twin', 'Twin', 'King + Extra bed'],
        'Nhom_Phong': ['1', '2', '2', '3'],
        'Bus': ['Bus 1', 'Bus 2', 'Bus 2', 'Bus 3'],
        'Ngay_Den': ['03-10-26', '03-10-26', '03-10-26', '03-10-26'],
        'Ngay_Di': ['04-10-26', '04-10-26', '04-10-26', '05-10-26'],
        'Ghi_Chu': ['VIP', '', '', 'Extra bed']
    }
    df_template = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Danh_Sach_Doan')
        ws = writer.sheets['Danh_Sach_Doan']
        for col, width in zip(['A','B','C','D','E','F','G','H'], [15, 30, 18, 15, 15, 18, 18, 25]):
            ws.column_dimensions[col].width = width
    return output.getvalue()

# ==============================================================================
# 3. GIAO DIỆN CHÍNH STREAMLIT
# ==============================================================================
st.title("Phần mềm Đối chiếu Rooming List (Ultimate Master)")

with st.info("Hỗ trợ đọc thẳng mọi định dạng file Tour (Company Trip & Hội nghị). Chuẩn hóa cắt key: KING = 2 key (1 bìa) | TWIN = 2 key (2 bìa riêng) | Phòng Extra Bed & 6 Phòng Exbed sẵn (0518, 0618, 0718, 0720, 0818, 0820) ở ≥3 khách = 3 key."):
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
    xml_bytes = xml_file.read()
    xls_inspect = pd.ExcelFile(BytesIO(excel_bytes))
    sheet_list = xls_inspect.sheet_names
    
    selected_sheet = sheet_list[0]
    if len(sheet_list) > 1:
        selected_sheet = st.selectbox("📑 File Excel có nhiều Sheet, vui lòng chọn Sheet chứa Rooming List:", sheet_list)

    current_file_sig = f"{xml_file.name}_{len(xml_bytes)}_{excel_file.name}_{len(excel_bytes)}_{selected_sheet}"

    _, btn_col, _ = st.columns([1, 1, 1])
    with btn_col:
        if st.button("🚀 Bắt đầu Quét Đa Luồng", type="primary", use_container_width=True):
            st.session_state['active_sig'] = current_file_sig

    if st.session_state.get('active_sig') == current_file_sig:
        with st.spinner("Hệ thống đang phân tích cấu trúc file, nhận diện phòng Exbed (3 key), nối booking và đối chiếu dữ liệu..."):
            try:
                # ---------------- BƯỚC 1: ĐỌC XML OPERA ----------------
                tree = ET.parse(BytesIO(xml_bytes))
                root = tree.getroot()
                opera_guests = []
                cf_rooms_total = 0
                for guest in root.iter('G_CONFIRMATION_NO'):
                    status = guest.find('RESV_STATUS').text if guest.find('RESV_STATUS') is not None else ''
                    if 'CANCEL' in str(status).upper(): continue
                        
                    room = guest.find('ROOM').text if guest.find('ROOM') is not None else ''
                    name = guest.find('FULL_NAME').text if guest.find('FULL_NAME') is not None else ''
                    room_cat = guest.find('ROOM_CATEGORY_LABEL').text if guest.find('ROOM_CATEGORY_LABEL') is not None else ''
                    arrival = guest.find('ARRIVAL').text if guest.find('ARRIVAL') is not None else ''
                    departure = guest.find('DEPARTURE').text if guest.find('DEPARTURE') is not None else ''
                    vip_code = guest.find('VIP').text if guest.find('VIP') is not None else ''
                    spec_req = guest.find('SPECIAL_REQUESTS').text if guest.find('SPECIAL_REQUESTS') is not None else ''
                    products = guest.find('PRODUCTS').text if guest.find('PRODUCTS') is not None else ''
                    adults = guest.find('ADULTS').text if guest.find('ADULTS') is not None else '0'
                    cf_rm = guest.find('CF_NO_OF_ROOMS').text if guest.find('CF_NO_OF_ROOMS') is not None else '0'
                    if str(cf_rm).strip().isdigit():
                        cf_rooms_total += int(str(cf_rm).strip())
                    
                    opera_guests.append({
                        'So_Phong': normalize_room(room, is_opera=True),
                        'Ten_Khach_Opera': normalize_name(name),
                        'Loai_Phong_Opera': str(room_cat).strip(),
                        'Ngay_Den_Opera': normalize_date(arrival),
                        'Ngay_Di_Opera': normalize_date(departure),
                        'VIP_Opera': str(vip_code).strip() if pd.notna(vip_code) else '',
                        'Req_Opera': str(spec_req).strip() if pd.notna(spec_req) else '',
                        'Products_Opera': str(products).strip() if pd.notna(products) else '',
                        'Adults_Opera': int(str(adults).strip()) if str(adults).strip().isdigit() else 0
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

                df_excel = df_excel[df_excel[col_name].apply(is_valid_guest_row)].copy()
                df_excel['Ten_Khach_Excel'] = df_excel[col_name].apply(normalize_name)

                if col_room and is_real_room_column(df_excel[col_room]):
                    df_excel['So_Phong'] = df_excel[col_room].apply(lambda x: normalize_room(x, is_opera=False))
                else:
                    df_excel['So_Phong'] = ""

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

                raw_notes_series = pd.Series([""] * len(df_excel), index=df_excel.index)
                if col_note:
                    raw_notes_series = raw_notes_series + " " + df_excel[col_note].fillna("")
                if col_room_type and col_room_type != col_note:
                    raw_notes_series = raw_notes_series + " " + df_excel[col_room_type].fillna("")
                df_excel['Raw_Note_Excel'] = raw_notes_series

                df_excel['Foreign_Flag_Excel'] = df_excel[col_foreign].fillna("") if col_foreign else ""
                df_excel['Bus_Excel'] = df_excel[col_bus].fillna("") if col_bus else ""
                df_excel['Ngay_Den_Excel'] = df_excel[col_in].apply(normalize_date) if col_in else ""
                df_excel['Ngay_Di_Excel'] = df_excel[col_out].apply(normalize_date) if col_out else ""

                # ---------------- BƯỚC 3: SMART MAPPING (ƯU TIÊN KHỚP NGÀY CHO KHÁCH B2B / ĐỔI PHÒNG) ----------------
                name_to_records = defaultdict(list)
                for _, op_row in df_opera.iterrows():
                    name_to_records[op_row['Ten_Khach_Opera']].append(
                        (op_row['So_Phong'], op_row['Ngay_Den_Opera'], op_row['Ngay_Di_Opera'])
                    )
                    
                fuzzy_name_map = defaultdict(list)
                for k in name_to_records.keys():
                    fuzzy_name_map[frozenset(k.split())].append(k)

                def get_matched_opera_keys(excel_name):
                    if excel_name in name_to_records and name_to_records[excel_name]:
                        return [excel_name]
                    excel_set = frozenset(excel_name.split())
                    if excel_set in fuzzy_name_map:
                        valid_keys = [k for k in fuzzy_name_map[excel_set] if name_to_records[k]]
                        if valid_keys: return valid_keys
                    for op_name, rec_list in name_to_records.items():
                        if rec_list and SequenceMatcher(None, excel_name, op_name).ratio() >= 0.88:
                            return [op_name]
                    return []

                def get_possible_rooms(excel_name, ex_arr="", ex_dep=""):
                    keys = get_matched_opera_keys(excel_name)
                    if not keys: return []
                    all_recs = []
                    for k in keys:
                        all_recs.extend(name_to_records[k])
                    if ex_arr or ex_dep:
                        exact_date_rooms = [rm for rm, a, d in all_recs if (not ex_arr or a == ex_arr) and (not ex_dep or d == ex_dep)]
                        if exact_date_rooms:
                            return exact_date_rooms
                        partial_date_rooms = [rm for rm, a, d in all_recs if (ex_arr and a == ex_arr) or (ex_dep and d == ex_dep)]
                        if partial_date_rooms:
                            return partial_date_rooms
                    return [rm for rm, _, _ in all_recs]

                def consume_room(excel_name, room_to_consume, ex_arr="", ex_dep=""):
                    keys = get_matched_opera_keys(excel_name)
                    for k in keys:
                        recs = name_to_records[k]
                        for idx_r, (rm, a, d) in enumerate(recs):
                            if rm == room_to_consume and ((not ex_arr or a == ex_arr) and (not ex_dep or d == ex_dep)):
                                recs.pop(idx_r)
                                return
                        for idx_r, (rm, a, d) in enumerate(recs):
                            if rm == room_to_consume:
                                recs.pop(idx_r)
                                return

                if df_excel['Nhom_Phong'].str.strip().astype(bool).any():
                    group_to_room = {}
                    for grp, grp_df in df_excel.groupby('Nhom_Phong'):
                        if pd.isna(grp) or str(grp).strip().upper() in ["", "NAN", "NAT", "NONE"]: continue
                        
                        possible_rooms = []
                        for _, r_ex in grp_df.iterrows():
                            possible_rooms.extend(get_possible_rooms(r_ex['Ten_Khach_Excel'], r_ex['Ngay_Den_Excel'], r_ex['Ngay_Di_Excel']))
                                
                        if possible_rooms:
                            best_room = pd.Series(possible_rooms).mode()[0]
                            group_to_room[grp] = best_room
                            for _, r_ex in grp_df.iterrows():
                                consume_room(r_ex['Ten_Khach_Excel'], best_room, r_ex['Ngay_Den_Excel'], r_ex['Ngay_Di_Excel'])
                                    
                    def assign_by_group(row):
                        grp = row['Nhom_Phong']
                        if pd.notna(grp) and str(grp).strip().upper() not in ["", "NAN", "NAT", "NONE"]:
                            return group_to_room.get(grp, "Chưa gán được")
                        rooms = get_possible_rooms(row['Ten_Khach_Excel'], row['Ngay_Den_Excel'], row['Ngay_Di_Excel'])
                        if rooms:
                            rm = rooms[0]
                            consume_room(row['Ten_Khach_Excel'], rm, row['Ngay_Den_Excel'], row['Ngay_Di_Excel'])
                            return rm
                        return "Chưa gán được"
                    
                    df_excel['So_Phong'] = np.where(df_excel['So_Phong'].replace(['', 'NAN', 'NONE'], pd.NA).isna(), df_excel.apply(assign_by_group, axis=1), df_excel['So_Phong'])
                else:
                    def auto_assign_room(row):
                        current_room = str(row['So_Phong']).strip().upper()
                        if current_room in ["", "NAN", "NAT", "NONE"]:
                            rooms = get_possible_rooms(row['Ten_Khach_Excel'], row['Ngay_Den_Excel'], row['Ngay_Di_Excel'])
                            if rooms:
                                rm = rooms[0]
                                consume_room(row['Ten_Khach_Excel'], rm, row['Ngay_Den_Excel'], row['Ngay_Di_Excel'])
                                return rm
                            return "Chưa gán được"
                        return current_room
                    df_excel['So_Phong'] = df_excel.apply(auto_assign_room, axis=1)

                # ---------------- BƯỚC 4: ĐỐI CHIẾU, PHÁT HIỆN EXBED / NỐI BOOKING / ĐỔI PHÒNG ----------------
                df_opera = df_opera.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])
                df_excel = df_excel.sort_values(by=['So_Phong', 'Ten_Khach_Excel'])

                df_compare = pair_guests_in_rooms(df_opera, df_excel)

                # 4.0. Phân tích từng phòng xem có thuộc diện 3 Key (Exbed sẵn 518, 618, 718, 720, 818, 820 hoặc Kê Extra Bed)
                room_3k_status = {}
                for rm, r_grp in df_compare.groupby('So_Phong'):
                    if rm in ["Chưa gán được", UNASSIGNED_OPERA_LABEL] or pd.isna(rm):
                        room_3k_status[rm] = (False, False, False)
                    else:
                        room_3k_status[rm] = is_room_3_keys(r_grp)

                # 4.1. Phát hiện khách có nhiều booking (Nối Booking cùng phòng HOẶC Đổi phòng Room Move)
                df_compare['Canonical_Name_Key'] = df_compare.apply(
                    lambda r: ' '.join(sorted(str(r.get('Ten_Khach_Opera') if pd.notna(r.get('Ten_Khach_Opera')) and str(r.get('Ten_Khach_Opera')).strip() not in ['', 'NAN'] else r.get('Ten_Khach_Excel', '')).split())),
                    axis=1
                )
                
                b2b_guest_map = {}
                for cname, cgrp in df_compare.groupby('Canonical_Name_Key'):
                    if not cname or cname in ['', 'NAN']: continue
                    if len(cgrp) > 1:
                        stays = []
                        for _, crow in cgrp.iterrows():
                            rm = str(crow.get('So_Phong', '')).strip()
                            a = str(crow.get('Ngay_Den_Opera', '')).strip() if pd.notna(crow.get('Ngay_Den_Opera')) and str(crow.get('Ngay_Den_Opera')).strip() not in ['', 'NAN'] else str(crow.get('Ngay_Den_Excel', '')).strip()
                            d = str(crow.get('Ngay_Di_Opera', '')).strip() if pd.notna(crow.get('Ngay_Di_Opera')) and str(crow.get('Ngay_Di_Opera')).strip() not in ['', 'NAN'] else str(crow.get('Ngay_Di_Excel', '')).strip()
                            stays.append((rm, a, d))
                        unique_stays = list(dict.fromkeys(stays))
                        unique_dates = set((a, d) for _, a, d in unique_stays if a or d)
                        if len(unique_dates) > 1:
                            unique_rms = list(dict.fromkeys([rm for rm, _, _ in unique_stays if rm not in ['', 'NAN', 'Chưa gán được']]))
                            if len(unique_rms) == 1:
                                itinerary = " ➡️ ".join([f"({a}->{d})" for _, a, d in unique_stays])
                                b2b_guest_map[cname] = f"🔗 Nối Booking cùng P.{unique_rms[0]}: {itinerary} (Gia hạn key xuyên suốt)"
                            else:
                                itinerary = " ➡️ ".join([f"P.{rm} ({a}->{d})" for rm, a, d in unique_stays])
                                b2b_guest_map[cname] = f"🔄 Đổi Phòng (Room Move): {itinerary}"

                # 4.2. Xác định Lịch trình chuẩn của Đoàn
                valid_arrs = [str(x).strip() for x in df_compare['Ngay_Den_Opera'].dropna() if str(x).strip() not in ['', 'NAN']]
                if not valid_arrs:
                    valid_arrs = [str(x).strip() for x in df_compare['Ngay_Den_Excel'].dropna() if str(x).strip() not in ['', 'NAN']]
                valid_deps = [str(x).strip() for x in df_compare['Ngay_Di_Opera'].dropna() if str(x).strip() not in ['', 'NAN']]
                if not valid_deps:
                    valid_deps = [str(x).strip() for x in df_compare['Ngay_Di_Excel'].dropna() if str(x).strip() not in ['', 'NAN']]
                
                main_group_arr = pd.Series(valid_arrs).mode()[0] if valid_arrs else ""
                main_group_dep = pd.Series(valid_deps).mode()[0] if valid_deps else ""

                # 4.3. Kiểm tra khách ở chung phòng nhưng LỆCH NGÀY C/I hoặc C/O
                sharer_split_date_map = {}
                for room, group in df_compare.groupby('So_Phong'):
                    if room in ["Chưa gán được", UNASSIGNED_OPERA_LABEL] or pd.isna(room) or len(group) <= 1:
                        continue
                    g_schedules = []
                    for _, r in group.iterrows():
                        gn = str(r.get('Ten_Khach_Opera', '')) if pd.notna(r.get('Ten_Khach_Opera')) and str(r.get('Ten_Khach_Opera')).strip() not in ['', 'NAN'] else str(r.get('Ten_Khach_Excel', ''))
                        ga = str(r.get('Ngay_Den_Opera', '')) if pd.notna(r.get('Ngay_Den_Opera')) and str(r.get('Ngay_Den_Opera')).strip() not in ['', 'NAN'] else str(r.get('Ngay_Den_Excel', ''))
                        gd = str(r.get('Ngay_Di_Opera', '')) if pd.notna(r.get('Ngay_Di_Opera')) and str(r.get('Ngay_Di_Opera')).strip() not in ['', 'NAN'] else str(r.get('Ngay_Di_Excel', ''))
                        g_schedules.append((gn, ga, gd))
                    u_arrs = set(a for _, a, _ in g_schedules if a and a != 'NAN')
                    u_deps = set(d for _, _, d in g_schedules if d and d != 'NAN')
                    if len(u_arrs) > 1 or len(u_deps) > 1:
                        detail = " | ".join([f"{gn} ({ga}->{gd})" for gn, ga, gd in g_schedules])
                        sharer_split_date_map[room] = f"⚠️ Chung phòng LỆCH NGÀY: {detail}"

                # 4.4. Cảnh báo khác tuyến xe Bus trong cùng 1 phòng
                bus_conflicts = []
                for room, group in df_compare.groupby('So_Phong'):
                    if room in ["Chưa gán được", UNASSIGNED_OPERA_LABEL] or pd.isna(room): continue
                    valid_buses = group['Bus_Excel'].astype(str).replace(['NAN', '', 'NONE', 'NAT'], pd.NA).dropna().unique()
                    if len(valid_buses) > 1: bus_conflicts.append(room)
                
                df_compare['Canh_Bao_Bus'] = ""
                df_compare.loc[df_compare['So_Phong'].isin(bus_conflicts), 'Canh_Bao_Bus'] = df_compare['Bus_Excel'].apply(
                    lambda x: f"⚠️ Khác Bus ({x})" if pd.notna(x) and str(x).strip().upper() not in ["", "NAN"] else ""
                )

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

                def build_operational_and_key_notes(row):
                    notes = []
                    rm = str(row.get('So_Phong', '')).strip()
                    is_3k, is_bi, ex_req = room_3k_status.get(rm, (False, False, False))
                    
                    base_note = extract_special_notes(
                        row.get('Raw_Note_Excel', ''),
                        row.get('VIP_Opera', ''),
                        row.get('Req_Opera', ''),
                        row.get('Products_Opera', ''),
                        rm_code=rm,
                        is_3k=is_3k,
                        is_builtin_exb=is_bi,
                        exb_requested=ex_req
                    )
                    if base_note: notes.append(base_note)
                    
                    cname = row.get('Canonical_Name_Key', '')
                    if cname in b2b_guest_map:
                        notes.append(b2b_guest_map[cname])
                        
                    if rm in sharer_split_date_map:
                        notes.append(sharer_split_date_map[rm])
                    else:
                        eff_arr = str(row.get('Ngay_Den_Opera', '')).strip() if pd.notna(row.get('Ngay_Den_Opera')) and str(row.get('Ngay_Den_Opera')).strip() not in ['', 'NAN'] else str(row.get('Ngay_Den_Excel', '')).strip()
                        eff_dep = str(row.get('Ngay_Di_Opera', '')).strip() if pd.notna(row.get('Ngay_Di_Opera')) and str(row.get('Ngay_Di_Opera')).strip() not in ['', 'NAN'] else str(row.get('Ngay_Di_Excel', '')).strip()
                        if main_group_dep and eff_dep and eff_dep != main_group_dep:
                            notes.append(f"⚡ C/O sớm/khác đoàn (Out: {eff_dep})")
                        if main_group_arr and eff_arr and eff_arr != main_group_arr:
                            notes.append(f"⏳ C/I khác ngày đoàn (In: {eff_arr})")
                            
                    return " | ".join(dict.fromkeys(notes))

                df_compare['Ghi_Chu_Van_Hanh'] = df_compare.apply(build_operational_and_key_notes, axis=1)

                df_compare['Quoc_Tich'] = df_compare.apply(
                    lambda row: guess_nationality(
                        row['Ten_Khach_Opera'] if pd.notna(row.get('Ten_Khach_Opera')) and str(row.get('Ten_Khach_Opera')).strip() != "" else row.get('Ten_Khach_Excel', ''),
                        row.get('Foreign_Flag_Excel', '')
                    ), axis=1
                )

                def deep_check_status(row):
                    rm_str = str(row['So_Phong']).strip()
                    cname = row.get('Canonical_Name_Key', '')
                    is_unassigned_opera = (rm_str == UNASSIGNED_OPERA_LABEL)
                    if rm_str == "Chưa gán được": return "Lỗi nặng: Khách Excel không có trên hệ thống Opera"
                    if row['_merge'] == 'left_only':
                        if cname in b2b_guest_map:
                            return "🔄 Chặng Nối Booking / Đổi Phòng trên Opera"
                        return "Hủy phòng / Excel gửi sót khách này"
                    if row['_merge'] == 'right_only':
                        if cname in b2b_guest_map:
                            return "🔄 Chặng Nối Booking / Đổi Phòng trong Excel"
                        return "Phát sinh thêm người ghép" if row['Guest_Index'] > 0 else "Phòng mới (Có trong Excel, chưa có trên Opera)"
                    
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
                        if is_unassigned_opera:
                            return "📌 Chưa gán số phòng trên Opera (Đã khớp tên & ngày)"
                        return "Khớp hoàn toàn"
                    return "Cần kiểm tra thủ công"

                df_compare['Trang_Thai'] = df_compare.apply(deep_check_status, axis=1)
                
                # ---------------- BƯỚC 5: DASHBOARD & XUẤT 3 SHEETS ----------------
                tong_khach_opera = len(df_opera)
                assigned_rooms_opera = df_opera[~df_opera['So_Phong'].isin(['', 'NAN', 'NAT', 'NONE', 'CHƯA GÁN ĐƯỢC', UNASSIGNED_OPERA_LABEL])]['So_Phong'].nunique()
                tong_phong_opera = cf_rooms_total if cf_rooms_total > 0 else assigned_rooms_opera
                tong_khach_excel = len(df_excel)
                tong_phong_da_gan = df_excel[~df_excel['So_Phong'].isin(['', 'NAN', 'NAT', 'NONE', 'CHƯA GÁN ĐƯỢC', UNASSIGNED_OPERA_LABEL])]['So_Phong'].nunique()

                st.markdown(f"### 📊 THỐNG KÊ TỔNG QUAN ĐOÀN *(Lịch chuẩn đoàn: {main_group_arr} ➡️ {main_group_dep})*")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Tổng Khách (Opera)", f"{tong_khach_opera} pax")
                m2.metric("Tổng Phòng (Opera)", f"{tong_phong_opera} phòng")
                m3.metric("Tổng Khách (Excel)", f"{tong_khach_excel} pax", delta=f"{tong_khach_excel - tong_khach_opera} khách" if tong_khach_excel != tong_khach_opera else None)
                m4.metric("Phòng Đã Xếp Số", f"{tong_phong_da_gan}/{tong_phong_opera} phòng", delta=f"-{tong_phong_opera - tong_phong_da_gan} phòng chưa gán số" if tong_phong_opera > tong_phong_da_gan else None)

                # Sheet 3: Bảng Cắt Key - Chia Bus (Chuẩn khách sạn: KING = 2 key (1 bìa); TWIN = 2 key (2 bìa); Phòng Exbed / 518, 618, 718, 720, 818, 820 ở >=3 khách = 3 key)
                df_valid_rooms = df_compare[df_compare['So_Phong'].astype(str).str.strip() != "Chưa gán được"].copy()
                df_valid_rooms['Bus_Excel'] = df_valid_rooms['Bus_Excel'].replace(['', 'NAN', 'NAT', 'NONE'], 'Tự túc / KHÔNG CÓ XE')
                df_valid_rooms['Ten_Hien_Thi'] = df_valid_rooms.apply(
                    lambda r: str(r['Ten_Khach_Opera']) if pd.notna(r['Ten_Khach_Opera']) and str(r['Ten_Khach_Opera']).strip() not in ['', 'NAN'] else str(r['Ten_Khach_Excel']), axis=1
                )
                df_valid_rooms['Eff_Arr'] = df_valid_rooms.apply(
                    lambda r: str(r['Ngay_Den_Opera']) if pd.notna(r['Ngay_Den_Opera']) and str(r['Ngay_Den_Opera']).strip() not in ['', 'NAN'] else str(r['Ngay_Den_Excel']), axis=1
                )
                df_valid_rooms['Eff_Dep'] = df_valid_rooms.apply(
                    lambda r: str(r['Ngay_Di_Opera']) if pd.notna(r['Ngay_Di_Opera']) and str(r['Ngay_Di_Opera']).strip() not in ['', 'NAN'] else str(r['Ngay_Di_Excel']), axis=1
                )

                def make_room_display_key(row):
                    sp = str(row['So_Phong']).strip()
                    if sp == UNASSIGNED_OPERA_LABEL:
                        nhom = str(row.get('Nhom_Phong', '')).replace('.0', '').strip()
                        if nhom and nhom.upper() not in ['', 'NAN', 'NONE']:
                            return f"⏳ Chưa gán phòng (Nhóm/STT {nhom})"
                        return f"⏳ Chưa gán phòng ({row['Ten_Hien_Thi']})"
                    return sp

                df_valid_rooms['Phong_Chia_Key'] = df_valid_rooms.apply(make_room_display_key, axis=1)
                full_room_groups = {rm_k: grp for rm_k, grp in df_valid_rooms.groupby('Phong_Chia_Key')}
                room_bus_counts = df_valid_rooms.groupby('Phong_Chia_Key')['Bus_Excel'].nunique().to_dict()

                key_cut_records = []
                for (bus_val, rm_key), sub_grp in df_valid_rooms.groupby(['Bus_Excel', 'Phong_Chia_Key']):
                    arrs = sorted(list(set(a for a in sub_grp['Eff_Arr'] if a and a not in ['', 'NAN'])))
                    ci_str = " / ".join(arrs) if arrs else "Chưa rõ"
                    
                    pairs = sorted(list(set(f"{a} -> {d}" for a, d in zip(sub_grp['Eff_Arr'], sub_grp['Eff_Dep']) if a or d)))
                    validity_str = " / ".join(pairs)
                    
                    n_keys, n_booklets, rule_desc = calc_key_and_booklet_rule(
                        sub_grp,
                        full_room_groups[rm_key],
                        room_bus_counts.get(rm_key, 1)
                    )
                    guests_str = '  |  '.join([str(i) for i in sub_grp['Ten_Hien_Thi'] if str(i).strip() not in ["", "NAN"]])
                    notes_str = ' | '.join(dict.fromkeys([str(i) for i in sub_grp['Ghi_Chu_Van_Hanh'] if str(i).strip() not in ["", "NAN"]]))
                    loai_p = str(sub_grp['Loai_Phong_Opera'].iloc[0]).strip() or str(sub_grp['Loai_Phong_Excel'].iloc[0]).strip()
                    
                    key_cut_records.append({
                        'Chuyến Xe (Bus)': bus_val,
                        'Số Phòng': rm_key,
                        'Ngày Check-in': ci_str,
                        'Loại Phòng': loai_p,
                        'Số Khách': len(sub_grp),
                        'Số Thẻ Cần Cắt': n_keys,
                        'Số Bìa Key': n_booklets,
                        'Quy Cách Cắt & Đóng Bìa Key': rule_desc,
                        'Hạn Cắt Key (In -> Out)': validity_str,
                        'Danh Sách Khách Trong Phòng': guests_str,
                        'Lưu Ý Cắt Key & Vận Hành': notes_str
                    })

                df_key_cut = pd.DataFrame(key_cut_records)
                if not df_key_cut.empty:
                    df_key_cut = df_key_cut.sort_values(by=['Ngày Check-in', 'Chuyến Xe (Bus)', 'Số Phòng'])

                cols = [
                    'So_Phong', 'Quoc_Tich', 'Trang_Thai', 'Canh_Bao_Bus', 'Canh_Bao_Loai_Phong', 'Ghi_Chu_Van_Hanh',
                    'Ten_Khach_Opera', 'Ten_Khach_Excel', 'Loai_Phong_Opera', 'Loai_Phong_Excel',
                    'Bus_Excel', 'Nhom_Phong', 'Ngay_Den_Opera', 'Ngay_Di_Opera', 'Ngay_Den_Excel', 'Ngay_Di_Excel'
                ]
                for c in cols:
                    if c not in df_compare.columns: df_compare[c] = ""
                df_compare = df_compare[cols]

                df_errors = df_compare[
                    (df_compare['Trang_Thai'] != 'Khớp hoàn toàn') | 
                    (df_compare['Canh_Bao_Bus'] != "") |
                    (df_compare['Canh_Bao_Loai_Phong'] != "") |
                    (df_compare['Ghi_Chu_Van_Hanh'] != "") |
                    (df_compare['Quoc_Tich'] == '🌍 Nước Ngoài')
                ].sort_values(by=['So_Phong', 'Trang_Thai'])
                
                df_full = df_compare.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])

                st.markdown("---")
                st.success("✅ Phân tích hoàn tất! Đã áp dụng quy chuẩn 2 key/phòng và tự động cắt 3 key cho phòng Extra Bed hoặc 6 phòng Exbed sẵn (0518, 0618, 0718, 0720, 0818, 0820) ở ≥3 khách.")
                
                def highlight_special(s):
                    status_str = str(s.get('Trang_Thai', ''))
                    note_str = str(s.get('Ghi_Chu_Van_Hanh', ''))
                    if '📌 Chưa gán số phòng' in status_str: return ['background-color: #cce5ff'] * len(s)
                    if status_str != 'Khớp hoàn toàn': return ['background-color: #ffb3b3'] * len(s)
                    if '⚠️' in str(s.get('Canh_Bao_Loai_Phong', '')) or '⚠️ Chung phòng LỆCH NGÀY' in note_str: return ['background-color: #ffff99'] * len(s)
                    if '⚠️' in str(s.get('Canh_Bao_Bus', '')) or '🔄' in note_str or '🔗' in note_str: return ['background-color: #ffe5b4'] * len(s)
                    if note_str != '': return ['background-color: #d9f2d9'] * len(s)
                    if s.get('Quoc_Tich') == '🌍 Nước Ngoài': return ['background-color: #ffcccc'] * len(s)
                    return [''] * len(s)
                
                tab1, tab2, tab3 = st.tabs(["⚠️ BẢNG CẦN LƯU Ý", "📋 DANH SÁCH TỔNG ĐOÀN", "🔑 BẢNG LÀM THẺ KHÓA (CHIA BUS)"])
                with tab1:
                    if df_errors.empty: st.info("🎉 Tuyệt vời! Đoàn khớp 100%, đã gán đủ phòng, đi chung Bus, đúng loại giường và hoàn toàn thuần Việt.")
                    else: st.dataframe(df_errors.style.apply(highlight_special, axis=1), use_container_width=True)
                with tab2:
                    st.dataframe(df_full.style.apply(highlight_special, axis=1), use_container_width=True)
                with tab3:
                    st.info("💡 Chuẩn cắt key: Phòng KING = 2 key (1 bìa) | Phòng TWIN = 2 key (2 bìa riêng) | Phòng Extra Bed hoặc 6 phòng Exbed sẵn (0518, 0618, 0718, 0720, 0818, 0820) ở ≥3 khách = 3 key.")
                    f_col1, f_col2 = st.columns(2)
                    unique_ci_dates = sorted(df_key_cut['Ngày Check-in'].unique().tolist())
                    unique_buses = sorted(df_key_cut['Chuyến Xe (Bus)'].unique().tolist())
                    
                    with f_col1:
                        ci_options = ["Tất cả các ngày"] + [
                            f"{d} ({len(df_key_cut[df_key_cut['Ngày Check-in']==d])} phòng - {df_key_cut[df_key_cut['Ngày Check-in']==d]['Số Thẻ Cần Cắt'].sum()} thẻ - {df_key_cut[df_key_cut['Ngày Check-in']==d]['Số Bìa Key'].sum()} bìa)"
                            for d in unique_ci_dates
                        ]
                        selected_ci_label = st.selectbox("📅 Lọc theo Ngày Check-in:", ci_options)
                    with f_col2:
                        bus_options = ["Tất cả các xe"] + unique_buses
                        selected_bus_label = st.selectbox("🚌 Lọc theo Chuyến Xe (Bus):", bus_options)

                    df_key_filtered = df_key_cut.copy()
                    if selected_ci_label != "Tất cả các ngày":
                        chosen_date = selected_ci_label.split(" (")[0]
                        df_key_filtered = df_key_filtered[df_key_filtered['Ngày Check-in'] == chosen_date]
                    if selected_bus_label != "Tất cả các xe":
                        df_key_filtered = df_key_filtered[df_key_filtered['Chuyến Xe (Bus)'] == selected_bus_label]

                    k1, k2, k3, k4 = st.columns(4)
                    k1.metric("Số Phòng (Đang lọc)", f"{len(df_key_filtered)} phòng")
                    k2.metric("Tổng Số Thẻ Từ Cần Cắt", f"{df_key_filtered['Số Thẻ Cần Cắt'].sum()} thẻ")
                    k3.metric("Tổng Số Bìa Đựng Key", f"{df_key_filtered['Số Bìa Key'].sum()} bìa")
                    exb_rooms_count = len(df_key_filtered[df_key_filtered['Số Thẻ Cần Cắt'] >= 3])
                    k4.metric("Phòng Cắt 3 Key (Exbed)", f"{exb_rooms_count} phòng")

                    st.dataframe(df_key_filtered, use_container_width=True)

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
                            worksheet.column_dimensions[column[0].column_letter].width = min(max_length + 3, 60)

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