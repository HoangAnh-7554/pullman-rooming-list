import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO

# Cấu hình giao diện Streamlit
st.set_page_config(page_title="Rooming List Checker", page_icon="🏨", layout="wide")

# ==============================================================================
# HÀM TẠO FILE EXCEL MẪU
# ==============================================================================
def create_template_excel():
    data = {
        'So_Phong': ['', '0201', '0201', '0301'],
        'Ten_Khach': ['Khach auto gan phong', 'Huynh Duc Phong', 'Nguyen Tai Uyen', 'Khach Moi'],
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
st.title("Phần mềm Đối chiếu Rooming List (Pro Version)")

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
        start_check = st.button("🚀 Bắt đầu đối chiếu đa luồng", type="primary", use_container_width=True)

    if start_check:
        with st.spinner("Đang tự động quét lỗi, gán phòng và xử lý khách ghép (Share room)..."):
            try:
                # ---------------------------------------------------------
                # BƯỚC 1: BÓC TÁCH FILE XML (OPERA)
                # ---------------------------------------------------------
                tree = ET.parse(xml_file)
                root = tree.getroot()
                
                opera_guests = []
                for guest in root.iter('G_CONFIRMATION_NO'):
                    room = guest.find('ROOM').text if guest.find('ROOM') is not None else ''
                    name = guest.find('FULL_NAME').text if guest.find('FULL_NAME') is not None else ''
                    arrival = guest.find('ARRIVAL').text if guest.find('ARRIVAL') is not None else ''
                    departure = guest.find('DEPARTURE').text if guest.find('DEPARTURE') is not None else ''
                    
                    clean_name = str(name).replace('*', '').replace(',', ' ').strip().upper()
                    
                    opera_guests.append({
                        'So_Phong': str(room).strip(),
                        'Ten_Khach_Opera': clean_name,
                        'Ngay_Den_Opera': str(arrival).strip(),
                        'Ngay_Di_Opera': str(departure).strip()
                    })
                df_opera = pd.DataFrame(opera_guests)
                
                # Trích xuất danh sách tất cả các tên trên Opera để kiểm tra chéo
                opera_all_names = df_opera['Ten_Khach_Opera'].tolist()

                # ---------------------------------------------------------
                # BƯỚC 2: ĐỌC FILE EXCEL (KHÁCH ĐOÀN)
                # ---------------------------------------------------------
                df_excel = pd.read_excel(excel_file, dtype=str)
                if 'Ten_Khach' in df_excel.columns:
                    df_excel['Ten_Khach_Excel'] = df_excel['Ten_Khach'].astype(str).str.strip().str.upper()
                else:
                    st.error("Lỗi: File Excel của bạn không có cột 'Ten_Khach'.")
                    st.stop()
                
                df_excel = df_excel.rename(columns={'Ngay_Den': 'Ngay_Den_Excel', 'Ngay_Di': 'Ngay_Di_Excel'})
                if 'So_Phong' not in df_excel.columns:
                    df_excel['So_Phong'] = ""

                # ---------------------------------------------------------
                # BƯỚC 3: XỬ LÝ TRÙNG TÊN VÀ TỰ ĐỘNG GÁN SỐ PHÒNG
                # ---------------------------------------------------------
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

                # ---------------------------------------------------------
                # BƯỚC 4: SẮP XẾP KHÁCH TRONG PHÒNG (SHARE ROOM)
                # ---------------------------------------------------------
                df_opera = df_opera.sort_values(by=['So_Phong', 'Ten_Khach_Opera'])
                df_excel = df_excel.sort_values(by=['So_Phong', 'Ten_Khach_Excel'])

                # Đánh số thứ tự người trong phòng (0, 1, 2...) để so khớp 1-1
                df_opera['Guest_Index'] = df_opera.groupby('So_Phong').cumcount()
                df_excel['Guest_Index'] = df_excel.groupby('So_Phong').cumcount()

                df_compare = pd.merge(df_opera, df_excel, on=['So_Phong', 'Guest_Index'], how='outer', indicator=True)
                
                # ---------------------------------------------------------
                # BƯỚC 5: BỘ LỌC KIỂM TRA ĐA TRƯỜNG HỢP (EDGE CASES)
                # ---------------------------------------------------------
                def deep_check_status(row):
                    merge_status = row['_merge']
                    room = str(row['So_Phong'])
                    
                    name_op = str(row['Ten_Khach_Opera'])
                    name_ex = str(row['Ten_Khach_Excel'])
                    date_in_op = str(row['Ngay_Den_Opera'])
                    date_out_op = str(row['Ngay_Di_Opera'])
                    date_in_ex = str(row['Ngay_Den_Excel']).strip().upper()
                    date_out_ex = str(row['Ngay_Di_Excel']).strip().upper()
                    
                    # 1. Trống/Thiếu số phòng không map được
                    if room == "Chưa gán được":
                        return "Lỗi nặng: Khách có trong Excel nhưng không tồn tại trên hệ thống Opera"
                        
                    # 2. Khách bị bỏ sót bên file Excel / Hủy phòng
                    if merge_status == 'left_only':
                        return "Hủy phòng / Excel gửi sót khách này"
                        
                    # 3. Tồn tại trong Excel nhưng không map được vào Opera ở vị trí này
                    if merge_status == 'right_only':
                        if date_in_ex in ['NAN', ''] or date_out_ex in ['NAN', '']:
                            return "Lỗi định dạng: File Excel bị trống ngày In/Out"
                            
                        # Kiểm tra xem tên khách này có đi lạc sang phòng khác không?
                        if name_ex in opera_all_names:
                            return "Sai số phòng (Khách này đang được xếp ở một phòng khác trên Opera)"
                        
                        # Khách dư ra trong phòng (Ví dụ phòng ở 2, đoàn nhét thành 3)
                        if row['Guest_Index'] > 0:
                            return "Phát sinh thêm người ghép phòng (Dư pax so với Opera)"
                        else:
                            return "Khách mới hoàn toàn (Chưa có trên hệ thống)"
                            
                    # 4. Khớp được cả Số phòng và Index ghép phòng
                    if merge_status == 'both':
                        if date_in_ex in ['NAN', ''] or date_out_ex in ['NAN', '']:
                            return "Lỗi định dạng: File Excel bị trống ngày In/Out"

                        name_mismatch = (name_op != name_ex)
                        date_mismatch = (date_in_op != date_in_ex) or (date_out_op != date_out_ex)
                        
                        if name_mismatch and date_mismatch:
                            return "Sai lệch KÉP (Sai tên & Sai cả Lịch trình)"
                        elif name_mismatch:
                            return "Sai lệch Tên khách"
                        elif date_mismatch:
                            return "Sai lệch ngày In/Out"
                        else:
                            return "Khớp hoàn toàn"
                            
                    return "Cần kiểm tra thủ công"

                # Áp dụng bộ lọc cho từng dòng dữ liệu
                df_compare['Trang_Thai'] = df_compare.apply(deep_check_status, axis=1)
                
                # Làm gọn bảng kết quả
                df_compare = df_compare.drop(columns=['_merge', 'Guest_Index'])
                
                # Ưu tiên xếp các dòng có chữ "Khớp hoàn toàn" xuống chót, đẩy các dòng lỗi lên đầu
                df_compare['Is_Perfect'] = df_compare['Trang_Thai'] == 'Khớp hoàn toàn'
                df_compare = df_compare.sort_values(by=['Is_Perfect', 'So_Phong'])
                df_compare = df_compare.drop(columns=['Is_Perfect'])

                # ---------------------------------------------------------
                # BƯỚC 6: XUẤT KẾT QUẢ VÀ HIỂN THỊ
                # ---------------------------------------------------------
                st.markdown("---")
                st.success("✅ Đã xử lý tất cả các trường hợp ngoại lệ thành công!")
                
                st.dataframe(df_compare, use_container_width=True)

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
                        file_name="Bao_Cao_Rooming_List_Nang_Cao.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Đã xảy ra lỗi hệ thống: {e}")