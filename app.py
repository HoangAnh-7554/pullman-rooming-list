import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import numpy as np
from io import BytesIO

# Cấu hình giao diện Streamlit
st.set_page_config(page_title="Rooming List Checker", page_icon="🏨", layout="wide")

# ==============================================================================
# HÀM TẠO FILE EXCEL MẪU TRỰC TIẾP TỪ CODE ĐỂ TẢI XUỐNG
# ==============================================================================
def create_template_excel():
    # Tạo dữ liệu mẫu
    data = {
        'So_Phong': ['', '0201', ''],
        'Ten_Khach': ['Nguyen Van A', 'Tran Thi B', 'Le Van C'],
        'Ngay_Den': ['03-10-26', '03-10-26', '03-10-26'],
        'Ngay_Di': ['04-10-26', '05-10-26', '04-10-26']
    }
    df_template = pd.DataFrame(data)
    
    # Tạo buffer xuất ra bộ nhớ
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Danh_Sach_Doan')
        
        # Tự động điều chỉnh độ rộng cột cho đẹp
        worksheet = writer.sheets['Danh_Sach_Doan']
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['B'].width = 30
        worksheet.column_dimensions['C'].width = 20
        worksheet.column_dimensions['D'].width = 20
        
    return output.getvalue()

# ==============================================================================
# PHẦN GIAO DIỆN CHÍNH
# ==============================================================================
st.title("Phần mềm Đối chiếu Rooming List (Opera vs Excel)")

# Thêm khu vực tải file mẫu có viền màu xanh (info)
with st.info("Tải file Excel mẫu về máy để điền thông tin đoàn trước khi tải lên (Nếu bạn chưa có form chuẩn)."):
    st.download_button(
        label="📥 Tải File Excel Mẫu (Template)",
        data=create_template_excel(),
        file_name="Template_Rooming_List_Doan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Định dạng chuẩn gồm 4 cột: So_Phong, Ten_Khach, Ngay_Den, Ngay_Di"
    )

st.markdown("---")
st.markdown("### VUI LÒNG TẢI FILE LÊN ĐỂ KIỂM TRA")

# Tạo 2 cột để upload file
col1, col2 = st.columns(2)
with col1:
    xml_file = st.file_uploader("1. Tải lên file XML (Từ hệ thống Opera)", type=['xml'])
with col2:
    excel_file = st.file_uploader("2. Tải lên file Excel (Khách đoàn đã điền theo Mẫu)", type=['xlsx', 'xls'])

# Bắt đầu xử lý khi cả 2 file đã được tải lên
if xml_file and excel_file:
    # Nút bấm căn giữa bằng columns
    _, btn_col, _ = st.columns([1, 1, 1])
    with btn_col:
        start_check = st.button("🚀 Bắt đầu đối chiếu", type="primary", use_container_width=True)

    if start_check:
        with st.spinner("Đang tự động gán phòng và so khớp dữ liệu..."):
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
                    
                    # Làm sạch tên
                    clean_name = str(name).replace('*', '').replace(',', ' ').strip().upper()
                    
                    opera_guests.append({
                        'So_Phong': str(room).strip(),
                        'Ten_Khach_Opera': clean_name,
                        'Ngay_Den_Opera': str(arrival).strip(),
                        'Ngay_Di_Opera': str(departure).strip()
                    })
                df_opera = pd.DataFrame(opera_guests)

                # ---------------------------------------------------------
                # BƯỚC 2: ĐỌC VÀ CHUẨN HÓA FILE EXCEL
                # ---------------------------------------------------------
                df_excel = pd.read_excel(excel_file, dtype=str)
                
                # Kiểm tra cột Tên khách
                if 'Ten_Khach' in df_excel.columns:
                    df_excel['Ten_Khach_Excel'] = df_excel['Ten_Khach'].astype(str).str.strip().str.upper()
                else:
                    st.error("Lỗi: File Excel của bạn tải lên không đúng định dạng mẫu (Thiếu cột 'Ten_Khach').")
                    st.stop()
                
                df_excel = df_excel.rename(columns={
                    'Ngay_Den': 'Ngay_Den_Excel',
                    'Ngay_Di': 'Ngay_Di_Excel'
                })

                # ---------------------------------------------------------
                # BƯỚC 3: TỰ ĐỘNG GÁN SỐ PHÒNG (NẾU THIẾU)
                # ---------------------------------------------------------
                room_mapping = dict(zip(df_opera['Ten_Khach_Opera'], df_opera['So_Phong']))
                
                if 'So_Phong' not in df_excel.columns:
                    df_excel['So_Phong'] = ""

                def auto_assign_room(row):
                    if pd.isna(row['So_Phong']) or str(row['So_Phong']).strip() in ["", "NAN", "NAT"]:
                        return room_mapping.get(row['Ten_Khach_Excel'], "Chưa gán được")
                    return str(row['So_Phong']).strip()

                df_excel['So_Phong'] = df_excel.apply(auto_assign_room, axis=1)

                # ---------------------------------------------------------
                # BƯỚC 4: ĐỐI CHIẾU DỮ LIỆU
                # ---------------------------------------------------------
                df_compare = pd.merge(df_opera, df_excel, on='So_Phong', how='outer', indicator=True)
                
                conditions = [
                    (df_compare['_merge'] == 'left_only'),
                    (df_compare['So_Phong'] == "Chưa gán được"), 
                    (df_compare['_merge'] == 'right_only'),
                    (df_compare['_merge'] == 'both') & (df_compare['Ten_Khach_Opera'] != df_compare['Ten_Khach_Excel']),
                    (df_compare['_merge'] == 'both') & ((df_compare['Ngay_Den_Opera'] != df_compare['Ngay_Den_Excel']) | (df_compare['Ngay_Di_Opera'] != df_compare['Ngay_Di_Excel'])),
                    (df_compare['_merge'] == 'both')
                ]
                
                choices = [
                    'Hủy phòng (Có trong Opera, không có trong Excel)',
                    'Lỗi gán phòng (Có tên Excel, không map được vào Opera)',
                    'Phòng mới thêm vào (Có trong Excel, chưa nhập Opera)',
                    'Tên khách chưa khớp',
                    'Sai lệch ngày In/Out',
                    'Khớp hoàn toàn'
                ]
                
                df_compare['Trang_Thai'] = np.select(conditions, choices, default='Cần kiểm tra lại')
                df_compare = df_compare.drop(columns=['_merge']).sort_values(by=['Trang_Thai', 'So_Phong'])

                # ---------------------------------------------------------
                # BƯỚC 5: HIỂN THỊ KẾT QUẢ VÀ TẢI XUỐNG
                # ---------------------------------------------------------
                st.markdown("---")
                st.success("✅ Đã xử lý đối chiếu và tự động gán phòng thành công!")
                
                st.dataframe(df_compare, use_container_width=True)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_compare.to_excel(writer, index=False, sheet_name='Ket_Qua_Doi_Chieu')
                    
                    # Tự động canh độ rộng cột cho báo cáo xuất ra
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
                        adjusted_width = (max_length + 2)
                        worksheet.column_dimensions[column_letter].width = adjusted_width

                _, dwn_col, _ = st.columns([1, 1, 1])
                with dwn_col:
                    st.download_button(
                        label="📥 TẢI FILE KẾT QUẢ ĐỐI CHIẾU (EXCEL)",
                        data=output.getvalue(),
                        file_name="Ket_Qua_Doi_Chieu_Rooming_List.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Đã xảy ra lỗi: {e}")
                st.info("💡 Vui lòng đảm bảo bạn sử dụng đúng file Excel Template được tải từ hệ thống này.")