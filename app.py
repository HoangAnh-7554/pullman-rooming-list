import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import numpy as np
from io import BytesIO

# Cấu hình giao diện Streamlit
st.set_page_config(page_title="Rooming List Checker", page_icon="🏨", layout="wide")
st.title("Phần mềm Đối chiếu Rooming List (Opera vs Excel)")
st.markdown("""
**HƯỚNG DẪN SỬ DỤNG:**
1. File Excel từ khách đoàn cần có các cột: `Ten_Khach`, `So_Phong`, `Ngay_Den`, `Ngay_Di`.
2. Nếu file Excel chưa có số phòng, để trống cột `So_Phong`, phần mềm sẽ **tự động gán** dựa trên tên khách từ hệ thống Opera.
3. Kéo thả file XML (từ Opera) và file Excel (từ khách đoàn) vào bên dưới.
""")

# Tạo 2 cột để upload file
col1, col2 = st.columns(2)
with col1:
    xml_file = st.file_uploader("1. Tải lên file XML (Từ Opera)", type=['xml'])
with col2:
    excel_file = st.file_uploader("2. Tải lên file Excel (Từ Khách đoàn)", type=['xlsx', 'xls'])

# Bắt đầu xử lý khi cả 2 file đã được tải lên
if xml_file and excel_file:
    if st.button("Bắt đầu đối chiếu", type="primary"):
        with st.spinner("Đang xử lý dữ liệu..."):
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
                    
                    # Làm sạch tên khách (Bỏ dấu *, thay phẩy thành khoảng trắng, viết hoa)
                    clean_name = str(name).replace('*', '').replace(',', ' ').strip().upper()
                    
                    opera_guests.append({
                        'So_Phong': str(room).strip(),
                        'Ten_Khach_Opera': clean_name,
                        'Ngay_Den_Opera': str(arrival).strip(),
                        'Ngay_Di_Opera': str(departure).strip()
                    })
                df_opera = pd.DataFrame(opera_guests)

                # ---------------------------------------------------------
                # BƯỚC 2: ĐỌC VÀ CHUẨN HÓA FILE EXCEL (KHÁCH ĐOÀN)
                # ---------------------------------------------------------
                df_excel = pd.read_excel(excel_file, dtype=str)
                
                # Làm sạch Tên khách bên Excel để đồng bộ định dạng với Opera
                if 'Ten_Khach' in df_excel.columns:
                    df_excel['Ten_Khach_Excel'] = df_excel['Ten_Khach'].astype(str).str.strip().str.upper()
                else:
                    st.error("Lỗi: File Excel của bạn không có cột 'Ten_Khach'. Vui lòng kiểm tra lại.")
                    st.stop()
                
                # Đổi tên các cột ngày tháng để dễ so khớp
                df_excel = df_excel.rename(columns={
                    'Ngay_Den': 'Ngay_Den_Excel',
                    'Ngay_Di': 'Ngay_Di_Excel'
                })

                # ---------------------------------------------------------
                # BƯỚC 3: TỰ ĐỘNG GÁN SỐ PHÒNG (NẾU THIẾU)
                # ---------------------------------------------------------
                # Khởi tạo "từ điển" map từ Tên Khách Opera -> Số Phòng Opera
                room_mapping = dict(zip(df_opera['Ten_Khach_Opera'], df_opera['So_Phong']))
                
                # Đảm bảo file Excel luôn có cột 'So_Phong' để xử lý
                if 'So_Phong' not in df_excel.columns:
                    df_excel['So_Phong'] = ""

                # Hàm kiểm tra và điền số phòng
                def auto_assign_room(row):
                    # Nếu ô Số phòng đang trống hoặc chứa giá trị rỗng (NaN, NaT, None)
                    if pd.isna(row['So_Phong']) or str(row['So_Phong']).strip() in ["", "NAN", "NAT"]:
                        # Tìm trong Opera, nếu thấy Tên thì gán phòng, không thấy thì ghi "Chưa gán được"
                        return room_mapping.get(row['Ten_Khach_Excel'], "Chưa gán được")
                    # Nếu đã có sẵn số phòng trong Excel thì giữ nguyên
                    return str(row['So_Phong']).strip()

                df_excel['So_Phong'] = df_excel.apply(auto_assign_room, axis=1)

                # ---------------------------------------------------------
                # BƯỚC 4: ĐỐI CHIẾU DỮ LIỆU BẰNG OUTER JOIN
                # ---------------------------------------------------------
                df_compare = pd.merge(df_opera, df_excel, on='So_Phong', how='outer', indicator=True)
                
                # Lọc ra các điều kiện sai lệch
                conditions = [
                    (df_compare['_merge'] == 'left_only'),
                    (df_compare['So_Phong'] == "Chưa gán được"), # Nhận diện những khách Excel không tự map được phòng
                    (df_compare['_merge'] == 'right_only'),
                    (df_compare['_merge'] == 'both') & (df_compare['Ten_Khach_Opera'] != df_compare['Ten_Khach_Excel']),
                    (df_compare['_merge'] == 'both') & ((df_compare['Ngay_Den_Opera'] != df_compare['Ngay_Den_Excel']) | (df_compare['Ngay_Di_Opera'] != df_compare['Ngay_Di_Excel'])),
                    (df_compare['_merge'] == 'both')
                ]
                
                choices = [
                    'Hủy phòng (Có trong Opera, không có trong Excel)',
                    'Lỗi gán phòng (Có tên Excel, không map được vào Opera)',
                    'Phòng mới (Có trong Excel, không có trong Opera)',
                    'Sai lệch Tên Khách (Kiểm tra lại)',
                    'Sai lịch trình In/Out',
                    'Khớp hoàn toàn'
                ]
                
                df_compare['Trang_Thai'] = np.select(conditions, choices, default='Cần kiểm tra lại')
                
                # Xóa cột kỹ thuật _merge và sắp xếp dữ liệu, ưu tiên hiển thị lỗi lên đầu
                df_compare = df_compare.drop(columns=['_merge']).sort_values(by=['Trang_Thai', 'So_Phong'])

                # ---------------------------------------------------------
                # BƯỚC 5: HIỂN THỊ VÀ XUẤT KẾT QUẢ ĐỂ TẢI VỀ
                # ---------------------------------------------------------
                st.success("✅ Đã xử lý đối chiếu và tự động gán phòng thành công!")
                
                # Bảng xem trước trên Web
                st.dataframe(df_compare, use_container_width=True)

                # Tạo buffer xuất file Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_compare.to_excel(writer, index=False, sheet_name='Ket_Qua_Doi_Chieu')
                
                # Nút Download
                st.download_button(
                    label="📥 TẢI FILE BÁO CÁO (EXCEL)",
                    data=output.getvalue(),
                    file_name="Bao_Cao_Rooming_List_Pullman.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Đã xảy ra lỗi: {e}")
                st.info("💡 Lời khuyên: Đảm bảo file Excel của bạn ĐÃ CHUẨN HÓA các tiêu đề cột đúng y hệt như sau: Ten_Khach, Ngay_Den, Ngay_Di, So_Phong (chữ viết hoa đầu từ, ngăn cách bởi dấu gạch dưới).")
