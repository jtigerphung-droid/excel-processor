import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 1. Cấu hình tiêu đề trang Web hiển thị
st.set_page_config(page_title="Công cụ Xử lý Excel Tự động", layout="centered")
st.title("📊 ỨNG DỤNG GỘP VÀ LÀM SẠCH DỮ LIỆU EXCEL")
st.write("Giải pháp thay thế hoàn toàn cho file `.bas` (VBA) cũ.")

# 2. Thành phần Giao diện: Cho phép người dùng kéo thả file Excel vào
uploaded_file = st.file_uploader("Bước 1: Chọn file Excel thô cần xử lý (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success("Đã tải file lên thành công!")
    
    # Tạo nút bấm để kích hoạt xử lý dữ liệu
    if st.button("Bước 2: Tiến hành Gộp & Làm sạch dữ liệu"):
        try:
            with st.spinner("Đang xử lý dữ liệu ngầm, vui lòng đợi trong giây lát..."):
                
                # Đọc tất cả các sheet từ file excel tải lên (trả về một thư viện các sheet)
                excel_file = pd.ExcelFile(uploaded_file)
                all_sheets_data = []
                
                # Duyệt qua từng sheet trong file Excel giống như vòng lặp For Each trong VBA
                for sheet_name in excel_file.sheet_names:
                    # Đọc sheet, bỏ qua 8 dòng hành chính đầu tiên (Dữ liệu bảng bắt đầu từ dòng 9, tức index 8)
                    # Không lấy tiêu đề tự động để tránh bị lệch cột
                    df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=8, header=None)
                    
                    # Nếu sheet không có dữ liệu thì bỏ qua
                    if df_sheet.empty:
                        continue
                        
                    # Đặt tên cột chuẩn cho bảng (Tương ứng với mẫu ảnh số 2 của bạn)
                    df_sheet.columns = ['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']
                    
                    # --- BẮT ĐẦU LÀM SẠCH DỮ LIỆU (Thay thế cho file BAS số 2) ---
                    # Loại bỏ các dòng trống hoàn toàn
                    df_sheet = df_sheet.dropna(subset=['Tên sách', 'Mã số'], how='all')
                    
                    # Ép kiểu cột Tên sách về dạng chuỗi văn bản để lọc chính xác
                    df_sheet['Tên sách'] = df_sheet['Tên sách'].astype(str).str.strip()
                    
                    # Loại bỏ dòng Tiêu đề lặp lại nếu có (dòng chứa chữ "Tên sách")
                    df_sheet = df_sheet[df_sheet['Tên sách'] != 'Tên sách']
                    
                    # Loại bỏ 3 dòng tổng kết cuối bảng dựa vào từ khóa
                    list_xoa = ['Tổng cộng:', 'Thuế VAT:', 'Thành tiền:', 'Tổng cộng', 'Thuế VAT']
                    df_sheet = df_sheet[~df_sheet['Tên sách'].isin(list_xoa)]
                    
                    # Lọc thêm các dòng chứa chữ trống hoặc rác hệ sinh ra
                    df_sheet = df_sheet[df_sheet['Tên sách'] != 'nan']
                    
                    # Đưa dữ liệu đã sạch của sheet này vào danh sách gộp
                    all_sheets_data.append(df_sheet)
                
                # Gộp tất cả các sheet đã làm sạch lại thành 1 bảng duy nhất
                if all_sheets_data:
                    df_final = pd.concat(all_sheets_data, ignore_index=True)
                    
                    # Đánh lại cột STT tự động từ 1 tăng dần cho toàn bộ file tổng
                    df_final['STT'] = range(1, len(df_final) + 1)
                    
                    st.success("🎉 Xử lý hoàn thành xuất sắc!")
                    
                    # Hiển thị bản xem trước (Preview) 5 dòng đầu và 5 dòng cuối ngay trên Web
                    st.write("### Bản xem trước dữ liệu kết quả:")
                    st.dataframe(df_final.head(10)) # Hiển thị 10 dòng đầu để kiểm tra
                    
                    # Chuyển đổi dữ liệu kết quả thành file Excel để người dùng tải về
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name="DuLieuGop_Sach")
                    processed_data = output.getvalue()
                    
                    # Tạo nút bấm tải file kết quả về máy tính
                    thoigian_hientai = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        label="📥 Bước 3: Tải file Excel kết quả hoàn chỉnh",
                        data=processed_data,
                        file_name=f"Ket_Qua_Gop_Sach_{thoigian_hientai}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("Không tìm thấy dữ liệu hợp lệ trong các sheet.")
                    
        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {str(e)}")
