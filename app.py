import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 1. Cấu hình tiêu đề trang Web hiển thị
st.set_page_config(page_title="Công cụ Xử lý Excel Tự động", layout="centered")
st.title("📊 ỨNG DỤNG GỘP VÀ LÀM SẠCH DỮ LIỆU EXCEL")
st.write("Phiên bản cập nhật: Bộ lọc thông minh triệt để dòng thừa.")

# 2. Thành phần Giao diện: Cho phép người dùng kéo thả file Excel vào
uploaded_file = st.file_uploader("Bước 1: Chọn file Excel thô cần xử lý (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success("Đã tải file lên thành công!")
    
    # Tạo nút bấm để kích hoạt xử lý dữ liệu
    if st.button("Bước 2: Tiến hành Gộp & Làm sạch dữ liệu"):
        try:
            with st.spinner("Đang xử lý dữ liệu ngầm, vui lòng đợi trong giây lát..."):
                
                excel_file = pd.ExcelFile(uploaded_file)
                all_sheets_data = []
                
                # Duyệt qua từng sheet trong file Excel
                for sheet_name in excel_file.sheet_names:
                    # Đọc sheet, bỏ qua 8 dòng hành chính đầu tiên
                    df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=8, header=None)
                    
                    if df_sheet.empty:
                        continue
                        
                    # Đặt tên cột tạm thời để dễ xử lý lọc
                    df_sheet.columns = ['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']
                    
                    # --- BỘ LỌC THÔNG MINH CẬP NHẬT (Xử lý triệt để ảnh lỗi) ---
                    
                    # Cắt khoảng trắng thừa ở tất cả các cột dạng chữ để tránh sót lỗi
                    for col in ['Tên sách', 'Mã số', 'ĐVT']:
                        df_sheet[col] = df_sheet[col].astype(str).str.strip()
                    
                    # Lọc 1: Xóa ngay dòng nếu cột "Mã số" trống, hoặc chứa chữ "Mã số" (tiêu đề lặp)
                    df_sheet = df_sheet[df_sheet['Mã số'].notna() & (df_sheet['Mã số'] != '') & (df_sheet['Mã số'] != 'nan')]
                    df_sheet = df_sheet[df_sheet['Mã số'] != 'Mã số']
                    
                    # Lọc 2: Xóa ngay dòng nếu cột "Tên sách" chứa các từ khóa hành chính hoặc tổng kết
                    tu_khoa_xoa = [
                        'Tổng cộng:', 'Thuế VAT:', 'Thành tiền:', 'Tổng cộng', 'Thuế VAT', 
                        'Phụ trách cung tiêu', 'Người giao hàng', 'Thủ Kho', 'nan', 'STT', 'Tên sách'
                    ]
                    # Kiểm tra xem cột Tên sách có chứa bất kỳ từ khóa rác nào ở trên không
                    df_sheet = df_sheet[~df_sheet['Tên sách'].isin(tu_khoa_xoa)]
                    
                    # Lọc 3: Một lớp khóa phụ - nếu dòng nào có "Giá bìa" hoặc "Số lượng" không phải là số, loại bỏ luôn
                    df_sheet = df_sheet[pd.to_numeric(df_sheet['Số lượng'], errors='coerce').notna()]
                    
                    # Đưa dữ liệu đã sạch của sheet này vào danh sách gộp
                    all_sheets_data.append(df_sheet)
                
                # Gộp tất cả các sheet đã làm sạch lại thành 1 bảng duy nhất
                if all_sheets_data:
                    df_final = pd.concat(all_sheets_data, ignore_index=True)
                    
                    # Ép kiểu dữ liệu chuẩn cho các cột số để tính toán đẹp mắt
                    df_final['Số lượng'] = pd.to_numeric(df_final['Số lượng'])
                    df_final['Giá bìa'] = pd.to_numeric(df_final['Giá bìa'])
                    df_final['Đơn giá'] = pd.to_numeric(df_final['Đơn giá'])
                    df_final['Thành tiền'] = pd.to_numeric(df_final['Thành tiền'])
                    
                    # Đánh lại cột STT tự động từ 1 tăng dần cho toàn bộ file tổng [cite: 1]
                    df_final['STT'] = range(1, len(df_final) + 1)
                    
                    st.success("🎉 Đã làm sạch toàn bộ dòng thừa!")
                    
                    # Hiển thị bản xem trước (Preview) ngay trên Web
                    st.write("### Bản xem trước dữ liệu kết quả đã sửa lỗi:")
                    st.dataframe(df_final.head(15))
                    
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
                        file_name=f"Ket_Qua_Gop_Sach_SuaLoi_{thoigian_hientai}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("Không tìm thấy dữ liệu hợp lệ trong các sheet sau khi lọc.")
                    
        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {str(e)}")
