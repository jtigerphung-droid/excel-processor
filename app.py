import streamlit as st
import pandas as pd
import io
from datetime import datetime
# Import thêm các công cụ định dạng giao diện Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 1. Cấu hình tiêu đề trang Web hiển thị
st.set_page_config(page_title="Công cụ Xử lý Excel Tự động", layout="centered")
st.title("📊 ỨNG DỤNG GỘP VÀ LÀM SẠCH DỮ LIỆU EXCEL")
st.write("Phiên bản nâng cấp: Tự động kẻ bảng, tô màu tiêu đề và định dạng số phân cách hàng ngàn.")

# 2. Thành phần Giao diện: Cho phép người dùng kéo thả file Excel vào
uploaded_file = st.file_uploader("Bước 1: Chọn file Excel thô cần xử lý (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success("Đã tải file lên thành công!")
    
    # Tạo nút bấm để kích hoạt xử lý dữ liệu
    if st.button("Bước 2: Tiến hành Gộp & Định dạng báo cáo"):
        try:
            with st.spinner("Đang xử lý dữ liệu và thiết kế bảng tính ngầm..."):
                
                excel_file = pd.ExcelFile(uploaded_file)
                all_sheets_data = []
                
                # Duyệt qua từng sheet trong file Excel
                for sheet_name in excel_file.sheet_names:
                    df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=8, header=None)
                    
                    if df_sheet.empty:
                        continue
                        
                    df_sheet.columns = ['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']
                    
                    # --- BỘ LỌC LÀM SẠCH ---
                    for col in ['Tên sách', 'Mã số', 'ĐVT']:
                        df_sheet[col] = df_sheet[col].astype(str).str.strip()
                    
                    df_sheet = df_sheet[df_sheet['Mã số'].notna() & (df_sheet['Mã số'] != '') & (df_sheet['Mã số'] != 'nan')]
                    df_sheet = df_sheet[df_sheet['Mã số'] != 'Mã số']
                    
                    tu_khoa_xoa = [
                        'Tổng cộng:', 'Thuế VAT:', 'Thành tiền:', 'Tổng cộng', 'Thuế VAT', 
                        'Phụ trách cung tiêu', 'Người giao hàng', 'Thủ Kho', 'nan', 'STT', 'Tên sách'
                    ]
                    df_sheet = df_sheet[~df_sheet['Tên sách'].isin(tu_khoa_xoa)]
                    df_sheet = df_sheet[pd.to_numeric(df_sheet['Số lượng'], errors='coerce').notna()]
                    
                    all_sheets_data.append(df_sheet)
                
                # Gộp tất cả các sheet lại thành 1 bảng duy nhất
                if all_sheets_data:
                    df_final = pd.concat(all_sheets_data, ignore_index=True)
                    
                    # Chuyển đổi dữ liệu chuẩn số để sẵn sàng định dạng tiền tệ
                    df_final['Số lượng'] = pd.to_numeric(df_final['Số lượng'])
                    df_final['Giá bìa'] = pd.to_numeric(df_final['Giá bìa'])
                    df_final['CK'] = pd.to_numeric(df_final['CK'])
                    df_final['Đơn giá'] = pd.to_numeric(df_final['Đơn giá'])
                    df_final['Thành tiền'] = pd.to_numeric(df_final['Thành tiền'])
                    
                    # Đánh lại cột STT tự động từ 1 tăng dần
                    df_final['STT'] = range(1, len(df_final) + 1)
                    
                    # --- QUY TRÌNH XUẤT FILE VÀ TRANG TRÍ ĐỊNH DẠNG (OPENPYXL) ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # Ghi dữ liệu thô vào sheet trước
                        df_final.to_excel(writer, index=False, sheet_name="DuLieuGop_Chuan")
                        
                        # Lấy đối tượng workbook và worksheet để trang trí
                        workbook = writer.book
                        worksheet = writer.sheets["DuLieuGop_Chuan"]
                        
                        # Định nghĩa các "Style" trang trí bảng
                        font_tieude = Font(name="Arial", size=11, bold=True, color="FFFFFF") # Chữ tiêu đề: Màu trắng, in đậm
                        fill_tieude = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Nền tiêu đề: Xanh dương đậm
                        font_noidung = Font(name="Arial", size=11) # Chữ nội dung: Arial 11 thông thường
                        
                        # Định nghĩa viền kẻ bảng (Border) mỏng, màu xám chuyên nghiệp
                        vien_mong = Side(border_style="thin", color="D9D9D9")
                        border_o = Border(left=vien_mong, right=vien_mong, top=vien_mong, bottom=vien_mong)
                        
                        # Cấu hình căn lề (Alignment)
                        cai_le_trai = Alignment(horizontal="left", vertical="center")
                        cai_le_giua = Alignment(horizontal="center", vertical="center")
                        cai_le_phai = Alignment(horizontal="right", vertical="center")
                        
                        # 1. Định dạng riêng cho Dòng tiêu đề (Dòng 1)
                        worksheet.row_dimensions[1].height = 26 # Tăng độ cao dòng tiêu đề cho thoáng
                        for col_idx in range(1, len(df_final.columns) + 1):
                            cell = worksheet.cell(row=1, column=col_idx)
                            cell.font = font_tieude
                            cell.fill = fill_tieude
                            cell.alignment = cai_le_giua
                            cell.border = border_o
                        
                        # 2. Định dạng cho các Dòng nội dung dữ liệu (Từ dòng 2 trở đi)
                        dinh_dang_phan_ngan = "#,##0" # Định dạng số chuẩn quốc tế (Có phân cách hàng ngàn)
                        
                        for row_idx in range(2, worksheet.max_row + 1):
                            worksheet.row_dimensions[row_idx].height = 20 # Độ cao dòng vừa vặn
                            
                            for col_idx in range(1, len(df_final.columns) + 1):
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                cell.font = font_noidung
                                cell.border = border_o
                                
                                # Căn lề và Định dạng số theo từng loại cột
                                col_name = df_final.columns[col_idx - 1]
                                
                                if col_name in ['STT', 'Mã số', 'ĐVT']:
                                    cell.alignment = cai_le_giua
                                elif col_name in ['Tên sách']:
                                    cell.alignment = cai_le_trai
                                elif col_name in ['Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']:
                                    cell.alignment = cai_le_phai
                                    cell.number_format = dinh_dang_phan_ngan # Áp dụng dấu phân cách hàng ngàn tại đây
                        
                        # 3. Tự động giãn rộng các cột thông minh theo độ dài chữ để không bị lỗi hiển thị "###"
                        for col in worksheet.columns:
                            max_len = max(len(str(cell.value or '')) for cell in col)
                            col_letter = get_column_letter(col[0].column)
                            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
                            
                    processed_data = output.getvalue()
                    
                    st.success("🎉 Đã hoàn thành gộp và thiết kế giao diện bảng tính đẹp mắt!")
                    
                    # Hiển thị bản xem trước trên Web
                    st.write("### Bản xem trước cấu trúc dữ liệu:")
                    st.dataframe(df_final.head(15))
                    
                    # Tạo nút bấm tải file kết quả đã có định dạng đẹp về máy
                    thoigian_hientai = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        label="📥 Bước 3: Tải file Excel kết quả (Đã kẻ bảng & Định dạng số)",
                        data=processed_data,
                        file_name=f"Bao_Cao_Gop_Dinh_Dang_{thoigian_hientai}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("Không tìm thấy dữ liệu hợp lệ để xử lý.")
                    
        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {str(e)}")
