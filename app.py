import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CẤU HÌNH GIAO DIỆN WEB ---
st.set_page_config(page_title="Hệ thống Xử lý & Phân tích Dữ liệu Phát hành", layout="wide")
st.title("📊 HỆ THỐNG XỬ LÝ DỮ LIỆU PHÁT HÀNH SÁCH TỰ ĐỘNG")
st.write("Tích hợp quy trình: Làm sạch dữ liệu dòng thừa ➡️ Phân tích, gom nhóm mã hàng ➡️ Tách hóa đơn tự động.")

# --- HÀM TRANG TRÍ EXCEL CHUYÊN NGHIỆP ---
def trang_tri_sheet(worksheet, tieude_color, has_vat_summary=False, total_row_type="standard"):
    font_tieude = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill_tieude = PatternFill(start_color=tieude_color, end_color=tieude_color, fill_type="solid")
    font_noidung = Font(name="Arial", size=10)
    font_tongket = Font(name="Arial", size=10, bold=True)
    
    vien_mong = Side(border_style="thin", color="D9D9D9")
    border_o = Border(left=vien_mong, right=vien_mong, top=vien_mong, bottom=vien_mong)
    
    # 1. Định dạng dòng tiêu đề
    worksheet.row_dimensions[1].height = 25
    for col_idx in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.font = font_tieude
        cell.fill = fill_tieude
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_o

    # 2. Định dạng nội dung và số phân ngàn
    max_r = worksheet.max_row
    for row_idx in range(2, max_r + 1):
        worksheet.row_dimensions[row_idx].height = 19
        
        is_total_row = False
        if total_row_type == "grand" and row_idx == max_r:
            is_total_row = True
        elif total_row_type == "standard" and has_vat_summary and row_idx >= max_r - 2:
            is_total_row = True
        elif total_row_type == "standard" and not has_vat_summary and row_idx == max_r:
            is_total_row = True

        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = border_o
            
            if is_total_row:
                cell.font = font_tongket
                if total_row_type == "grand":
                    cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            else:
                cell.font = font_noidung

            # Phân loại canh lề dữ liệu sách
            if col_idx in [1, 3]:  # STT, Mã số
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx == 3: cell.number_format = '@'
            elif col_idx == 2:  # Tên sách
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 4:  # ĐVT
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:  # Các cột số tiền, số lượng
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"

    # 3. Tự động giãn cột sinh động
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

# --- KHỞI CHẠY QUY TRÌNH XỬ LÝ ---
uploaded_file = st.file_uploader("Kéo thả file Excel thô của hệ thống vào đây:", type=["xlsx"])

if uploaded_file is not None:
    st.success("Đã nhận file thô thành công!")
    
    if st.button("🚀 BẮT ĐẦU XỬ LÝ VÀ PHÂN TÍCH MASTER PROCESS"):
        try:
            with st.spinner("Hệ thống đang chạy bộ lọc và phân tích cấu trúc tài chính nâng cao..."):
                excel_file = pd.ExcelFile(uploaded_file)
                all_sheets_data = []
                
                # --- BƯỚC 1: ĐỌC VÀ LÀM SẠCH DỮ LIỆU CHUẨN HOÁ NGOẠI LAI ---
                for sheet_name in excel_file.sheet_names:
                    df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=8, header=None)
                    if df_sheet.empty: continue
                    
                    # Giới hạn chuẩn hóa đúng số lượng cột của bảng dữ liệu gốc
                    df_sheet = df_sheet.iloc[:, 0:9]
                    df_sheet.columns = ['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']
                    
                    for col in ['Tên sách', 'Mã số', 'ĐVT']:
                        df_sheet[col] = df_sheet[col].astype(str).str.strip()
                    
                    df_sheet = df_sheet[df_sheet['Mã số'].notna() & (df_sheet['Mã số'] != '') & (df_sheet['Mã số'] != 'nan') & (df_sheet['Mã số'] != 'Mã số')]
                    tu_khoa_xoa = ['Tổng cộng:', 'Thuế VAT:', 'Thành tiền:', 'Tổng cộng', 'Thuế VAT', 'Phụ trách cung tiêu', 'Người giao hàng', 'Thủ Kho', 'nan', 'STT', 'Tên sách']
                    df_sheet = df_sheet[~df_sheet['Tên sách'].isin(tu_khoa_xoa)]
                    
                    all_sheets_data.append(df_sheet)
                
                if not all_sheets_data:
                    st.error("Không trích xuất được dữ liệu hợp lệ từ file Excel nạp vào.")
                    st.stop()
                    
                df_raw = pd.concat(all_sheets_data, ignore_index=True)
                
                # Ép kiểu dữ liệu số chính xác chống tràn bộ nhớ
                for col in ['Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']:
                    df_raw[col] = pd.to_numeric(df_raw[col], errors='coerce').fillna(0)
                
                # Loại bỏ dòng lỗi logic không có số lượng thực tế
                df_raw = df_raw[df_raw['Số lượng'] != 0]
                df_raw['STT'] = range(1, len(df_raw) + 1)
                
                # --- BƯỚC 2: PHÂN TÍCH THEO PHƯƠNG PHÁP DICTIONARY (FILE .BAS) ---
                df_duong = df_raw[df_raw['Số lượng'] >= 0].copy()
                df_am = df_raw[df_raw['Số lượng'] < 0].copy()
                
                # 2.1 Tổng hợp mã hàng bán (Gom theo Mã số, lấy đơn giá bìa làm gốc) [cite: 78, 83]
                df_tonghop_ban = df_duong.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False).agg({'Số lượng': 'sum'})
                df_tonghop_ban['CK'] = 0
                df_tonghop_ban['Đơn giá'] = df_tonghop_ban['Giá bìa']
                df_tonghop_ban['Thành tiền'] = df_tonghop_ban['Số lượng'] * df_tonghop_ban['Đơn giá']
                df_tonghop_ban = df_tonghop_ban[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                df_tonghop_ban.insert(0, 'STT', range(1, len(df_tonghop_ban) + 1))
                
                # 2.2 Tổng hợp hàng trả (Chuyển dấu Âm sang Dương - theo Giá bìa) [cite: 82, 86]
                df_tonghop_tra = df_am.copy()
                df_tonghop_tra['Số lượng'] = df_tonghop_tra['Số lượng'].abs()
                df_tonghop_tra = df_tonghop_tra.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False).agg({'Số lượng': 'sum'})
                df_tonghop_tra['CK'] = 0
                df_tonghop_tra['Đơn giá'] = df_tonghop_tra['Giá bìa']
                df_tonghop_tra['Thành tiền'] = df_tonghop_tra['Số lượng'] * df_tonghop_tra['Đơn giá']
                df_tonghop_tra = df_tonghop_tra[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                df_tonghop_tra.insert(0, 'STT', range(1, len(df_tonghop_tra) + 1))
                
                # 2.3 Chi tiết hàng trả (Giữ nguyên dấu và tỷ lệ chiết khấu thực tế) [cite: 80, 89]
                df_chitiet_am = df_am.copy()
                df_chitiet_am['Thành tiền'] = df_chitiet_am['Số lượng'] * df_chitiet_am['Đơn giá']
                df_chitiet_am['STT'] = range(1, len(df_chitiet_am) + 1)
                df_chitiet_am = df_chitiet_am[['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                
                # 2.4 Tổng hợp toàn bộ danh mục xuất bán ra [cite: 91, 92]
                df_tonghop_banra = df_duong.copy()
                df_tonghop_banra['Thành tiền'] = df_tonghop_banra['Số lượng'] * df_tonghop_banra['Đơn giá']
                df_tonghop_banra['STT'] = range(1, len(df_tonghop_banra) + 1)
                df_tonghop_banra = df_tonghop_banra[['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]

                # --- BƯỚC 3: ĐÓNG GÓI VÀ TẠO FILE EXCEL REPORT ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    grand_summary_data = []
                    
                    def ghi_sheet_kem_tong(df, name, color, has_vat=False):
                        df.to_excel(writer, index=False, sheet_name=name)
                        ws = writer.sheets[name]
                        r_idx = len(df) + 2
                        
                        sum_sl = df['Số lượng'].sum()
                        sum_tt = df['Thành tiền'].sum()
                        
                        ws.cell(row=r_idx, column=6).value = "Tổng cộng:"
                        ws.cell(row=r_idx, column=7).value = sum_sl
                        ws.cell(row=r_idx, column=9).value = sum_tt
                        
                        vat, sau_thue = 0.0, 0.0
                        if has_vat:
                            vat = sum_tt * 0.05
                            sau_thue = sum_tt * 1.05
                            ws.cell(row=r_idx+1, column=8).value = "VAT 5%:"
                            ws.cell(row=r_idx+1, column=9).value = vat
                            ws.cell(row=r_idx+2, column=8).value = "Sau Thuế:"
                            ws.cell(row=r_idx+2, column=9).value = sau_thue
                            
                        trang_tri_sheet(ws, color, has_vat_summary=has_vat)
                        return len(df), sum_sl, sum_tt, vat, sau_thue

                    # Ghi các sheet chức năng theo mã màu nhận diện nghiệp vụ [cite: 85, 87, 90, 92]
                    if not df_tonghop_ban.empty:
                        ln, sl, tt, _, _ = ghi_sheet_kem_tong(df_tonghop_ban, "Tong_Hop_Ma_Hang_Ban", "003366")
                        grand_summary_data.append(["Tong Hop Ma Hang Ban", "Duong", ln, sl, tt, 0, tt, "Gia Bia (Khong VAT)"])
                        
                    if not df_tonghop_tra.empty:
                        ln, sl, tt, _, _ = ghi_sheet_kem_tong(df_tonghop_tra, "Tong_Hop_Hang_Tra", "FF6600")
                        grand_summary_data.append(["Tong Hop Hang Tra", "Duong", ln, sl, tt, 0, tt, "Gia Bia (Khong VAT)"])
                        
                    if not df_chitiet_am.empty:
                        ln, sl, tt, vt, st = ghi_sheet_kem_tong(df_chitiet_am, "Chi_Tiet_Am", "C00000", has_vat=True)
                        grand_summary_data.append(["Chi Tiet Hang Tra", "Am", ln, sl, tt, vt, st, "Thuc te sau CK"])
                        
                    if not df_tonghop_banra.empty:
                        ln, sl, tt, vt, st = ghi_sheet_kem_tong(df_tonghop_banra, "Tong_Hop_Ban_Ra", "00B050", has_vat=True)
                        grand_summary_data.append(["Tong Hop Ban Ra", "Duong", ln, sl, tt, vt, st, "Toan bộ du lieu duong"])

                    # Tiến hành thuật toán cắt lát chia nhỏ Hóa Đơn (Batch Size = 1000 dòng) [cite: 73, 93]
                    batch_size = 1000
                    sheet_index = 1
                    for i in range(0, len(df_tonghop_banra), batch_size):
                        df_batch = df_tonghop_banra.iloc[i:i+batch_size].copy()
                        df_batch['STT'] = range(1, len(df_batch) + 1)
                        hd_name = f"HD {sheet_index}"
                        ln, sl, tt, vt, st = ghi_sheet_kem_tong(df_batch, hd_name, "0070C0", has_vat=True)
                        grand_summary_data.append([f"Hoa Don {sheet_index}", "Duong", ln, sl, tt, vt, st, "Tach le 1000 dong"])
                        sheet_index += 1

                    # Tạo báo cáo tài chính "Tong_Ket_Chung" lên trang đầu [cite: 96, 97]
                    df_grand = pd.DataFrame(grand_summary_data, columns=["Ten Sheet / Hang muc", "Loai", "So dong", "So luong", "Truoc Thue", "VAT (5%)", "Sau Thue", "Ghi chu"])
                    df_grand.to_excel(writer, index=False, sheet_name="Tong_Ket_Chung")
                    ws_grand = writer.sheets["Tong_Ket_Chung"]
                    
                    # Công thức tính doanh thu thực tế tuyệt đối sau bù trừ hàng trả [cite: 97]
                    thuong_sl = df_tonghop_banra['Số lượng'].sum() + df_chitiet_am['Số lượng'].sum()
                    thuong_tt = df_tonghop_banra['Thành tiền'].sum() + df_chitiet_am['Thành tiền'].sum()
                    thuong_vat = thuong_tt * 0.05
                    thuong_st = thuong_tt * 1.05
                    
                    r_grand = len(df_grand) + 2
                    ws_grand.cell(row=r_grand, column=1).value = "DOANH THU THUC TE (BAN - TRA)"
                    ws_grand.cell(row=r_grand, column=4).value = thuong_sl
                    ws_grand.cell(row=r_grand, column=5).value = thuong_tt
                    ws_grand.cell(row=r_grand, column=6).value = thuong_vat
                    ws_grand.cell(row=r_grand, column=7).value = thuong_st
                    
                    trang_tri_sheet(ws_grand, "000000", total_row_type="grand")
                    
                    # Sắp xếp lại vị trí đưa sheet tổng kết lên hàng đầu
                    wb = writer.book
                    order = [wb.sheetnames[-1]] + wb.sheetnames[:-1]
                    wb._sheets = [wb._sheets[wb.sheetnames.index(name)] for name in order]

                processed_data = output.getvalue()
                st.success("🎉 PHÂN TÍCH THÀNH CÔNG MASTER PROCESS!")
                
                # --- HIỂN THỊ DANH MỤC TRỰC QUAN TRÊN PHẦN MỀM WEB ---
                tab1, tab2, tab3, tab4 = st.tabs(["📋 Tổng Kết Chung (Báo cáo)", "📈 Tổng Hợp Mã Bán", "📉 Tổng Hợp Mã Trả", "📦 Chi Tiết Hàng Âm"])
                with tab1:
                    st.dataframe(df_grand, use_container_width=True)
                    st.metric(label="Doanh Thu Thuần Thực Tế Của Hệ Thống (Sau Thuế)", value=f"{thuong_st:,.0f} đ")
                with tab2:
                    st.dataframe(df_tonghop_ban, use_container_width=True)
                with tab3:
                    st.dataframe(df_tonghop_tra, use_container_width=True)
                with tab4:
                    st.dataframe(df_chitiet_am, use_container_width=True)

                # Nút kết xuất dữ liệu
                st.download_button(
                    label="📥 TẢI FILE EXCEL MASTER REPORT (FULL SHEETS)",
                    data=processed_data,
                    file_name=f"Master_Report_Sach_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as lỗi_he_thong:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý dữ liệu: {lỗi_he_thong}")
