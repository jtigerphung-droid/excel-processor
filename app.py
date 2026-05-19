import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CẤU HÌNH GIAO DIỆN WEB ---
st.set_page_config(page_title="Hệ thống Phát Hành Sách 2 Giai Đoạn", layout="wide")
st.title("📊 HỆ THỐNG XỬ LÝ DỮ LIỆU PHÁT HÀNH SÁCH")
st.write("Phiên bản Vá Lỗi Khớp Cột (Column Mismatch): Tự động bù cột thiếu hụt và phân tách 2 giai đoạn.")

# --- HÀM TRANG TRÍ EXCEL QUY CHUẨN ---
def trang_tri_sheet(worksheet, tieude_color, has_vat_summary=False, total_row_type="standard"):
    font_tieude = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill_tieude = PatternFill(start_color=tieude_color, end_color=tieude_color, fill_type="solid")
    font_noidung = Font(name="Arial", size=10)
    font_tongket = Font(name="Arial", size=10, bold=True)
    vien_mong = Side(border_style="thin", color="D9D9D9")
    border_o = Border(left=vien_mong, right=vien_mong, top=vien_mong, bottom=vien_mong)
    
    worksheet.row_dimensions[1].height = 25
    for col_idx in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.font = font_tieude
        cell.fill = fill_tieude
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_o

    max_r = worksheet.max_row
    for row_idx in range(2, max_r + 1):
        worksheet.row_dimensions[row_idx].height = 19
        is_total_row = (total_row_type == "grand" and row_idx == max_r) or \
                       (total_row_type == "standard" and has_vat_summary and row_idx >= max_r - 2) or \
                       (total_row_type == "standard" and not has_vat_summary and row_idx == max_r)
        
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = border_o
            if is_total_row:
                cell.font = font_tongket
                if total_row_type == "grand": cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            else:
                cell.font = font_noidung

            if col_idx in [1, 3]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx == 3: cell.number_format = '@'
            elif col_idx == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 4:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"

    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)


# --- ĐIỀU HƯỚNG 2 GIAI ĐOẠN QUA TAB ---
tab_giai_doan_1, tab_giai_doan_2 = st.tabs(["🔄 GIAI ĐOẠN 1: Gộp & Làm Sạch File Thô", "🧮 GIAI ĐOẠN 2: Tính Toán Master Process"])

# ==========================================================================================
# GIAI ĐOẠN 1: GỘP VÀ LÀM SẠCH DỮ LIỆU THÔ (VÁ LỖI THIẾU CỘT VẬT LÝ)
# ==========================================================================================
with tab_giai_doan_1:
    st.header("Bước 1: Gộp Nhiều Sheet & Loại Bỏ 100% Dòng Rác")
    st.info("Tính năng này tự động chuẩn hóa số lượng cột, gộp các sheet và loại bỏ toàn bộ dòng rác, dòng tiêu đề lặp lại.")
    
    file_tho = st.file_uploader("Tải lên file Excel THÔ (File có nhiều sheet hệ thống):", type=["xlsx"], key="file_tho")
    
    if file_tho is not None:
        if st.button("🛠️ THỰC HIỆN GỘP VÀ LÀM SẠCH DỮ LIỆU"):
            try:
                excel_file = pd.ExcelFile(file_tho)
                list_df_sheets = []
                
                for sheet_name in excel_file.sheet_names:
                    # Đọc toàn bộ dữ liệu dưới dạng chuỗi văn bản thô
                    df_raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None).astype(str)
                    if df_raw.empty: continue
                    
                    # Tìm dòng chứa tiêu đề mấu chốt để xác định điểm cắt dữ liệu
                    header_row_idx = None
                    for idx, row in df_raw.iterrows():
                        row_list = row.str.strip().tolist()
                        if any("Mã số" in s or "Tên sách" in s for s in row_list):
                            header_row_idx = idx
                            break
                    
                    # Cắt lấy phần nội dung phía dưới dòng tiêu đề tìm được
                    if header_row_idx is not None:
                        df_data = df_raw.iloc[header_row_idx + 1:].copy()
                    else:
                        df_data = df_raw.iloc[8:].copy()
                    
                    # THUẬT TOÁN BÙ CỘT THÔNG MINH: Đảm bảo bảng luôn có ít nhất 9 cột trước khi gán tên
                    current_cols = df_data.shape[1]
                    if current_cols < 9:
                        for c_add in range(9 - current_cols):
                            df_data[f"custom_col_add_{c_add}"] = "0"  # Thêm cột trống định dạng chuỗi
                    
                    # Trích xuất chuẩn xác 9 cột từ chỉ số 0 đến 8
                    df_data = df_data.iloc[:, 0:9]
                    df_data.columns = ['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']
                    list_df_sheets.append(df_data)
                
                if not list_df_sheets:
                    st.error("Không tìm thấy dữ liệu hợp lệ trong các sheet.")
                    st.stop()

                # Tiến hành gộp dữ liệu diện rộng
                df_combined = pd.concat(list_df_sheets, ignore_index=True)
                
                # Làm sạch khoảng trắng dư thừa
                for col in df_combined.columns:
                    df_combined[col] = df_combined[col].str.strip()
                
                # Loại bỏ các dòng trống hoặc mã trống không hợp lệ
                df_combined = df_combined[df_combined['Mã số'].notna() & (df_combined['Mã số'] != '') & (df_combined['Mã số'] != 'nan') & (df_combined['Mã số'] != 'Mã số')]
                
                # Bộ lọc triệt tiêu các từ khóa dòng tổng kết trung gian và dòng chữ ký kế toán
                tu_khoa_xoa = ['Tổng cộng:', 'Thuế VAT:', 'Thành tiền:', 'Tổng cộng', 'Thuế VAT', 'Phụ trách cung tiêu', 'Người giao hàng', 'Thủ Kho', 'nan', 'STT', 'Tên sách', 'Cộng trước:', 'Cộng trước', 'Người lập']
                df_combined = df_combined[~df_combined['Tên sách'].isin(tu_khoa_xoa)]
                df_combined = df_combined[~df_combined['Mã số'].isin(tu_khoa_xoa)]
                
                # Ép toàn bộ định dạng tiền tệ và số lượng về dạng số xử lý được
                for col in ['Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']:
                    df_combined[col] = pd.to_numeric(df_combined[col].str.replace(',', '').str.replace('.', ''), errors='coerce').fillna(0)
                
                # Loại bỏ dòng không có biến động số lượng phát hành
                df_combined = df_combined[df_combined['Số lượng'] != 0]
                df_combined['STT'] = range(1, len(df_combined) + 1)
                
                # Khởi tạo tệp tin xuất ra cho Giai đoạn 2
                out_sach = io.BytesIO()
                with pd.ExcelWriter(out_sach, engine='openpyxl') as writer:
                    df_combined.to_excel(writer, index=False, sheet_name="Du_Lieu_Sach_100")
                    trang_tri_sheet(writer.sheets["Du_Lieu_Sach_100"], "003366")
                
                st.success("🎉 GIAI ĐOẠN 1 HOÀN THÀNH: ĐÃ GỘP VÀ LÀM SẠCH 100% FILE DỮ LIỆU THÔ!")
                st.dataframe(df_combined, use_container_width=True)
                
                st.download_button(
                    label="📥 TẢI FILE DỮ LIỆU SẠCH (Dùng cho Giai đoạn 2)",
                    data=out_sach.getvalue(),
                    file_name=f"DuLieu_Gop_Va_LamSach_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Lỗi xảy ra ở Giai đoạn 1: {str(e)}")

# ==========================================================================================
# GIAI ĐOẠN 2: TÍNH TOÁN VÀ PHÂN TÍCH MASTER PROCESS (TƯƠNG ĐƯƠNG FILE B3.BAS)
# ==========================================================================================
with tab_giai_doan_2:
    st.header("Bước 2: Phân Tích Danh Mục, Xử Lý Âm Dương & Tách Hóa Đơn")
    st.warning("⚠️ LƯU Ý: Vui lòng nạp file Excel đã được LÀM SẠCH tải ra từ Giai đoạn 1 vào đây để tính toán.")
    
    file_sach = st.file_uploader("Tải lên file Excel ĐÃ LÀM SẠCH:", type=["xlsx"], key="file_sach")
    
    if file_sach is not None:
        if st.button("🧮 KHỞI CHẠY THUẬT TOÁN TÍNH TOÁN MASTER REPORT"):
            try:
                df_master = pd.read_excel(file_sach, sheet_name="Du_Lieu_Sach_100")
                
                df_duong = df_master[df_master['Số lượng'] >= 0].copy()
                df_am = df_master[df_master['Số lượng'] < 0].copy()
                
                # 2.1 Tổng hợp mã hàng bán (Gom theo Mã số, Đơn giá = Giá bìa gốc)
                if not df_duong.empty:
                    df_tonghop_ban = df_duong.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False).agg({'Số lượng': 'sum'})
                    df_tonghop_ban['CK'] = 0
                    df_tonghop_ban['Đơn giá'] = df_tonghop_ban['Giá bìa']
                    df_tonghop_ban['Thành tiền'] = df_tonghop_ban['Số lượng'] * df_tonghop_ban['Đơn giá']
                    df_tonghop_ban = df_tonghop_ban[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                    df_tonghop_ban.insert(0, 'STT', range(1, len(df_tonghop_ban) + 1))
                else:
                    df_tonghop_ban = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                
                # 2.2 Tổng hợp hàng trả (Chuyển đổi dấu âm sang dấu dương - Tính theo Giá bìa)
                if not df_am.empty:
                    df_tonghop_tra = df_am.copy()
                    df_tonghop_tra['Số lượng'] = df_tonghop_tra['Số lượng'].abs()
                    df_tonghop_tra = df_tonghop_tra.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False).agg({'Số lượng': 'sum'})
                    df_tonghop_tra['CK'] = 0
                    df_tonghop_tra['Đơn giá'] = df_tonghop_tra['Giá bìa']
                    df_tonghop_tra['Thành tiền'] = df_tonghop_tra['Số lượng'] * df_tonghop_tra['Đơn giá']
                    df_tonghop_tra = df_tonghop_tra[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                    df_tonghop_tra.insert(0, 'STT', range(1, len(df_tonghop_tra) + 1))
                    
                    # 2.3 Chi tiết hàng trả âm (Giữ nguyên dấu âm và tỷ lệ chiết khấu thực tế)
                    df_chitiet_am = df_am.copy()
                    df_chitiet_am['Thành tiền'] = df_chitiet_am['Số lượng'] * df_chitiet_am['Đơn giá']
                    df_chitiet_am['STT'] = range(1, len(df_chitiet_am) + 1)
                    df_chitiet_am = df_chitiet_am[['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                else:
                    df_tonghop_tra = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                    df_chitiet_am = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                
                # 2.4 Toàn bộ mảng dữ liệu bán ra dương thực tế
                df_tonghop_banra = df_duong.copy()
                if not df_tonghop_banra.empty:
                    df_tonghop_banra['Thành tiền'] = df_tonghop_banra['Số lượng'] * df_tonghop_banra['Đơn giá']
                    df_tonghop_banra['STT'] = range(1, len(df_tonghop_banra) + 1)
                    df_tonghop_banra = df_tonghop_banra[['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                else:
                    df_tonghop_banra = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])

                # --- BƯỚC 3: ĐÓNG GÓI RA FILE EXCEL MULTI-SHEETS ---
                out_report = io.BytesIO()
                with pd.ExcelWriter(out_report, engine='openpyxl') as writer:
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
                        grand_summary_data.append(["Tong Hop Ban Ra", "Duong", ln, sl, tt, vt, st, "Toan bo du lieu duong"])

                    # Phân tách nhỏ hóa đơn (Batch Size = 1000 dòng theo file .bas)
                    batch_size = 1000
                    sheet_index = 1
                    if not df_tonghop_banra.empty:
                        for i in range(0, len(df_tonghop_banra), batch_size):
                            df_batch = df_tonghop_banra.iloc[i:i+batch_size].copy()
                            df_batch['STT'] = range(1, len(df_batch) + 1)
                            hd_name = f"HD {sheet_index}"
                            ln, sl, tt, vt, st = ghi_sheet_kem_tong(df_batch, hd_name, "0070C0", has_vat=True)
                            grand_summary_data.append([f"Hoa Don {sheet_index}", "Duong", ln, sl, tt, vt, st, "Tach le 1000 dong"])
                            sheet_index += 1

                    # Báo cáo tổng quan Dashboard đưa lên trang đầu
                    df_grand = pd.DataFrame(grand_summary_data, columns=["Ten Sheet / Hang muc", "Loai", "So dong", "So luong", "Truoc Thue", "VAT (5%)", "Sau Thue", "Ghi chu"])
                    df_grand.to_excel(writer, index=False, sheet_name="Tong_Ket_Chung")
                    ws_grand = writer.sheets["Tong_Ket_Chung"]
                    
                    thuong_sl = df_tonghop_banra['Số lượng'].sum() + df_chitiet_am['Số lượng'].sum() if not df_chitiet_am.empty else df_tonghop_banra['Số lượng'].sum()
                    thuong_tt = df_tonghop_banra['Thành tiền'].sum() + df_chitiet_am['Thành tiền'].sum() if not df_chitiet_am.empty else df_tonghop_banra['Thành tiền'].sum()
                    thuong_vat = thuong_tt * 0.05
                    thuong_st = thuong_tt * 1.05
                    
                    r_grand = len(df_grand) + 2
                    ws_grand.cell(row=r_grand, column=1).value = "DOANH THU THUC TE (BAN - TRA)"
                    ws_grand.cell(row=r_grand, column=4).value = thuong_sl
                    ws_grand.cell(row=r_grand, column=5).value = thuong_tt
                    ws_grand.cell(row=r_grand, column=6).value = thuong_vat
                    ws_grand.cell(row=r_grand, column=7).value = thuong_st
                    
                    trang_tri_sheet(ws_grand, "000000", total_row_type="grand")
                    
                    wb = writer.book
                    order = [wb.sheetnames[-1]] + wb.sheetnames[:-1]
                    wb._sheets = [wb._sheets[wb.sheetnames.index(name)] for name in order]

                st.success("🎉 PHÂN TÍCH TÍNH TOÁN MASTER REPORT THÀNH CÔNG RỰC RỠ!")
                
                tab_res1, tab_res2, tab_res3, tab_res4 = st.tabs(["📋 Bảng Tổng Kết Doanh Thu", "📈 Tổng Hợp Mã Bán", "📉 Tổng Hợp Mã Trả", "📦 Chi Tiết Hàng Âm"])
                with tab_res1:
                    st.dataframe(df_grand, use_container_width=True)
                    st.metric(label="Doanh Thu Thuần Thực Tế Đối Soát Hệ Thống (Sau Thuế)", value=f"{thuong_st:,.0f} đ")
                with tab_res2:
                    st.dataframe(df_tonghop_ban, use_container_width=True)
                with tab_res3:
                    st.dataframe(df_tonghop_tra, use_container_width=True)
                with tab_res4:
                    st.dataframe(df_chitiet_am, use_container_width=True)

                st.download_button(
                    label="📥 TẢI FILE EXCEL MASTER REPORT CUỐI CÙNG",
                    data=out_report.getvalue(),
                    file_name=f"Master_Report_Final_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Lỗi xảy ra ở Giai đoạn 2: {str(e)}")
