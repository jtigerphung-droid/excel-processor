import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CẤU HÌNH GIAO DIỆN HỆ THỐNG ---
st.set_page_config(page_title="Hệ thống Phát Hành Sách 2 Giai Đoạn", layout="wide")
st.title("📊 HỆ THỐNG XỬ LÝ DỮ LIỆU PHÁT HÀNH SÁCH")
st.write("Phiên bản chuẩn hóa nghiệp vụ Giai đoạn 2 - Loại bỏ tư duy VBA, tối ưu thuật toán Python độc lập.")

# --- HÀM TRANG TRÍ EXCEL THEO QUY CHUẨN KẾ TOÁN ---
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


# --- ĐIỀU HƯỚNG TÍNH NĂNG QUA TAB ---
tab_giai_doan_1, tab_giai_doan_2 = st.tabs(["🔄 GIAI ĐOẠN 1: Gộp & Làm Sạch (Bản Ổn Định)", "🧮 GIAI ĐOẠN 2: Tính Toán Master Process"])

# ==========================================================================================
# GIAI ĐOẠN 1: GIỮ NGUYÊN BẢN THUẬT TOÁN GỐC ỔN ĐỊNH 100%
# ==========================================================================================
with tab_giai_doan_1:
    st.header("Bước 1: Gộp Nhiều Sheet & Làm Sạch Diện Rộng")
    st.info("Sử dụng thuật toán dò quét gốc ổn định để trích xuất file dữ liệu sạch.")
    
    uploaded_file = st.file_uploader("Kéo thả file Excel thô của hệ thống vào đây:", type=["xlsx"], key="file_tho_goc")

    if uploaded_file is not None:
        if st.button("🚀 THỰC HIỆN GỘP VÀ LÀM SẠCH PHIÊN BẢN CHUẨN"):
            try:
                excel_file = pd.ExcelFile(uploaded_file)
                all_cleaned_rows = []
                
                for sheet_name in excel_file.sheet_names:
                    df_raw_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                    if df_raw_sheet.empty: continue
                    
                    for idx, row in df_raw_sheet.iterrows():
                        row_vals = row.fillna("").astype(str).str.strip().tolist()
                        
                        if not row_vals or len(row_vals) < 7: continue
                        if any(k in "".join(row_vals) for k in ["Tổng cộng", "Thuế VAT", "Thành tiền", "Người lập", "Thủ kho", "Giám đốc"]): continue
                        if "Mã số" in row_vals or "Tên sách" in row_vals or "STT" in row_vals: continue
                        
                        idx_ma = -1
                        idx_ten = -1
                        
                        for c_i, val in enumerate(row_vals):
                            if val and not " " in val and len(val) >= 4 and idx_ma == -1:
                                if re.match(r'^[A-Za-z0-9\-_.]+$', val): idx_ma = c_i
                            elif val and " " in val and len(val) > 10 and idx_ten == -1:
                                idx_ten = c_i
                        
                        if idx_ma == -1 or idx_ten == -1:
                            if row_vals[0].isdigit() or (row_vals[2] != "" and len(row_vals[2]) >= 4):
                                idx_ma = 2
                                idx_ten = 1
                            else:
                                continue
                                
                        try:
                            base_idx = min(idx_ten, idx_ma)
                            if base_idx > 0 and row_vals[base_idx-1].isdigit():
                                stt_val = row_vals[base_idx-1]
                            else:
                                stt_val = ""
                                
                            ten_sach = row_vals[idx_ten]
                            ma_so = row_vals[idx_ma]
                            dvt = row_vals[idx_ma + 1] if (idx_ma + 1) < len(row_vals) else "Cuốn"
                            
                            num_fields = []
                            for v in row_vals[idx_ma + 2:]:
                                if v == "" or v == "nan": num_fields.append(0)
                                else:
                                    v_clean = v.replace(",", "").replace(".", "").strip()
                                    if v_clean.replace("-", "").isdigit(): num_fields.append(float(v_clean))
                                    else: num_fields.append(0)
                                    
                            while len(num_fields) < 5: num_fields.append(0)
                            
                            gia_bia = num_fields[0]
                            ck = num_fields[1]
                            sl = num_fields[2]
                            don_gia = num_fields[3]
                            thanh_tien = num_fields[4]
                            
                            if sl == 0 or ma_so == "" or ma_so == "nan": continue
                            
                            all_cleaned_rows.append([
                                stt_val, ten_sach, ma_so, dvt, gia_bia, ck, sl, don_gia, thanh_tien
                            ])
                        except Exception:
                            continue

                if not all_cleaned_rows:
                    st.error("❌ LỖI: Không trích xuất được dòng dữ liệu sách hợp lệ nào.")
                    st.stop()
                    
                df_master = pd.DataFrame(all_cleaned_rows, columns=['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                df_master.insert(0, 'STT', range(1, len(df_master) + 1))
                
                out_sach = io.BytesIO()
                with pd.ExcelWriter(out_sach, engine='openpyxl') as writer:
                    df_master.to_excel(writer, index=False, sheet_name="Du_Lieu_Sach_100")
                    trang_tri_sheet(writer.sheets["Du_Lieu_Sach_100"], "003366")
                    
                st.success("🎉 GIAI ĐOẠN 1 HOÀN THÀNH: PHỤC HỒI KẾT QUẢ CHUẨN THÀNH CÔNG!")
                st.dataframe(df_master, use_container_width=True)
                
                st.download_button(
                    label="📥 TẢI FILE DỮ LIỆU SẠCH (Dùng cho Giai đoạn 2)",
                    data=out_sach.getvalue(),
                    file_name=f"DuLieu_Gop_Va_LamSach_Chuan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Lỗi hệ thống phục hồi G1: {str(e)}")

# ==========================================================================================
# GIAI ĐOẠN 2: THUẬT TOÁN TÍNH TOÁN ĐỘC LẬP CHUẨN HÓA NGHIỆP VỤ KẾ TOÁN
# ==========================================================================================
with tab_giai_doan_2:
    st.header("Bước 2: Phân Tích Danh Mục, Xử Lý Âm Dương & Tách Hóa Đơn")
    st.warning("⚠️ LƯU Ý: Vui lòng nạp file Excel thu được từ Giai đoạn 1 vào đây.")
    
    file_sach = st.file_uploader("Tải lên file Excel ĐÃ LÀM SẠCH CHUẨN:", type=["xlsx"], key="file_sach_nghiep_vu")
    
    if file_sach is not None:
        if st.button("🧮 KHỞI CHẠY THUẬT TOÁN TÍNH TOÁN MASTER REPORT"):
            try:
                # Đọc dữ liệu từ sheet "Du_Lieu_Sach_100"
                df_input = pd.read_excel(file_sach, sheet_name="Du_Lieu_Sach_100")
                
                # ÉP CHUẨN KIỂU DỮ LIỆU ĐỂ TRÁNH LỖI KEYERROR / ĐỊNH DẠNG CHUỖI
                df_input['Mã số'] = df_input['Mã số'].astype(str).str.strip()
                df_input['Tên sách'] = df_input['Tên sách'].astype(str).str.strip()
                df_input['ĐVT'] = df_input['ĐVT'].astype(str).str.strip()
                
                for col_num in ['Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']:
                    df_input[col_num] = pd.to_numeric(df_input[col_num], errors='coerce').fillna(0)
                
                # PHÂN TÁCH MẢNG DỮ LIỆU ĐỂ XỬ LÝ
                df_duong_goc = df_input[df_input['Số lượng'] > 0].copy()
                df_am_goc = df_input[df_input['Số lượng'] < 0].copy()
                
                # --- THUẬT TOÁN 1: TỔNG HỢP MÃ HÀNG BÁN (Gom nhóm mảng dương, Đơn giá = Giá bìa) ---
                if not df_duong_goc.empty:
                    df_th_ban = df_duong_goc.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False)['Số lượng'].sum()
                    df_th_ban['CK'] = 0
                    df_th_ban['Đơn giá'] = df_th_ban['Giá bìa']
                    df_th_ban['Thành tiền'] = df_th_ban['Số lượng'] * df_th_ban['Đơn giá']
                    df_th_ban = df_th_ban[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                    df_th_ban.insert(0, 'STT', range(1, len(df_th_ban) + 1))
                else:
                    df_th_ban = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                
                # --- THUẬT TOÁN 2: TỔNG HỢP MÃ HÀNG TRẢ (Gom nhóm mảng âm, đảo dấu thành Dương, Đơn giá = Giá bìa) ---
                if not df_am_goc.empty:
                    df_am_dao_dau = df_am_goc.copy()
                    df_am_dao_dau['Số lượng'] = df_am_dao_dau['Số lượng'].abs()
                    
                    df_th_tra = df_am_dao_dau.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False)['Số lượng'].sum()
                    df_th_tra['CK'] = 0
                    df_th_tra['Đơn giá'] = df_th_tra['Giá bìa']
                    df_th_tra['Thành tiền'] = df_th_tra['Số lượng'] * df_th_tra['Đơn giá']
                    df_th_tra = df_th_tra[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                    df_th_tra.insert(0, 'STT', range(1, len(df_th_tra) + 1))
                    
                    # --- THUẬT TOÁN 3: CHI TIẾT HÀNG TRẢ ÂM (Giữ nguyên dấu âm và chiết khấu thực tế) ---
                    df_ct_am = df_am_goc.copy()
                    df_ct_am['Thành tiền'] = df_ct_am['Số lượng'] * df_ct_am['Đơn giá']
                    df_ct_am = df_ct_am[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                    df_ct_am.insert(0, 'STT', range(1, len(df_ct_am) + 1))
                else:
                    df_th_tra = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                    df_ct_am = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                
                # --- THUẬT TOÁN 4: CHI TIẾT TOÀN BỘ MẢNG BÁN RA DƯƠNG ---
                df_th_banra = df_duong_goc.copy()
                if not df_th_banra.empty:
                    df_th_banra['Thành tiền'] = df_th_banra['Số lượng'] * df_th_banra['Đơn giá']
                    df_th_banra = df_th_banra[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                    df_th_banra.insert(0, 'STT', range(1, len(df_th_banra) + 1))
                else:
                    df_th_banra = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])

                # --- ĐÓNG GÓI DỮ LIỆU VÀ XUẤT FILE EXCEL MULTI-SHEET ---
                out_report = io.BytesIO()
                with pd.ExcelWriter(out_report, engine='openpyxl') as writer:
                    grand_summary_list = []
                    
                    def xu_ly_ghi_sheet(df_sheet, sheet_name, color_theme, has_vat=False):
                        df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)
                        ws = writer.sheets[sheet_name]
                        
                        total_row = len(df_sheet) + 2
                        sum_qty = df_sheet['Số lượng'].sum()
                        sum_amount = df_sheet['Thành tiền'].sum()
                        
                        # Điền dòng tổng kết
                        ws.cell(row=total_row, column=6).value = "Tổng cộng:"
                        ws.cell(row=total_row, column=7).value = sum_qty
                        ws.cell(row=total_row, column=9).value = sum_amount
                        
                        v_tax, v_after_tax = 0.0, 0.0
                        if has_vat:
                            v_tax = sum_amount * 0.05
                            v_after_tax = sum_amount * 1.05
                            ws.cell(row=total_row + 1, column=8).value = "VAT 5%:"
                            ws.cell(row=total_row + 1, column=9).value = v_tax
                            ws.cell(row=total_row + 2, column=8).value = "Sau Thuế:"
                            ws.cell(row=total_row + 2, column=9).value = v_after_tax
                            
                        trang_tri_sheet(ws, color_theme, has_vat_summary=has_vat)
                        return len(df_sheet), sum_qty, sum_amount, v_tax, v_after_tax

                    # Ghi các sheet tổng hợp cơ bản
                    if not df_th_ban.empty:
                        l, q, a, _, _ = xu_ly_ghi_sheet(df_th_ban, "Tong_Hop_Ma_Hang_Ban", "003366")
                        grand_summary_list.append(["Tong Hop Ma Hang Ban", "Duong", l, q, a, 0, a, "Theo Giá Bìa Gốc"])
                        
                    if not df_th_tra.empty:
                        l, q, a, _, _ = xu_ly_ghi_sheet(df_th_tra, "Tong_Hop_Hang_Tra", "FF6600")
                        grand_summary_list.append(["Tong Hop Hang Tra", "Duong (Đảo dấu)", l, q, a, 0, a, "Theo Giá Bìa Gốc"])
                        
                    if not df_ct_am.empty:
                        l, q, a, vt, st = xu_ly_ghi_sheet(df_ct_am, "Chi_Tiet_Am", "C00000", has_vat=True)
                        grand_summary_list.append(["Chi Tiet Hang Tra", "Am", l, q, a, vt, st, "Thực tế sau Chiết Khấu"])
                        
                    if not df_th_banra.empty:
                        l, q, a, vt, st = xu_ly_ghi_sheet(df_th_banra, "Tong_Hop_Ban_Ra", "00B050", has_vat=True)
                        grand_summary_list.append(["Tong Hop Ban Ra", "Duong", l, q, a, vt, st, "Thực tế sau Chiết Khấu"])

                    # --- THUẬT TOÁN 5: TỰ ĐỘNG CHẶT NHỎ HÓA ĐƠN 1000 DÒNG ---
                    batch_size = 1000
                    hd_index = 1
                    if not df_th_banra.empty:
                        for start_idx in range(0, len(df_th_banra), batch_size):
                            df_batch = df_th_banra.iloc[start_idx : start_idx + batch_size].copy()
                            df_batch['STT'] = range(1, len(df_batch) + 1)
                            
                            sheet_hd_name = f"HD {hd_index}"
                            l, q, a, vt, st = xu_ly_ghi_sheet(df_batch, sheet_hd_name, "0070C0", has_vat=True)
                            grand_summary_list.append([f"Hoa Don {hd_index}", "Duong Split", l, q, a, vt, st, "Tách lẻ quy chuẩn 1000 dòng"])
                            hd_index += 1

                    # --- THUẬT TOÁN 6: XÂY DỰNG TRANG DASHBOARD ĐỐI SOÁT TỔNG ---
                    df_grand_final = pd.DataFrame(grand_summary_list, columns=["Ten Sheet / Hang muc", "Loai", "So dong", "So luong", "Truoc Thue", "VAT (5%)", "Sau Thue", "Ghi chu"])
                    df_grand_final.to_excel(writer, index=False, sheet_name="Tong_Ket_Chung")
                    ws_grand = writer.sheets["Tong_Ket_Chung"]
                    
                    # Tính toán Doanh thu thuần thực tế sau khi bù trừ xuất - trả toàn cục
                    net_qty = df_th_banra['Số lượng'].sum() + df_ct_am['Số lượng'].sum()
                    net_amount = df_th_banra['Thành tiền'].sum() + df_ct_am['Thành tiền'].sum()
                    net_vat = net_amount * 0.05
                    net_after_tax = net_amount * 1.05
                    
                    r_grand_total = len(df_grand_final) + 2
                    ws_grand.cell(row=r_grand_total, column=1).value = "DOANH THU THUC TE (BAN - TRA)"
                    ws_grand.cell(row=r_grand_total, column=4).value = net_qty
                    ws_grand.cell(row=r_grand_total, column=5).value = net_amount
                    ws_grand.cell(row=r_grand_total, column=6).value = net_vat
                    ws_grand.cell(row=r_grand_total, column=7).value = net_after_tax
                    
                    trang_tri_sheet(ws_grand, "000000", total_row_type="grand")
                    
                    # Đẩy sheet tổng kết doanh thu lên vị trí đầu tiên trong file Excel
                    wb = writer.book
                    sheets_order = [wb.sheetnames[-1]] + wb.sheetnames[:-1]
                    wb._sheets = [wb._sheets[wb.sheetnames.index(name)] for name in sheets_order]

                st.success("🎉 PHÂN TÍCH TÍNH TOÁN MASTER REPORT THÀNH CÔNG!")
                
                # Hiển thị trực quan dữ liệu lên giao diện Streamlit
                t1, t2, t3, t4 = st.tabs(["📋 Bảng Doanh Thu Thuần", "📈 Tổng Hợp Mã Bán", "📉 Tổng Hợp Mã Trả", "📦 Chi Tiết Hàng Trả"])
                with t1:
                    st.dataframe(df_grand_final, use_container_width=True)
                    st.metric(label="Doanh Thu Thuần Thực Tế Cuối Cùng (Sau Thuế 5%)", value=f"{net_after_tax:,.0f} đ")
                with t2:
                    st.dataframe(df_th_ban, use_container_width=True)
                with t3:
                    st.dataframe(df_th_tra, use_container_width=True)
                with t4:
                    st.dataframe(df_ct_am, use_container_width=True)

                st.download_button(
                    label="📥 TẢI FILE EXCEL MASTER REPORT CUỐI CÙNG",
                    data=out_report.getvalue(),
                    file_name=f"Master_Report_Final_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Lỗi xảy ra ở thuật toán xử lý Giai đoạn 2: {str(e)}")
