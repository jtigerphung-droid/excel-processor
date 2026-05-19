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
st.write("Phiên bản sửa lỗi khởi động - Tối ưu hóa luồng biên dịch Streamlit.")

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
            elif col_idx == 6:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "0.00%" if total_row_type != "grand" else "#,##0"
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"

    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)


# --- ĐIỀU HƯỚNG TÍNH NĂNG QUA TAB ---
tab_giai_doan_1, tab_giai_doan_2 = st.tabs(["🔄 GIAI ĐOẠN 1: Gộp & Làm Sạch (Bảo Lưu)", "🧮 GIAI ĐOẠN 2: Tính Toán Phân Tách Thuộc Tính"])

# ==========================================================================================
# GIAI ĐOẠN 1: BẢO LƯU NGUYÊN BẢN 100% CODE GỐC CHẠY HOÀN CHỈNH CỦA USER
# ==========================================================================================
with tab_giai_doan_1:
    st.header("Bước 1: Gộp Nhiều Sheet & Làm Sạch Diện Rộng")
    st.info("Mã nguồn phần này được bảo lưu nguyên vẹn theo cấu trúc chạy ổn định trước đó.")
    
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
                    
                df_master = pd.DataFrame(all_cleaned_rows, columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                df_master['STT'] = range(1, len(df_master) + 1)
                
                out_sach = io.BytesIO()
                with pd.ExcelWriter(out_sach, engine='openpyxl') as writer:
                    df_master.to_excel(writer, index=False, sheet_name="Du_Lieu_Sach_100")
                    trang_tri_sheet(writer.sheets["Du_Lieu_Sach_100"], "003366")
                    
                st.success("🎉 GIAI ĐOẠN 1 HOÀN THÀNH: XUẤT FILE 9 CỘT CHUẨN!")
                st.dataframe(df_master, use_container_width=True)
                
                st.download_button(
                    label="📥 TẢI FILE DỮ LIỆU SẠCH (Dùng cho Giai đoạn 2)",
                    data=out_sach.getvalue(),
                    file_name=f"DuLieu_Gop_Va_LamSach_Chuan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Lỗi hệ thống G1: {str(e)}")

# ==========================================================================================
# GIAI ĐOẠN 2: THUẬT TOÁN TÍNH TOÁN PHÂN TÁCH THEO THUỘC TÍNH NGHIỆP VỤ MỚI
# ==========================================================================================
with tab_giai_doan_2:
    st.header("Bước 2: Phân Tích Thuộc Tính & Tách Hóa Đơn Quy Chuẩn")
    st.warning("⚠️ LƯU Ý: Vui lòng nạp chính xác file Excel kết quả từ Giai đoạn 1 vào đây.")
    
    file_sach = st.file_uploader("Tải lên file Excel ĐÃ LÀM SẠCH CHUẨN:", type=["xlsx"], key="file_sach_nghiep_vu_moi")
    
    if file_sach is not None:
        if st.button("🧮 KHỞI CHẠY THUẬT TOÁN PHÂN TÁCH MASTER REPORT"):
            try:
                # Đọc bỏ qua dòng chữ tiêu đề thô để tránh lệch ma trận, lấy trực tiếp mảng 9 cột cố định
                df_raw_g2 = pd.read_excel(file_sach, sheet_name="Du_Lieu_Sach_100", skiprows=1, header=None)
                
                # Ép chặt tiêu đề cố định theo đúng vị trí từ 0 đến 8
                df_input = df_raw_g2.iloc[:, 0:9].copy()
                df_input.columns = ['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']
                
                # Ép kiểu dữ liệu số và chuỗi tường minh chống khoảng trắng ẩn
                df_input['Mã số'] = df_input['Mã số'].astype(str).str.strip()
                df_input['Tên sách'] = df_input['Tên sách'].astype(str).str.strip()
                df_input['ĐVT'] = df_input['ĐVT'].astype(str).str.strip()
                for col_name in ['Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']:
                    df_input[col_name] = pd.to_numeric(df_input[col_name], errors='coerce').fillna(0)
                
                # TÁCH HAI MẢNG ĐỘC LẬP ÂM / DƯƠNG
                df_duong = df_input[df_input['Số lượng'] > 0].copy()
                df_am = df_input[df_input['Số lượng'] < 0].copy()
                
                # ----------------------------------------------------------------------------------
                # THUẬT TOÁN S2: SHEET TỔNG HỢP HÀNG BÁN THEO GIÁ BÌA
                # ----------------------------------------------------------------------------------
                if not df_duong.empty:
                    df_th_gia_bia = df_duong.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa'], as_index=False)['Số lượng'].sum()
                    df_th_gia_bia['CK'] = 0.0
                    df_th_gia_bia['Đơn giá'] = df_th_gia_bia['Giá bìa']
                    df_th_gia_bia['Thành tiền'] = df_th_gia_bia['Số lượng'] * df_th_gia_bia['Đơn giá']
                    df_th_gia_bia = df_th_gia_bia[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                    df_th_gia_bia.insert(0, 'STT', range(1, len(df_th_gia_bia) + 1))
                else:
                    df_th_gia_bia = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                
                # ----------------------------------------------------------------------------------
                # THUẬT TOÁN S3: SHEET TỔNG HỢP HÀNG BÁN THEO CHIẾT KHẤU
                # ----------------------------------------------------------------------------------
                if not df_duong.empty:
                    df_th_chiet_khau = df_duong.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa', 'CK'], as_index=False)['Số lượng'].sum()
                    df_th_chiet_khau['Đơn giá'] = df_th_chiet_khau.apply(
                        lambda r: r['Giá bìa'] * (1 - r['CK'] if r['CK'] <= 1 else 1 - r['CK']/100), axis=1
                    )
                    df_th_chiet_khau['Thành tiền'] = df_th_chiet_khau['Số lượng'] * df_th_chiet_khau['Đơn giá']
                    df_th_chiet_khau = df_th_chiet_khau[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                    df_th_chiet_khau.insert(0, 'STT', range(1, len(df_th_chiet_khau) + 1))
                else:
                    df_th_chiet_khau = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])
                
                # ----------------------------------------------------------------------------------
                # THUẬT TOÁN S4: SHEET TỔNG HỢP HÀNG TRẢ (Số lượng < 0, đảo dấu thành Dương)
                # ----------------------------------------------------------------------------------
                if not df_am.empty:
                    df_am_duong_hoa = df_am.copy()
                    df_am_duong_hoa['Số lượng'] = df_am_duong_hoa['Số lượng'].abs()
                    
                    df_th_hang_tra = df_am_duong_hoa.groupby(['Mã số', 'Tên sách', 'ĐVT', 'Giá bìa', 'CK'], as_index=False)['Số lượng'].sum()
                    df_th_hang_tra['Đơn giá'] = df_th_hang_tra.apply(
                        lambda r: r['Giá bìa'] * (1 - r['CK'] if r['CK'] <= 1 else 1 - r['CK']/100), axis=1
                    )
                    df_th_hang_tra['Thành tiền'] = df_th_hang_tra['Số lượng'] * df_th_hang_tra['Đơn giá']
                    df_th_hang_tra = df_th_hang_tra[['Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                    df_th_hang_tra.insert(0, 'STT', range(1, len(df_th_hang_tra) + 1))
                else:
                    df_th_hang_tra = pd.DataFrame(columns=['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền'])

                # --- ĐÓNG GÓI RA FILE EXCEL ĐA SHEET VÀ CONFIG MÀU SẮC RIÊNG BIỆT ---
                out_report = io.BytesIO()
                with pd.ExcelWriter(out_report, engine='openpyxl') as writer:
                    grand_summary_list = []
                    
                    def ghi_sheet_quy_trinh(df_sheet, name, color, has_vat=False):
                        df_sheet.to_excel(writer, index=False, sheet_name=name)
                        ws = writer.sheets[name]
                        r_idx = len(df_sheet) + 2
                        
                        s_sl = df_sheet['Số lượng'].sum()
                        s_tt = df_sheet['Thành tiền'].sum()
                        
                        ws.cell(row=r_idx, column=6).value = "Tổng cộng:"
                        ws.cell(row=r_idx, column=7).value = s_sl
                        ws.cell(row=r_idx, column=9).value = s_tt
                        
                        v_tax, v_after = 0.0, 0.0
                        if has_vat:
                            v_tax = s_tt * 0.05
                            v_after = s_tt * 1.05
                            ws.cell(row=r_idx+1, column=8).value = "VAT 5%:"
                            ws.cell(row=r_idx+1, column=9).value = v_tax
                            ws.cell(row=r_idx+2, column=8).value = "Sau Thuế:"
                            ws.cell(row=r_idx+2, column=9).value = v_after
                            
                        trang_tri_sheet(ws, color, has_vat_summary=has_vat)
                        return len(df_sheet), s_sl, s_tt, v_tax, v_after

                    # S2. TH_Hang_Ban_Gia_Bia (Màu Xanh Dương Đậm)
                    if not df_th_gia_bia.empty:
                        l, q, a, _, _ = ghi_sheet_quy_trinh(df_th_gia_bia, "TH_Hang_Ban_Gia_Bia", "002060", has_vat=False)
                        grand_summary_list.append(["TH Hàng Bán Theo Giá Bìa", "Dương", l, q, a, 0, a, "Tính theo Giá Bìa gốc"])

                    # S3. TH_Hang_Ban_Chiet_Khau (Màu Xanh Lá Cây)
                    if not df_th_chiet_khau.empty:
                        l, q, a, vt, st = ghi_sheet_quy_trinh(df_th_chiet_khau, "TH_Hang_Ban_Chiet_Khau", "339933", has_vat=True)
                        grand_summary_list.append(["TH Hàng Bán Chiết Khấu", "Dương", l, q, a, vt, st, "Giá thực tế sau CK"])

                    # S4. Tong_Hop_Hang_Tra (Màu Đỏ Đô)
                    if not df_th_hang_tra.empty:
                        l, q, a, vt, st = ghi_sheet_quy_trinh(df_th_hang_tra, "Tong_Hop_Hang_Tra", "C00000", has_vat=True)
                        grand_summary_list.append(["Tổng Hợp Hàng Trả", "Dương (Đảo dấu)", l, q, a, vt, st, "Hàng trả quy đổi dương"])

                    # S5. Tách hóa đơn từ sheet Chiết Khấu gốc (Màu Xanh Ngọc)
                    batch_size = 1000
                    hd_count = 1
                    if not df_th_chiet_khau.empty:
                        for start_pos in range(0, len(df_th_chiet_khau), batch_size):
                            df_batch = df_th_chiet_khau.iloc[start_pos : start_pos + batch_size].copy()
                            df_batch['STT'] = range(1, len(df_batch) + 1)
                            
                            hd_name = f"HD {hd_count}"
                            l, q, a, vt, st = ghi_sheet_quy_trinh(df_batch, hd_name, "008080", has_vat=True)
                            grand_summary_list.append([f"Hóa Đơn {hd_count}", "Phân tách", l, q, a, vt, st, f"Tách từ dòng {start_pos+1}"])
                            hd_count += 1

                    # S1. THÀNH PHẦN SHEET TỔNG HỢP THÔNG TIN (Màu Đen Kế Toán) - SỬA LỖI TYPO TẠI ĐÂY
                    df_grand_final = pd.DataFrame(grand_summary_list, columns=[
                        "Tên Sheet / Hạng mục", "Loại thuộc tính", "Số dòng", "Số lượng", "Trước Thuế", "VAT (5%)", "Sau Thuế", "Ghi chú"
                    ])
                    df_grand_final.to_excel(writer, index=False, sheet_name="Tong_Ket_Chung")
                    ws_grand = writer.sheets["Tong_Ket_Chung"]
                    
                    ban_tt = df_th_chiet_khau['Thành tiền'].sum() if not df_th_chiet_khau.empty else 0
                    tra_tt = df_th_hang_tra['Thành tiền'].sum() if not df_th_hang_tra.empty else 0
                    net_amount = ban_tt - tra_tt  # Đã sửa lỗi chính tả 'tra' thành 'tra_tt' ở đây
                    
                    ban_sl = df_th_chiet_khau['Số lượng'].sum() if not df_th_chiet_khau.empty else 0
                    tra_sl = df_th_hang_tra['Số lượng'].sum() if not df_th_hang_tra.empty else 0
                    net_qty = ban_sl - tra_sl
                    
                    net_vat = net_amount * 0.05
                    net_after_tax = net_amount * 1.05
                    
                    r_total = len(df_grand_final) + 2
                    ws_grand.cell(row=r_total, column=1).value = "DOANH THU THỰC TẾ (BÁN - TRẢ)"
                    ws_grand.cell(row=r_total, column=2).value = "Bù trừ dòng"
                    ws_grand.cell(row=r_total, column=4).value = net_qty
                    ws_grand.cell(row=r_total, column=5).value = net_amount
                    ws_grand.cell(row=r_total, column=6).value = net_vat
                    ws_grand.cell(row=r_total, column=7).value = net_after_tax
                    
                    trang_tri_sheet(ws_grand, "1F1F1F", total_row_type="grand")
                    
                    # Đưa sheet tổng hợp lên đầu tiên
                    wb = writer.book
                    sheets_order = [wb.sheetnames[-1]] + wb.sheetnames[:-1]
                    wb._sheets = [wb._sheets[wb.sheetnames.index(name)] for name in sheets_order]

                st.success("🎉 ỨNG DỤNG ĐÃ KHỞI ĐỘNG VÀ SỬA LỖI THÀNH CÔNG!")
                
                # Tab hiển thị nhanh kết quả trên Web
                t_total, t_bia, t_ck, t_tra = st.tabs(["📋 Tổng Kết Chung", "📈 Theo Giá Bìa", "📉 Theo Chiết Khấu", "📦 Hàng Trả"])
                with t_total:
                    st.dataframe(df_grand_final, use_container_width=True)
                    st.metric(label="Doanh Thu Thuần Thực Tế Đối Soát Hệ Thống (Sau Thuế)", value=f"{net_after_tax:,.0f} đ")
                with t_bia:
                    st.dataframe(df_th_gia_bia, use_container_width=True)
                with t_ck:
                    st.dataframe(df_th_chiet_khau, use_container_width=True)
                with t_tra:
                    st.dataframe(df_th_hang_tra, use_container_width=True)

                st.download_button(
                    label="📥 TẢI FILE MASTER REPORT HOÀN CHỈNH (ĐA SHEET MÀU SẮC)",
                    data=out_report.getvalue(),
                    file_name=f"Master_Report_Properties_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Lỗi nghiêm trọng xảy ra ở Giai đoạn 2: {str(e)}")
