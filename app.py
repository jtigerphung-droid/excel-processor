import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import zipfile

# --- CẤU HÌNH GIAO DIỆN HỆ THỐNG ĐẬM CHẤT DASHBOARD ---
st.set_page_config(
    page_title="Hệ thống xử lý dữ liệu 3 Giai Đoạn", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Phong cách hóa tiêu đề ứng dụng trực quan hơn
st.markdown("""
    <div style="background-color: #0f172a; padding: 20px; border-radius: 10px; margin-bottom: 25px; border-left: 5px solid #0284c7;">
        <h1 style="color: #ffffff; margin: 0; font-size: 28px; font-family: 'Arial';">📊 HỆ THỐNG XỬ LÝ DỮ LIỆU</h1>
        <p style="color: #94a3b8; margin: 5px 0 0 0; font-size: 14px;">Phiên bản V4 Pro Interface — Tối ưu hóa trải nghiệm Dashboard Kế toán & Phân tách Doanh thu PAB21</p>
    </div>
""", unsafe_allow_html=True)

# --- HÀM TRANG TRÍ EXCEL THEO QUY CHUẨN KẾ TOÁN (GIỮ NGUYÊN 100% CỐT LÕI) ---
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
                if total_row_type == "grand":
                    cell.number_format = "#,##0"
                else:
                    cell.number_format = "0.00"
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"

    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)


# --- ĐIỀU HƯỚNG TÍNH NĂNG QUA TAB ---
tab_giai_doan_1, tab_giai_doan_2, tab_giai_doan_3 = st.tabs([
    "📥 GIAI ĐOẠN 1: Gộp & Làm Sạch Dữ Liệu", 
    "🧮 GIAI ĐOẠN 2: Tính Toán Phân Tách PAB21",
    "📝 GIAI ĐOẠN 3: Đóng Gói File Hóa Đơn Mẫu"
])

# ==========================================================================================
# GIAI ĐOẠN 1: GỘP & LÀM SẠCH DỮ LIỆU THÔ
# ==========================================================================================
with tab_giai_doan_1:
    st.markdown("### 🗂️ Quy trình xử lý cấu trúc dữ liệu thô diện rộng")
    st.info("💡 Hệ thống sẽ tự động quét qua toàn bộ các Sheet, bóc tách cấu trúc, bỏ dòng thừa/dòng ký duyệt để gom về một bảng Master duy nhất.")
    
    uploaded_file = st.file_uploader("Chọn file Excel thô xuất từ hệ thống nội bộ (.xlsx):", type=["xlsx"], key="g1_raw")

    if uploaded_file is not None:
        st.markdown("---")
        if st.button("🚀 KHỞI CHẠY LÀM SẠCH VÀ CHUẨN HÓA DỮ LIỆU", use_container_width=True):
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
                
                st.markdown("#### ✨ KẾT QUẢ ĐÃ TRÍCH XUẤT THÀNH CÔNG")
                c1, c2 = st.columns(2)
                with c1:
                    st.metric(label="Tổng số hạng mục sách tìm thấy", value=f"{len(df_master)} dòng")
                with c2:
                    st.metric(label="Tổng sản lượng sách gộp", value=f"{int(df_master['Số lượng'].sum()):,} cuốn")
                
                st.dataframe(df_master, use_container_width=True)
                
                st.markdown('<div style="margin-top:15px;"></div>', unsafe_allow_html=True)
                st.download_button(
                    label="📥 TẢI XUỐNG TỆP TIN DỮ LIỆU SẠCH (.XLSX)",
                    data=out_sach.getvalue(),
                    file_name=f"DuLieu_Gop_Va_LamSach_Chuan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Lỗi hệ thống G1: {str(e)}")

# ==========================================================================================
# GIAI ĐOẠN 2: TÍNH TOÁN PHÂN TÁCH PAB21 (CÓ THÊM CARD SAU THUẾ & ĐỔI SANG 'CUỐN')
# ==========================================================================================
with tab_giai_doan_2:
    st.markdown("### 🏛️ Phân tách ma trận thuộc tính doanh thu PAB21 & Tô màu định dạng")
    st.info("💡 Hệ thống tách riêng biệt luồng hàng bán (Dương) và hàng trả (Âm), tự động gom nhóm theo chiết khấu riêng biệt và chia nhỏ thành các gói HĐ giới hạn tối đa 1000 dòng.")
    
    file_sach = st.file_uploader("Tải lên file Excel ĐÃ LÀM SẠCH CHUẨN từ Giai đoạn 1:", type=["xlsx"], key="g2_cleaned")
    
    if file_sach is not None:
        st.markdown("---")
        if st.button("🧮 KÍCH HOẠT THUẬT TOÁN TÍNH TOÁN & THIẾT LẬP REPORT", use_container_width=True):
            try:
                wb_in = openpyxl.load_workbook(file_sach, data_only=True)
                ws_in = wb_in["Du_Lieu_Sach_100"]
                
                mảng_bán_dương = []
                mảng_trả_âm = []
                
                for row_idx in range(2, ws_in.max_row + 1):
                    row_vals = [ws_in.cell(row=row_idx, column=c).value for c in range(1, 10)]
                    if not row_vals[2] or row_vals[2] == "None": continue
                    
                    val_sl = float(row_vals[6] or 0)
                    if val_sl == 0: continue
                    
                    formatted_row = [
                        row_vals[0], str(row_vals[1] or ""), str(row_vals[2]), str(row_vals[3] or "Cuốn"),
                        float(row_vals[4] or 0), float(row_vals[5] or 0), val_sl, 
                        float(row_vals[7] or 0), float(row_vals[8] or 0)
                    ]
                    
                    if val_sl > 0: mảng_bán_dương.append(formatted_row)
                    else: mảng_trả_âm.append(formatted_row)

                dic_gia_bia_ban = {}
                for r in mảng_bán_dương:
                    ma = r[2]
                    if ma not in dic_gia_bia_ban: dic_gia_bia_ban[ma] = {"ten": r[1], "dvt": r[3], "bia": r[4], "sl": 0.0}
                    dic_gia_bia_ban[ma]["sl"] += r[6]
                bảng_gia_bia_ban = [[i+1, v["ten"], k, v["dvt"], v["bia"], 0.0, v["sl"], v["bia"], v["sl"]*v["bia"]] for i, (k, v) in enumerate(dic_gia_bia_ban.items())]

                dic_gia_bia_tra = {}
                for r in mảng_trả_âm:
                    ma = r[2]
                    if ma not in dic_gia_bia_tra: dic_gia_bia_tra[ma] = {"ten": r[1], "dvt": r[3], "bia": r[4], "sl": 0.0}
                    dic_gia_bia_tra[ma]["sl"] += r[6]
                bảng_gia_bia_tra = [[i+1, v["ten"], k, v["dvt"], v["bia"], 0.0, v["sl"], v["bia"], v["sl"]*v["bia"]] for i, (k, v) in enumerate(dic_gia_bia_tra.items())]

                dic_ck_ban = {}
                for r in mảng_bán_dương:
                    key = f"{r[2]}_{r[5]}"
                    if key not in dic_ck_ban: dic_ck_ban[key] = {"ten": r[1], "ma": r[2], "dvt": r[3], "bia": r[4], "ck": r[5], "sl": 0.0, "dg": r[7]}
                    dic_ck_ban[key]["sl"] += r[6]
                bảng_ck_ban = [[i+1, v["ten"], v["ma"], v["dvt"], v["bia"], v["ck"], v["sl"], v["dg"], v["sl"]*v["dg"]] for i, (k, v) in enumerate(dic_ck_ban.items())]

                dic_ck_tra = {}
                for r in mảng_trả_âm:
                    key = f"{r[2]}_{r[5]}"
                    if key not in dic_ck_tra: dic_ck_tra[key] = {"ten": r[1], "ma": r[2], "dvt": r[3], "bia": r[4], "ck": r[5], "sl": 0.0, "dg": r[7]}
                    dic_ck_tra[key]["sl"] += r[6]
                bảng_ck_tra = [[i+1, v["ten"], v["ma"], v["dvt"], v["bia"], v["ck"], v["sl"], v["dg"], v["sl"]*v["dg"]] for i, (k, v) in enumerate(dic_ck_tra.items())]

                out_report = io.BytesIO()
                cols = ['STT', 'Tên sách', 'Mã số', 'ĐVT', 'Giá bìa', 'CK', 'Số lượng', 'Đơn giá', 'Thành tiền']
                
                with pd.ExcelWriter(out_report, engine='openpyxl') as writer:
                    summary_data = []
                    
                    def ghi_sheet(data, name, h_color, tab_color, has_vat=False):
                        df = pd.DataFrame(data, columns=cols)
                        df.to_excel(writer, index=False, sheet_name=name)
                        ws = writer.sheets[name]
                        ws.sheet_properties.tabColor = tab_color
                        
                        r_last = len(df) + 2
                        s_sl, s_tt = df['Số lượng'].sum(), df['Thành tiền'].sum()
                        ws.cell(row=r_last, column=6).value = "Tổng cộng:"
                        ws.cell(row=r_last, column=7).value = s_sl
                        ws.cell(row=r_last, column=9).value = s_tt
                        
                        v_val, st_val = 0.0, 0.0
                        if has_vat:
                            v_val = s_tt * 0.05
                            st_val = s_tt + v_val
                            ws.cell(row=r_last+1, column=8).value = "VAT 5%:"
                            ws.cell(row=r_last+1, column=9).value = v_val
                            ws.cell(row=r_last+2, column=8).value = "Sau Thuế:"
                            ws.cell(row=r_last+2, column=9).value = st_val
                        
                        trang_tri_sheet(ws, h_color, has_vat_summary=has_vat)
                        return len(df), s_sl, s_tt, v_val, st_val

                    ghi_sheet(mảng_bán_dương, "Du_Lieu_Ban_Duong", "595959", "595959")
                    ghi_sheet(mảng_trả_âm, "Du_Lieu_Tra_Am", "595959", "3B3838")
                    
                    l, q, a, _, _ = ghi_sheet(bảng_gia_bia_ban, "TH_Hang_Ban_Gia_Bia", "002060", "002060")
                    summary_data.append(["TH Hàng Bán Giá Bìa", "Dương", l, q, a, 0, a, "Giá Bìa Gốc"])
                    
                    l, q, a, _, _ = ghi_sheet(bảng_gia_bia_tra, "TH_Hang_Tra_Gia_Bia", "800000", "800000")
                    summary_data.append(["TH Hàng Trả Giá Bìa", "Âm", l, q, a, 0, a, "Bảo toàn dấu âm"])

                    l_b, q_b, a_b, v_b, s_b = ghi_sheet(bảng_ck_ban, "TH_Hang_Ban_Chiet_Khau", "339933", "339933", has_vat=True)
                    summary_data.append(["TH Hàng Bán Chiết Khấu", "Dương", l_b, q_b, a_b, v_b, s_b, "Bán thực tế"])

                    l_t, q_t, a_t, v_t, s_t = ghi_sheet(bảng_ck_tra, "Tong_Hop_Hang_Tra", "C00000", "C00000", has_vat=True)
                    summary_data.append(["Tổng Hợp Hàng Trả", "Âm", l_t, q_t, a_t, v_t, s_t, "Trả thực tế (Âm)"])

                    num_invoices = 0
                    if bảng_ck_ban:
                        for i, start in enumerate(range(0, len(bảng_ck_ban), 1000), 1):
                            batch = [list(r) for r in bảng_ck_ban[start:start+1000]]
                            for idx, row in enumerate(batch, 1): row[0] = idx
                            hd_name = f"HD {i}"
                            l, q, a, v, s = ghi_sheet(batch, hd_name, "008080", "008080", has_vat=True)
                            summary_data.append([hd_name, "Tách HĐ", l, q, a, v, s, f"Từ dòng {start+1}"])
                            num_invoices += 1

                    df_sum = pd.DataFrame(summary_data, columns=["Tên Sheet / Hạng mục", "Loại", "Số dòng", "Số lượng", "Trước Thuế", "VAT (5%)", "Sau Thuế", "Ghi chú"])
                    df_sum.to_excel(writer, index=False, sheet_name="Tong_Ket_Chung")
                    ws_sum = writer.sheets["Tong_Ket_Chung"]
                    ws_sum.sheet_properties.tabColor = "1F1F1F"
                    
                    net_q = q_b + q_t
                    net_a = a_b + a_t
                    net_v = v_b + v_t
                    net_s = s_b + s_t
                    
                    ws_sum.cell(row=len(df_sum) + 2, column=1).value = "DOANH THU THỰC TẾ (BÁN + TRẢ)"
                    ws_sum.cell(row=len(df_sum) + 2, column=4).value = net_q
                    ws_sum.cell(row=len(df_sum) + 2, column=5).value = net_a
                    ws_sum.cell(row=len(df_sum) + 2, column=6).value = net_v
                    ws_sum.cell(row=len(df_sum) + 2, column=7).value = net_s
                    trang_tri_sheet(ws_sum, "1F1F1F", total_row_type="grand")
                    
                    wb = writer.book
                    wb._sheets = [wb._sheets[wb.sheetnames.index(n)] for n in (["Tong_Ket_Chung"] + [sn for sn in wb.sheetnames if sn != "Tong_Ket_Chung"])]

                # GIAO DIỆN KHỐI KPI TỔNG QUAN HIỆN ĐẠI (ĐÃ SỬA ĐƠN VỊ TÍNH VÀ THÊM KHỐI SAU THUẾ)
                st.markdown("#### 📊 TỔNG HỢP BIẾN ĐỘNG DOANH THU HỆ THỐNG")
                kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
                with kpi1:
                    st.metric(label="Tổng Sản Lượng Thực Tế", value=f"{int(net_q):,} cuốn")
                with kpi2:
                    st.metric(label="Doanh Thu Trước Thuế (Net)", value=f"{net_a:,.0f} đ")
                with kpi3:
                    st.metric(label="Tổng Thuế GTGT VAT (5%)", value=f"{net_v:,.0f} đ")
                with kpi4:
                    st.metric(label="Tổng Cộng Sau Thuế (Gross)", value=f"{net_s:,.0f} đ")
                with kpi5:
                    st.metric(label="Số lượng HĐ tách", value=f"{num_invoices} file", delta="Sẵn sàng xuất")
                
                st.success("🎉 THUẬT TOÁN ĐÃ PHÂN TÁCH XONG MA TRẬN SHEET TAB!")
                st.dataframe(df_sum, use_container_width=True)
                
                st.markdown('<div style="margin-top:15px;"></div>', unsafe_allow_html=True)
                st.download_button(
                    label="📥 TẢI FILE PHÂN TÍCH TỔNG HỢP MASTER REPORT G2 (.XLSX)", 
                    data=out_report.getvalue(), 
                    file_name=f"Master_PAB21_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Lỗi G2: {str(e)}")

# ==========================================================================================
# GIAI ĐOẠN 3: XUẤT FILE MẪU CHUẨN ĐỊNH DẠNG V4 GỐC
# ==========================================================================================
with tab_giai_doan_3:
    st.markdown("### 📝 Kết xuất ma trận sang biểu mẫu Hóa đơn điện tử Viettel")
    st.info("💡 Hệ thống đọc cấu trúc tệp Master G2 và tự động ánh xạ dữ liệu chuẩn xác lên Form mẫu .xlsx của nhà cung cấp Viettel, cam kết giữ nguyên 100% định dạng khung viền gốc.")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        g3_result_file = st.file_uploader("1. Tải file Kết quả Master Report G2:", type=["xlsx"], key="g3_res")
    with col_f2:
        g3_template_file = st.file_uploader("2. Tải FILE MẪU TRẮNG (.xlsx giữ form):", type=["xlsx"], key="g3_tpl")
    
    if g3_result_file is not None and g3_template_file is not None:
        st.markdown("---")
        if st.button("📝 TIẾN HÀNH ĐÓNG GÓI HÓA ĐƠN HÀNG LOẠT", use_container_width=True):
            try:
                wb_res = openpyxl.load_workbook(g3_result_file, data_only=True)
                template_bytes = g3_template_file.read()
                
                zip_buffer = io.BytesIO()
                success_count = 0
                
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for sheet_name in wb_res.sheetnames:
                        if "HD " in sheet_name:
                            ws_res = wb_res[sheet_name]
                            
                            last_row = ws_res.max_row
                            data_end_row = last_row - 3
                            num_data_rows = data_end_row - 1
                            
                            if num_data_rows < 1: continue
                            
                            val_ae11 = round(float(ws_res.cell(row=last_row - 1, column=9).value or 0), 0)
                            
                            tpl_io = io.BytesIO(template_bytes)
                            wb_tpl = openpyxl.load_workbook(tpl_io)
                            ws_tpl = wb_tpl.worksheets[0]
                            
                            val_b11 = ws_tpl.cell(row=11, column=2).value
                            
                            for offset in range(num_data_rows):
                                r_res = 2 + offset
                                r_tpl = 11 + offset
                                
                                ws_tpl.cell(row=r_tpl, column=22).value = ws_res.cell(row=r_res, column=2).value # Tên sách -> V
                                ws_tpl.cell(row=r_tpl, column=21).value = ws_res.cell(row=r_res, column=3).value # Mã sách -> U
                                ws_tpl.cell(row=r_tpl, column=26).value = ws_res.cell(row=r_res, column=4).value # ĐVT -> Z
                                ws_tpl.cell(row=r_tpl, column=23).value = ws_res.cell(row=r_res, column=5).value # Giá bìa -> W
                                ws_tpl.cell(row=r_tpl, column=24).value = ws_res.cell(row=r_res, column=6).value # Chiết khấu -> X
                                ws_tpl.cell(row=r_tpl, column=27).value = ws_res.cell(row=r_res, column=7).value # Số lượng -> AA
                                ws_tpl.cell(row=r_tpl, column=28).value = ws_res.cell(row=r_res, column=8).value # Đơn giá -> AB
                                ws_tpl.cell(row=r_tpl, column=29).value = ws_res.cell(row=r_res, column=9).value # Thành tiền -> AC
                                
                                ws_tpl.cell(row=r_tpl, column=1).value = offset + 1
                                if val_b11:
                                    ws_tpl.cell(row=r_tpl, column=2).value = val_b11
                                    
                            ws_tpl.cell(row=11, column=31).value = val_ae11
                            
                            out_single = io.BytesIO()
                            wb_tpl.save(out_single)
                            
                            out_filename = f"Up_Hoa_Don_{sheet_name}_{datetime.now().strftime('%H%m')}.xls"
                            zip_file.writestr(out_filename, out_single.getvalue())
                            success_count += 1
                            
                if success_count == 0:
                    st.warning("⚠️ Hệ thống không phát hiện dữ liệu hóa đơn dạng tách hợp lệ.")
                    st.stop()
                    
                st.balloons()
                st.success(f"🔥 XUẤT HOÀN TẤT: Đã tạo thành công bộ {success_count} hóa đơn điện tử chuẩn mẫu Viettel!")
                
                st.markdown('<div style="margin-top:15px;"></div>', unsafe_allow_html=True)
                st.download_button(
                    label="📥 TẢI XUỐNG TOÀN BỘ BỘ FILE HÓA ĐƠN (.ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Bo_Hoa_Don_Goc_V4_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Lỗi hệ thống G3: {str(e)}")
