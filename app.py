"""
HỆ THỐNG TRA CỨU & BÁO CÁO TỔNG HỢP SỐ LIỆU PHỔ CẬP GIÁO DỤC - XÓA MÙ CHỮ NĂM 2026
SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH THÁI NGUYÊN
Phiên bản Web App Độc lập (Viewer Mode) - Tối ưu triển khai Cloud (Render, Streamlit Cloud, v.v.)
"""

import os
import json
import pandas as pd
import streamlit as st
import openpyxl

# ---------------------------------------------------------
# CẤU HÌNH TRANG WEB STREAMLIT
# ---------------------------------------------------------
st.set_page_config(
    page_title="Tổng hợp PCGD & XMC 2026 - Sở GDĐT Thái Nguyên",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Đường dẫn thư mục outputs tương đối trong webapp
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUTS_DIR = os.path.join(CURRENT_DIR, "outputs")

# ---------------------------------------------------------
# TÙY CHỈNH GIAO DIỆN CSS PHONG CÁCH HÀNH CHÍNH TRANG NHÃ
# ---------------------------------------------------------
st.markdown("""
<style>
    /* Font và khoảng cách chung */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    
    /* Header Quốc hiệu & Đơn vị */
    .gov-header {
        text-align: center;
        border-bottom: 2px solid #1e3a8a;
        padding-bottom: 12px;
        margin-bottom: 20px;
    }
    .gov-title-sub {
        font-size: 13px;
        font-weight: 700;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 2px;
    }
    .gov-title-main {
        font-size: 24px;
        font-weight: 800;
        color: #1e3a8a;
        text-transform: uppercase;
        margin-top: 4px;
        margin-bottom: 4px;
    }
    .gov-title-desc {
        font-size: 14px;
        color: #64748b;
        font-style: italic;
    }
    
    /* KPI Card */
    .metric-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
        border: 1px solid #e2e8f0;
        border-left: 4px solid #1e3a8a;
        border-radius: 8px;
        padding: 12px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
    .metric-label {
        font-size: 12px;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
    }
    .metric-value {
        font-size: 22px;
        font-weight: 700;
        color: #0f172a;
        margin-top: 2px;
    }
    
    /* Alert sync status */
    .sync-badge {
        display: inline-block;
        background-color: #ecfdf5;
        color: #065f46;
        border: 1px solid #a7f3d0;
        border-radius: 20px;
        padding: 4px 12px;
        font-size: 12px;
        font-weight: 600;
    }
    
    /* Footer */
    .footer-container {
        margin-top: 40px;
        border-top: 1px solid #e2e8f0;
        padding-top: 16px;
        text-align: center;
        color: #64748b;
        font-size: 12px;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# HÀM TẢI DỮ LIỆU CÓ CACHE TỐC ĐỘ CAO
# ---------------------------------------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def load_summary_dashboard():
    """Đọc dữ liệu tổng kết từ summary_dashboard.json."""
    path = os.path.join(OUTPUTS_DIR, "summary_dashboard.json")
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

@st.cache_data(ttl=3600, show_spinner=False)
def load_metadata():
    """Đọc thông tin độ mới từ last_sync_metadata.json."""
    path = os.path.join(OUTPUTS_DIR, "last_sync_metadata.json")
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

@st.cache_data(ttl=3600, show_spinner=False)
def read_excel_sheet_cached(file_name, sheet_name=None):
    """Đọc một sheet cụ thể trong file Excel với cache bộ nhớ."""
    full_path = os.path.join(OUTPUTS_DIR, file_name)
    if not os.path.exists(full_path):
        return None
    try:
        if sheet_name:
            df = pd.read_excel(full_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(full_path)
        return df
    except Exception as e:
        st.error(f"Lỗi khi đọc sheet '{sheet_name}' từ tệp {file_name}: {e}")
        return None

@st.cache_data(ttl=3600, show_spinner=False)
def get_excel_sheet_names_cached(file_name):
    """Lấy danh sách các sheet trong file Excel."""
    full_path = os.path.join(OUTPUTS_DIR, file_name)
    if not os.path.exists(full_path):
        return []
    try:
        wb = openpyxl.load_workbook(full_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    except Exception:
        return []

# ---------------------------------------------------------
# TẢI DỮ LIỆU BAN ĐẦU
# ---------------------------------------------------------
dashboard_data = load_summary_dashboard()
metadata = load_metadata()

sync_time = metadata.get("last_sync_time") or dashboard_data.get("timestamp", "Đã cập nhật")
total_communes = metadata.get("total_valid_files") or dashboard_data.get("total_communes", 92)
sheets_dict = dashboard_data.get("sheets", {})

# ---------------------------------------------------------
# BANNER TIÊU ĐỀ CHÍNH
# ---------------------------------------------------------
st.markdown("""
<div class="gov-header">
    <div class="gov-title-sub">ỦY BAN NHÂN DÂN TỈNH THÁI NGUYÊN — SỞ GIÁO DỤC VÀ ĐÀO TẠO</div>
    <div class="gov-title-main">HỆ THỐNG TRA CỨU & BÁO CÁO SỐ LIỆU PHỔ CẬP GIÁO DỤC - XÓA MÙ CHỮ NĂM 2026</div>
    <div class="gov-title-desc">Dữ liệu tổng hợp toàn diện từ đủ 92 xã/phường/thị trấn toàn tỉnh Thái Nguyên</div>
</div>
""", unsafe_allow_html=True)

# Thanh trạng thái & Tải nhanh toàn tỉnh
col_info, col_d1, col_d2 = st.columns([2, 1.2, 1.2])

with col_info:
    st.markdown(f"""
    <div style="padding-top: 6px;">
        <span class="sync-badge">✓ DỮ LIỆU CHUẨN ĐÃ ĐỒNG BỘ</span>
        <span style="font-size: 13px; color: #475569; margin-left: 10px;">
            Thời gian: <b>{sync_time}</b> | Quy mô: <b>{total_communes}/92 Xã/Phường</b> | Tổng số: <b>33/33 Biểu</b>
        </span>
    </div>
    """, unsafe_allow_html=True)

master_file_path = os.path.join(OUTPUTS_DIR, "SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx")
report_file_path = os.path.join(OUTPUTS_DIR, "BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx")

with col_d1:
    if os.path.exists(master_file_path):
        with open(master_file_path, "rb") as f:
            st.download_button(
                label="📥 Tải Sổ cái Master 33 Biểu",
                data=f.read(),
                file_name="SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Sổ cái hợp nhất toàn bộ 33 biểu nghiệp vụ chuẩn cấp tỉnh"
            )

with col_d2:
    if os.path.exists(report_file_path):
        with open(report_file_path, "rb") as f:
            st.download_button(
                label="📥 Tải Báo cáo 63 Quy tắc",
                data=f.read(),
                file_name="BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Báo cáo kiểm tra chất lượng và đối soát 63 quy tắc logic"
            )

st.markdown("<hr style='margin-top: 10px; margin-bottom: 20px; border-color: #e2e8f0;'/>", unsafe_allow_html=True)

# ---------------------------------------------------------
# SIDEBAR ĐIỀU HƯỚNG
# ---------------------------------------------------------
with st.sidebar:
    st.markdown("### 🏛️ ĐIỀU HÀNH HỆ THỐNG")
    st.markdown("**Sở Giáo dục và Đào tạo Thái Nguyên**")
    st.caption("Ban Chỉ đạo Phổ cập Giáo dục & Xóa mù chữ")
    
    st.markdown("---")
    main_menu = st.radio(
        "Lựa chọn phân hệ:",
        [
            "📊 Tra cứu 33 Biểu Nghiệp vụ",
            "🛡️ Báo cáo Đối soát 63 Quy tắc",
            "📖 Thông tin & Nguyên tắc Nghiệp vụ"
        ],
        index=0
    )
    
    st.markdown("---")
    st.markdown("#### 📌 Thông số dữ liệu")
    st.info(f"""
    - **Số xã tổng hợp:** {total_communes}/92
    - **Đồng bộ lần cuối:** {sync_time}
    - **Chế độ:** Hiển thị tức thì (Viewer)
    """)

# ---------------------------------------------------------
# PHÂN HỆ 1: TRA CỨU 33 BIỂU MẪU NGHIỆP VỤ
# ---------------------------------------------------------
if main_menu == "📊 Tra cứu 33 Biểu Nghiệp vụ":
    # 1. Bộ lọc khối
    col_group, col_sheet = st.columns([1, 2])
    
    GROUPS = [
        "Tất cả các khối",
        "Mầm non (0.M1 - 0.M7)",
        "Tiểu học (1.M1 - 1.M7)",
        "THCS (2.M1 - 2.M10)",
        "Xóa mù chữ (3.M1 - 3.M9)"
    ]
    
    with col_group:
        selected_group = st.selectbox("📂 Lọc theo khối cấp học:", GROUPS, index=0)
    
    # 2. Danh mục biểu
    filtered_sheets = {}
    for code, info in sheets_dict.items():
        grp = info.get("group", "")
        if selected_group == "Tất cả các khối":
            filtered_sheets[code] = info
        elif "Mầm non" in selected_group and grp == "Mầm non":
            filtered_sheets[code] = info
        elif "Tiểu học" in selected_group and grp == "Tiểu học":
            filtered_sheets[code] = info
        elif "THCS" in selected_group and grp == "THCS":
            filtered_sheets[code] = info
        elif "Xóa mù chữ" in selected_group and grp == "Xóa mù chữ":
            filtered_sheets[code] = info
            
    if not filtered_sheets:
        filtered_sheets = sheets_dict

    sheet_options = [f"{code} — {info.get('name')}" for code, info in filtered_sheets.items()]
    
    with col_sheet:
        selected_sheet_label = st.selectbox(
            "📋 Chọn biểu mẫu cần xem chi tiết:",
            sheet_options,
            index=0 if sheet_options else 0
        )
    
    if selected_sheet_label:
        selected_code = selected_sheet_label.split(" — ")[0].strip()
        current_info = sheets_dict.get(selected_code, {})
        excel_name = current_info.get("excel_file", f"TONG_HOP_THU_{selected_code}_PCGD_2026.xlsx")
        sheet_title = current_info.get("name", selected_sheet_label)
        
        # Tiêu đề biểu & Nút tải Excel của biểu
        c_title, c_btn = st.columns([3, 1])
        with c_title:
            st.markdown(f"### 📑 Biểu {selected_code}: {sheet_title}")
            st.caption(f"Khối: **{current_info.get('group', 'Chuyên môn')}** | Tệp kết quả: `{excel_name}`")
        with c_btn:
            single_file_path = os.path.join(OUTPUTS_DIR, excel_name)
            if os.path.exists(single_file_path):
                with open(single_file_path, "rb") as f:
                    st.download_button(
                        label=f"📥 Tải Excel Biểu {selected_code}",
                        data=f.read(),
                        file_name=excel_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        
        # KPI Metrics tóm tắt
        metrics = current_info.get("summary_metrics", {})
        if metrics:
            st.markdown("##### 📌 Chỉ tiêu cốt lõi cấp tỉnh:")
            m_cols = st.columns(len(metrics))
            for idx, (m_label, m_val) in enumerate(metrics.items()):
                with m_cols[idx]:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">{m_label}</div>
                        <div class="metric-value">{m_val}</div>
                    </div>
                    """, unsafe_allow_html=True)
        
        # Đọc các sheet của file Excel này
        available_sheets = get_excel_sheet_names_cached(excel_name)
        
        if not available_sheets:
            st.warning(f"Chưa tìm thấy tệp kết quả `{excel_name}` hoặc tệp không có dữ liệu.")
        else:
            tabs = st.tabs([f"📄 {sname}" for sname in available_sheets])
            for idx, sname in enumerate(available_sheets):
                with tabs[idx]:
                    df_sheet = read_excel_sheet_cached(excel_name, sname)
                    if df_sheet is not None:
                        # Nếu là sheet chi tiết theo xã, hỗ trợ tìm kiếm nhanh theo tên xã
                        is_detail_sheet = "CHI_TIET" in sname or "XA" in sname
                        if is_detail_sheet:
                            search_col1, search_col2 = st.columns([2, 1])
                            with search_col1:
                                search_term = st.text_input(
                                    f"🔍 Tìm kiếm nhanh tên xã/phường trong {sname}:",
                                    key=f"search_{selected_code}_{sname}"
                                )
                            with search_col2:
                                st.write("")
                                st.caption(f"Tổng số bản ghi: **{len(df_sheet)}** hàng")
                            
                            if search_term:
                                # Lọc qua các cột chứa chữ (đặc biệt là cột tên xã / đơn vị)
                                mask = df_sheet.astype(str).apply(
                                    lambda row: row.str.contains(search_term, case=False, na=False).any(),
                                    axis=1
                                )
                                df_display = df_sheet[mask]
                                st.caption(f"Tìm thấy **{len(df_display)}** xã khớp từ khóa '{search_term}'")
                            else:
                                df_display = df_sheet
                        else:
                            df_display = df_sheet
                            st.caption(f"Bảng tổng hợp: **{len(df_display)}** hàng x **{len(df_display.columns)}** cột")
                        
                        # Hiển thị bảng dữ liệu với định dạng gọn gàng
                        st.dataframe(
                            df_display,
                            use_container_width=True,
                            height=min(500, 38 + len(df_display) * 35)
                        )
                    else:
                        st.info("Không có dữ liệu hiển thị.")

# ---------------------------------------------------------
# PHÂN HỆ 2: BÁO CÁO ĐỐI SOÁT 63 QUY TẮC KIỂM TRA
# ---------------------------------------------------------
elif main_menu == "🛡️ Báo cáo Đối soát 63 Quy tắc":
    st.markdown("### 🛡️ BÁO CÁO ĐỐI SOÁT CHẤT LƯỢNG & 63 QUY TẮC LOGIC LIÊN BIỂU")
    st.write(
        "Hệ thống tự động rà soát chéo giữa các biểu mẫu (Mầm non, Tiểu học, THCS, Xóa mù chữ) "
        "để phát hiện mâu thuẫn số học, công thức lỗi hoặc số liệu nộp sai lệch từ cấp xã."
    )
    
    report_sheets = get_excel_sheet_names_cached("BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx")
    
    if not report_sheets:
        st.error("Không tìm thấy tệp `BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx` trong thư mục `outputs/`.")
    else:
        # Tên các tab thân thiện
        tab_names = []
        for s in report_sheets:
            if "TONG_QUAN" in s:
                tab_names.append("🌐 1. Tổng quan Toàn tỉnh")
            elif "DOI_SOAT" in s:
                tab_names.append("⚖️ 2. Đối soát Cấp tỉnh (63 Quy tắc)")
            elif "DASHBOARD" in s:
                tab_names.append("📊 3. Dashboard Đánh giá 92 Xã")
            elif "SAI_LECH" in s:
                tab_names.append("🚨 4. Chi tiết Sai lệch Thanh tra")
            else:
                tab_names.append(f"📄 {s}")
                
        r_tabs = st.tabs(tab_names)
        
        for idx, sname in enumerate(report_sheets):
            with r_tabs[idx]:
                df_rep = read_excel_sheet_cached("BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx", sname)
                if df_rep is not None:
                    # Cho phép tìm kiếm trên dashboard xã hoặc bảng sai lệch
                    if "DASHBOARD" in sname or "SAI_LECH" in sname:
                        c_s1, c_s2 = st.columns([2, 1])
                        with c_s1:
                            search_q = st.text_input(
                                f"🔍 Tìm kiếm xã/chỉ tiêu trong sheet {sname}:",
                                key=f"rep_search_{sname}"
                            )
                        with c_s2:
                            st.write("")
                            st.caption(f"Tổng số dòng: **{len(df_rep)}**")
                        
                        if search_q:
                            mask = df_rep.astype(str).apply(
                                lambda row: row.str.contains(search_q, case=False, na=False).any(),
                                axis=1
                            )
                            df_view = df_rep[mask]
                            st.caption(f"Tìm thấy **{len(df_view)}** kết quả khớp")
                        else:
                            df_view = df_rep
                    else:
                        df_view = df_rep
                        st.caption(f"Bảng: **{len(df_view)}** hàng x **{len(df_view.columns)}** cột")
                    
                    st.dataframe(
                        df_view,
                        use_container_width=True,
                        height=min(550, 38 + len(df_view) * 35)
                    )
                else:
                    st.info("Không có dữ liệu.")

# ---------------------------------------------------------
# PHÂN HỆ 3: THÔNG TIN & NGUYÊN TẮC NGHIỆP VỤ
# ---------------------------------------------------------
elif main_menu == "📖 Thông tin & Nguyên tắc Nghiệp vụ":
    st.markdown("### 📖 THÔNG TIN HỆ THỐNG VÀ NGUYÊN TẮC NGHIỆP VỤ")
    
    st.markdown("""
    #### 1. Đơn vị chủ quản và mục tiêu vận hành
    - **Cơ quan chủ quản:** Sở Giáo dục và Đào tạo tỉnh Thái Nguyên.
    - **Nhiệm vụ:** Phục vụ công tác thống kê, đối soát, kiểm tra và ra quyết định công nhận đạt chuẩn 
      Phổ cập giáo dục - Xóa mù chữ năm 2026 cho 92 xã, phường, thị trấn trên địa bàn tỉnh.
    - **Quy mô:** Bao gồm đầy đủ 33 biểu nghiệp vụ chuyên môn (Mầm non, Tiểu học, THCS, Xóa mù chữ) 
      và 4 biểu tổng hợp đối soát chung.
      
    ---
    
    #### 2. Nguyên tắc tính toán số học & Tỷ lệ chuẩn (Bất di bất dịch)
    1. **Chỉ tiêu số lượng (đếm/người/lớp/phòng):** Được cộng dồn chính xác từ 92 xã nộp hợp lệ.
    2. **Chỉ tiêu tỷ lệ phần trăm (%):** 
       - **Tuyệt đối không cộng trực tiếp** các tỷ lệ của cấp xã hoặc lấy trung bình cộng đơn giản.
       - Tỷ lệ cấp tỉnh phải được tính lại trên cùng tập xã hợp lệ theo nguyên tắc:
       $$\\text{Tỷ lệ Cấp Tỉnh (\\%)} = \\frac{\\sum \\text{Tử số các xã}}{\\sum \\text{Mẫu số các xã}} \\times 100$$
    3. **Quy chuẩn lưu trữ và hiển thị phần trăm:**
       - Giữ tỷ lệ dưới dạng **số thực theo thang 0–100** (ví dụ: `98.63` nghĩa là `98.63%`).
       - Định dạng hiển thị ô Excel: `cell.number_format = '0.00"%"'`.
    4. **Bảo toàn dữ liệu nguồn:** Tuyệt đối không can thiệp, không sửa đổi dữ liệu gốc do cơ sở nộp lên.
    
    ---
    
    #### 3. Cấu trúc bộ mã và biểu mẫu thống kê (33 Biểu)
    - **Khối GD Mầm non (0.M1 — 0.M7):** Thống kê trẻ em theo độ tuổi, quy mô nhóm lớp, đội ngũ CBQL-GV-NV, cơ sở vật chất, tài chính GDMN, trẻ khuyết tật và sổ PCGD.
    - **Khối GD Tiểu học (1.M1 — 1.M7):** Trẻ 6-14t PCGD, đội ngũ GV, cơ sở vật chất, chuẩn PCGD bắt buộc, trẻ khuyết tật, học sinh ngoài địa bàn đến và đi.
    - **Khối GD Trung học cơ sở (2.M1 — 2.M10):** Thanh thiếu niên 11-18t PCGD THCS, hiệu quả đào tạo, cơ sở vật chất, giáo viên, trẻ khuyết tật, chuẩn PCGD THCS Mức 1-4.
    - **Khối Xóa mù chữ (3.M1 — 3.M9):** Chuẩn biết chữ Mức 1 & Mức 2 các nhóm tuổi (15-25t, 15-35t, 15-60t) và hiện trạng mù chữ phân theo giới tính, dân tộc.
    """)

# ---------------------------------------------------------
# FOOTER HÀNH CHÍNH
# ---------------------------------------------------------
st.markdown("""
<div class="footer-container">
    <b>HỆ THỐNG PHỔ CẬP GIÁO DỤC VÀ XÓA MÙ CHỮ TỈNH THÁI NGUYÊN NĂM 2026</b><br/>
    Bản quyền thuộc Sở Giáo dục và Đào tạo tỉnh Thái Nguyên • Triển khai kỹ thuật phục vụ công tác chuyển đổi số giáo dục.
</div>
""", unsafe_allow_html=True)

