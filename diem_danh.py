import os
import warnings
import locale
import io
import pandas as pd
import qrcode
import streamlit as st

# Tắt các thông báo cảnh báo không cần thiết trên terminal
warnings.filterwarnings("ignore", category=UserWarning)

# Cấu hình locale tiếng Việt để sort chữ cái chuẩn (D và Đ)
try:
    locale.setlocale(locale.LC_COLLATE, "vi_VN.UTF-8")
except Exception:
    try:
        locale.setlocale(locale.LC_COLLATE, "Vietnamese_Vietnam.1258")
    except Exception:
        pass

# File lưu danh sách gốc, danh sách tiêu đề và dữ liệu điểm danh
EXCEL_FILE = "danh_sach_nhan_su.xlsx"
ATTENDANCE_FILE = "ket_qua_diem_danh.xlsx"
TITLES_FILE = "danh_sach_tieu_de.xlsx"


def load_data():
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
    else:
        df = pd.DataFrame({
            "STT": [1, 2, 3],
            "Họ tên": ["Nguyễn Văn A", "Trần Thị B", "Lê Văn C"],
            "Phòng ban": ["Phòng Tổ chức", "Phòng Đào tạo", "Ban Giám đốc"],
            "Chức vụ": ["Nhân viên", "Chuyên viên", "Trưởng phòng"],
        })
        df.to_excel(EXCEL_FILE, index=False)
    
    # Sắp xếp toàn bộ dataframe nhân sự theo bảng chữ cái tiếng Việt chuẩn dựa trên cột "Họ tên"
    if "Họ tên" in df.columns:
        df = df.copy()
        df["_sort_key"] = df["Họ tên"].astype(str).apply(lambda x: locale.strxfrm(x.split()[-1] if x.strip() else ""))
        df = df.sort_values(by="_sort_key").drop(columns=["_sort_key"]).reset_index(drop=True)
        df["STT"] = range(1, len(df) + 1)
        
    return df


def load_titles():
    if os.path.exists(TITLES_FILE):
        try:
            return pd.read_excel(TITLES_FILE)
        except Exception:
            return pd.DataFrame(columns=["Tên Tiêu đề", "Ngày học"])
    else:
        return pd.DataFrame(columns=["Tên Tiêu đề", "Ngày học"])


def save_titles(df):
    df.to_excel(TITLES_FILE, index=False)


def apply_custom_css():
    st.markdown("""
        <style>
            .stApp {
                background-color: #F8FAFC;
            }

            .block-container {
                padding-top: 2.5rem !important;
                padding-bottom: 2rem !important;
                max-width: 95% !important;
            }

            /* Header Tiêu đề ứng dụng cao cấp */
            .app-header {
                background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
                padding: 20px 16px;
                border-radius: 12px;
                text-align: center;
                box-shadow: 0 6px 10px -2px rgba(0, 0, 0, 0.1);
                margin-bottom: 15px;
                border-bottom: 3px solid #F97316;
            }
            .main-title {
                color: #F97316;
                font-weight: 900;
                font-size: 25px !important;
                text-transform: uppercase;
                margin: 0;
                letter-spacing: 1px;
            }
            .sub-title {
                color: #F8FAFC;
                font-weight: 700;
                font-size: 17px !important;
                text-transform: uppercase;
                margin-top: 8px;
                margin-bottom: 0;
                letter-spacing: 0.5px;
            }

            /* TÙY CHỈNH TÊN TAB */
            .stTabs [data-baseweb="tab-list"] {
                gap: 12px;
                background-color: #E2E8F0;
                padding: 6px;
                border-radius: 12px;
                margin-bottom: -10px;
            }
            .stTabs [data-baseweb="tab"] {
                background-color: #FFFFFF;
                border-radius: 8px;
                color: #334155;
                font-size: 22px !important;
                font-weight: 800 !important;
                padding: 10px 24px !important;
                box-shadow: 0 1px 3px rgba(0,0,0,0.05);
                border: none;
            }
            .stTabs [aria-selected="true"] {
                background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%) !important;
                color: #F97316 !important;
            }

            /* Ép sát các khoảng trống bên trong tab */
            div[data-testid="stVerticalBlock"] div[data-testid="stVerticalBlock"] {
                gap: 0.4rem !important;
            }

            /* ĐIỀU CHỈNH KÍCH THƯỚC CHỮ */
            label, .stTextInput label, .stSelectbox label, .stDateInput label {
                font-size: 17px !important;
                font-weight: 700 !important;
                color: #1E293B !important;
            }
            input, div[data-baseweb="select"] span {
                font-size: 14px !important;
            }

            /* Tiêu đề mục bên trong tab ép sát lên trên */
            h3 {
                color: #0F172A !important;
                font-size: 24px !important;
                font-weight: 800 !important;
                margin-top: -10px !important;
                margin-bottom: 5px !important;
            }

            /* Cố định chiều ngang 280px cho tất cả các nút ở cột phải giao diện chính */
            div[data-testid="stColumn"] div[data-testid="stButton"] > button,
            div[data-testid="stDownloadButton"] > button {
                width: 280px !important;
                display: block;
                margin: 0 auto;
                background: linear-gradient(135deg, #F97316 0%, #EA580C 100%);
                color: white;
                border-radius: 10px;
                font-weight: 800 !important;
                font-size: 18px !important;
                padding: 10px 20px !important;
                border: none;
                box-shadow: 0 4px 6px rgba(249, 115, 22, 0.2);
            }
            div[data-testid="stColumn"] div[data-testid="stButton"] > button:hover,
            div[data-testid="stDownloadButton"] > button:hover {
                background: linear-gradient(135deg, #EA580C 0%, #C2410C 100%);
            }

            /* FIX CHO CÁC NÚT TRONG DIALOG (POPUP) HIỂN THỊ DẠNG XẾP DỌC TRÊN - DƯỚI */
            div[data-baseweb="modal"] div[data-testid="stButton"] > button {
                width: 100% !important;
                min-width: unset !important;
                margin: 0 !important;
                background: linear-gradient(135deg, #F97316 0%, #EA580C 100%) !important;
                color: white !important;
                border-radius: 8px !important;
                font-weight: 700 !important;
                padding: 10px 20px !important;
                border: none !important;
            }

            /* Bảng dữ liệu (Dataframe) */
            .stDataFrame {
                font-size: 18px !important;
            }
        </style>
    """, unsafe_allow_html=True)


# Định nghĩa Popup Modal Xác Nhận Xóa (Xếp dọc nút trên - dưới)
@st.dialog("⚠️ Xác Nhận Xóa Tiêu Đề")
def delete_confirmation_dialog(target_title):
    st.markdown(f"Bạn có chắc chắn muốn xóa tiêu đề **'{target_title}'** này không?")
    st.write("")
    
    if st.button("🗑️ Đồng ý xóa", use_container_width=True, key="btn_confirm_delete"):
        df_titles = load_titles()
        df_titles = df_titles[df_titles["Tên Tiêu đề"] != target_title]
        save_titles(df_titles)
        st.success(f"Đã xóa thành công tiêu đề '{target_title}'.")
        st.rerun()
        
    st.write("")
    
    if st.button("❌ Hủy bỏ", use_container_width=True, key="btn_cancel_delete"):
        st.rerun()


def main():
    st.set_page_config(
        page_title="TTTM SATRA Phạm Hùng - Hệ Thống Điểm Danh", layout="wide"
    )
    apply_custom_css()

    st.markdown("""
        <div class="app-header">
            <p class="main-title">🏢 CHI BỘ TTTM SATRA PHẠM HÙNG</p>
            <p class="sub-title">📋 HỆ THỐNG QUẢN LÝ ĐIỂM DANH</p>
        </div>
    """, unsafe_allow_html=True)

    query_params = st.query_params
    is_checkin_page = "nq" in query_params

    df_nhansu = load_data()

    if is_checkin_page:
        nq_title = query_params.get("nq", "Học nghị quyết")
        nq_date = query_params.get("date", "")

        st.markdown(f"""
            <div style="background-color: #FFFFFF; padding: 35px; border-radius: 16px; text-align: center; margin: 0 auto; max-width: 700px; box-shadow: 0 10px 25px rgba(0,0,0,0.05); border: 1px solid #E2E8F0;">
                <h2 style="color: #0F172A; margin: 0; font-size: 30px !important;">📌 {nq_title}</h2>
                <p style="color: #475569; font-size: 22px !important; margin-top: 15px; font-weight: 600;">📅 Ngày học: {nq_date}</p>
                <hr style="margin: 20px 0; border: none; border-top: 1px solid #E2E8F0;">
            </div>
        """, unsafe_allow_html=True)

        col_c1, col_c2, col_c3 = st.columns([1, 2, 1])
        with col_c2:
            all_names = df_nhansu["Họ tên"].dropna().unique().tolist()

            # GIẢI PHÁP CHO IPHONE: Thêm ô gõ tìm kiếm nhanh trước, sau đó mới đến selectbox
            search_keyword = st.text_input("🔍 Gõ tên để lọc nhanh (hỗ trợ iPhone):", "", placeholder="Nhập tên hoặc họ...")
            
            if search_keyword.strip():
                # Lọc danh sách theo từ khóa người dùng gõ vào (không phân biệt hoa thường)
                filtered_names = [name for name in all_names if search_keyword.strip().lower() in name.lower()]
            else:
                filtered_names = all_names

            selected_name = st.selectbox(
                "👉 Gõ hoặc chọn họ tên của đồng chí:", ["-- Chọn họ tên --"] + filtered_names
            )

            default_pb = ""
            default_cv = ""

            if selected_name != "-- Chọn họ tên --":
                matched_row = df_nhansu[df_nhansu["Họ tên"] == selected_name]
                if not matched_row.empty:
                    default_pb = str(matched_row.iloc[0]["Phòng ban"])
                    default_cv = str(matched_row.iloc[0]["Chức vụ"])

            st.markdown("<p style='font-size: 17px; font-weight: 700; margin-top: 15px; margin-bottom: 5px; color: #1E293B;'>🏢 Phòng ban:</p>", unsafe_allow_html=True)
            st.info(f"**{default_pb}**" if default_pb else "Chưa chọn tên...")

            st.markdown("<p style='font-size: 17px; font-weight: 700; margin-top: 10px; margin-bottom: 5px; color: #1E293B;'>💼 Chức vụ:</p>", unsafe_allow_html=True)
            st.info(f"**{default_cv}**" if default_cv else "Chưa chọn tên...")

            st.write("")
            if st.button("✅ XÁC NHẬN ĐIỂM DANH"):
                if selected_name == "-- Chọn họ tên --":
                    st.error("⚠️ Vui lòng chọn hoặc gõ tìm họ tên của đồng chí!")
                else:
                    new_record = pd.DataFrame([{
                        "Nội dung Nghị quyết": nq_title,
                        "Ngày học": nq_date,
                        "Họ tên": selected_name,
                        "Phòng ban": default_pb,
                        "Chức vụ": default_cv,
                        "Thời gian điểm danh": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                    }])

                    if os.path.exists(ATTENDANCE_FILE):
                        df_att = pd.read_excel(ATTENDANCE_FILE)
                        df_att = pd.concat([df_att, new_record], ignore_index=True)
                    else:
                        df_att = new_record

                    df_att.to_excel(ATTENDANCE_FILE, index=False)
                    st.success("🎉 Cảm ơn Đồng chí! Điểm danh thành công.")
                    st.balloons()

    else:
        tab1, tab2, tab3 = st.tabs([
            "🎯 1. Tạo QR Điểm danh", 
            "📊 2. Xem Danh sách Điểm danh", 
            "👥 3. Quản lý Nhân sự"
        ])

        with tab1:
            col_left, col_right = st.columns([2, 1], gap="large")

            df_titles = load_titles()

            with col_left:
                st.markdown("### 🛠️ Thiết Lập Mã QR Điểm Danh")
                st.write("")

                selected_rows = st.session_state.get("titles_dataframe", {}).get("selection", {}).get("rows", [])
                default_title_val = ""
                if selected_rows and not df_titles.empty:
                    selected_idx = selected_rows[0]
                    if selected_idx < len(df_titles):
                        default_title_val = str(df_titles.iloc[selected_idx]["Tên Tiêu đề"])

                col_lbl1, col_input1 = st.columns([1.5, 8.5], gap="small")
                with col_lbl1:
                    st.markdown("<p style='margin-top: 8px; font-size: 17px; font-weight: 700; color: #1E293B;'>Tiêu đề:</p>", unsafe_allow_html=True)
                with col_input1:
                    nq_title_input = st.text_input("Nhập tiêu đề", value=default_title_val, placeholder="Nhập tên tiêu đề sự kiện / nghị quyết...", label_visibility="collapsed")

                st.write("")

                col_lbl2, col_date2 = st.columns([1.5, 8.5], gap="small")
                with col_lbl2:
                    st.markdown("<p style='margin-top: 8px; font-size: 17px; font-weight: 700; color: #1E293B;'>Ngày học:</p>", unsafe_allow_html=True)
                with col_date2:
                    nq_date_input = st.date_input("Chọn ngày học", label_visibility="collapsed", format="DD/MM/YYYY")

                formatted_date_str = nq_date_input.strftime("%d/%m/%Y")

                st.write("")

                st.markdown("#### 📋 Danh Sách Tiêu Đề Đã Thiết Lập")

                if df_titles.empty:
                    st.info("ℹ️ Chưa có tiêu đề nào được thêm.")
                else:
                    event = st.dataframe(
                        df_titles, 
                        width="stretch", 
                        height=200, 
                        selection_mode="single-row", 
                        on_select="rerun",
                        key="titles_dataframe"
                    )

            with col_right:
                create_qr_clicked = st.button("🚀 Tạo mã QRCode")

                if create_qr_clicked:
                    title_input = nq_title_input.strip()
                    if not title_input:
                        st.warning("⚠️ Vui lòng nhập tiêu đề!")
                    else:
                        current_host = "https://app-diem-danh-nx2uwapdvmixmcuze7cjzn.streamlit.app"  
                        qr_url = f"{current_host}/?nq={title_input}&date={formatted_date_str}"

                        st.session_state["qr_url"] = qr_url
                        st.session_state["nq_title"] = title_input

                        qr = qrcode.QRCode(version=1, box_size=10, border=5)
                        qr.add_data(qr_url)
                        qr.make(fit=True)
                        img = qr.make_image(fill_color="black", back_color="white")
                        img.save("temp_qr.png")

                        df_titles_current = load_titles()
                        if title_input in df_titles_current["Tên Tiêu đề"].values:
                            st.success("✨ Đã tạo mã QR thành công! (Tiêu đề này đã có sẵn trong danh sách).")
                        else:
                            new_row = pd.DataFrame([{"Tên Tiêu đề": title_input, "Ngày học": formatted_date_str}])
                            df_titles_current = pd.concat([df_titles_current, new_row], ignore_index=True)
                            save_titles(df_titles_current)
                            st.success("✨ Đã tạo mã QR và thêm tiêu đề thành công!")
                            st.rerun()

                st.write("")

                if "temp_qr.png" in os.listdir():
                    st.image("temp_qr.png", caption=st.session_state.get("nq_title", ""), width=280)
                    st.write("")
                    with open("temp_qr.png", "rb") as file:
                        st.download_button(
                            label="📥 Tải xuống mã QRCode",
                            data=file,
                            file_name="qr_diem_danh.png",
                            mime="image/png",
                        )
                else:
                    st.info("ℹ️ Chưa có mã QR nào được tạo.")

                st.write("")

                delete_title_clicked = st.button("🗑️ Xóa Tiêu đề")
                if delete_title_clicked:
                    selected_rows = st.session_state.get("titles_dataframe", {}).get("selection", {}).get("rows", [])
                    if not selected_rows:
                        st.warning("⚠️ Vui lòng nhấp chọn một hàng trong bảng danh sách bên trái để xóa!")
                    else:
                        selected_idx = selected_rows[0]
                        target_title = df_titles.iloc[selected_idx]["Tên Tiêu đề"]
                        
                        has_transaction = False
                        if os.path.exists(ATTENDANCE_FILE):
                            df_att_check = pd.read_excel(ATTENDANCE_FILE)
                            if "Nội dung Nghị quyết" in df_att_check.columns:
                                if target_title in df_att_check["Nội dung Nghị quyết"].values:
                                    has_transaction = True

                        if has_transaction:
                            st.error(f"❌ Không thể xóa tiêu đề '{target_title}' vì tiêu đề này đã có người điểm danh.")
                        else:
                            delete_confirmation_dialog(target_title)

        with tab2:
            st.markdown("### 📈 Thống Kê & Báo Cáo Điểm Danh")
            if os.path.exists(ATTENDANCE_FILE):
                df_att = pd.read_excel(ATTENDANCE_FILE)

                list_nq = df_att["Nội dung Nghị quyết"].unique().tolist()
                selected_filter = st.selectbox(
                    "🔍 Lọc theo nội dung nghị quyết:", ["Tất cả"] + list_nq
                )

                df_filtered = df_att[df_att["Nội dung Nghị quyết"] == selected_filter] if selected_filter != "Tất cả" else df_att

                st.write("")
                st.dataframe(df_filtered, width="stretch", height=450)

                st.write("")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_filtered.to_excel(writer, index=False)
                excel_data = output.getvalue()

                file_name_download = f"bao_cao_{selected_filter}.xlsx" if selected_filter != "Tất cả" else "bao_cao_tat_ca_diem_danh.xlsx"

                st.download_button(
                    label="📥 Tải Xuống Báo Cáo Điểm Danh (Excel)",
                    data=excel_data,
                    file_name=file_name_download,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("ℹ️ Hiện tại chưa có dữ liệu điểm danh nào được ghi nhận.")

        with tab3:
            st.markdown("### 📂 Quản Lý Danh Sách Nhân Sự TTTM")
            st.dataframe(df_nhansu, width="stretch", height=450)
            st.warning(
                "💡 **Lưu ý:** Bạn có thể thay thế file `danh_sach_nhan_su.xlsx` bằng danh sách thực tế của đơn vị với đúng tên các cột tương ứng."
            )

if __name__ == "__main__":
    main()