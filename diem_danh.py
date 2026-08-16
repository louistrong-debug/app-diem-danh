import os
import warnings
import locale
import io
import base64
import unicodedata
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import qrcode
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image

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

# File lưu danh sách gốc, danh sách tiêu đề, dữ liệu điểm danh và danh sách user
EXCEL_FILE = "danh_sach_nhan_su.xlsx"
ATTENDANCE_FILE = "ket_qua_diem_danh.xlsx"
TITLES_FILE = "danh_sach_tieu_de.xlsx"
USER_FILE = "danh_sach_user.xlsx"

SHEET_NAME = "QuanLyDiemDanh" 
CREDENTIALS_FILE = "credentials.json"
FILES_MAP = {
    "danh_sach_nhan_su": EXCEL_FILE,
    "danh_sach_tieu_de": TITLES_FILE,
    "ket_qua_diem_danh": ATTENDANCE_FILE,
    "danh_sach_user": USER_FILE
}


def convert_image_to_base64(image_bytes):
    """Hàm nén, thu nhỏ và chuyển đổi bytes ảnh chụp thành chuỗi Base64 nhỏ gọn (vừa vặn với Excel)"""
    try:
        image = Image.open(io.BytesIO(image_bytes))
        
        max_width = 300
        if image.width > max_width:
            ratio = max_width / image.width
            new_height = int(image.height * ratio)
            image = image.resize((max_width, new_height), Image.Resampling.LANCZOS)
            
        buffered = io.BytesIO()
        image.save(buffered, format="JPEG", quality=60)
        encoded = base64.b64encode(buffered.getvalue()).decode('utf-8')
        return f"data:image/jpeg;base64,{encoded}"
    except Exception as e:
        st.error(f"❌ Lỗi mã hóa ảnh: {e}")
        return ""


def sync_to_google():
    """Hàm đồng bộ toàn bộ dữ liệu lên Google Sheets, giữ nguyên vẹn chuỗi Base64 của ảnh"""
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        
        client = gspread.authorize(creds)
        spreadsheet = client.open(SHEET_NAME)
        
        for sheet_key, filename in FILES_MAP.items():
            if os.path.exists(filename):
                df = pd.read_excel(filename)
            else:
                df = pd.DataFrame()
                
            try:
                ws = spreadsheet.worksheet(sheet_key)
            except Exception:
                ws = spreadsheet.add_worksheet(title=sheet_key, rows="100", cols="20")
            
            # Xóa dữ liệu cũ trên sheet trước khi cập nhật
            ws.clear()
            
            df = df.fillna("")
            if not df.empty:
                # Đẩy toàn bộ dữ liệu (bao gồm cả chuỗi mã ảnh dài) lên Google Sheets
                data_to_update = [df.columns.values.tolist()] + df.astype(str).values.tolist()
                ws.update(data_to_update)
            else:
                if len(df.columns) > 0:
                    ws.update([df.columns.values.tolist()])
        return True
    except Exception as e:
        return False


def sync_from_google_to_local():
    """Tải dữ liệu từ Google Sheets về lại file Excel cục bộ nhưng vẫn bảo vệ chuỗi ảnh Base64 local"""
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open(SHEET_NAME)
        
        for sheet_key, filename in FILES_MAP.items():
            try:
                ws = spreadsheet.worksheet(sheet_key)
                data = ws.get_all_records()
                if data:
                    df_cloud = pd.DataFrame(data)
                    
                    if sheet_key == "ket_qua_diem_danh" and os.path.exists(filename):
                        df_local = pd.read_excel(filename)
                        if "Mã Ảnh Drive" in df_local.columns and "Mã Ảnh Drive" in df_cloud.columns:
                            if len(df_local) == len(df_cloud):
                                df_cloud["Mã Ảnh Drive"] = df_local["Mã Ảnh Drive"].values
                    
                    df_cloud.to_excel(filename, index=False)
                else:
                    all_values = ws.get_all_values()
                    if all_values:
                        df_cloud = pd.DataFrame(all_values[1:], columns=all_values[0])
                        df_cloud.to_excel(filename, index=False)
            except Exception:
                pass
    except Exception:
        pass


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
    
    if "Họ tên" in df.columns:
        df = df.copy()
        
        def vn_sort_key(full_name):
            if not isinstance(full_name, str) or not full_name.strip():
                return ""
            parts = full_name.strip().split()
            name_reversed = [parts[-1]] + parts[:-1]
            key_str = " ".join(name_reversed)
            
            key_str = key_str.replace('Đ', 'Dba').replace('đ', 'dba')
            
            nfkd_form = unicodedata.normalize('NFKD', key_str)
            ascii_str = "".join([c for c in nfkd_form if not unicodedata.combining(c)]).lower()
            return ascii_str

        df["_sort_key"] = df["Họ tên"].apply(vn_sort_key)
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
    # Lưu trực tiếp vào file Excel cục bộ
    df.to_excel(TITLES_FILE, index=False)
    
    # Đồng bộ trực tiếp và làm sạch dữ liệu bảng tiêu đề trên Google Sheets ngay lập tức
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open(SHEET_NAME)
        ws = spreadsheet.worksheet("danh_sach_tieu_de")
        ws.clear()
        
        df_clean = df.fillna("")
        if not df_clean.empty:
            ws.update([df_clean.columns.values.tolist()] + df_clean.astype(str).values.tolist())
        else:
            if len(df_clean.columns) > 0:
                ws.update([df_clean.columns.values.tolist()])
    except Exception:
        pass


def load_users():
    if os.path.exists(USER_FILE):
        try:
            df_u = pd.read_excel(USER_FILE)
            if "Mật khẩu" in df_u.columns:
                df_u["Mật khẩu"] = df_u["Mật khẩu"].astype(str)
            return df_u
        except Exception:
            pass
    default_user = pd.DataFrame({
        "Tên đăng nhập": ["admin"],
        "Mật khẩu": ["123456"],
        "Quyền hạn": ["Quản trị viên (Admin)"]
    })
    default_user.to_excel(USER_FILE, index=False)
    return default_user


def save_users(df):
    df.to_excel(USER_FILE, index=False)
    sync_to_google()


def sync_single_record_to_google(record_dict):
    try:
        # Gọi thẳng hàm đồng bộ tổng để đẩy file Excel đã lưu lên thẳng Google Sheets
        return sync_to_google()
    except Exception:
        return False


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
                margin-bottom: 4px;
                letter-spacing: 0.5px;
            }
            .ai-badge {
                color: #94A3B8;
                font-weight: 500;
                font-size: 13px !important;
                font-style: italic;
                margin: 0;
                letter-spacing: 0.3px;
            }

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
                font-size: 19px !important;
                font-weight: 800 !important;
                padding: 10px 18px !important;
                box-shadow: 0 1px 3px rgba(0,0,0,0.05);
                border: none;
            }
            .stTabs [aria-selected="true"] {
                background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%) !important;
                color: #F97316 !important;
            }

            div[data-testid="stVerticalBlock"] div[data-testid="stVerticalBlock"] {
                gap: 0.4rem !important;
            }

            label, .stTextInput label, .stSelectbox label, .stDateInput label {
                font-size: 17px !important;
                font-weight: 700 !important;
                color: #1E293B !important;
            }
            input, div[data-baseweb="select"] span {
                font-size: 14px !important;
            }

            h3 {
                color: #0F172A !important;
                font-size: 24px !important;
                font-weight: 800 !important;
                margin-top: -10px !important;
                margin-bottom: 5px !important;
            }

            div[data-testid="stColumn"] div[data-testid="stButton"] > button,
            div[data-testid="stDownloadButton"] > button {
                width: 100% !important;
                background: linear-gradient(135deg, #F97316 0%, #EA580C 100%);
                color: white;
                border-radius: 10px;
                font-weight: 800 !important;
                font-size: 16px !important;
                padding: 10px 20px !important;
                border: none;
                box-shadow: 0 4px 6px rgba(249, 115, 22, 0.2);
            }
            div[data-testid="stColumn"] div[data-testid="stButton"] > button:hover,
            div[data-testid="stDownloadButton"] > button:hover {
                background: linear-gradient(135deg, #EA580C 0%, #C2410C 100%);
            }

            .stDataFrame {
                font-size: 18px !important;
            }
        </style>
    """, unsafe_allow_html=True)

# 🟢 DÁN ĐOẠN CODE NÀY VÀO ĐÂY:
@st.dialog("🎉 Ghi nhận Điểm Danh 🎉")
def success_attendance_dialog(selected_name, nq_title, formatted_time):
    st.markdown(f"""
        <div style="text-align: center; padding: 10px;">
            <h3 style="color: #1E40AF !important; margin-bottom: 9px;">Xác nhận thành công! 👌</h2>
            <p style="font-size: 16px; color: #1E293B; margin-bottom: 5px;">Chi bộ cảm ơn đồng chí: <b>{selected_name}</b></p>
            <p style="font-size: 14px; color: #64748B;">Thời gian: {formatted_time}</p>
            <p style="font-size: 15px; color: #B91C1C; font-style: italic; margin-top: 10px;">"Chúc đồng chí sức khỏe và hoàn thành tốt nhiệm vụ!"</p>
        </div>
    """, unsafe_allow_html=True)
    
    st.write("")
    if st.button("🚀 Hoàn tất 🚀", use_container_width=True, key="btn_close_success_popup"):
        st.rerun()

@st.dialog("⚠️ Xác Nhận Xóa Tiêu Đề")
def delete_confirmation_dialog(target_title):
    st.markdown(f"Bạn có chắc chắn muốn xóa tiêu đề **'{target_title}'** này không?")
    st.write("")
    
    if st.button("🗑️ Đồng ý xóa", use_container_width=True, key="btn_confirm_delete"):
        df_titles = load_titles()
        df_titles = df_titles[df_titles["Tên Tiêu đề"] != target_title]
        save_titles(df_titles)
        st.success(f"Đã xóa thành công tiêu đề '{target_title}' (đã đồng bộ Cloud).")
        st.rerun()
        
    st.write("")
    if st.button("❌ Hủy bỏ", use_container_width=True, key="btn_cancel_delete"):
        st.rerun()


@st.dialog("⚠️ Xác Nhận Xóa Lượt Điểm Danh")
def delete_single_attendance_dialog(row_index_to_delete, row_data):
    st.markdown(f"Bạn có chắc chắn muốn xóa lượt điểm danh của đồng chí **{row_data['Họ tên']}** trong sự kiện **'{row_data['Nội dung']}'** không?")
    st.write("")
    
    if st.button("🗑️ Đồng ý xóa", use_container_width=True, key="btn_confirm_delete_single"):
        if os.path.exists(ATTENDANCE_FILE):
            df_att = pd.read_excel(ATTENDANCE_FILE)
            if "Nội dung Nghị quyết" in df_att.columns:
                df_att = df_att.rename(columns={"Nội dung Nghị quyết": "Nội dung"})
                
            if row_index_to_delete in df_att.index:
                df_att = df_att.drop(index=row_index_to_delete).reset_index(drop=True)
                df_att.to_excel(ATTENDANCE_FILE, index=False)
                
                # Đồng bộ giữ nguyên chuỗi Base64 đầy đủ lên Google Sheets
                try:
                    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                    creds_dict = dict(st.secrets["gcp_service_account"])
                    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
                    client = gspread.authorize(creds)
                    spreadsheet = client.open(SHEET_NAME)
                    ws = spreadsheet.worksheet("ket_qua_diem_danh")
                    ws.clear()
                    
                    df_cloud = df_att.copy()
                    df_cloud = df_cloud.fillna("")
                    
                    if not df_cloud.empty:
                        ws.update([df_cloud.columns.values.tolist()] + df_cloud.astype(str).values.tolist())
                    else:
                        if len(df_cloud.columns) > 0:
                            ws.update([df_cloud.columns.values.tolist()])
                except Exception:
                    pass

                st.success(f"Đã xóa thành công lượt điểm danh của: {row_data['Họ tên']}")
                st.rerun()
            else:
                st.error("❌ Không tìm thấy dòng dữ liệu cần xóa trong file!")
        
    st.write("")
    if st.button("❌ Hủy bỏ", use_container_width=True, key="btn_cancel_delete_single"):
        st.rerun()


@st.dialog("⚠️ Cảnh Báo: Xóa Tất Cả Điểm Danh")
def delete_all_attendance_dialog():
    st.markdown("Bạn có thực sự muốn **XÓA TẤT CẢ** dữ liệu điểm danh của toàn bộ các sự kiện không? Hành động này không thể hoàn tác!")
    st.write("")
    
    if st.button("🚨 Đồng ý xóa sạch", use_container_width=True, key="btn_confirm_delete_all"):
        if os.path.exists(ATTENDANCE_FILE):
            df_att = pd.read_excel(ATTENDANCE_FILE)
            if "Nội dung Nghị quyết" in df_att.columns:
                df_att = df_att.rename(columns={"Nội dung Nghị quyết": "Nội dung"})
            df_empty = pd.DataFrame(columns=df_att.columns)
            df_empty.to_excel(ATTENDANCE_FILE, index=False)
        else:
            df_empty = pd.DataFrame(columns=["Nội dung", "Ngày học", "Họ tên", "Phòng ban", "Chức vụ", "Thời gian điểm danh", "Mã Ảnh Drive"])
            df_empty.to_excel(ATTENDANCE_FILE, index=False)

        # Làm sạch và đồng bộ trạng thái trống lên Google Sheets
        try:
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            client = gspread.authorize(creds)
            spreadsheet = client.open(SHEET_NAME)
            ws = spreadsheet.worksheet("ket_qua_diem_danh")
            ws.clear()
            if len(df_empty.columns) > 0:
                ws.update([df_empty.columns.values.tolist()])
        except Exception:
            pass

        st.success("Đã xóa toàn bộ lịch sử điểm danh thành công.")
        st.rerun()
        
    st.write("")
    if st.button("❌ Hủy bỏ", use_container_width=True, key="btn_cancel_delete_all"):
        st.rerun()


def main():
    st.set_page_config(
        page_title="TTTM SATRA Phạm Hùng - Hệ Thống Điểm Danh", layout="wide"
    )
    
    sync_from_google_to_local()
    apply_custom_css()

    query_params = st.query_params

    if "logged_in" not in st.session_state:
        if query_params.get("logged_in") == "true":
            st.session_state["logged_in"] = True
            st.session_state["username"] = query_params.get("username", "admin")
            st.session_state["role"] = query_params.get("role", "Quản trị viên (Admin)")
        else:
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["role"] = ""

    st.markdown("""
        <div class="app-header">
            <p class="main-title">🏢 CHI BỘ TTTM SATRA PHẠM HÙNG</p>
            <p class="sub-title">📋 HỆ THỐNG QUẢN LÝ ĐIỂM DANH</p>
            <p class="ai-badge">✨ Ứng dụng hỗ trợ bởi Trí tuệ Nhân tạo (AI)</p>
        </div>
    """, unsafe_allow_html=True)

    is_checkin_page = "nq" in query_params
    df_nhansu = load_data()

    # ========================== GIAO DIỆN ĐIỂM DANH QUA QR ==========================
    if is_checkin_page:
        nq_title = query_params.get("nq", "Học nghị quyết")
        nq_date = query_params.get("date", "")
        exp_timestamp = query_params.get("exp", "")

        is_expired = False
        if exp_timestamp:
            try:
                exp_time = datetime.fromtimestamp(float(exp_timestamp), ZoneInfo("Asia/Ho_Chi_Minh"))
                current_time = datetime.now(ZoneInfo("Asia/Ho_Chi_Minh"))
                if current_time > exp_time:
                    is_expired = True
            except Exception:
                pass

        st.markdown(f"""
            <div style="background-color: #FFFFFF; padding: 35px; border-radius: 16px; text-align: center; margin: 0 auto; max-width: 700px; box-shadow: 0 10px 25px rgba(0,0,0,0.05); border: 1px solid #E2E8F0;">
                <h2 style="color: #0F172A; margin: 0; font-size: 30px !important;">📌 {nq_title}</h2>
                <p style="color: #475569; font-size: 22px !important; margin-top: 15px; font-weight: 600;">📅 Ngày học: {nq_date}</p>
                <hr style="margin: 20px 0; border: none; border-top: 1px solid #E2E8F0;">
            </div>
        """, unsafe_allow_html=True)

        if is_expired:
            st.error("🚨 **Mã QR này đã hết hạn hiệu lực (quá 15 phút).** Vui lòng liên hệ ban tổ chức để quét mã QR mới nhất!")
        else:
            col_c1, col_c2, col_c3 = st.columns([1, 2, 1])
            with col_c2:
                all_names = df_nhansu["Họ tên"].dropna().unique().tolist()

                search_keyword = st.text_input("🔍 Gõ tên để lọc nhanh (hỗ trợ iPhone):", "", placeholder="Nhập tên hoặc họ...")
                
                if search_keyword.strip():
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
                st.markdown("<p style='font-size: 17px; font-weight: 700; margin-top: 10px; margin-bottom: 5px; color: #1E293B;'>📸 Chụp ảnh xác thực khuôn mặt:</p>", unsafe_allow_html=True)
                
                camera_photo = st.camera_input("Đưa khuôn mặt vào giữa khung hình và nhấn 'Take Photo'")

                st.write("")
                if st.button("✅ XÁC NHẬN ĐIỂM DANH", use_container_width=True):
                    if selected_name == "-- Chọn họ tên --":
                        st.error("⚠️ Vui lòng chọn hoặc gõ tìm họ tên của đồng chí!")
                    elif camera_photo is None:
                        st.warning("⚠️ Đồng chí vui lòng chụp ảnh khuôn mặt trước khi bấm xác nhận!")
                    else:
                        vn_time = datetime.now(ZoneInfo("Asia/Ho_Chi_Minh"))
                        formatted_time = vn_time.strftime("%d/%m/%Y %H:%M:%S")
                        
                        with st.spinner("Đang xử lý và đồng bộ lên Cloud..."):
                            image_bytes = camera_photo.getvalue()
                            image_base64 = convert_image_to_base64(image_bytes)

                            record_data = {
                                "Nội dung": nq_title,
                                "Ngày học": nq_date,
                                "Họ tên": selected_name,
                                "Phòng ban": default_pb,
                                "Chức vụ": default_cv,
                                "Thời gian điểm danh": formatted_time,
                                "Mã Ảnh Drive": image_base64 if image_base64 else "Chưa lưu"
                            }
                            
                            # 1. Lưu xuống file Excel cục bộ trước
                            if os.path.exists(ATTENDANCE_FILE):
                                df_att = pd.read_excel(ATTENDANCE_FILE)
                                if "Nội dung Nghị quyết" in df_att.columns:
                                    df_att = df_att.rename(columns={"Nội dung Nghị quyết": "Nội dung"})
                                df_att = pd.concat([df_att, pd.DataFrame([record_data])], ignore_index=True)
                            else:
                                df_att = pd.DataFrame([record_data])
                            df_att.to_excel(ATTENDANCE_FILE, index=False)
                            
                            # 2. Gọi trực tiếp hàm sync_to_google() để đồng bộ toàn bộ lên Cloud an toàn tuyệt đối
                            sync_success = sync_to_google()

                        # 3. Chỉ khi đồng bộ thành công trên Cloud mới hiển thị Popup và hiệu ứng ngay trong 1 lần bấm duy nhất
                        if sync_success:
                            # Xóa bộ nhớ tạm của camera để người tiếp theo điểm danh không bị kẹt ảnh cũ
                            for key in list(st.session_state.keys()):
                                if "camera_input" in key:
                                    st.session_state.pop(key)
                                    
                            st.balloons()
                            success_attendance_dialog(selected_name, nq_title, formatted_time)
                        else:
                            st.error("❌ Đồng bộ lên Google Sheets thất bại! Vui lòng kiểm tra lại mạng và bấm xác nhận lại.")

    # ========================== GIAO DIỆN QUẢN TRỊ VIÊN ==========================
    else:
        if not st.session_state["logged_in"]:
            col_l1, col_l2, col_l3 = st.columns([1, 1.2, 1])
            with col_l2:
                st.markdown("<h3 style='text-align: center; margin-top: 20px;'>🔐 ĐĂNG NHẬP HỆ THỐNG</h3>", unsafe_allow_html=True)
                with st.form("login_form"):
                    input_user = st.text_input("Tên đăng nhập:")
                    input_pass = st.text_input("Mật khẩu:", type="password")
                    remember_me = st.checkbox("Duy trì đăng nhập", value=True)
                    submit_login = st.form_submit_button("🚀 Đăng Nhập", use_container_width=True)

                    if submit_login:
                        df_users = load_users()
                        matched_u = df_users[(df_users["Tên đăng nhập"] == input_user.strip()) & (df_users["Mật khẩu"] == input_pass.strip())]
                        if not matched_u.empty:
                            st.session_state["logged_in"] = True
                            st.session_state["username"] = input_user.strip()
                            st.session_state["role"] = str(matched_u.iloc[0]["Quyền hạn"])
                            
                            if remember_me:
                                st.query_params["logged_in"] = "true"
                                st.query_params["username"] = st.session_state["username"]
                                st.query_params["role"] = st.session_state["role"]
                                if "tab" not in st.query_params:
                                    st.query_params["tab"] = "0"

                            st.success("🎉 Đăng nhập thành công!")
                            st.rerun()
                        else:
                            st.error("❌ Tên đăng nhập hoặc mật khẩu không chính xác!")
            return

        col_top1, col_top2 = st.columns([8, 2])
        with col_top1:
            st.info(f"👤 Xin chào: **{st.session_state['username']}** | Phân quyền: **{st.session_state['role']}**")
        with col_top2:
            if st.button("🚪 Đăng Xuất", use_container_width=True):
                st.session_state["logged_in"] = False
                st.session_state["username"] = ""
                st.session_state["role"] = ""
                st.query_params.clear()
                st.rerun()

        if st.session_state["role"] == "Quản trị viên (Admin)":
            tab_labels = [
                "🎯 1. Tạo QR", 
                "📊 2. Điểm Danh", 
                "👥 3. Nhân Sự", 
                "🔐 4. Quản Trị User"
            ]
        else:
            tab_labels = [
                "🎯 1. Tạo QR", 
                "📊 2. Điểm Danh"
            ]

        current_tab_str = query_params.get("tab", "0")
        try:
            current_tab_idx = int(current_tab_str)
            if current_tab_idx < 0 or current_tab_idx >= len(tab_labels):
                current_tab_idx = 0
        except Exception:
            current_tab_idx = 0

        rendered_tabs = st.tabs(tab_labels)

        # ------------------ TAB 1: TẠO MÃ QR (GIỮ NGUYÊN FORM GỐC TRÁI - PHẢI) ------------------
        with rendered_tabs[0]:
            col_left, col_right = st.columns([2, 1], gap="large")

            df_titles = load_titles()

            current_vn_date = datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).date()

            # Khởi tạo session state lưu tiêu đề và ngày chọn
            if "selected_title_input" not in st.session_state:
                st.session_state["selected_title_input"] = ""
            if "selected_date_input" not in st.session_state:
                st.session_state["selected_date_input"] = current_vn_date

            with col_left:
                st.markdown("### 🛠️ Thiết Lập Mã QR Điểm Danh")
                st.write("")

                col_lbl1, col_input1 = st.columns([1.5, 8.5], gap="small")
                with col_lbl1:
                    st.markdown("<p style='margin-top: 8px; font-size: 17px; font-weight: 700; color: #1E293B;'>Tiêu đề:</p>", unsafe_allow_html=True)
                with col_input1:
                    nq_title_input = st.text_input("Nhập tiêu đề", value=st.session_state["selected_title_input"], placeholder="Nhập tên tiêu đề sự kiện / nghị quyết...", label_visibility="collapsed")

                st.write("")

                col_lbl2, col_date2 = st.columns([1.5, 8.5], gap="small")
                with col_lbl2:
                    st.markdown("<p style='margin-top: 8px; font-size: 17px; font-weight: 700; color: #1E293B;'>Ngày:</p>", unsafe_allow_html=True)
                with col_date2:
                    nq_date_input = st.date_input("Chọn ngày/tháng/năm", value=st.session_state["selected_date_input"], label_visibility="collapsed", format="DD/MM/YYYY")

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
                    
                    # Lắng nghe sự kiện click chọn dòng trong bảng để cập nhật ngay vào session state
                    selected_rows = event.get("selection", {}).get("rows", [])
                    if selected_rows:
                        selected_idx = selected_rows[0]
                        if selected_idx < len(df_titles):
                            new_title = str(df_titles.iloc[selected_idx]["Tên Tiêu đề"])
                            raw_date = str(df_titles.iloc[selected_idx]["Ngày học"])
                            try:
                                new_date = datetime.strptime(raw_date.strip(), "%d/%m/%Y").date()
                            except Exception:
                                new_date = current_vn_date
                            
                            # Cập nhật nếu có thay đổi để tránh vòng lặp rerun
                            if st.session_state["selected_title_input"] != new_title or st.session_state["selected_date_input"] != new_date:
                                st.session_state["selected_title_input"] = new_title
                                st.session_state["selected_date_input"] = new_date
                                st.rerun()

            with col_right:
                create_qr_clicked = st.button("🚀 Tạo mã QRCode (Hiệu lực 15 phút)", use_container_width=True)

                if create_qr_clicked:
                    title_input = nq_title_input.strip()
                    if not title_input:
                        st.warning("⚠️ Vui lòng nhập tiêu đề!")
                    elif nq_date_input < current_vn_date:
                        st.error(f"🚨 Tiêu đề '{title_input}' có ngày học ({formatted_date_str}) đã nhỏ hơn ngày hiện tại ({current_vn_date.strftime('%d/%m/%Y')}). Sự kiện này đã kết thúc, không thể tạo mã QR!")
                    else:
                        expire_time = datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")) + timedelta(minutes=15)
                        expire_timestamp = expire_time.timestamp()

                        current_host = "https://app-diem-danh-nx2uwapdvmixmcuze7cjzn.streamlit.app"
                        qr_url = f"{current_host}/?nq={title_input}&date={formatted_date_str}&exp={expire_timestamp}"

                        st.session_state["qr_url"] = qr_url
                        st.session_state["nq_title"] = title_input
                        st.session_state["expire_timestamp"] = expire_timestamp

                        qr = qrcode.QRCode(version=1, box_size=10, border=5)
                        qr.add_data(qr_url)
                        qr.make(fit=True)
                        img = qr.make_image(fill_color="black", back_color="white")
                        img.save("temp_qr.png")

                        df_titles_current = load_titles()
                        if title_input in df_titles_current["Tên Tiêu đề"].values:
                            st.success("✨ Đã tạo mã QR mới thành công!")
                        else:
                            new_row = pd.DataFrame([{"Tên Tiêu đề": title_input, "Ngày học": formatted_date_str}])
                            df_titles_current = pd.concat([df_titles_current, new_row], ignore_index=True)
                            
                            # 🟢 Lưu trực tiếp xuống file Excel cục bộ trước, sau đó mới gọi hàm đồng bộ
                            df_titles_current.to_excel(TITLES_FILE, index=False)
                            save_titles(df_titles_current)
                            
                            st.success("✨ Đã tạo mã QR và đồng bộ lên Cloud thành công!")
                            st.rerun()

                st.write("")

                if "temp_qr.png" in os.listdir() and "expire_timestamp" in st.session_state:
                    st.image("temp_qr.png", caption=st.session_state.get("nq_title", ""), width=280)
                    
                    exp_ms = int(st.session_state["expire_timestamp"] * 1000)
                    countdown_html = """
                    <div style="text-align: center; font-size: 15px; font-weight: bold; color: #DC2626; background-color: #FEF2F2; padding: 8px; border-radius: 8px; border: 1px solid #FCA5A5; margin-bottom: 10px;">
                        ⏳ Mã QR sẽ hết hạn sau: <span id="countdown" style="font-size: 16px;">--:--</span>
                    </div>
                    <script>
                        var countDownDate = %s;
                        var x = setInterval(function() {
                            var now = new Date().getTime();
                            var distance = countDownDate - now;
                            var minutes = Math.floor((distance %% (1000 * 60 * 60)) / (1000 * 60));
                            var seconds = Math.floor((distance %% (1000 * 60)) / 1000);
                            
                            if (minutes < 10) minutes = "0" + minutes;
                            if (seconds < 10) seconds = "0" + seconds;

                            if (distance < 0) {
                                clearInterval(x);
                                document.getElementById("countdown").innerHTML = "ĐÃ HẾT HẠN!";
                                document.getElementById("countdown").style.color = "red";
                            } else {
                                document.getElementById("countdown").innerHTML = minutes + " phút " + seconds + " giây";
                            }
                        }, 1000);
                    </script>
                    """ % exp_ms
                    
                    countdown_html = countdown_html.replace("%%", "%")
                    components.html(countdown_html, height=50)

                    with open("temp_qr.png", "rb") as file:
                        st.download_button(
                            label="📥 Tải xuống mã QRCode",
                            data=file,
                            file_name="qr_diem_danh.png",
                            mime="image/png",
                            use_container_width=True,
                        )
                else:
                    st.info("ℹ️ Chưa có mã QR nào được tạo.")

                st.write("")

                delete_title_clicked = st.button("🗑️ Xóa Tiêu đề", use_container_width=True)
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
                            col_check = "Nội dung" if "Nội dung" in df_att_check.columns else ("Nội dung Nghị quyết" if "Nội dung Nghị quyết" in df_att_check.columns else None)
                            if col_check and target_title in df_att_check[col_check].values:
                                has_transaction = True

                        if has_transaction:
                            st.error(f"❌ Không thể xóa tiêu đề '{target_title}' vì tiêu đề này đã có người điểm danh.")
                        else:
                            delete_confirmation_dialog(target_title)

                st.write("")

                if not df_titles.empty:
                    output_titles = io.BytesIO()
                    with pd.ExcelWriter(output_titles, engine='openpyxl') as writer:
                        df_titles.to_excel(writer, index=False)
                    titles_excel_data = output_titles.getvalue()

                    st.download_button(
                        label="📥 Tải DS Tiêu Đề (Excel)",
                        data=titles_excel_data,
                        file_name="danh_sach_tieu_de.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

        # ------------------ TAB 2: ĐIỂM DANH & BÁO CÁO ------------------
        with rendered_tabs[1]:
            st.markdown("### 📈 Thống Kê & Báo Cáo Điểm Danh")
            if os.path.exists(ATTENDANCE_FILE):
                df_att = pd.read_excel(ATTENDANCE_FILE)

                if "Nội dung Nghị quyết" in df_att.columns:
                    df_att = df_att.rename(columns={"Nội dung Nghị quyết": "Nội dung"})
                    df_att.to_excel(ATTENDANCE_FILE, index=False)

                if "Thời gian điểm danh" in df_att.columns:
                    df_att["Thời gian điểm danh"] = pd.to_datetime(
                        df_att["Thời gian điểm danh"], dayfirst=True, errors='coerce'
                    ).dt.strftime("%d/%m/%Y %H:%M:%S").fillna(df_att["Thời gian điểm danh"].astype(str))

                list_nq = df_att["Nội dung"].unique().tolist()
                selected_filter = st.selectbox(
                    "🔍 Lọc theo nội dung:", ["Tất cả"] + list_nq
                )

                df_filtered = df_att[df_att["Nội dung"] == selected_filter] if selected_filter != "Tất cả" else df_att

                st.write("")
                st.markdown("💡 *Bấm chọn vào dòng cần xóa hoặc xem ảnh xác thực trong bảng dưới đây:*")
                
                event_att = st.dataframe(
                    df_filtered, 
                    width="stretch", 
                    height=280, 
                    selection_mode="single-row", 
                    on_select="rerun",
                    key="attendance_dataframe"
                )

                # --- KHUNG HIỂN THỊ ẢNH XÁC THỰC CỦA DÒNG ĐƯỢC CHỌN ---
                selected_att_rows = st.session_state.get("attendance_dataframe", {}).get("selection", {}).get("rows", [])
                if selected_att_rows:
                    selected_idx_in_filtered = selected_att_rows[0]
                    if selected_idx_in_filtered < len(df_filtered):
                        row_selected = df_filtered.iloc[selected_idx_in_filtered]
                        img_data = row_selected.get("Mã Ảnh Drive", "")
                        
                        st.write("---")
                        col_img1, col_img2 = st.columns([1, 2], gap="large")
                        with col_img1:
                            st.markdown(f"#### 📸 Ảnh xác thực")
                            st.markdown(f"👤 **Họ tên:** {row_selected['Họ tên']}")
                            st.markdown(f"🏢 **Phòng ban:** {row_selected.get('Phòng ban', '')}")
                            st.markdown(f"⏱️ **Thời gian:** {row_selected['Thời gian điểm danh']}")
                        with col_img2:
                            if str(img_data).startswith("data:image/"):
                                st.markdown(f'<img src="{img_data}" width="220" style="border-radius: 10px; border: 2px solid #CBD5E1; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">', unsafe_allow_html=True)
                            else:
                                st.info("ℹ️ Không có ảnh xác thực hoặc định dạng cũ.")

                st.write("---")
                col_btn1, col_btn2, col_btn3 = st.columns(3, gap="small")
                
                with col_btn1:
                    if st.button("🗑️ Xóa dòng đã chọn", use_container_width=True):
                        if not selected_att_rows:
                            st.warning("⚠️ Vui lòng nhấp chọn một dòng điểm danh trong bảng phía trên!")
                        else:
                            selected_idx_in_filtered = selected_att_rows[0]
                            row_to_delete = df_filtered.iloc[selected_idx_in_filtered]
                            original_index = row_to_delete.name
                            delete_single_attendance_dialog(original_index, row_to_delete)

                with col_btn2:
                    btn_label = "🔥 Xóa tất cả điểm danh" if selected_filter == "Tất cả" else f"🔥 Xóa điểm danh sự kiện này"
                    if st.button(btn_label, use_container_width=True):
                        if selected_filter == "Tất cả":
                            delete_all_attendance_dialog()
                        else:
                            @st.dialog("⚠️ Xác Nhận Xóa Điểm Danh Theo Sự Kiện")
                            def delete_specific_event_dialog(target_event):
                                st.markdown(f"Bạn có chắc chắn muốn xóa **toàn bộ dữ liệu điểm danh** của sự kiện **'{target_event}'** không? Hành động này không thể hoàn tác!")
                                st.write("")
                                if st.button("🚨 Đồng ý xóa", use_container_width=True, key="btn_confirm_delete_specific"):
                                    if os.path.exists(ATTENDANCE_FILE):
                                        df_att_all = pd.read_excel(ATTENDANCE_FILE)
                                        if "Nội dung Nghị quyết" in df_att_all.columns:
                                            df_att_all = df_att_all.rename(columns={"Nội dung Nghị quyết": "Nội dung"})
                                        
                                        df_remaining = df_att_all[df_att_all["Nội dung"] != target_event]
                                        df_remaining.to_excel(ATTENDANCE_FILE, index=False)
                                        sync_to_google() 
                                        st.success(f"Đã xóa toàn bộ điểm danh của sự kiện '{target_event}' thành công.")
                                        st.rerun()
                                st.write("")
                                if st.button("❌ Hủy bỏ", use_container_width=True, key="btn_cancel_delete_specific"):
                                    st.rerun()
                            
                            delete_specific_event_dialog(selected_filter)

                with col_btn3:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_filtered.to_excel(writer, index=False)
                    excel_data = output.getvalue()

                    file_name_download = f"bao_cao_{selected_filter}.xlsx" if selected_filter != "Tất cả" else "bao_cao_tat_ca_diem_danh.xlsx"

                    st.download_button(
                        label="📥 Tải Xuống Báo Cáo",
                        data=excel_data,
                        file_name=file_name_download,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            else:
                st.info("ℹ️ Hiện tại chưa có dữ liệu điểm danh nào được ghi nhận.")

        # ------------------ TAB 3 & 4: CHỈ HIỂN THỊ VỚI ADMIN ------------------
        if st.session_state["role"] == "Quản trị viên (Admin)":
            with rendered_tabs[2]:
                st.markdown("### 📂 Quản Lý Danh Sách Nhân Sự TTTM")
                st.dataframe(df_nhansu, width="stretch", height=450)
                st.warning(
                    "💡 **Lưu ý:** Bạn có thể thay thế file `danh_sach_nhan_su.xlsx` bằng danh sách thực tế của đơn vị với đúng tên các cột tương ứng."
                )

            with rendered_tabs[3]:
                st.markdown("### 🔐 Quản Trị Hệ Thống Người Dùng")
                st.write("")

                if st.button("🔄 Đồng bộ dữ liệu thủ công ngay", use_container_width=True):
                    with st.spinner("Đang kết nối và đẩy dữ liệu lên Google Sheets..."):
                        if sync_to_google():
                            st.success("✅ Đồng bộ dữ liệu lên Google Sheets thành công!")
                        else:
                            st.error("❌ Lỗi đồng bộ Google Sheets.")
                
                st.write("---")

                df_users = load_users()

                col_u1, col_u2 = st.columns(2, gap="large")

                with col_u1:
                    st.markdown("#### ➕ Tạo Tài Khoản Mới")
                    with st.form("form_create_user"):
                        new_username = st.text_input("Tên đăng nhập:")
                        new_password = st.text_input("Mật khẩu:", type="password")
                        new_role = st.selectbox("Phân quyền:", ["Quản trị viên (Admin)", "Nhân sự (User)"])
                        
                        submit_create = st.form_submit_button("💾 Lưu Tài Khoản Mới", use_container_width=True)
                        if submit_create:
                            if not new_username.strip() or not new_password.strip():
                                st.error("⚠️ Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu!")
                            elif new_username.strip() in df_users["Tên đăng nhập"].values:
                                st.error(f"❌ Tên đăng nhập '{new_username}' đã tồn tại!")
                            else:
                                new_u_row = pd.DataFrame([{
                                    "Tên đăng nhập": new_username.strip(),
                                    "Mật khẩu": str(new_password).strip(),
                                    "Quyền hạn": new_role
                                }])
                                df_users = pd.concat([df_users, new_u_row], ignore_index=True)
                                save_users(df_users)
                                st.success(f"✨ Đã tạo thành công tài khoản: **{new_username.strip()}** (đã đồng bộ Cloud)")
                                st.rerun()

                    st.markdown("#### 🔑 Thay Đổi Mật Khẩu")
                    with st.form("form_change_password"):
                        target_user = st.selectbox("Chọn tài khoản cần đổi mật khẩu:", df_users["Tên đăng nhập"].tolist())
                        new_pwd = st.text_input("Mật khẩu mới:", type="password")
                        
                        submit_change = st.form_submit_button("🔄 Cập Nhật Mật Khẩu", use_container_width=True)
                        if submit_change:
                            if not new_pwd.strip():
                                st.error("⚠️ Vui lòng nhập mật khẩu mới!")
                            else:
                                df_users.loc[df_users["Tên đăng nhập"] == target_user, "Mật khẩu"] = str(new_pwd).strip()
                                save_users(df_users)
                                st.success(f"✨ Đã đổi mật khẩu thành công cho tài khoản: **{target_user}** (đã đồng bộ Cloud)")
                                st.rerun()

                with col_u2:
                    st.markdown("#### 👥 Danh Sách Tài Khoản Hiện Tại")
                    st.dataframe(df_users, width="stretch", height=320)

                    st.write("")
                    st.markdown("#### 🗑️ Xóa Tài Khoản")
                    with st.form("form_delete_user"):
                        del_user = st.selectbox("Chọn tài khoản cần xóa:", ["-- Chọn tài khoản --"] + df_users["Tên đăng nhập"].tolist(), key="del_selectbox")
                        submit_del = st.form_submit_button("🚨 Xóa Tài Khoản Này", use_container_width=True)
                        if submit_del:
                            if del_user == "-- Chọn tài khoản --":
                                st.warning("⚠️ Vui lòng chọn tài khoản cần xóa!")
                            elif del_user == "admin":
                                st.error("❌ Không thể xóa tài khoản Quản trị viên gốc (admin)!")
                            else:
                                df_users = df_users[df_users["Tên đăng nhập"] != del_user]
                                save_users(df_users)
                                st.success(f"🗑️ Đã xóa tài khoản '{del_user}' thành công (đã đồng bộ Cloud).")
                                st.rerun()


if __name__ == "__main__":
    main()