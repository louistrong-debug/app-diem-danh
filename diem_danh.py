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
import plotly.express as px

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
            
            ws.clear()
            
            df = df.fillna("")
            if not df.empty:
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
            df_t = pd.read_excel(TITLES_FILE)
            if "Tên Tiêu đề" in df_t.columns:
                df_t = df_t.rename(columns={"Tên Tiêu đề": "Sự kiện"})
            if "Ngày học" in df_t.columns:
                df_t = df_t.rename(columns={"Ngày học": "Ngày tổ chức"})
            if "Sự kiện" not in df_t.columns:
                df_t["Sự kiện"] = []
            if "Ngày tổ chức" not in df_t.columns:
                df_t["Ngày tổ chức"] = []
            return df_t
        except Exception:
            return pd.DataFrame(columns=["Sự kiện", "Ngày tổ chức"])
    else:
        return pd.DataFrame(columns=["Sự kiện", "Ngày tổ chức"])


def save_titles(df):
    df_save = df.copy()
    if "Sự kiện" in df_save.columns:
        df_save = df_save.rename(columns={"Sự kiện": "Tên Tiêu đề"})
    if "Ngày tổ chức" in df_save.columns:
        df_save = df_save.rename(columns={"Ngày tổ chức": "Ngày học"})
    df_save.to_excel(TITLES_FILE, index=False)
    
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
        </style>
    """, unsafe_allow_html=True)


@st.dialog("🟢 Ghi nhận Điểm Danh")
def success_attendance_dialog(selected_name, nq_title, formatted_time):
    st.markdown(f"""
        <div style="text-align: center; padding: 9px;">
            <h3 style="color: #1E40AF !important; margin-bottom: 9px;">Xác nhận thành công! 👌</h2>
            <p style="font-size: 16px; color: #1E293B; margin-bottom: 5px;">Cảm ơn đồng chí: <b>{selected_name}</b></p>
            <p style="font-size: 14px; color: #64748B;">Thời gian: {formatted_time}</p>
            <p style="font-size: 15px; color: #B91C1C; font-style: italic; margin-top: 10px;">"Chúc đồng chí sức khỏe và hoàn thành tốt nhiệm vụ!"</p>
        </div>
    """, unsafe_allow_html=True)
    
    st.write("")
    if st.button("🚀 Hoàn tất 🚀", use_container_width=True, key="btn_close_success_popup"):
        for key in list(st.session_state.keys()):
            if "camera_input" in key or key == "selected_name":
                st.session_state.pop(key, None)
        st.rerun()


@st.dialog("⚠️ Xác Nhận Xóa Sự Kiện")
def delete_confirmation_dialog(target_title):
    st.markdown(f"Bạn có chắc chắn muốn xóa sự kiện **'{target_title}'** này không?")
    st.write("")
    
    if st.button("🗑️ Đồng ý xóa", use_container_width=True, key="btn_confirm_delete"):
        df_titles = load_titles()
        df_titles = df_titles[df_titles["Sự kiện"] != target_title]
        save_titles(df_titles)
        st.success(f"Đã xóa thành công sự kiện '{target_title}' (đã đồng bộ Cloud).")
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
                
                try:
                    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                    creds_dict = dict(st.secrets["gcp_service_account"])
                    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
                    client = gspread.authorize(creds)
                    spreadsheet = client.open(SHEET_NAME)
                    ws = spreadsheet.worksheet("ket_qua_diem_danh")
                    ws.clear()
                    
                    df_cloud = df_att.copy().fillna("")
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
        nq_title = query_params.get("nq", "Họp chi bộ")
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
                <p style="color: #475569; font-size: 22px !important; margin-top: 15px; font-weight: 600;">📅 Ngày tổ chức: {nq_date}</p>
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
                            
                            if os.path.exists(ATTENDANCE_FILE):
                                df_att = pd.read_excel(ATTENDANCE_FILE)
                                if "Nội dung Nghị quyết" in df_att.columns:
                                    df_att = df_att.rename(columns={"Nội dung Nghị quyết": "Nội dung"})
                                df_att = pd.concat([df_att, pd.DataFrame([record_data])], ignore_index=True)
                            else:
                                df_att = pd.DataFrame([record_data])
                            df_att.to_excel(ATTENDANCE_FILE, index=False)
                            
                            sync_success = sync_to_google()

                        if sync_success:
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
                "🔐 4. Quản Trị User",
                "📈 5. Báo Cáo Chuyên Sâu"
            ]
        else:
            tab_labels = [
                "🎯 1. Tạo QR", 
                "📊 2. Điểm Danh"
            ]

        rendered_tabs = st.tabs(tab_labels)

        # ------------------ TAB 1: TẠO MÃ QR ------------------
        with rendered_tabs[0]:
            col_left, col_right = st.columns([2, 1], gap="large")

            df_titles = load_titles()
            current_vn_date = datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).date()

            if "selected_title_input" not in st.session_state:
                st.session_state["selected_title_input"] = ""
            if "selected_date_input" not in st.session_state:
                st.session_state["selected_date_input"] = current_vn_date

            with col_left:
                st.markdown("### 🛠️ Thiết Lập Mã QR Điểm Danh")
                st.write("")

                col_lbl1, col_input1 = st.columns([2.2, 7.8], gap="small")
                with col_lbl1:
                    st.markdown("<p style='margin-top: 8px; font-size: 17px; font-weight: 700; color: #1E293B;'>Sự kiện:</p>", unsafe_allow_html=True)
                with col_input1:
                    nq_title_input = st.text_input("Nhập sự kiện", value=st.session_state["selected_title_input"], placeholder="Nhập tên buổi họp, kết nạp Đảng, sự kiện...", label_visibility="collapsed")

                st.write("")

                # ĐÃ MỞ RỘNG CỘT LABEL THÀNH [2.2, 7.8] ĐỂ "Ngày tổ chức" NẰM TRÊN 1 HÀNG
                col_lbl2, col_date2 = st.columns([2.2, 7.8], gap="small")
                with col_lbl2:
                    st.markdown("<p style='margin-top: 8px; font-size: 17px; font-weight: 700; color: #1E293B;'>Ngày tổ chức:</p>", unsafe_allow_html=True)
                with col_date2:
                    nq_date_input = st.date_input("Chọn ngày", value=st.session_state["selected_date_input"], label_visibility="collapsed", format="DD/MM/YYYY")

                formatted_date_str = nq_date_input.strftime("%d/%m/%Y")

                st.write("")
                st.markdown("#### 📋 Danh Sách Sự Kiện Đã Thiết Lập")

                if df_titles.empty or "Sự kiện" not in df_titles.columns:
                    st.info("ℹ️ Chưa có sự kiện nào được thêm.")
                else:
                    df_titles_show = df_titles.copy().reset_index(drop=True)
                    df_titles_show.insert(0, "STT", range(1, len(df_titles_show) + 1))
                    
                    event = st.dataframe(
                        df_titles_show, 
                        width="stretch", 
                        height=200, 
                        selection_mode="single-row", 
                        on_select="rerun",
                        key="titles_dataframe",
                        hide_index=True
                    )
                    
                    selected_rows = event.get("selection", {}).get("rows", [])
                    if selected_rows:
                        selected_idx = selected_rows[0]
                        if selected_idx < len(df_titles):
                            new_title = str(df_titles.iloc[selected_idx]["Sự kiện"])
                            raw_date = str(df_titles.iloc[selected_idx]["Ngày tổ chức"])
                            try:
                                new_date = datetime.strptime(raw_date.strip(), "%d/%m/%Y").date()
                            except Exception:
                                new_date = current_vn_date
                            
                            if st.session_state["selected_title_input"] != new_title or st.session_state["selected_date_input"] != new_date:
                                st.session_state["selected_title_input"] = new_title
                                st.session_state["selected_date_input"] = new_date
                                st.rerun()

            with col_right:
                create_qr_clicked = st.button("🚀 Tạo mã QRCode (Hiệu lực 15 phút)", use_container_width=True)

                if create_qr_clicked:
                    title_input = nq_title_input.strip()
                    if not title_input:
                        st.warning("⚠️ Vui lòng nhập tên sự kiện!")
                    elif nq_date_input < current_vn_date:
                        st.error(f"🚨 Sự kiện '{title_input}' có ngày tổ chức ({formatted_date_str}) đã nhỏ hơn ngày hiện tại ({current_vn_date.strftime('%d/%m/%Y')}). Sự kiện này đã kết thúc, không thể tạo mã QR!")
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
                        if not df_titles_current.empty and "Sự kiện" in df_titles_current.columns and title_input in df_titles_current["Sự kiện"].values:
                            st.success("✨ Đã tạo mã QR mới thành công!")
                        else:
                            new_row = pd.DataFrame([{"Sự kiện": title_input, "Ngày tổ chức": formatted_date_str}])
                            df_titles_current = pd.concat([df_titles_current, new_row], ignore_index=True)
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

                delete_title_clicked = st.button("🗑️ Xóa Sự Kiện", use_container_width=True)
                if delete_title_clicked:
                    selected_rows = st.session_state.get("titles_dataframe", {}).get("selection", {}).get("rows", [])
                    if not selected_rows:
                        st.warning("⚠️ Vui lòng nhấp chọn một hàng trong bảng danh sách bên trái để xóa!")
                    else:
                        selected_idx = selected_rows[0]
                        target_title = df_titles.iloc[selected_idx]["Sự kiện"]
                        
                        has_transaction = False
                        if os.path.exists(ATTENDANCE_FILE):
                            df_att_check = pd.read_excel(ATTENDANCE_FILE)
                            col_check = "Nội dung" if "Nội dung" in df_att_check.columns else ("Nội dung Nghị quyết" if "Nội dung Nghị quyết" in df_att_check.columns else None)
                            if col_check and target_title in df_att_check[col_check].values:
                                has_transaction = True

                        if has_transaction:
                            st.error(f"❌ Không thể xóa sự kiện '{target_title}' vì sự kiện này đã có người điểm danh.")
                        else:
                            delete_confirmation_dialog(target_title)

                st.write("")

                if not df_titles.empty:
                    output_titles = io.BytesIO()
                    with pd.ExcelWriter(output_titles, engine='openpyxl') as writer:
                        df_titles.to_excel(writer, index=False)
                    titles_excel_data = output_titles.getvalue()

                    st.download_button(
                        label="📥 Tải DS Sự Kiện (Excel)",
                        data=titles_excel_data,
                        file_name="danh_sach_su_kien.xlsx",
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

                list_nq = df_att["Nội dung"].unique().tolist() if "Nội dung" in df_att.columns else []
                selected_filter = st.selectbox(
                    "🔍 Lọc theo sự kiện:", ["Tất cả"] + list_nq
                )

                df_filtered = df_att[df_att["Nội dung"] == selected_filter].copy() if selected_filter != "Tất cả" else df_att.copy()

                st.write("")
                st.markdown("💡 *Bấm chọn vào dòng cần xóa hoặc xem ảnh xác thực trong bảng dưới đây:*")
                
                if not df_filtered.empty:
                    df_filtered = df_filtered.reset_index(drop=True)
                    df_filtered.insert(0, "STT", range(1, len(df_filtered) + 1))

                event_att = st.dataframe(
                    df_filtered, 
                    width="stretch", 
                    height=280, 
                    selection_mode="single-row", 
                    on_select="rerun",
                    key="attendance_dataframe",
                    hide_index=True
                )

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

        # ------------------ TAB 3: NHÂN SỰ (ADMIN) ------------------
        if st.session_state["role"] == "Quản trị viên (Admin)":
            with rendered_tabs[2]:
                st.markdown("### 📂 Quản Lý Danh Sách Nhân Sự TTTM")
                df_nhansu_show = df_nhansu.copy().reset_index(drop=True)
                df_nhansu_show["STT"] = range(1, len(df_nhansu_show) + 1)
                st.dataframe(df_nhansu_show[["STT", "Họ tên", "Phòng ban", "Chức vụ"]], width="stretch", height=450, hide_index=True)
                st.warning(
                    "💡 **Lưu ý:** Bạn có thể thay thế file `danh_sach_nhan_su.xlsx` bằng danh sách thực tế của đơn vị với đúng tên các cột tương ứng."
                )

            # ------------------ TAB 4: QUẢN TRỊ USER (ADMIN) ------------------
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
                    df_users_show = df_users.copy().reset_index(drop=True)
                    df_users_show.insert(0, "STT", range(1, len(df_users_show) + 1))
                    st.dataframe(df_users_show, width="stretch", height=320, hide_index=True)

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

            # ------------------ TAB 5: BÁO CÁO CHUYÊN SÂU (ADMIN) ------------------
            with rendered_tabs[4]:
                st.markdown("### 📈 Dashboard Phân Tích & Báo Cáo Chuyên Sâu")
                st.write("")

                if not os.path.exists(ATTENDANCE_FILE) or not os.path.exists(TITLES_FILE) or not os.path.exists(EXCEL_FILE):
                    st.info("ℹ️ Chưa đủ dữ liệu từ các file (Điểm danh, Sự kiện, Nhân sự) để tổng hợp báo cáo.")
                else:
                    df_att_raw = pd.read_excel(ATTENDANCE_FILE)
                    df_titles = pd.read_excel(TITLES_FILE)
                    df_nhansu = pd.read_excel(EXCEL_FILE)

                    if "Nội dung Nghị quyết" in df_att_raw.columns:
                        df_att_raw = df_att_raw.rename(columns={"Nội dung Nghị quyết": "Nội dung"})

                    # LỌC BỎ TRÙNG LẶP: Nếu 1 người điểm danh nhiều lần trong cùng 1 sự kiện, chỉ giữ lại 1 lần duy nhất
                    df_att = df_att_raw.drop_duplicates(subset=["Nội dung", "Họ tên"]).copy()

                    total_events = len(df_titles)
                    total_staff = len(df_nhansu)
                    total_attendance_records = len(df_att)

                    # A. PHẦN TỔNG QUAN (TOP WIDGETS - KPI)
                    col_kpi1, col_kpi2, col_kpi3 = st.columns(3, gap="medium")
                    
                    with col_kpi1:
                        st.metric(label="📋 Tổng Số Lượt Điểm Danh (Đã lọc trùng)", value=f"{total_attendance_records} lượt")

                    with col_kpi2:
                        max_possible_attendance = total_staff * total_events if (total_staff > 0 and total_events > 0) else 1
                        avg_attendance_rate = (total_attendance_records / max_possible_attendance) * 100
                        if avg_attendance_rate > 100: 
                            avg_attendance_rate = 100.0
                        st.metric(label="📊 Tỉ Lệ Chuyên Cần Trung Bình", value=f"{avg_attendance_rate:.1f}%")

                    with col_kpi3:
                        latest_event = df_titles.iloc[-1]["Sự kiện"] if not df_titles.empty and "Sự kiện" in df_titles.columns else "Chưa có"
                        latest_count = len(df_att[df_att["Nội dung"] == latest_event]) if not df_titles.empty else 0
                        latest_rate = (latest_count / total_staff * 100) if total_staff > 0 else 0
                        st.metric(label="🔥 Sự Kiện Gần Nhất", value=f"{latest_rate:.1f}%", help=f"Sự kiện: {latest_event} ({latest_count}/{total_staff} nhân sự)")

                    st.write("---")

                    # B. PHẦN BIỂU ĐỒ (VISUALIZATIONS - PLOTLY)
                    col_chart1, col_chart2 = st.columns(2, gap="large")

                    with col_chart1:
                        st.markdown("#### 🏢 Tỉ Lệ Chuyên Cần Theo Phòng Ban")
                        if not df_att.empty and not df_nhansu.empty:
                            df_dept_att = df_att.groupby("Phòng ban")["Họ tên"].count().reset_index()
                            df_dept_att = df_dept_att.rename(columns={"Họ tên": "Số lượt tham gia"})
                            
                            df_dept_total = df_nhansu.groupby("Phòng ban")["Họ tên"].count().reset_index()
                            df_dept_total = df_dept_total.rename(columns={"Họ tên": "Tổng nhân sự"})

                            df_merged_dept = pd.merge(df_dept_total, df_dept_att, on="Phòng ban", how="left").fillna(0)
                            max_dept_slots = total_events if total_events > 0 else 1
                            df_merged_dept["Tỉ lệ (%)"] = (df_merged_dept["Số lượt tham gia"] / (df_merged_dept["Tổng nhân sự"] * max_dept_slots) * 100).round(1)

                            fig_bar = px.bar(
                                df_merged_dept, 
                                x="Phòng ban", 
                                y="Tỉ lệ (%)", 
                                text="Tỉ lệ (%)",
                                color="Phòng ban",
                                title="Hiệu suất tham gia theo phòng ban"
                            )
                            fig_bar.update_traces(texttemplate='%{text}%', textposition='outside')
                            fig_bar.update_layout(xaxis_title="Phòng Ban", yaxis_title="Tỉ Lệ Tham Gia (%)", showlegend=False)
                            st.plotly_chart(fig_bar, use_container_width=True)
                        else:
                            st.info("ℹ️ Chưa đủ dữ liệu để vẽ biểu đồ phòng ban.")

                    with col_chart2:
                        st.markdown("#### 📈 Xu Hướng Điểm Danh Theo Sự Kiện")
                        if not df_titles.empty and not df_att.empty and "Sự kiện" in df_titles.columns:
                            event_counts = []
                            for idx, row in df_titles.iterrows():
                                ev_title = row["Sự kiện"]
                                count = len(df_att[df_att["Nội dung"] == ev_title])
                                event_counts.append({"Sự kiện": ev_title, "Ngày tổ chức": row.get("Ngày tổ chức", ""), "Số người tham gia": count})
                            
                            df_trend = pd.DataFrame(event_counts)
                            
                            fig_line = px.line(
                                df_trend, 
                                x="Sự kiện", 
                                y="Số người tham gia", 
                                markers=True,
                                title="Biểu đồ biến động số lượng tham gia sự kiện"
                            )
                            fig_line.update_layout(xaxis_title="Tên Sự Kiện", yaxis_title="Số Lượng Người Tham Gia")
                            st.plotly_chart(fig_line, use_container_width=True)
                        else:
                            st.info("ℹ️ Chưa đủ dữ liệu xu hướng.")

                    st.write("---")

                    # C. PHẦN THỐNG KÊ CHI TIẾT TỪNG SỰ KIỆN (ĐÃ LỌC TRÙNG)
                    st.markdown("### 🔎 Thống Kê Chi Tiết Theo Từng Sự Kiện")
                    if not df_titles.empty and "Sự kiện" in df_titles.columns:
                        list_all_events = df_titles["Sự kiện"].tolist()
                        selected_detail_event = st.selectbox("👉 Chọn sự kiện cần xem chi tiết:", list_all_events, key="select_detail_event")

                        if selected_detail_event:
                            df_event_att = df_att[df_att["Nội dung"] == selected_detail_event]
                            event_attendees_count = len(df_event_att)
                            event_rate = (event_attendees_count / total_staff * 100) if total_staff > 0 else 0

                            col_ev1, col_ev2, col_ev3 = st.columns(3, gap="medium")
                            with col_ev1:
                                st.metric(label="👥 Tổng Số Người Tham Gia (Duy nhất)", value=f"{event_attendees_count} / {total_staff} nhân sự")
                            with col_ev2:
                                st.metric(label="📊 Tỉ Lệ Tham Gia Sự Kiện", value=f"{event_rate:.1f}%")
                            with col_ev3:
                                missing_count = total_staff - event_attendees_count
                                st.metric(label="❌ Số Lượng Vắng Mặt", value=f"{missing_count} nhân sự")

                            st.write("")

                            col_ev_chart, col_ev_table = st.columns([1, 1.2], gap="large")

                            with col_ev_chart:
                                st.markdown(f"#### 📊 Tỉ Lệ Phòng Ban Trong Sự Kiện")
                                if not df_event_att.empty and not df_nhansu.empty:
                                    df_ev_dept_att = df_event_att.groupby("Phòng ban")["Họ tên"].count().reset_index()
                                    df_ev_dept_att = df_ev_dept_att.rename(columns={"Họ tên": "Tham gia"})
                                    df_ev_dept_total = df_nhansu.groupby("Phòng ban")["Họ tên"].count().reset_index()
                                    df_ev_dept_total = df_ev_dept_total.rename(columns={"Họ tên": "Tổng"})

                                    df_ev_merged = pd.merge(df_ev_dept_total, df_ev_dept_att, on="Phòng ban", how="left").fillna(0)
                                    df_ev_merged["Tỉ lệ (%)"] = (df_ev_merged["Tham gia"] / df_ev_merged["Tổng"] * 100).round(1)

                                    fig_ev_bar = px.bar(
                                        df_ev_merged, 
                                        x="Phòng ban", 
                                        y="Tỉ lệ (%)", 
                                        text="Tỉ lệ (%)",
                                        color="Phòng ban",
                                        title=f"Mức độ tham gia: {selected_detail_event}"
                                    )
                                    fig_ev_bar.update_traces(texttemplate='%{text}%', textposition='outside')
                                    fig_ev_bar.update_layout(xaxis_title="Phòng Ban", yaxis_title="Tỉ Lệ (%)", showlegend=False)
                                    st.plotly_chart(fig_ev_bar, use_container_width=True)
                                else:
                                    st.info("ℹ️ Chưa có dữ liệu tham gia cho sự kiện này.")

                            with col_ev_table:
                                st.markdown(f"#### ❌ Danh Sách Vắng Mặt Sự Kiện Này")
                                if not df_nhansu.empty:
                                    attended_names = df_event_att["Họ tên"].tolist()
                                    df_missing_staff = df_nhansu[~df_nhansu["Họ tên"].isin(attended_names)].copy()
                                    if not df_missing_staff.empty:
                                        df_missing_staff = df_missing_staff.reset_index(drop=True)
                                        df_missing_staff["STT"] = range(1, len(df_missing_staff) + 1)
                                        df_missing_staff = df_missing_staff[["STT", "Họ tên", "Phòng ban", "Chức vụ"]]
                                        st.dataframe(df_missing_staff, use_container_width=True, height=280, hide_index=True)
                                    else:
                                        st.success("🎉 Tuyệt vời! Sự kiện này không có ai vắng mặt.")
                                else:
                                    st.info("ℹ️ Không có dữ liệu nhân sự.")

                    st.write("---")

                    # D. PHẦN TỔNG HỢP TOÀN BỘ NHÂN SỰ
                    st.markdown("### 📋 Bảng Tổng Hợp Tình Hình Tham Gia Sinh Hoạt")

                    if total_events > 0 and not df_nhansu.empty:
                        attendance_counts = df_att.groupby("Họ tên").size().reset_index(name="Số buổi tham gia")
                        df_summary = pd.merge(df_nhansu, attendance_counts, on="Họ tên", how="left").fillna({"Số buổi tham gia": 0})
                        df_summary["Tổng Sự kiện"] = total_events
                        df_summary["Số buổi vắng"] = total_events - df_summary["Số buổi tham gia"]
                        df_summary["Tỉ lệ tham gia (%)"] = ((df_summary["Số buổi tham gia"] / total_events) * 100).round(1)

                        # Sắp xếp mặc định Tỉ lệ tham gia (%) tăng dần (từ 0% lên cao nhất)
                        df_summary = df_summary.sort_values(by="Tỉ lệ tham gia (%)", ascending=True).reset_index(drop=True)
                        df_summary["STT"] = range(1, len(df_summary) + 1)

                        # Sắp xếp lại thứ tự cột cho đúng yêu cầu
                        df_summary_show = df_summary[["STT", "Họ tên", "Phòng ban", "Chức vụ", "Tổng Sự kiện", "Số buổi vắng", "Tỉ lệ tham gia (%)"]]

                        st.dataframe(df_summary_show, use_container_width=True, height=400, hide_index=True)
                    else:
                        st.info("ℹ️ Cần có ít nhất 1 sự kiện và danh sách nhân sự để phân tích chi tiết.")


if __name__ == "__main__":
    main()