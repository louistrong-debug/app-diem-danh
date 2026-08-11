import os
import warnings
import locale
import io
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import qrcode
import streamlit as st
import streamlit.components.v1 as components

# Tắt các thông báo cảnh báo không cần thiết
warnings.filterwarnings("ignore", category=UserWarning)

# Cấu hình locale
try:
    locale.setlocale(locale.LC_COLLATE, "vi_VN.UTF-8")
except:
    pass

# File lưu danh sách
EXCEL_FILE = "danh_sach_nhan_su.xlsx"
ATTENDANCE_FILE = "ket_qua_diem_danh.xlsx"
TITLES_FILE = "danh_sach_tieu_de.xlsx"
USER_FILE = "danh_sach_user.xlsx"

SHEET_NAME = "QuanLyDiemDanh" 
FILES_MAP = {
    "danh_sach_nhan_su": EXCEL_FILE,
    "danh_sach_tieu_de": TITLES_FILE,
    "ket_qua_diem_danh": ATTENDANCE_FILE,
    "danh_sach_user": USER_FILE
}

def sync_to_google():
    try:
        # Lấy credentials từ Streamlit Secrets
        if "gcp_service_account" not in st.secrets:
            return "❌ Thiếu cấu hình Secrets trên Streamlit Cloud!"
            
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        
        client = gspread.authorize(creds)
        spreadsheet = client.open(SHEET_NAME)
        
        for sheet_key, filename in FILES_MAP.items():
            if os.path.exists(filename):
                df = pd.read_excel(filename)
                try:
                    ws = spreadsheet.worksheet(sheet_key)
                except:
                    ws = spreadsheet.add_worksheet(title=sheet_key, rows="100", cols="20")
                ws.clear()
                ws.update([df.columns.values.tolist()] + df.values.tolist())
        return "✅ Đồng bộ dữ liệu lên Google Sheets thành công!"
    except Exception as e: 
        return f"❌ Lỗi đồng bộ: {str(e)}"

# --- CÁC HÀM LOAD/SAVE DỮ LIỆU CŨ CỦA BẠN ---
def load_data():
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
    else:
        df = pd.DataFrame({"STT": [1], "Họ tên": ["Mẫu"], "Phòng ban": ["Mẫu"], "Chức vụ": ["Mẫu"]})
        df.to_excel(EXCEL_FILE, index=False)
    return df

def load_titles():
    return pd.read_excel(TITLES_FILE) if os.path.exists(TITLES_FILE) else pd.DataFrame(columns=["Tên Tiêu đề", "Ngày học"])

def save_titles(df): df.to_excel(TITLES_FILE, index=False)

def load_users():
    if os.path.exists(USER_FILE):
        return pd.read_excel(USER_FILE)
    default = pd.DataFrame({"Tên đăng nhập": ["admin"], "Mật khẩu": ["123456"], "Quyền hạn": ["Quản trị viên (Admin)"]})
    default.to_excel(USER_FILE, index=False)
    return default

def save_users(df): df.to_excel(USER_FILE, index=False)

def apply_custom_css():
    st.markdown("""<style>
        .app-header { background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%); padding: 20px; border-radius: 12px; text-align: center; color: white; }
        .main-title { color: #F97316; font-weight: 900; font-size: 25px; }
    </style>""", unsafe_allow_html=True)

# --- PHẦN GIAO DIỆN CHÍNH ---
def main():
    st.set_page_config(page_title="TTTM SATRA Phạm Hùng", layout="wide")
    apply_custom_css()
    
    # ... (giữ nguyên logic session_state đăng nhập của bạn) ...
    if "logged_in" not in st.session_state: st.session_state["logged_in"] = False
    
    st.markdown('<div class="app-header"><p class="main-title">🏢 CHI BỘ TTTM SATRA PHẠM HÙNG</p></div>', unsafe_allow_html=True)

    # Nơi đặt logic hiển thị tabs (như code gốc của bạn)
    # Lưu ý: Khi gọi hàm sync_to_google(), nó đã tự dùng st.secrets rồi.
    
    # Ví dụ trong Tab 4 Admin:
    # if st.button("🔄 Đồng bộ dữ liệu ngay"):
    #     res = sync_to_google()
    #     st.write(res)

if __name__ == "__main__":
    main()