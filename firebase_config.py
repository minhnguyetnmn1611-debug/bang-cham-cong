import firebase_admin
from firebase_admin import credentials
from firebase_admin import db
import os
import streamlit as st
from log_config import logger

# Lấy đường dẫn tuyệt đối đến file JSON
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CRED_PATH = os.path.join(BASE_DIR, "vmos---app-firebase-adminsdk-fbsvc-2bad41e050.json")

# URL của Realtime Database (đã ping thử thành công)
DB_URL = "https://vmos---app-default-rtdb.asia-southeast1.firebasedatabase.app/"

@st.cache_resource
def init_firebase():
    """Khởi tạo Firebase Admin SDK. Sử dụng st.cache_resource để không bị lặp lại."""
    if not firebase_admin._apps:
        try:
            cred = credentials.Certificate(CRED_PATH)
            firebase_admin.initialize_app(cred, {
                'databaseURL': DB_URL
            })
            logger.info("Firebase initialized successfully!")
        except Exception as e:
            logger.error(f"Error initializing Firebase: {e}")
            return False
    return True

# Gọi hàm khởi tạo
init_firebase()

def get_db_ref(path="/"):
    """Lấy reference tới một đường dẫn trong Firebase Realtime DB"""
    return db.reference(path)
