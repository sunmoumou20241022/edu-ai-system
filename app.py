import streamlit as st
import requests
import io
import json
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from PIL import Image
import pytesseract

# ================= 配置区 =================
API_KEY = st.secrets.get("ZHIPU_API_KEY", "你的API_KEY填这里")
API_URL = "https://open.bigmodel.cn/api/paas/v4/chat/completions"

st.set_page_config(page_title="智盾·天机教学提炼系统", layout="centered")

# ================= 沉浸式 UI =================
st.markdown("""
<meta name="google" content="notranslate">
<style>
    .stApp { background-color: #F8F9FA; color: #212529; }
    .block-container { padding: 2rem 1rem !important; }
    header { visibility: hidden !important; }
    footer { visibility: hidden !important; }
    .stButton button { 
        width: 100%; height: 50px; font-size: 16px; font-weight: bold;
        border-radius: 8px; background-color: #0056B3 !important;
        color: #FFFFFF !important; border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stTextArea textarea, .stTextInput input { 
        font-size: 15px !important; background-color: #FFFFFF !important;
        color: #212529 !important; border: 1px solid #CED4DA !important; border-radius: 8px;
    }
    [data-testid="stFileUploader"] {
        background-color: #FFFFFF; border: 1px dashed #0056B3; border-radius: 8px;
    }
    .question-card { background: #fff; padding: 15px; border-radius: 8px; border: 1px solid #e0e0e0; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);}
</style>
""", unsafe_allow_html=True)

# 初始化 Session State
if "extracted_points" not in st.session_state:
    st.session_state.extracted_points = ""
if "quiz_data" not in st.session_state:
    st.session_state.quiz_data = None

# ================= 工具函数 =================
def extract_text_from_docx(file):
    doc = Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_text_from_image(file):
    img = Image.open(file)
    img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
    return pytesseract.image_to_string(img, lang='chi_sim+eng')

def call_glm_api(system_prompt, user_content):
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    data = {
        "model": "glm-4-flash",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content}
        ],
        "temperature": 0.3, "max_tokens": 4096 
    }
    try:
        response = requests.post(API_URL, headers=headers, json=data)
        return response.json()['choices'][0]['message']['content']
    except Exception as e:
        return f"API异常: {e}"

def extract_json_from_text(text):
    match = re.search(r"
http://googleusercontent.com/immersive_entry_chip/0
