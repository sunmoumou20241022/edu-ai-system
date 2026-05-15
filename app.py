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
API_KEY = st.secrets.get("ZHIPU_API_KEY", "你的API_KEY填这里") # 防止本地运行报错
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
    st.session_state.quiz_data = None  # 存储结构化的题目数据

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

def call_glm_api_stream(system_prompt, user_content):
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    data = {
        "model": "glm-4-flash",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content}
        ],
        "temperature": 0.3, "max_tokens": 8192,
        "stream": True 
    }
    try:
        response = requests.post(API_URL, headers=headers, json=data, stream=True)
        for line in response.iter_lines():
            if line:
                decoded_line = line.decode('utf-8')
                if decoded_line.startswith("data: "):
                    content = decoded_line[6:]
                    if content == "[DONE]": break
                    chunk = json.loads(content)
                    yield chunk['choices'][0]['delta'].get('content', '')
    except Exception as e:
        yield f"\n[网络流异常: {e}]"

def extract_json_from_text(text):
    """鲁棒的JSON提取器"""
    match = re.search(r'
http://googleusercontent.com/immersive_entry_chip/0

### 🌟 这次更新带来了哪些颠覆性改变？

1. **结构化底层逻辑 (JSON Engine)**
   我放弃了单纯渲染纯文本流，而是强制让大模型以 **JSON** 格式输出考卷数据。这让题目彻底变成了可被 Python 代码单独管理、修改和渲染的“数字资产”。
2. **可视化题目微调控制台 (步骤3)**
   你现在能看到每一个生成的题目被单独折叠在卡片里（`st.expander`）。你可以直接在输入框里修改错别字、改答案、或者修正选项顺序。**改完不需要点保存，直接点击最下方的下载，Word 文档里的字就是你刚刚改过的！**
3. **混合题型生成面板 (步骤2)**
   老师可以自由定义：**选择题出几道（单题几分）、填空题出几道（单题几分）、简答题出几道（单题几分）**。甚至能自定义试卷的大标题。
4. **企业级排版 + 双版本导出 (步骤4)**
   彻底摆脱简陋的 Word。利用深度的 `python-docx` 接口，这份代码会自动为你设置：
   * 极窄试卷页边距（省纸）。
   * 姓名/班级/考号 的下划线密封区。
   * 选择/填空题排版优化。
   * 提供了**【学生版】**（纯白卷）和**【教师版】**（带全套红字解析和翻页设计）两个下载按钮。

直接将整段代码复制覆盖你的 `app.py`，你的系统立刻就能变身为一款**商业级**的教师提效 SaaS 平台！
