import streamlit as st
import requests
import io
import time
from docx import Document
from PIL import Image
import pytesseract
# 强制指定底层 OCR 引擎的绝对路径（请根据你实际安装的位置进行修改）

# ================= 配置区 =================
API_KEY = st.secrets["ZHIPU_API_KEY"]
API_URL = "https://open.bigmodel.cn/api/paas/v4/chat/completions"

st.set_page_config(page_title="智盾·天机教学提炼系统", layout="wide")

# ================= 手机端 App 沉浸式 UI 与底层伪装载荷 =================
st.markdown("""
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="mobile-web-app-capable" content="yes">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">

<style>
    /* 全局暗黑背景，防止滑动时露出白边 */
    .stApp { background-color: #1A1A1A; color: #E0E0E0; }
    
    /* 强行抹除网页端特有的宽边距，100% 贴合手机屏幕边缘 */
    .block-container { padding: 1rem 0.5rem !important; max-width: 100vw; }
    
    /* 隐藏 Streamlit 默认的导航栏、页脚菜单、以及右上角的三条杠 */
    header { visibility: hidden !important; }
    footer { visibility: hidden !important; }
    [data-testid="stToolbar"] { visibility: hidden !important; }
    
    /* 增大按钮的触控面积，防止手机端误触 */
    .stButton button { 
        width: 100%; 
        height: 50px; 
        font-size: 16px; 
        font-weight: bold;
        border-radius: 10px; 
    }
    
    /* 调整输入框在手机上的文字可读性，防止 iOS 自动放大屏幕 */
    .stTextArea textarea { font-size: 16px !important; }
    
    /* 隐藏手机自带的滚动条，增加沉浸感 */
    ::-webkit-scrollbar { display: none; }
</style>
""", unsafe_allow_html=True)
if "extracted_points" not in st.session_state:
    st.session_state.extracted_points = ""
if "generated_exercises" not in st.session_state:
    st.session_state.generated_exercises = ""

# ================= 工具函数 =================
def extract_text_from_docx(file):
    doc = Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_text_from_image(file):
    img = Image.open(file)
    # 使用 OCR 提取文字
    return pytesseract.image_to_string(img, lang='chi_sim+eng')

def call_glm_api(system_prompt, user_content):
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    data = {
        "model": "glm-4-flash",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content}
        ],
        "temperature": 0.3,
        "max_tokens": 4096 
    }
    try:
        response = requests.post(API_URL, headers=headers, json=data)
        return response.json()['choices'][0]['message']['content']
    except Exception as e:
        return f"API异常: {e}"

# ================= 主界面 =================
st.title(" 智能教学重难点提炼与出题系统")

# ================= 阶段一：多模态素材摄入 =================
st.header("步骤1：教学素材深度剖析")
col1, col2 = st.columns([2, 1])

with col1:
    tab1, tab2 = st.tabs(["📄 文件上传 (Word/图片)", "⌨️ 手动粘贴"])
    
    with tab1:
        uploaded_file = st.file_uploader("上传 Word 课件或讲义照片", type=['docx', 'png', 'jpg', 'jpeg'])
        parsed_text = ""
        if uploaded_file:
            with st.spinner("正在解析文件内容..."):
                if uploaded_file.name.endswith('.docx'):
                    parsed_text = extract_text_from_docx(uploaded_file)
                else:
                    parsed_text = extract_text_from_image(uploaded_file)
            
            # 【优化点1】：针对无字图片的友好提示
            if len(parsed_text.strip()) == 0:
                st.warning("解析完毕：未提取到任何文字。大模型无法从纯图形中提取教学重难点。")
            else:
                st.success(f"解析成功：检测到 {len(parsed_text)} 个字符")
                
    with tab2:
        manual_text = st.text_area("或者在此处直接输入文本：", height=150)
    
    final_source = parsed_text if uploaded_file else manual_text

with col2:
    st.info(" **防御监控焦点**：系统已启用多模态清洗网关。解析外部文件时，将强制执行 15 字符最小长度校验及 AI 防泄漏回退机制。")
    
    # 【优化点2】：注入 key="extract_btn"，彻底解决 DuplicateElementId 崩溃！
    if st.button(" 一键提取核心重难点", type="primary", key="extract_btn"):
        if len(final_source.strip()) < 15:
            st.error(" 拦截异常输入：有效文本不足 15 个字符。请确保上传的图片中包含清晰的教学文字，而非纯图形。")
        else:
            with st.spinner("AI 正在萃取知识骨架并执行排版清洗..."):
                sys_prompt = """你是一名特级教师。请深度剖析文本并提取核心概念、教学难点、考点预测。
                【最高排版指令】：
                你必须且只能输出纯文本！绝对禁止使用任何 Markdown 符号（如加粗的星号、标题的井号、列表的减号）。
                
                请严格参照以下纯净格式输出（使用中文数字和普通标点）：
                一、核心概念
                1. 概念说明：这里写具体内容。
                
                二、教学难点
                1. 难点说明：这里写具体内容。
                
                三、考点预测
                1. 考点说明：这里写具体内容。
                
                【安全指令】：若输入内容无意义，仅输出： 未检测到有效教学内容。"""
                
                raw_response = call_glm_api(sys_prompt, final_source)
                
                # 物理清洗残留的 Markdown 符号
                clean_response = raw_response.replace("**", "").replace("*", "").replace("###", "").replace("##", "").replace("#", "")
                clean_response = clean_response.replace("- ", "  ") 
                
                st.session_state.extracted_points = clean_response
                st.rerun()
# 阶段二：预览、修改与组卷 (保持平整布局)
st.header("步骤2：重难点确认与组卷配置")
col3, col4 = st.columns([1, 1])

with col3:
    edited_points = st.text_area(" 知识点预览与修改 (建议删除解析错误或冗余信息)：", 
                                 value=st.session_state.extracted_points, height=350)

with col4:
    q_type = st.multiselect("选择题型：", ["单项选择题", "多项选择题", "填空题"], default=["单项选择题"])
    diff_cfg = {
        "基础 (3题)": 3,
        "进阶 (8题)": 8,
        "实战 (20题)": 20
    }
    selected_diff = st.radio("难度与题量：", list(diff_cfg.keys()), horizontal=True)
    
    if st.button("定向生成 Word 版习题卷", type="primary"):
        if edited_points:
            with st.spinner(f"正在严格按照重难点命题..."):
                count = diff_cfg[selected_diff]
                sys_prompt = f"你是个命题专家。基于给定知识点，精确生成 {count} 道题目。题型：{q_type}。格式锁死：1. 答案：X。错因解析：..."
                st.session_state.generated_exercises = call_glm_api(sys_prompt, edited_points)
        else:
            st.error("请先完成步骤1。")

# ================= 阶段三：成品输出与 Word 下载 =================
output_placeholder = st.empty()
if st.session_state.generated_exercises:
    with output_placeholder.container():
        st.divider()
        st.subheader("Step 3: 最终教学讲义与习题卷")
        
        # 将生成的文本实时转换为 Word 文档流
        def create_word_docx(text):
            doc = Document()
            doc.add_heading('智能教学讲义与习题卷', level=1)
            for line in text.split('\n'):
                # 简单过滤可能残留的markdown符号
                clean_line = line.replace('**', '').replace('##', '')
                doc.add_paragraph(clean_line)
            bio = io.BytesIO()
            doc.save(bio)
            return bio.getvalue()

        docx_file = create_word_docx(st.session_state.generated_exercises)
        
        # 放置醒目的下载按钮
        st.download_button(
            label="📄 一键下载为 Word 文档 (.docx) 方便打印",
            data=docx_file,
            file_name="教学讲义与习题卷.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary"
        )
        
        st.markdown(st.session_state.generated_exercises)
