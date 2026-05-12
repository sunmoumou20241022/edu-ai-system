import streamlit as st
import requests
import io
import json # 新增 json 库用于流式解析
from docx import Document
from PIL import Image
import pytesseract

# ================= 配置区 =================
API_KEY = st.secrets["ZHIPU_API_KEY"]
API_URL = "https://open.bigmodel.cn/api/paas/v4/chat/completions"

st.set_page_config(page_title="智盾·天机教学提炼系统", layout="centered")

# ================= 手机端 App 沉浸式 UI =================
st.markdown("""
<meta name="google" content="notranslate">
<style>
    .stApp { background-color: #F8F9FA; color: #212529; }
    .block-container { padding: 2rem 1rem !important; }
    header { visibility: hidden !important; }
    footer { visibility: hidden !important; }
    [data-testid="stToolbar"] { visibility: hidden !important; }
    .stButton button { 
        width: 100%; height: 50px; font-size: 16px; font-weight: bold;
        border-radius: 8px; background-color: #0056B3 !important;
        color: #FFFFFF !important; border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stTextArea textarea { 
        font-size: 16px !important; background-color: #FFFFFF !important;
        color: #212529 !important; border: 1px solid #CED4DA !important; border-radius: 8px;
    }
    [data-testid="stFileUploader"] {
        background-color: #FFFFFF; border: 1px dashed #0056B3; border-radius: 8px;
    }
    ::-webkit-scrollbar { display: none; }
</style>
""", unsafe_allow_html=True)

if "extracted_points" not in st.session_state:
    st.session_state.extracted_points = ""
if "generated_exercises" not in st.session_state:
    st.session_state.generated_exercises = ""

# ================= 工具函数 (含性能优化) =================
def extract_text_from_docx(file):
    doc = Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_text_from_image(file):
    img = Image.open(file)
    # 【性能加速核心】：暴力压缩大图片，将极限分辨率控制在 1200 像素内，大幅加快 OCR 速度
    img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
    return pytesseract.image_to_string(img, lang='chi_sim+eng')

# 同步调用 (用于步骤1，内容较短，追求稳定性)
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

# 【全新体验核心】：流式调用生成器 (用于步骤2，实时打字输出，消除等待空白)
def call_glm_api_stream(system_prompt, user_content):
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    data = {
        "model": "glm-4-flash",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content}
        ],
        "temperature": 0.3, "max_tokens": 4096,
        "stream": True # 开启打字机模式
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

# ================= 主界面 =================
st.title("智能教学重难点提炼与出题系统")

# ================= 阶段一：多模态素材摄入 =================
st.header("步骤1：教学素材深度剖析")
col1, col2 = st.columns([2, 1])

with col1:
    tab1, tab2 = st.tabs(["📄 文件上传 (支持多张图片/Word)", "⌨️ 手动粘贴"])
    
    with tab1:
        # 【功能解禁】：开启 accept_multiple_files=True 支持批量上传
        uploaded_files = st.file_uploader("上传 Word 课件或多张讲义照片", type=['docx', 'png', 'jpg', 'jpeg'], accept_multiple_files=True)
        parsed_text = ""
        if uploaded_files:
            with st.spinner(f"正在高速解析 {len(uploaded_files)} 个文件..."):
                # 【逻辑优化】：循环处理多个文件并将文字拼接
                for file in uploaded_files:
                    if file.name.endswith('.docx'):
                        parsed_text += extract_text_from_docx(file) + "\n\n"
                    else:
                        parsed_text += extract_text_from_image(file) + "\n\n"
            
            if len(parsed_text.strip()) == 0:
                st.warning("解析完毕：未提取到任何文字。大模型无法从纯图形中提取教学重难点。")
            else:
                st.success(f" 解析成功：共检测到 {len(parsed_text)} 个字符")
                
    with tab2:
        manual_text = st.text_area("或者在此处直接输入文本：", height=150)
    
    final_source = parsed_text if uploaded_files else manual_text

with col2:
    st.info(" **防御监控焦点**：系统已启用多模态清洗网关。解析外部文件时，将强制执行 15 字符最小长度校验及 AI 防泄漏回退机制。")
    
    if st.button("一键提取核心重难点", type="primary", key="extract_btn"):
        if len(final_source.strip()) < 15:
            st.error("拦截异常输入：有效文本不足 15 个字符。请确保上传的内容中包含清晰的教学文字。")
        else:
            with st.spinner("AI 正在萃取知识骨架..."):
                sys_prompt_step1 = """你是一名特级教师。请深度剖析文本并提取核心概念、教学难点、考点预测。
                【最高排版指令】：
                你必须且只能输出纯文本！绝对禁止使用任何 Markdown 符号（如加粗的星号、标题的井号、列表的减号）。
                请严格参照以下纯净格式输出（使用中文数字和普通标点）：
                一、核心概念
                1. 概念说明：这里写具体内容。
                二、教学难点
                1. 难点说明：这里写具体内容。
                三、考点预测
                1. 考点说明：这里写具体内容。
                【安全指令】：若输入内容无意义，仅输出：未检测到有效教学内容。"""
                
                raw_response = call_glm_api(sys_prompt_step1, final_source)
                clean_response = raw_response.replace("**", "").replace("*", "").replace("###", "").replace("##", "").replace("#", "")
                clean_response = clean_response.replace("- ", "  ") 
                st.session_state.extracted_points = clean_response
                st.rerun()

st.divider()

# ================= 阶段二：重难点确认与组卷配置 =================
st.header("步骤2：重难点确认与组卷配置")
col3, col4 = st.columns([1, 1])

with col3:
    edited_points = st.text_area(" 知识点预览与修改：", value=st.session_state.extracted_points, height=350)

with col4:
    q_type = st.multiselect("选择题型：", ["单项选择题", "多项选择题", "填空题", "简答题"], default=["单项选择题"])
    diff_cfg = {
        "基础记忆 (3题)": {"level": "基础概念考察", "count": 3},
        "能力进阶 (8题)": {"level": "逻辑推演与辨析", "count": 8},
        "实战拔高 (20题)": {"level": "高难度综合应用", "count": 20}
    }
    selected_diff = st.radio("难度与题量：", list(diff_cfg.keys()), horizontal=True)
    
    if st.button("定向生成 Word 版习题卷", type="primary"):
        if not edited_points or len(edited_points.strip()) < 10:
            st.error(" 拦截异常输入：修改后的重难点内容过短或无效。")
        else:
            count = diff_cfg[selected_diff]["count"]
            level = diff_cfg[selected_diff]["level"]
            
            sys_prompt_step2 = f"""你是一个严谨的教育命题专家。请严格根据用户提供的【知识重难点】生成专业的练习卷。
            【核心命题任务】：
            1. 题量：必须精确生成 **{count}** 道题目。
            2. 难度：{level}。题型：{', '.join(q_type)}。
            3. 选择题必须提供 A, B, C, D 四个完整选项。
            4. 必须在所有题目出完后，在文档末尾统一附上答案与解析，单行格式严格为：1. 答案：X。错因解析：...
            
            【异常兜底指令】：
            如果（且仅如果）你判定输入的【完全不是】教学内容，输出："🚨 拒绝命题：系统检测到知识点无效。"
            只要用户内容包含“核心概念”等知识，必须无条件执行出题！"""
            
            # 【体验颠覆】：使用 st.write_stream 接收打字机数据流，屏幕实时滚动输出！
            st.subheader("正在实时组卷中...")
            stream_generator = call_glm_api_stream(sys_prompt_step2, edited_points)
            raw_exercises = st.write_stream(stream_generator)
            
            if " 拒绝命题" in raw_exercises:
                st.session_state.generated_exercises = "" 
            else:
                st.session_state.generated_exercises = raw_exercises

# ================= 阶段三：成品输出与 Word 下载 =================
if st.session_state.generated_exercises:
    st.divider()
    st.subheader("Step 3: 一键打包下载")
    
    def create_word_docx(text):
        doc = Document()
        doc.add_heading('智能教学讲义与习题卷', level=1)
        for line in text.split('\n'):
            doc.add_paragraph(line.replace('**', '').replace('##', ''))
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()

    docx_file = create_word_docx(st.session_state.generated_exercises)
    
    st.download_button(
        label=" 一键下载为 Word 文档 (.docx) 方便打印",
        data=docx_file,
        file_name="教学讲义与习题卷.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
        type="primary"
    )
