import os
import io
import streamlit as st
import requests
import json

# Try to import necessary libraries
try:
    from docx import Document
except ImportError:
    st.error("❌ 缺少依赖库：python-docx。请运行 `pip install python-docx`。")
    st.stop()

# --- Page Configuration ---
st.set_page_config(page_title="DeepSeek x Word 智能助手", page_icon="🐼", layout="wide")

# --- Persistent Configuration Management ---
CONFIG_FILE = "config.json"

def load_config():
    """优先级读取 API Key：1. Session 2. 环境变量 3. 本地文件"""
    # 优先从 Streamlit Session State 获取（防止页面刷新丢失）
    if "saved_api_key" in st.session_state:
        return {"api_key": st.session_state.saved_api_key, "base_url": st.session_state.get("saved_base_url", "https://api.deepseek.com")}

    # 1. 尝试环境变量 (Streamlit Cloud Secrets)
    env_key = os.environ.get("DEEPSEEK_API_KEY", "")
    if env_key:
        return {"api_key": env_key, "base_url": "https://api.deepseek.com"}
    
    # 2. 尝试本地配置文件
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except:
            pass
    return {"api_key": "", "base_url": "https://api.deepseek.com"}

def save_config(api_key, base_url):
    """保存到 Session 和本地文件"""
    st.session_state.saved_api_key = api_key
    st.session_state.saved_base_url = base_url
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump({"api_key": api_key, "base_url": base_url}, f)
    except:
        pass # 在云端环境可能无法写入文件，忽略即可

# --- Core Logic: Word Replacement ---
def _apply_replace(paragraph, replace_dict):
    """替换段落文本并保留格式"""
    for old_text, new_text in replace_dict.items():
        if old_text in paragraph.text:
            for run in paragraph.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)

def smart_replace(doc, replace_dict):
    """扫描全文和表格"""
    for p in doc.paragraphs:
        _apply_replace(p, replace_dict)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _apply_replace(p, replace_dict)
    return doc

def parse_replace_text(text):
    """解析 '旧内容 ==>> 新内容' 格式"""
    replace_dict = {}
    for line in text.split('\n'):
        if "==>>" in line:
            parts = line.split("==>>")
            if len(parts) == 2:
                replace_dict[parts[0].strip()] = parts[1].strip()
    return replace_dict

# --- AI Logic: Calling DeepSeek API ---
def get_deepseek_rules(api_key, base_url, doc_content, user_demand):
    """调用 DeepSeek API 生成规则"""
    try:
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        
        prompt = f"""
        你是一个文档处理专家。下面是一段 Word 模板的内容：
        ---
        {doc_content}
        ---
        用户的修改需求是："{user_demand}"
        
        请对比模板内容，找出需要被替换的精确原文字，并生成替换列表。
        要求：
        1. 严格遵守格式：旧内容 ==>> 新内容
        2. 每行一对，不要有任何多余的文字、序号或解释。
        3. 确保“旧内容”在模板文本中是完全一致的字符串。
        """
        
        data = {
            "model": "deepseek-chat",
            "messages": [
                {"role": "system", "content": "你是一个专业的文档分析助手。"},
                {"role": "user", "content": prompt}
            ],
            "stream": False
        }
        
        response = requests.post(f"{base_url}/chat/completions", headers=headers, json=data, timeout=30)
        response.raise_for_status()
        result = response.json()
        return result['choices'][0]['message']['content']
    except Exception as e:
        return f"Error: {str(e)}"

# --- UI Interface ---
def main():
    st.title("🐼 DeepSeek x Word 智能助手")
    
    # 初始化配置
    config = load_config()
    
    # 侧边栏设置
    with st.sidebar:
        st.header("⚙️ API 配置")
        input_key = st.text_input("DeepSeek API Key", value=config["api_key"], type="password")
        input_url = st.text_input("Base URL", value=config["base_url"])
        
        if st.button("💾 临时记住 Key"):
            save_config(input_key, input_url)
            st.success("已在此会话中记住设置")
            
        st.divider()
        st.caption("注：若在 Streamlit Cloud 无法找到 Secrets 选项，请尝试在 [share.streamlit.io](https://share.streamlit.io) 列表页点击 Settings。")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. 上传模板")
        uploaded_file = st.file_uploader("选择 Word 文件", type=["docx"])
        
        doc_sample = ""
        if uploaded_file:
            doc = Document(uploaded_file)
            doc_sample = "\n".join([p.text for p in doc.paragraphs])[:3000]
            st.success(f"✅ 已加载: {uploaded_file.name}")

        st.subheader("2. 描述需求")
        user_demand = st.text_area("告诉 AI 你想改什么？", placeholder="例如：把甲方换成华为，日期设为2026年4月")
        
        if "ai_rules" not in st.session_state:
            st.session_state.ai_rules = ""

        if st.button("✨ 生成替换规则", type="primary"):
            if not input_key:
                st.error("请先输入 API Key")
            elif not uploaded_file or not user_demand:
                st.warning("请上传文件并输入修改需求")
            else:
                with st.spinner("DeepSeek 正在分析并提取键值对..."):
                    result = get_deepseek_rules(input_key, input_url, doc_sample, user_demand)
                    st.session_state.ai_rules = result

    with col2:
        st.subheader("3. 确认与下载")
        rules_text = st.text_area(
            "最终替换规则 (旧内容 ==>> 新内容)",
            value=st.session_state.ai_rules,
            height=300
        )
        
        if st.button("🚀 执行替换并下载", use_container_width=True):
            if not uploaded_file or not rules_text:
                st.error("请补充完整信息")
            else:
                replacements = parse_replace_text(rules_text)
                with st.spinner("保留格式替换中..."):
                    doc = Document(uploaded_file)
                    processed_doc = smart_replace(doc, replacements)
                    
                    output = io.BytesIO()
                    processed_doc.save(output)
                    output.seek(0)
                    
                    st.download_button(
                        "📥 下载处理后的 Word",
                        data=output,
                        file_name=f"Smart_Fixed_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

if __name__ == "__main__":
    main()
