import os
import io
import streamlit as st
import requests
import json
import time

# 尝试导入必要库
try:
    from docx import Document
except ImportError:
    st.error("❌ 缺少依赖库：python-docx。请运行 `pip install python-docx`。")
    st.stop()

# --- 页面配置 ---
st.set_page_config(page_title="DeepSeek x Word 智能助手", page_icon="🐼", layout="wide")

# --- 持久化配置管理 ---
CONFIG_FILE = "config.json"

def load_config():
    """优先级读取 API Key：1. Session 2. 环境变量 3. 本地文件"""
    if "saved_api_key" in st.session_state:
        return {"api_key": st.session_state.saved_api_key, "base_url": st.session_state.get("saved_base_url", "https://api.deepseek.com")}

    env_key = os.environ.get("DEEPSEEK_API_KEY", "")
    if env_key:
        return {"api_key": env_key, "base_url": "https://api.deepseek.com"}
    
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
        pass 

# --- 核心逻辑：Word 替换 ---
def _apply_replace(paragraph, replace_dict):
    """替换段落文本并保留格式"""
    for old_text, new_text in replace_dict.items():
        if old_text in paragraph.text:
            for run in paragraph.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)

def smart_replace(doc, replace_dict):
    """
    全面扫描替换。
    说明：此函数仅处理正文段落和表格中的内容。
    由于不涉及 doc.sections[0].header/footer 的操作，原文档自带的页眉页脚将原封不动地保留。
    """
    # 1. 替换正文
    for p in doc.paragraphs:
        _apply_replace(p, replace_dict)
    
    # 2. 替换表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _apply_replace(p, replace_dict)
    return doc

def parse_replace_text(text):
    """解析 '旧内容 ==>> 新内容' 格式，清理 Markdown"""
    replace_dict = {}
    for line in text.split('\n'):
        # 移除可能干扰匹配的 Markdown 符号
        clean_line = line.replace('**', '').replace('`', '').strip()
        if "==>>" in clean_line:
            parts = clean_line.split("==>>")
            if len(parts) == 2:
                replace_dict[parts[0].strip()] = parts[1].strip()
    return replace_dict

# --- AI 逻辑：调用 DeepSeek API ---
def get_deepseek_rules(api_key, base_url, doc_content, user_demand):
    """调用 DeepSeek API 生成规则"""
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
    格式要求：
    1. 严格使用：旧内容 ==>> 新内容
    2. 禁止 Markdown 格式（不要加粗 **，不要代码块）。
    3. 确保“旧内容”在模板中真实存在。
    """
    
    data = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "system", "content": "你是一个专业的文档分析助手，只输出纯文本格式的替换规则。"},
            {"role": "user", "content": prompt}
        ],
        "stream": False
    }

    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = requests.post(f"{base_url}/chat/completions", headers=headers, json=data, timeout=60)
            response.raise_for_status()
            return response.json()['choices'][0]['message']['content']
        except:
            if attempt < max_retries - 1:
                time.sleep(2)
                continue
            return "❌ API 连接超时，请稍后再试。"

# --- UI 界面 ---
def main():
    st.title("🐼 DeepSeek x Word 智能助手")
    
    config = load_config()
    
    with st.sidebar:
        st.header("⚙️ 设置")
        input_key = st.text_input("DeepSeek API Key", value=config["api_key"], type="password")
        input_url = st.text_input("Base URL", value=config["base_url"])
        if st.button("💾 临时记住 Key"):
            save_config(input_key, input_url)
            st.success("已记住设置")
        
        st.divider()
        st.info("提示：此版本会原样保留您 Word 文档中原有的页眉和页脚。")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. 上传模板")
        uploaded_file = st.file_uploader("选择 Word 文件", type=["docx"])
        
        doc_sample = ""
        if uploaded_file:
            doc = Document(uploaded_file)
            # 获取前 3000 字作为 AI 参考
            doc_sample = "\n".join([p.text for p in doc.paragraphs])[:3000]
            st.success(f"✅ 已加载: {uploaded_file.name}")

        st.subheader("2. 描述需求")
        user_demand = st.text_area("告诉 AI 你想改什么？", placeholder="例如：把甲方换成华为...")
        
        if "ai_rules" not in st.session_state:
            st.session_state.ai_rules = ""

        if st.button("✨ 生成替换规则", type="primary"):
            if not input_key:
                st.error("请先输入 API Key")
            elif not uploaded_file:
                st.warning("请上传文件")
            else:
                with st.spinner("DeepSeek 分析中..."):
                    result = get_deepseek_rules(input_key, input_url, doc_sample, user_demand)
                    st.session_state.ai_rules = result

    with col2:
        st.subheader("3. 执行替换")
        rules_text = st.text_area("最终替换规则", value=st.session_state.ai_rules, height=300)
        
        if st.button("🚀 执行替换并下载", use_container_width=True):
            if not uploaded_file:
                st.error("请上传模板")
            else:
                replacements = parse_replace_text(rules_text)
                with st.spinner("正在处理正文和表格..."):
                    # 重新打开文档进行处理
                    doc = Document(uploaded_file)
                    # 执行替换逻辑（不触碰页眉页脚）
                    processed_doc = smart_replace(doc, replacements)
                    
                    output = io.BytesIO()
                    processed_doc.save(output)
                    output.seek(0)
                    
                    st.download_button(
                        "📥 下载处理后的 Word",
                        data=output,
                        file_name=f"Fixed_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

if __name__ == "__main__":
    main()
