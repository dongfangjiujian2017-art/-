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
st.set_page_config(page_title="DeepSeek x Word 智能助手 (终极修复版)", page_icon="🐼", layout="wide")

# --- 持久化配置管理 ---
CONFIG_FILE = "config.json"

def load_config():
    if "saved_api_key" in st.session_state:
        return {"api_key": st.session_state.saved_api_key, "base_url": st.session_state.get("saved_base_url", "https://api.deepseek.com")}
    env_key = os.environ.get("DEEPSEEK_API_KEY", "")
    if env_key:
        return {"api_key": env_key, "base_url": "https://api.deepseek.com"}
    return {"api_key": "", "base_url": "https://api.deepseek.com"}

def save_config(api_key, base_url):
    st.session_state.saved_api_key = api_key
    st.session_state.saved_base_url = base_url

# --- 核心逻辑：安全替换（不碰触域代码） ---

def safe_replace_in_paragraph(paragraph, replace_dict):
    """
    安全地在段落中替换文本。
    采用“段落级物理隔离”：只要段落包含域代码，整段放弃处理，防止 WPS 解析崩溃。
    """
    # --- 终极防御：段落级隔离 ---
    xml_str = paragraph._element.xml
    # 'w:fldChar' 和 'w:instrText' 是动态域（如页码）的独有标记
    if 'w:fldChar' in xml_str or 'w:instrText' in xml_str:
        return # 立即退出，该段落原封不动保留
        
    for old_text, new_text in replace_dict.items():
        if not old_text or old_text not in paragraph.text:
            continue
            
        for run in paragraph.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)

def process_all_parts(doc, replace_dict):
    """
    遍历文档的所有部分：正文、表格、页眉、页脚、文本框。
    只修改文本内容，不执行 XML 重写，从而保护页码。
    """
    # 1. 正文段落
    for p in doc.paragraphs:
        safe_replace_in_paragraph(p, replace_dict)
        
    # 2. 表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    safe_replace_in_paragraph(p, replace_dict)
                    
    # 3. 页眉页脚（仅当它们存在时）
    for section in doc.sections:
        # 处理页眉
        for p in section.header.paragraphs:
            safe_replace_in_paragraph(p, replace_dict)
        for table in section.header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        safe_replace_in_paragraph(p, replace_dict)
        
        # 处理页脚
        for p in section.footer.paragraphs:
            safe_replace_in_paragraph(p, replace_dict)
        for table in section.footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        safe_replace_in_paragraph(p, replace_dict)

    # 4. 文本框（通过底层 XML 遍历，但仅修改文本节点）
    for txbx in doc.element.xpath('//w:txbxContent//w:p'):
        from docx.text.paragraph import Paragraph
        p = Paragraph(txbx, doc)
        safe_replace_in_paragraph(p, replace_dict)

    return doc

def parse_replace_text(text):
    replace_dict = {}
    for line in text.split('\n'):
        clean_line = line.replace('**', '').replace('`', '').strip()
        if "==>>" in clean_line:
            parts = clean_line.split("==>>")
            if len(parts) == 2:
                replace_dict[parts[0].strip()] = parts[1].strip()
    return replace_dict

# --- AI 逻辑 ---
def get_deepseek_rules(api_key, base_url, doc_content, user_demand):
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    prompt = f"分析文档内容：{doc_content}\n需求：{user_demand}\n输出格式：旧内容 ==>> 新内容。禁止包含 PAGE 或 NUMPAGES 等页码字符。"
    data = {"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}], "stream": False}
    try:
        response = requests.post(f"{base_url}/chat/completions", headers=headers, json=data, timeout=60)
        return response.json()['choices'][0]['message']['content']
    except:
        return "❌ API 连接超时。"

# --- UI 界面 ---
def main():
    st.title("🐼 Word 智能助手 (终极安全版)")
    config = load_config()
    
    with st.sidebar:
        st.header("⚙️ 设置")
        input_key = st.text_input("DeepSeek API Key", value=config["api_key"], type="password")
        input_url = st.text_input("Base URL", value=config["base_url"])
        if st.button("💾 记住 Key"):
            save_config(input_key, input_url)
            st.success("已记住")
        st.divider()
        st.info("🛡️ 终极方案：此版本采用 Run 级别检测，会自动避开包含域代码（PAGE/NUMPAGES）的区域，只修改纯文本内容。")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. 上传模板")
        uploaded_file = st.file_uploader("选择 Word 模板", type=["docx"])
        doc_sample = ""
        if uploaded_file:
            doc = Document(uploaded_file)
            doc_sample = "\n".join([p.text for p in doc.paragraphs])[:2000]
            st.success(f"✅ 文件已就绪")

        st.subheader("2. 修改需求")
        user_demand = st.text_area("你想改什么？")
        
        if st.button("✨ 生成替换规则", type="primary"):
            if not input_key or not uploaded_file:
                st.warning("请检查配置")
            else:
                with st.spinner("AI 分析中..."):
                    result = get_deepseek_rules(input_key, input_url, doc_sample, user_demand)
                    st.session_state.ai_rules = result

    with col2:
        st.subheader("3. 执行并下载")
        rules_text = st.text_area("最终规则预览", value=st.session_state.get("ai_rules", ""), height=300)
        
        if st.button("🚀 执行精准替换", use_container_width=True):
            if not uploaded_file:
                st.error("请先上传文件")
            else:
                replacements = parse_replace_text(rules_text)
                with st.spinner("正在安全替换文本..."):
                    doc = Document(uploaded_file)
                    
                    # 执行全文档精准扫描
                    processed_doc = process_all_parts(doc, replacements)
                    
                    output = io.BytesIO()
                    processed_doc.save(output)
                    output.seek(0)
                    st.download_button("📥 下载修复后的文档", data=output, file_name=f"Fixed_{uploaded_file.name}", use_container_width=True)

if __name__ == "__main__":
    main()
