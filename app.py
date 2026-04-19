import os
import io
import streamlit as st
import requests
import json
import time
import copy

# 尝试导入必要库
try:
    from docx import Document
    from docx.oxml import parse_xml
except ImportError:
    st.error("❌ 缺少依赖库：python-docx。请运行 `pip install python-docx`。")
    st.stop()

# --- 页面配置 ---
st.set_page_config(page_title="DeepSeek x Word 智能助手 (增强保护版)", page_icon="🐼", layout="wide")

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

# --- 核心逻辑：记住、替换、恢复 ---

def backup_headers_footers(doc):
    """备份文档所有节的页眉页脚 XML 对象"""
    backups = []
    for i, section in enumerate(doc.sections):
        section_backup = {
            "index": i,
            "header_xml": copy.deepcopy(section.header._element),
            "footer_xml": copy.deepcopy(section.footer._element),
            "header_distance": section.header_distance,
            "footer_distance": section.footer_distance
        }
        backups.append(section_backup)
    return backups

def restore_headers_footers(doc, backups):
    """将备份的页眉页脚 XML 重新写回文档"""
    for backup in backups:
        idx = backup["index"]
        section = doc.sections[idx]
        
        # 移除旧的并同步备份的 XML 元素
        section.header._element.clear()
        for child in backup["header_xml"]:
            section.header._element.append(child)
            
        section.footer._element.clear()
        for child in backup["footer_xml"]:
            section.footer._element.append(child)
            
        # 恢复间距设置
        section.header_distance = backup["header_distance"]
        section.footer_distance = backup["footer_distance"]

def force_replace_body_xml(doc, replace_dict):
    """仅在主文档（Body）范围内执行强力替换"""
    for old_text, new_text in replace_dict.items():
        if not old_text: continue
        # 限制在 //w:body 路径下，避开页眉页脚定义的 XML 部分
        for element in doc.element.xpath('//w:body//w:t'):
            if element.text and old_text in element.text:
                element.text = element.text.replace(old_text, new_text)
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
    prompt = f"分析文档内容：{doc_content}\n需求：{user_demand}\n输出格式：旧内容 ==>> 新内容。禁止 Markdown。"
    data = {"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}]}
    try:
        response = requests.post(f"{base_url}/chat/completions", headers=headers, json=data, timeout=60)
        return response.json()['choices'][0]['message']['content']
    except:
        return "❌ API 连接超时。"

# --- UI 界面 ---
def main():
    st.title("🐼 Word 智能替换 (页眉页脚保护版)")
    config = load_config()
    
    with st.sidebar:
        st.header("⚙️ 设置")
        input_key = st.text_input("API Key", value=config["api_key"], type="password")
        input_url = st.text_input("Base URL", value=config["base_url"])
        if st.button("💾 记住 Key"):
            save_config(input_key, input_url)
            st.success("已记住")
        st.divider()
        st.info("🛡️ 运行逻辑：程序会先对页眉页脚进行【物理备份】，完成正文替换后再将其【原样还原】，确保页码、Logo等不受干扰。")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. 上传模板")
        uploaded_file = st.file_uploader("选择 Word 文件", type=["docx"])
        doc_sample = ""
        if uploaded_file:
            doc = Document(uploaded_file)
            doc_sample = "\n".join([p.text for p in doc.paragraphs])[:2000]
            st.success(f"✅ 文件已就绪")

        st.subheader("2. 修改需求")
        user_demand = st.text_area("你想改什么？")
        
        if st.button("✨ 生成替换规则", type="primary"):
            if not input_key or not uploaded_file:
                st.warning("请检查 Key 和文件")
            else:
                with st.spinner("AI 分析中..."):
                    result = get_deepseek_rules(input_key, input_url, doc_sample, user_demand)
                    st.session_state.ai_rules = result

    with col2:
        st.subheader("3. 执行并下载")
        rules_text = st.text_area("确认规则", value=st.session_state.get("ai_rules", ""), height=300)
        
        if st.button("🚀 执行安全替换", use_container_width=True):
            if not uploaded_file:
                st.error("请先上传文件")
            else:
                replacements = parse_replace_text(rules_text)
                with st.spinner("执行快照保护与替换..."):
                    doc = Document(uploaded_file)
                    
                    # 第一步：记住（备份）页眉页脚
                    backups = backup_headers_footers(doc)
                    
                    # 第二步：替换（仅主文档）
                    processed_doc = force_replace_body_xml(doc, replacements)
                    
                    # 第三步：恢复页眉页脚
                    restore_headers_footers(processed_doc, backups)
                    
                    output = io.BytesIO()
                    processed_doc.save(output)
                    output.seek(0)
                    st.download_button("📥 下载 (已还原页眉页脚)", data=output, file_name=f"Safe_Fixed_{uploaded_file.name}", use_container_width=True)

if __name__ == "__main__":
    main()
