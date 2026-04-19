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
st.set_page_config(page_title="DeepSeek x Word 智能助手 (全能保护版)", page_icon="🐼", layout="wide")

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

# --- 核心逻辑：精准 XML 替换 ---

def _xml_text_replace(element, replace_dict):
    """
    在 XML 元素中深度查找所有 w:t 节点并执行替换。
    这种方法可以安全地处理文本框内容，同时不会破坏域代码结构。
    """
    for old_text, new_text in replace_dict.items():
        if not old_text: continue
        # 查找当前元素下所有的文本节点
        for t_node in element.xpath('.//w:t'):
            if t_node.text and old_text in t_node.text:
                t_node.text = t_node.text.replace(old_text, new_text)

def backup_and_process_parts(doc, replace_dict):
    """
    备份页眉页脚，并在备份的 XML 副本中执行文字替换。
    """
    backups = []
    for i, section in enumerate(doc.sections):
        # 复制原始 XML
        h_xml = copy.deepcopy(section.header._element)
        f_xml = copy.deepcopy(section.footer._element)
        
        # 在副本中执行“手术级”替换（针对文本框等文字）
        _xml_text_replace(h_xml, replace_dict)
        _xml_text_replace(f_xml, replace_dict)
        
        backups.append({
            "index": i,
            "header_xml": h_xml,
            "footer_xml": f_xml,
            "header_distance": section.header_distance,
            "footer_distance": section.footer_distance
        })
    return backups

def restore_parts(doc, backups):
    """将处理后的 XML 重新写回"""
    for backup in backups:
        section = doc.sections[backup["index"]]
        
        # 恢复页眉
        section.header._element.clear()
        for child in backup["header_xml"]:
            section.header._element.append(child)
            
        # 恢复页脚
        section.footer._element.clear()
        for child in backup["footer_xml"]:
            section.footer._element.append(child)
            
        section.header_distance = backup["header_distance"]
        section.footer_distance = backup["footer_distance"]

def force_replace_body_xml(doc, replace_dict):
    """仅在主文档范围内执行替换"""
    for old_text, new_text in replace_dict.items():
        if not old_text: continue
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
    prompt = f"分析文档：{doc_content}\n需求：{user_demand}\n格式：旧内容 ==>> 新内容。禁止 Markdown。"
    data = {"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}], "stream": False}
    try:
        response = requests.post(f"{base_url}/chat/completions", headers=headers, json=data, timeout=60)
        return response.json()['choices'][0]['message']['content']
    except:
        return "❌ API 连接失败"

# --- UI 界面 ---
def main():
    st.title("🐼 Word 智能助手 (文本框+页码完美版)")
    config = load_config()
    
    with st.sidebar:
        st.header("⚙️ 配置")
        input_key = st.text_input("DeepSeek API Key", value=config["api_key"], type="password")
        input_url = st.text_input("Base URL", value=config["base_url"])
        if st.button("💾 记住 Key"):
            save_config(input_key, input_url)
            st.success("已记住")
        st.divider()
        st.info("🚀 修复逻辑：先备份页眉页脚 XML，在内存中精准替换文本框文字，最后整体还原。既保护了动态页码，又改掉了文本框里的文字。")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. 导入文档")
        uploaded_file = st.file_uploader("选择 Word 模板", type=["docx"])
        doc_sample = ""
        if uploaded_file:
            doc = Document(uploaded_file)
            doc_sample = "\n".join([p.text for p in doc.paragraphs])[:2000]
            st.success(f"已加载：{uploaded_file.name}")

        st.subheader("2. 修改需求")
        user_demand = st.text_area("你想改什么？")
        
        if st.button("✨ 生成规则", type="primary"):
            if not input_key or not uploaded_file:
                st.warning("请检查配置")
            else:
                with st.spinner("AI 分析中..."):
                    result = get_deepseek_rules(input_key, input_url, doc_sample, user_demand)
                    st.session_state.ai_rules = result

    with col2:
        st.subheader("3. 执行并下载")
        rules_text = st.text_area("确认规则", value=st.session_state.get("ai_rules", ""), height=300)
        
        if st.button("🚀 开始完美替换", use_container_width=True):
            if not uploaded_file:
                st.error("请先上传文件")
            else:
                replacements = parse_replace_text(rules_text)
                with st.spinner("正在进行多层 XML 替换..."):
                    doc = Document(uploaded_file)
                    
                    # 1. 备份页眉页脚并在 XML 副本中执行文本框内容的替换
                    backups = backup_and_process_parts(doc, replacements)
                    
                    # 2. 替换正文
                    processed_doc = force_replace_body_xml(doc, replacements)
                    
                    # 3. 恢复处理后的页眉页脚
                    restore_parts(processed_doc, backups)
                    
                    output = io.BytesIO()
                    processed_doc.save(output)
                    output.seek(0)
                    st.download_button("📥 下载完美修复版", data=output, file_name=f"Perfect_{uploaded_file.name}", use_container_width=True)

if __name__ == "__main__":
    main()
