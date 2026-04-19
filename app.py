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
    """
    if element is None:
        return
    for old_text, new_text in replace_dict.items():
        if not old_text: continue
        for t_node in element.xpath('.//w:t'):
            if t_node.text and old_text in t_node.text:
                t_node.text = t_node.text.replace(old_text, new_text)

def backup_and_process_parts(doc, replace_dict):
    """
    备份页眉页脚。如果原本就没有内容（空段落且无文本框），则标记为无效，不强制恢复。
    """
    backups = []
    for i, section in enumerate(doc.sections):
        # 检查是否有内容存在（避免强制生成空白页码）
        has_header = any(p.text.strip() for p in section.header.paragraphs) or len(section.header.tables) > 0 or section.header._element.xpath('.//w:txbxContent')
        has_footer = any(p.text.strip() for p in section.footer.paragraphs) or len(section.footer.tables) > 0 or section.footer._element.xpath('.//w:txbxContent')

        h_xml = copy.deepcopy(section.header._element) if has_header else None
        f_xml = copy.deepcopy(section.footer._element) if has_footer else None
        
        # 在有效的副本中执行替换
        if h_xml is not None:
            _xml_text_replace(h_xml, replace_dict)
        if f_xml is not None:
            _xml_text_replace(f_xml, replace_dict)
        
        backups.append({
            "index": i,
            "header_xml": h_xml,
            "footer_xml": f_xml,
            "header_distance": section.header_distance,
            "footer_distance": section.footer_distance,
            "has_h": has_header,
            "has_f": has_footer
        })
    return backups

def restore_parts(doc, backups):
    """仅还原原本就有内容的页眉页脚"""
    for backup in backups:
        section = doc.sections[backup["index"]]
        
        # 只有原本有内容的才去 clear() 和 append()
        if backup["has_h"] and backup["header_xml"] is not None:
            section.header._element.clear()
            for child in backup["header_xml"]:
                section.header._element.append(child)
                
        if backup["has_f"] and backup["footer_xml"] is not None:
            section.footer._element.clear()
            for child in backup["footer_xml"]:
                section.footer._element.append(child)
            
        section.header_distance = backup["header_distance"]
        section.footer_distance = backup["footer_distance"]

def force_replace_body_xml(doc, replace_dict):
    """仅在主文档范围内执行替换"""
    for old_text, new_text in replace_dict.items():
        if not old_text: continue
        # 严格限制在 w:body 下
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
    prompt = f"分析文档：{doc_content}\n需求：{user_demand}\n规则：旧内容 ==>> 新内容。禁止 Markdown。"
    data = {"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}], "stream": False}
    try:
        response = requests.post(f"{base_url}/chat/completions", headers=headers, json=data, timeout=60)
        return response.json()['choices'][0]['message']['content']
    except:
        return "❌ API 连接失败"

# --- UI 界面 ---
def main():
    st.title("🐼 Word 智能助手 (无损保护版)")
    config = load_config()
    
    with st.sidebar:
        st.header("⚙️ 设置")
        input_key = st.text_input("DeepSeek API Key", value=config["api_key"], type="password")
        input_url = st.text_input("Base URL", value=config["base_url"])
        if st.button("💾 记住 Key"):
            save_config(input_key, input_url)
            st.success("已记住")
        st.divider()
        st.info("🛠️ 更新说明：现在程序会自动检测页眉页脚是否有内容。如果原本是空的，程序绝不会强行生成任何页码或文字。")

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
        
        if st.button("🚀 开始处理", use_container_width=True):
            if not uploaded_file:
                st.error("请先上传文件")
            else:
                replacements = parse_replace_text(rules_text)
                with st.spinner("正在处理..."):
                    doc = Document(uploaded_file)
                    # 1. 检测并有条件地备份
                    backups = backup_and_process_parts(doc, replacements)
                    # 2. 替换正文
                    processed_doc = force_replace_body_xml(doc, replacements)
                    # 3. 有条件地还原
                    restore_parts(processed_doc, backups)
                    
                    output = io.BytesIO()
                    processed_doc.save(output)
                    output.seek(0)
                    st.download_button("📥 下载文件", data=output, file_name=f"Fixed_{uploaded_file.name}", use_container_width=True)

if __name__ == "__main__":
    main()
