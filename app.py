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
st.set_page_config(page_title="DeepSeek x Word 智能助手 (域代码保护版)", page_icon="🐼", layout="wide")

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
    """
    备份文档所有节的页眉页脚 XML 对象。
    使用 ._element 访问底层 XML，确保域代码（页码）被完整记录。
    """
    backups = []
    for i, section in enumerate(doc.sections):
        # 深度复制 XML 元素
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
    """
    将备份的页眉页脚 XML 重新写回文档，修复域代码显示问题。
    """
    for backup in backups:
        idx = backup["index"]
        section = doc.sections[idx]
        
        # 1. 恢复页眉
        header_el = section.header._element
        header_el.clear()
        for child in backup["header_xml"]:
            header_el.append(child)
            
        # 2. 恢复页脚
        footer_el = section.footer._element
        footer_el.clear()
        for child in backup["footer_xml"]:
            footer_el.append(child)
            
        # 3. 恢复间距设置
        section.header_distance = backup["header_distance"]
        section.footer_distance = backup["footer_distance"]

def force_replace_body_xml(doc, replace_dict):
    """
    仅在主文档（Body）范围内执行强力替换。
    使用 XPath 限制范围，绝对不触碰页眉页脚部分的 XML。
    """
    for old_text, new_text in replace_dict.items():
        if not old_text: continue
        # 仅针对 w:body 下的文本节点进行替换
        for element in doc.element.xpath('//w:body//w:t'):
            if element.text and old_text in element.text:
                element.text = element.text.replace(old_text, new_text)
    return doc

def parse_replace_text(text):
    """解析 AI 生成的规则"""
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
    prompt = f"""
    分析文档内容：{doc_content}
    用户需求：{user_demand}
    
    指令：
    1. 找出需要替换的文字对。
    2. 输出格式：旧内容 ==>> 新内容。
    3. 严禁 Markdown 格式。
    4. 不要尝试修改页码（如 PAGE, NUMPAGES），保持它们在正文之外。
    """
    data = {
        "model": "deepseek-chat", 
        "messages": [{"role": "system", "content": "你是一个专业的文档规则生成器。"}, {"role": "user", "content": prompt}],
        "stream": False
    }
    try:
        response = requests.post(f"{base_url}/chat/completions", headers=headers, json=data, timeout=60)
        return response.json()['choices'][0]['message']['content']
    except Exception as e:
        return f"❌ API 连接失败: {str(e)}"

# --- UI 界面 ---
def main():
    st.title("🐼 Word 智能替换 (页码域代码保护版)")
    config = load_config()
    
    with st.sidebar:
        st.header("⚙️ 配置中心")
        input_key = st.text_input("DeepSeek API Key", value=config["api_key"], type="password")
        input_url = st.text_input("Base URL", value=config["base_url"])
        if st.button("💾 记住设置"):
            save_config(input_key, input_url)
            st.success("配置已保存")
        st.divider()
        st.info("💡 修复说明：本版本会在处理前对页眉页脚执行『全量 XML 快照』，在完成正文替换后原样写回，从而保护 Word 页码域代码不被破坏。")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. 导入文档")
        uploaded_file = st.file_uploader("选择 Word 模板 (.docx)", type=["docx"])
        doc_sample = ""
        if uploaded_file:
            doc = Document(uploaded_file)
            # 仅提取正文前 2000 字给 AI 辅助分析
            doc_sample = "\n".join([p.text for p in doc.paragraphs])[:2000]
            st.success(f"已加载：{uploaded_file.name}")

        st.subheader("2. 交互需求")
        user_demand = st.text_area("输入修改指示", placeholder="例如：将甲方公司名称改为‘字节跳动’")
        
        if st.button("✨ 生成规则", type="primary"):
            if not input_key or not uploaded_file:
                st.warning("请检查 API Key 和文件上传状态")
            else:
                with st.spinner("AI 正在解析正文内容..."):
                    result = get_deepseek_rules(input_key, input_url, doc_sample, user_demand)
                    st.session_state.ai_rules = result

    with col2:
        st.subheader("3. 确认并执行")
        rules_text = st.text_area("替换预览 (旧内容 ==>> 新内容)", value=st.session_state.get("ai_rules", ""), height=300)
        
        if st.button("🚀 安全替换并导出", use_container_width=True):
            if not uploaded_file:
                st.error("请先上传文件")
            else:
                replacements = parse_replace_text(rules_text)
                with st.spinner("执行快照保护与正文替换..."):
                    # 重新读取文档以确保干净
                    doc = Document(uploaded_file)
                    
                    # [步骤1] 备份原始页眉页脚的 XML 结构（含域代码）
                    backups = backup_headers_footers(doc)
                    
                    # [步骤2] 仅针对 Body 区域执行全量文字替换
                    processed_doc = force_replace_body_xml(doc, replacements)
                    
                    # [步骤3] 还原备份的页眉页脚，彻底解决 {PAGE} 显示问题
                    restore_headers_footers(processed_doc, backups)
                    
                    # 导出
                    output = io.BytesIO()
                    processed_doc.save(output)
                    output.seek(0)
                    st.download_button(
                        "📥 点击下载处理后的文档", 
                        data=output, 
                        file_name=f"Safe_{uploaded_file.name}", 
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

if __name__ == "__main__":
    main()
