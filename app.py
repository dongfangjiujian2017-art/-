import os
import io
import streamlit as st
import requests
import copy
import time
from docx import Document
from docx.oxml.ns import qn

# --- 页面配置 ---
st.set_page_config(page_title="Word 无损替换助手 (正文提取版)", page_icon="🛡️", layout="wide")

# --- 配置读取逻辑 ---
def get_config():
    """
    自动从 Streamlit Secrets 读取配置
    """
    default_key = ""
    default_url = "https://api.deepseek.com"
    
    if "DEEPSEEK_API_KEY" in st.secrets:
        default_key = st.secrets["DEEPSEEK_API_KEY"]
    if "DEEPSEEK_BASE_URL" in st.secrets:
        default_url = st.secrets["DEEPSEEK_BASE_URL"]
        
    return default_key, default_url

def clone_only_body_content(source_doc, target_doc):
    """
    【绝对隔离方案】
    仅搬运正文中的段落和表格，显式忽略所有节属性(sectPr)。
    这样可以确保 target_doc 原始的页眉页脚定义不被污染。
    """
    # 1. 彻底清空目标文档正文（保留空壳）
    target_body = target_doc._element.body
    for child in list(target_body):
        # 保留最后的节属性(sectPr)，这是维持页脚正常的关键，除此之外全部删除
        if child.tag != qn('w:sectPr'):
            target_body.remove(child)

    # 2. 从源文档中仅提取 p (段落) 和 tbl (表格)
    source_body = source_doc._element.body
    for child in source_body:
        # 排除 sectPr，确保不把源文档的页脚逻辑带过来
        if child.tag in (qn('w:p'), qn('w:tbl')):
            new_element = copy.deepcopy(child)
            # 将新元素插入到目标文档节属性之前
            target_body.insert(-1, new_element)

def safe_replace_text(doc, replace_dict):
    """在文档对象中执行文本替换"""
    if not replace_dict:
        return
        
    for p in doc.paragraphs:
        for old_txt, new_txt in replace_dict.items():
            if old_txt in p.text:
                for run in p.runs:
                    if old_txt in run.text:
                        run.text = run.text.replace(old_txt, new_txt)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for old_txt, new_txt in replace_dict.items():
                        if old_txt in p.text:
                            for run in p.runs:
                                if old_txt in run.text:
                                    run.text = run.text.replace(old_txt, new_txt)

def get_deepseek_rules(api_key, base_url, doc_sample, user_demand):
    """请求 DeepSeek API 自动生成替换规则，带重试机制"""
    if not api_key:
        return "❌ 错误：未检测到 API Key。"

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    prompt = f"正文样本：{doc_sample}\n修改需求：{user_demand}\n请输出替换对。格式：旧内容 ==>> 新内容。每行一对。严禁修改页码相关的代码。"
    
    data = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3
    }
    
    max_retries = 5
    for attempt in range(max_retries):
        try:
            response = requests.post(f"{base_url}/chat/completions", headers=headers, json=data, timeout=60)
            response.raise_for_status()
            return response.json()['choices'][0]['message']['content']
        except Exception:
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)
                continue
            return "❌ API 连接超时，请稍后再试。"

def parse_rules(rules_text):
    """解析 AI 文本为字典"""
    replace_dict = {}
    if not rules_text: return replace_dict
    for line in rules_text.split('\n'):
        if "==>>" in line:
            parts = line.split("==>>")
            if len(parts) == 2:
                replace_dict[parts[0].strip()] = parts[1].strip()
    return replace_dict

def main():
    st.title("🛡️ Word 智能替换助手 (正文精准提取版)")
    
    def_key, def_url = get_config()
    
    with st.sidebar:
        st.header("⚙️ 账号配置")
        api_key = st.text_input("DeepSeek API Key", value=def_key, type="password")
        base_url = st.text_input("Base URL", value=def_url)
        
        st.divider()
        st.warning("🚀 核心技术：本版本通过【sectPr 强制隔离】技术，仅提取文档正文节点，绝不触碰页眉页脚底层定义。")

    uploaded_file = st.file_uploader("上传 Word 文档", type=["docx"])
    user_demand = st.text_area("修改需求", placeholder="例如：把甲方改为『某某公司』")
    
    col1, col2 = st.columns(2)

    with col1:
        if st.button("✨ 智能分析规则", type="primary", use_container_width=True):
            if not uploaded_file:
                st.warning("请先上传文件")
            else:
                with st.spinner("AI 正在分析正文..."):
                    content_bytes = uploaded_file.getvalue()
                    sample_doc = Document(io.BytesIO(content_bytes))
                    sample_text = "\n".join([p.text for p in sample_doc.paragraphs[:15]])
                    rules = get_deepseek_rules(api_key, base_url, sample_text, user_demand)
                    st.session_state.ai_rules = rules

    with col2:
        rules_text = st.text_area("规则预览", value=st.session_state.get("ai_rules", ""), height=150)
        
        if st.button("🚀 执行无损处理", use_container_width=True):
            if not uploaded_file or not rules_text:
                st.error("请确保已上传文件并生成规则")
            else:
                replacements = parse_rules(rules_text)
                with st.spinner("正在提取正文并注入干净容器..."):
                    content_bytes = uploaded_file.getvalue()
                    
                    # 1. 创建干净的容器（保留页脚）
                    container_doc = Document(io.BytesIO(content_bytes))
                    # 2. 创建用于修改的正文源
                    source_doc = Document(io.BytesIO(content_bytes))

                    # 3. 修改正文源的文字
                    safe_replace_text(source_doc, replacements)

                    # 4. 【关键】仅将改好的段落/表格搬运到容器，避开节属性
                    clone_only_body_content(source_doc, container_doc)

                    output = io.BytesIO()
                    container_doc.save(output)
                    output.seek(0)
                    
                    st.success("✅ 处理完成！正文已替换，页脚已强制隔离保护。")
                    st.download_button(
                        "📥 下载结果文件", 
                        data=output, 
                        file_name=f"Clean_{uploaded_file.name}",
                        use_container_width=True
                    )

if __name__ == "__main__":
    main()
