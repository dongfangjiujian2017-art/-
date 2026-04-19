import os
import io
import streamlit as st
import requests
import copy
import time
from docx import Document
from docx.oxml.ns import qn

# --- 页面配置 ---
st.set_page_config(page_title="Word 智能替换助手", page_icon="📝", layout="centered")

# --- 配置读取逻辑 ---
def get_config():
    """从 Secrets 或环境变量读取配置"""
    default_key = st.secrets.get("DEEPSEEK_API_KEY", "")
    default_url = st.secrets.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
    return default_key, default_url

def clone_only_body_content(source_doc, target_doc):
    """提取正文并保护页脚"""
    target_body = target_doc._element.body
    # 清空目标正文，保留 sectPr (页眉页脚定义)
    for child in list(target_body):
        if child.tag != qn('w:sectPr'):
            target_body.remove(child)

    source_body = source_doc._element.body
    for child in source_body:
        if child.tag in (qn('w:p'), qn('w:tbl')):
            new_element = copy.deepcopy(child)
            target_body.insert(-1, new_element)

def safe_replace_text(doc, replace_dict):
    """
    全量文本替换：涵盖正文、表格以及页眉(通常标题所在处)
    """
    if not replace_dict: return

    # 1. 替换所有页眉内容 (很多文档的标题在页眉)
    for section in doc.sections:
        header = section.header
        for p in header.paragraphs:
            _replace_in_paragraph(p, replace_dict)
        for table in header.tables:
            _replace_in_table(table, replace_dict)

    # 2. 替换正文段落
    for p in doc.paragraphs:
        _replace_in_paragraph(p, replace_dict)

    # 3. 替换表格内容
    for table in doc.tables:
        _replace_in_table(table, replace_dict)

def _replace_in_paragraph(p, replace_dict):
    """底层段落替换逻辑"""
    for old, new in replace_dict.items():
        if old in p.text:
            for run in p.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)

def _replace_in_table(table, replace_dict):
    """底层表格替换逻辑"""
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                _replace_in_paragraph(p, replace_dict)

def get_deepseek_rules(api_key, base_url, doc_sample, user_demand):
    """请求 AI 生成规则"""
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    # 强化 Prompt，要求 AI 必须覆盖标题
    prompt = f"正文样本：{doc_sample}\n修改需求：{user_demand}\n请输出替换对。格式：旧内容 ==>> 新内容。注意：必须包含标题中的关键信息替换。只输出对子，严禁其他文字。"
    data = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1
    }
    for _ in range(3):
        try:
            response = requests.post(f"{base_url}/chat/completions", headers=headers, json=data, timeout=40)
            response.raise_for_status()
            return response.json()['choices'][0]['message']['content']
        except:
            time.sleep(1)
    return ""

def parse_rules(rules_text):
    """解析键值对"""
    replace_dict = {}
    for line in rules_text.split('\n'):
        if "==>>" in line:
            parts = line.split("==>>")
            if len(parts) == 2:
                replace_dict[parts[0].strip()] = parts[1].strip()
    return replace_dict

def main():
    st.title("📝 Word 智能替换助手")
    
    def_key, def_url = get_config()
    with st.expander("🔑 API 设置"):
        api_key = st.text_input("DeepSeek API Key", value=def_key, type="password")
        base_url = st.text_input("Base URL", value=def_url)

    # 简化的使用说明
    st.markdown("""
    ### 📖 使用说明
    1. **上传模板**：上传您的 Word 文档。
    2. **输入需求**：描述您想修改的内容（如：把甲方改为华为）。
    3. **自动生成**：系统将自动替换正文及标题内容，并保留原始页脚。
    """)

    uploaded_file = st.file_uploader("1. 上传模板文件", type=["docx"])
    user_input = st.text_area("2. 描述修改需求", placeholder="例如：把文档中的甲方公司名改为『字节跳动』，合同金额改为『壹佰万元』...", height=120)
    
    if st.button("🚀 开始自动替换并生成", type="primary", use_container_width=True):
        if not uploaded_file or not user_input:
            st.error("请确保已上传文件并输入修改需求")
            return

        with st.spinner("正在处理文档..."):
            content_bytes = uploaded_file.getvalue()
            source_doc = Document(io.BytesIO(content_bytes))
            
            # 自动处理：如果用户粘贴了规则则直接解析，否则请求AI
            if "==>>" in user_input:
                replacements = parse_rules(user_input)
            else:
                if not api_key:
                    st.error("请在上方设置中输入 API Key 以启用智能分析")
                    return
                # 样本提取（包含前20段以确保抓到标题）
                sample_text = "\n".join([p.text for p in source_doc.paragraphs[:20]])
                ai_raw = get_deepseek_rules(api_key, base_url, sample_text, user_input)
                replacements = parse_rules(ai_raw)
            
            if not replacements:
                st.error("未能识别修改指令，请尝试更明确地描述。")
                return

            # 创建下载容器并执行替换逻辑
            container_doc = Document(io.BytesIO(content_bytes))
            
            # 在源文档上执行替换（包含页眉标题）
            safe_replace_text(source_doc, replacements)
            
            # 将改好的内容克隆进保留了页脚定义的容器
            clone_only_body_content(source_doc, container_doc)

            # 针对容器文档也要单独跑一次页眉替换，因为页眉定义在容器的 sectPr 里
            safe_replace_text(container_doc, replacements)

            output = io.BytesIO()
            container_doc.save(output)
            output.seek(0)
            
            st.success(f"✅ 处理成功！已识别并替换 {len(replacements)} 处关键信息。")
            st.download_button(
                "📥 点击下载新文档", 
                data=output, 
                file_name=f"已修改_{uploaded_file.name}",
                use_container_width=True
            )

if __name__ == "__main__":
    main()
