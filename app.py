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
    for child in list(target_body):
        if child.tag != qn('w:sectPr'):
            target_body.remove(child)

    source_body = source_doc._element.body
    for child in source_body:
        if child.tag in (qn('w:p'), qn('w:tbl')):
            new_element = copy.deepcopy(child)
            target_body.insert(-1, new_element)

def safe_replace_text(doc, replace_dict):
    """全量深度替换：涵盖页眉、正文、表格"""
    if not replace_dict: return
    for section in doc.sections:
        header = section.header
        for p in header.paragraphs:
            _refined_replace(p, replace_dict)
        for table in header.tables:
            _replace_in_table(table, replace_dict)
    for p in doc.paragraphs:
        _refined_replace(p, replace_dict)
    for table in doc.tables:
        _replace_in_table(table, replace_dict)

def _refined_replace(paragraph, replace_dict):
    """精细化替换逻辑：处理跨 Run 拆分问题"""
    p_text = paragraph.text
    needs_replace = False
    new_p_text = p_text
    for old_text, new_text in replace_dict.items():
        if old_text and old_text in new_p_text:
            new_p_text = new_p_text.replace(old_text, new_text)
            needs_replace = True
    
    if not needs_replace: return

    # 先尝试 Run 级别替换（保留格式）
    for old_text, new_text in replace_dict.items():
        for run in paragraph.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
    
    # 最终检查：如果还有残留（说明被拆分了），强制重写第一个 Run
    final_check = paragraph.text
    if any(old in final_check for old in replace_dict.keys()):
        final_t = final_check
        for old, new in replace_dict.items():
            final_t = final_t.replace(old, new)
        if paragraph.runs:
            paragraph.runs[0].text = final_t
            for i in range(1, len(paragraph.runs)):
                paragraph.runs[i].text = ""

def _replace_in_table(table, replace_dict):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                _refined_replace(p, replace_dict)

def get_deepseek_rules(api_key, base_url, doc_sample, user_demand):
    """增强版 AI 规则提取：强制要求精准对子"""
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    # 核心改进：在 Prompt 中明确要求只输出核心词汇对
    prompt = (
        f"你是一个文档替换专家。请分析以下内容并根据需求提取【关键词】替换对。\n"
        f"文档样本：{doc_sample}\n"
        f"需求：{user_demand}\n\n"
        f"重要指令：\n"
        f"1. 识别样本中提到的实体（如学校名、公司名、标题关键词）。\n"
        f"2. 严格按格式输出：旧词 ==>> 新词\n"
        f"3. 不要添加解释，不要加『学校』或『名称』等修饰词，除非原文就有。\n"
        f"4. 确保替换对非常精准，以便在页眉标题中也能匹配。"
    )
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
    replace_dict = {}
    for line in rules_text.split('\n'):
        if "==>>" in line:
            parts = line.split("==>>")
            if len(parts) == 2:
                k, v = parts[0].strip(), parts[1].strip()
                if k: replace_dict[k] = v
    return replace_dict

def main():
    st.title("📝 Word 智能替换助手")
    def_key, def_url = get_config()
    with st.expander("🔑 API 设置"):
        api_key = st.text_input("DeepSeek API Key", value=def_key, type="password")
        base_url = st.text_input("Base URL", value=def_url)

    st.markdown("### 📖 使用说明\n1. 上传模板文件。\n2. 描述修改需求（如：把第一中学改为高级中学）。\n3. 系统将尝试智能识别并同步替换标题及正文。")

    uploaded_file = st.file_uploader("1. 上传模板文件", type=["docx"])
    user_input = st.text_area("2. 描述修改需求", placeholder="例如：把第一中学改为高级中学...", height=120)
    
    if st.button("🚀 开始自动替换并生成", type="primary", use_container_width=True):
        if not uploaded_file or not user_input:
            st.error("请确保已上传文件并输入需求")
            return

        with st.spinner("正在精准扫描替换..."):
            content_bytes = uploaded_file.getvalue()
            source_doc = Document(io.BytesIO(content_bytes))
            container_doc = Document(io.BytesIO(content_bytes))
            
            if "==>>" in user_input:
                replacements = parse_rules(user_input)
            else:
                if not api_key:
                    st.error("请输入 API Key")
                    return
                # 增加样本采集范围，尽量抓到标题
                sample_text = "\n".join([p.text for p in source_doc.paragraphs[:30]])
                ai_raw = get_deepseek_rules(api_key, base_url, sample_text, user_input)
                replacements = parse_rules(ai_raw)
            
            if not replacements:
                st.error("未能识别有效规则。")
                return

            safe_replace_text(source_doc, replacements)
            clone_only_body_content(source_doc, container_doc)
            safe_replace_text(container_doc, replacements)

            output = io.BytesIO()
            container_doc.save(output)
            output.seek(0)
            
            st.success(f"✅ 处理完成！已执行替换：{list(replacements.keys())}")
            st.download_button("📥 下载结果文件", data=output, file_name=f"Fixed_{uploaded_file.name}", use_container_width=True)

if __name__ == "__main__":
    main()
