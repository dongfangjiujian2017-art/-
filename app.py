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
    # 移除目标文档正文中除 sectPr 以外的所有子节点
    for child in list(target_body):
        if child.tag != qn('w:sectPr'):
            target_body.remove(child)

    source_body = source_doc._element.body
    for child in source_body:
        if child.tag in (qn('w:p'), qn('w:tbl')):
            new_element = copy.deepcopy(child)
            # 插入到最后一个 sectPr 之前
            target_body.insert(-1, new_element)

def safe_replace_text(doc, replace_dict):
    """
    深度扫描并替换：涵盖页眉、正文、表格
    """
    if not replace_dict: return

    # 1. 扫描所有节的页眉 (处理标题)
    for section in doc.sections:
        header = section.header
        for p in header.paragraphs:
            _refined_replace(p, replace_dict)
        for table in header.tables:
            _replace_in_table(table, replace_dict)

    # 2. 扫描正文段落
    for p in doc.paragraphs:
        _refined_replace(p, replace_dict)

    # 3. 扫描表格内容
    for table in doc.tables:
        _replace_in_table(table, replace_dict)

def _refined_replace(paragraph, replace_dict):
    """
    精细化替换逻辑：解决 Run 拆分问题，同时防止重复替换
    """
    p_text = paragraph.text
    needs_replace = False
    
    # 先检查是否真的需要替换
    new_p_text = p_text
    for old_text, new_text in replace_dict.items():
        if old_text and old_text in new_p_text:
            new_p_text = new_p_text.replace(old_text, new_text)
            needs_replace = True
    
    if not needs_replace:
        return

    # 尝试在 Run 级别替换以保留格式
    for old_text, new_text in replace_dict.items():
        for run in paragraph.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
    
    # 关键检查：如果 Run 级别替换后，段落中依然残留旧文本（说明被拆分了）
    # 则执行“保命”重写：合并全文并写入第一个 Run，清空后续 Runs
    final_check_text = paragraph.text
    still_has_old = any(old in final_check_text for old in replace_dict.keys())
    
    if still_has_old:
        # 彻底执行一次性的最终替换
        final_text = final_check_text
        for old, new in replace_dict.items():
            final_text = final_text.replace(old, new)
        
        if len(paragraph.runs) > 0:
            paragraph.runs[0].text = final_text
            for i in range(1, len(paragraph.runs)):
                paragraph.runs[i].text = ""

def _replace_in_table(table, replace_dict):
    """表格递归替换"""
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                _refined_replace(p, replace_dict)

def get_deepseek_rules(api_key, base_url, doc_sample, user_demand):
    """AI 规则提取"""
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    prompt = f"正文样本：{doc_sample}\n需求：{user_demand}\n输出：旧文字 ==>> 新文字。每行一对，不要解释。"
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
    """解析 AI 返回的规则列表"""
    replace_dict = {}
    if not rules_text: return replace_dict
    for line in rules_text.split('\n'):
        if "==>>" in line:
            parts = line.split("==>>")
            if len(parts) == 2:
                key = parts[0].strip()
                val = parts[1].strip()
                if key: replace_dict[key] = val
    return replace_dict

def main():
    st.title("📝 Word 智能替换助手")
    
    def_key, def_url = get_config()
    with st.expander("🔑 API 设置"):
        api_key = st.text_input("DeepSeek API Key", value=def_key, type="password")
        base_url = st.text_input("Base URL", value=def_url)

    st.markdown("""
    ### 📖 使用说明
    1. **上传模板**：上传 Word 文档。
    2. **输入需求**：描述您想修改的内容（如：把第一中学改为高级中学）。
    3. **自动生成**：系统将精准同步标题与正文，并保护页脚。
    """)

    uploaded_file = st.file_uploader("1. 上传模板文件", type=["docx"])
    user_input = st.text_area("2. 描述修改需求", placeholder="例如：把文档中的第一中学改为高级中学...", height=120)
    
    if st.button("🚀 开始自动替换并生成", type="primary", use_container_width=True):
        if not uploaded_file or not user_input:
            st.error("请确保已上传文件并输入需求")
            return

        with st.spinner("正在进行精准扫描替换..."):
            content_bytes = uploaded_file.getvalue()
            
            # 建立双实例
            source_doc = Document(io.BytesIO(content_bytes))
            container_doc = Document(io.BytesIO(content_bytes))
            
            # 规则获取
            if "==>>" in user_input:
                replacements = parse_rules(user_input)
            else:
                if not api_key:
                    st.error("请输入 API Key")
                    return
                sample_text = "\n".join([p.text for p in source_doc.paragraphs[:15]])
                ai_raw = get_deepseek_rules(api_key, base_url, sample_text, user_input)
                replacements = parse_rules(ai_raw)
            
            if not replacements:
                st.error("未能识别修改指令。")
                return

            # 执行源文档替换
            safe_replace_text(source_doc, replacements)
            
            # 正文搬运
            clone_only_body_content(source_doc, container_doc)
            
            # 容器（页眉/标题）同步替换
            safe_replace_text(container_doc, replacements)

            output = io.BytesIO()
            container_doc.save(output)
            output.seek(0)
            
            st.success(f"✅ 处理完成！已完成 {len(replacements)} 类关键词的精准替换。")
            st.download_button(
                "📥 下载结果文件", 
                data=output, 
                file_name=f"Fixed_{uploaded_file.name}",
                use_container_width=True
            )

if __name__ == "__main__":
    main()
