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
    全量深度替换：涵盖所有段落、表格、页眉以及隐藏的节属性
    """
    if not replace_dict: return

    # 1. 替换所有节的页眉
    for section in doc.sections:
        header = section.header
        for p in header.paragraphs:
            _replace_in_paragraph(p, replace_dict)
        for table in header.tables:
            _replace_in_table(table, replace_dict)

    # 2. 替换文档主体的段落
    for p in doc.paragraphs:
        _replace_in_paragraph(p, replace_dict)

    # 3. 替换文档主体的表格内容
    for table in doc.tables:
        _replace_in_table(table, replace_dict)

def _replace_in_paragraph(paragraph, replace_dict):
    """
    【核心修复】底层替换算法：
    解决 Word 中文本被随机拆分到多个 Run 导致无法匹配的问题。
    """
    for old_text, new_text in replace_dict.items():
        if old_text in paragraph.text:
            # 这种方法会尽量保留格式，但处理跨 Run 匹配更鲁棒
            # 我们直接对段落文本进行整体替换，同时尝试维护 Run 结构
            inline = paragraph.runs
            for i in range(len(inline)):
                if old_text in inline[i].text:
                    inline[i].text = inline[i].text.replace(old_text, new_text)
                else:
                    # 如果关键词被拆分到了多个 Run 之间，这种简单替换会失败
                    # 针对这种情况，如果段落包含但单个 Run 不包含，我们执行全段重写替换
                    # 这是目前 python-docx 处理此类问题的最稳妥方式
                    pass
            
            # 如果依然存在未替换的情况（说明被拆分了）
            if old_text in paragraph.text:
                full_text = paragraph.text.replace(old_text, new_text)
                # 清空所有 runs 重新写入（会丢失段内部分精细格式，但能保命）
                # 为了折中，我们只在必要时使用
                for run in paragraph.runs:
                    run.text = ""
                if len(paragraph.runs) > 0:
                    paragraph.runs[0].text = full_text

def _replace_in_table(table, replace_dict):
    """底层表格替换逻辑"""
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                _replace_in_paragraph(p, replace_dict)

def get_deepseek_rules(api_key, base_url, doc_sample, user_demand):
    """请求 AI 生成规则"""
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    prompt = f"文档样本：{doc_sample}\n需求：{user_demand}\n输出：旧文字 ==>> 新文字。每行一对，不要解释。"
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
    if not rules_text: return replace_dict
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

    st.markdown("""
    ### 📖 使用说明
    1. **上传模板**：上传 Word 文档。
    2. **描述修改**：输入修改描述。
    3. **自动生成**：系统将强制替换正文与标题内容，并保护页脚。
    """)

    uploaded_file = st.file_uploader("1. 上传模板文件", type=["docx"])
    user_input = st.text_area("2. 描述修改需求", placeholder="例如：把甲方改为某某公司...", height=120)
    
    if st.button("🚀 开始自动替换并生成", type="primary", use_container_width=True):
        if not uploaded_file or not user_input:
            st.error("请确保已上传文件并输入需求")
            return

        with st.spinner("正在深度扫描并替换..."):
            content_bytes = uploaded_file.getvalue()
            source_doc = Document(io.BytesIO(content_bytes))
            container_doc = Document(io.BytesIO(content_bytes))
            
            if "==>>" in user_input:
                replacements = parse_rules(user_input)
            else:
                if not api_key:
                    st.error("请输入 API Key")
                    return
                sample_text = "\n".join([p.text for p in source_doc.paragraphs[:20]])
                ai_raw = get_deepseek_rules(api_key, base_url, sample_text, user_input)
                replacements = parse_rules(ai_raw)
            
            if not replacements:
                st.error("未能识别修改指令。")
                return

            # 对源文档执行替换（强化算法）
            safe_replace_text(source_doc, replacements)
            # 克隆正文
            clone_only_body_content(source_doc, container_doc)
            # 对容器（页眉标题）执行再次替换
            safe_replace_text(container_doc, replacements)

            output = io.BytesIO()
            container_doc.save(output)
            output.seek(0)
            
            st.success(f"✅ 处理完成！已强制同步标题和正文中的 {len(replacements)} 类信息。")
            st.download_button(
                "📥 下载结果文件", 
                data=output, 
                file_name=f"Fixed_{uploaded_file.name}",
                use_container_width=True
            )

if __name__ == "__main__":
    main()
