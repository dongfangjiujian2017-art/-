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
    # 清空目标正文，保留 sectPr
    for child in list(target_body):
        if child.tag != qn('w:sectPr'):
            target_body.remove(child)

    source_body = source_doc._element.body
    for child in source_body:
        if child.tag in (qn('w:p'), qn('w:tbl')):
            new_element = copy.deepcopy(child)
            target_body.insert(-1, new_element)

def safe_replace_text(doc, replace_dict):
    """执行文本替换"""
    if not replace_dict: return
    for p in doc.paragraphs:
        for old, new in replace_dict.items():
            if old in p.text:
                for run in p.runs:
                    if old in run.text:
                        run.text = run.text.replace(old, new)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for old, new in replace_dict.items():
                        if old in p.text:
                            for run in p.runs:
                                if old in run.text:
                                    run.text = run.text.replace(old, new)

def get_deepseek_rules(api_key, base_url, doc_sample, user_demand):
    """请求 AI 生成规则"""
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    prompt = f"正文样本：{doc_sample}\n需求：{user_demand}\n输出格式：旧内容 ==>> 新内容。每行一对，严禁其他文字。"
    data = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1
    }
    for _ in range(3): # 简化重试逻辑
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
    
    # 账号配置 (折叠隐藏)
    def_key, def_url = get_config()
    with st.expander("🔑 API 设置 (初次使用或 Key 失效时展开)"):
        api_key = st.text_input("DeepSeek API Key", value=def_key, type="password")
        base_url = st.text_input("Base URL", value=def_url)

    # 使用说明
    st.markdown("""
    ### 📖 使用说明
    1. **上传模板**：点击下方上传您的 Word 文档 (.docx)。
    2. **输入需求**：粘贴您在豆包、DeepSeek 中生成的规则（如：`A ==>> B`）或描述修改需求。
    3. **自动替换**：点击按钮，系统将自动保留页脚并生成新文档。
    """)

    # 主体操作区
    uploaded_file = st.file_uploader("1. 上传模板文件", type=["docx"])
    user_input = st.text_area("2. 粘贴规则或输入修改需求", placeholder="格式例：\n甲方 ==>> 未来科技\n金额 ==>> 壹佰万元\n（或者直接输入：把甲方改为未来科技...）", height=150)
    
    if st.button("🚀 开始自动替换并生成", type="primary", use_container_width=True):
        if not uploaded_file or not user_input:
            st.error("请确保已上传文件并输入需求")
            return

        with st.spinner("正在处理，请稍候..."):
            content_bytes = uploaded_file.getvalue()
            source_doc = Document(io.BytesIO(content_bytes))
            
            # 判断输入内容：如果是 AI 已经生成的格式，直接解析；否则请求 AI 生成
            if "==>>" in user_input:
                replacements = parse_rules(user_input)
            else:
                if not api_key:
                    st.error("请输入 API Key 以启用智能分析")
                    return
                sample_text = "\n".join([p.text for p in source_doc.paragraphs[:15]])
                ai_raw = get_deepseek_rules(api_key, base_url, sample_text, user_input)
                replacements = parse_rules(ai_raw)
            
            if not replacements:
                st.error("未能识别有效替换规则，请检查输入格式。")
                return

            # 执行注入
            container_doc = Document(io.BytesIO(content_bytes))
            safe_replace_text(source_doc, replacements)
            clone_only_body_content(source_doc, container_doc)

            # 保存下载
            output = io.BytesIO()
            container_doc.save(output)
            output.seek(0)
            
            st.success(f"✅ 成功替换 {len(replacements)} 处内容！")
            st.download_button(
                "📥 点击下载新文档", 
                data=output, 
                file_name=f"已替换_{uploaded_file.name}",
                use_container_width=True
            )

if __name__ == "__main__":
    main()
