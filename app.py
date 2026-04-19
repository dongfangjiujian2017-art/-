import os
import io
import streamlit as st
import requests
import copy
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn

# --- 页面配置 ---
st.set_page_config(page_title="Word 无损替换助手 (模板克隆版)", page_icon="🛡️", layout="wide")

# --- 配置读取逻辑 ---
def get_config():
    """
    优先从 streamlit secrets 读取配置，如果没有则返回空字符串。
    用户可以在 .streamlit/secrets.toml 中配置:
    DEEPSEEK_API_KEY = "your_key_here"
    DEEPSEEK_BASE_URL = "https://api.deepseek.com"
    """
    default_key = ""
    default_url = "https://api.deepseek.com"
    
    if "DEEPSEEK_API_KEY" in st.secrets:
        default_key = st.secrets["DEEPSEEK_API_KEY"]
    if "DEEPSEEK_BASE_URL" in st.secrets:
        default_url = st.secrets["DEEPSEEK_BASE_URL"]
        
    return default_key, default_url

def clone_content_to_template(source_doc, target_doc):
    """
    将 source_doc 的正文内容克隆到 target_doc 中。
    这样 target_doc 原有的页眉页脚（正确的）会被保留。
    """
    # 清空目标文档的正文段落
    for p in target_doc.paragraphs:
        p._element.getparent().remove(p._element)

    # 将源文档的所有段落添加到目标文档
    for paragraph in source_doc.paragraphs:
        new_p = target_doc.add_paragraph()
        new_p._element.clear_content()
        for child in paragraph._element.iterchildren():
            if child.tag.endswith('pPr') or child.tag.endswith('r'):
                new_p._element.append(copy.deepcopy(child))

    # 复制表格
    for table in source_doc.tables:
        target_doc._element.body.append(copy.deepcopy(table._element))

def safe_replace_text(doc, replace_dict):
    """在文档正文中执行替换"""
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
    """调用 DeepSeek API 获取替换规则"""
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    prompt = f"分析文档内容：{doc_sample}\n修改需求：{user_demand}\n请输出需要替换的对子，格式为：旧内容 ==>> 新内容。每行一对。不要输出任何其他解释。"
    
    data = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "stream": False
    }
    
    try:
        response = requests.post(f"{base_url}/chat/completions", headers=headers, json=data, timeout=60)
        response.raise_for_status()
        return response.json()['choices'][0]['message']['content']
    except Exception as e:
        return f"❌ 规则生成失败: {str(e)}"

def parse_rules(rules_text):
    """解析 AI 生成的规则"""
    replace_dict = {}
    for line in rules_text.split('\n'):
        if "==>>" in line:
            parts = line.split("==>>")
            if len(parts) == 2:
                replace_dict[parts[0].strip()] = parts[1].strip()
    return replace_dict

def main():
    st.title("🛡️ Word 无损替换助手 (模板克隆版)")
    
    # 读取默认配置
    def_key, def_url = get_config()
    
    with st.sidebar:
        st.header("⚙️ 账号配置")
        api_key = st.text_input("DeepSeek API Key", value=def_key, type="password", help="可在 .streamlit/secrets.toml 中预设")
        base_url = st.text_input("Base URL", value=def_url)
        st.divider()
        st.markdown("""
        **使用指南：**
        1. 在侧边栏配置 API Key。
        2. 上传一份**页脚显示正确**的文档（哪怕是只有页脚的空文档）。
        3. 上传你需要修改内容的文档。
        4. 点击生成并下载。
        """)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. 准备文件")
        template_file = st.file_uploader("上传【页脚正确】的模板", type=["docx"], key="tpl")
        content_file = st.file_uploader("上传【待修改内容】的文档", type=["docx"], key="cnt")
        
        st.subheader("2. 修改需求")
        user_demand = st.text_area("告诉 AI 你想修改什么内容？", placeholder="例如：将甲方改为华为，金额改为一千万元...")
        
        if st.button("✨ 智能生成规则", type="primary"):
            if not api_key or not content_file:
                st.warning("请检查 API Key 和内容文档是否已就绪")
            else:
                with st.spinner("AI 正在分析文档内容..."):
                    # 加载部分内容用于分析
                    sample_doc = Document(content_file)
                    sample_text = "\n".join([p.text for p in sample_doc.paragraphs[:20]])
                    rules = get_deepseek_rules(api_key, base_url, sample_text, user_demand)
                    st.session_state.ai_rules = rules

    with col2:
        st.subheader("3. 确认并执行")
        rules_text = st.text_area("替换规则预览 (可手动修改)", value=st.session_state.get("ai_rules", ""), height=250)
        
        if st.button("🚀 执行克隆注入", use_container_width=True):
            if not template_file or not content_file or not rules_text:
                st.error("请确保模板、内容文档和规则都已就绪")
            else:
                replacements = parse_rules(rules_text)
                with st.spinner("正在克隆正文并注入模板..."):
                    # 1. 加载文件
                    tpl_doc = Document(template_file)
                    cnt_doc = Document(content_file)

                    # 2. 在内容文档副本中执行替换
                    safe_replace_text(cnt_doc, replacements)

                    # 3. 将内容克隆进正确页脚的模板
                    clone_content_to_template(cnt_doc, tpl_doc)

                    # 4. 保存输出
                    output = io.BytesIO()
                    tpl_doc.save(output)
                    output.seek(0)
                    st.success("✅ 处理完成！页脚已成功保护。")
                    st.download_button("📥 下载最终成品", data=output, file_name="Protected_Footer_Doc.docx", use_container_width=True)

if __name__ == "__main__":
    main()
