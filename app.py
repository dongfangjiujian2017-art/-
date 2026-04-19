import os
import io
import streamlit as st
import requests
import copy
from docx import Document

# --- 页面配置 ---
st.set_page_config(page_title="Word 无损替换助手", page_icon="🛡️", layout="wide")

# --- 配置读取逻辑 ---
def get_config():
    """
    自动从 Streamlit Secrets 或环境变量中获取配置
    """
    default_key = ""
    default_url = "https://api.deepseek.com"
    
    # 优先从 st.secrets 读取 (Streamlit Cloud 推荐方式)
    if "DEEPSEEK_API_KEY" in st.secrets:
        default_key = st.secrets["DEEPSEEK_API_KEY"]
    if "DEEPSEEK_BASE_URL" in st.secrets:
        default_url = st.secrets["DEEPSEEK_BASE_URL"]
        
    return default_key, default_url

def clone_content_to_template(source_doc, target_doc):
    """
    核心无损克隆：将 source_doc 的正文搬运到 target_doc，
    target_doc 的页眉页脚（原始格式）将被 100% 保留。
    """
    # 移除目标文档的所有正文段落
    for p in target_doc.paragraphs:
        p._element.getparent().remove(p._element)

    # 搬运正文段落
    for paragraph in source_doc.paragraphs:
        new_p = target_doc.add_paragraph()
        new_p._element.clear_content()
        for child in paragraph._element.iterchildren():
            # 仅复制段落属性(pPr)和运行块(r)
            if child.tag.endswith('pPr') or child.tag.endswith('r'):
                new_p._element.append(copy.deepcopy(child))

    # 搬运表格内容
    for table in source_doc.tables:
        target_doc._element.body.append(copy.deepcopy(table._element))

def safe_replace_text(doc, replace_dict):
    """在文档正文和表格中执行文本替换"""
    if not replace_dict:
        return
        
    # 替换段落
    for p in doc.paragraphs:
        for old_txt, new_txt in replace_dict.items():
            if old_txt in p.text:
                for run in p.runs:
                    if old_txt in run.text:
                        run.text = run.text.replace(old_txt, new_txt)
    
    # 替换表格
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
    """请求 DeepSeek API 自动生成替换规则"""
    if not api_key:
        return "❌ 错误：未检测到 API Key，请在侧边栏输入或在 Secrets 中配置。"

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    prompt = f"文档样本内容：{doc_sample}\n修改需求：{user_demand}\n请输出需要替换的对子。格式要求：旧内容 ==>> 新内容。每行一对，不要包含任何额外文字或解释。"
    
    data = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3
    }
    
    try:
        response = requests.post(f"{base_url}/chat/completions", headers=headers, json=data, timeout=30)
        if response.status_code == 401:
            return "❌ API Key 效验失败 (401)。请检查配置是否正确。"
        response.raise_for_status()
        return response.json()['choices'][0]['message']['content']
    except Exception as e:
        return f"❌ 规则生成失败: {str(e)}"

def parse_rules(rules_text):
    """解析 AI 生成的文本为字典格式"""
    replace_dict = {}
    if not rules_text: return replace_dict
    for line in rules_text.split('\n'):
        if "==>>" in line:
            parts = line.split("==>>")
            if len(parts) == 2:
                replace_dict[parts[0].strip()] = parts[1].strip()
    return replace_dict

def main():
    st.title("🛡️ Word 智能替换助手 (单文件无损版)")
    
    # 获取默认配置
    def_key, def_url = get_config()
    
    with st.sidebar:
        st.header("⚙️ 账号配置")
        api_key = st.text_input("DeepSeek API Key", value=def_key, type="password")
        base_url = st.text_input("Base URL", value=def_url)
        
        if st.button("🔌 测试 API 连接"):
            if not api_key:
                st.error("请输入 API Key")
            else:
                test_headers = {"Authorization": f"Bearer {api_key}"}
                try:
                    res = requests.get(f"{base_url}/models", headers=test_headers, timeout=10)
                    if res.status_code == 200: st.success("连接成功！")
                    else: st.error(f"连接失败: {res.status_code}")
                except Exception as e:
                    st.error(f"连接异常: {e}")
        
        st.divider()
        st.info("💡 提示：本程序采用『自注入』技术。上传的文件既是模板也是内容源，程序会保护页码 XML 不受触碰。")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. 上传文件")
        uploaded_file = st.file_uploader("选择 Word 文档 (.docx)", type=["docx"])
        
        st.subheader("2. 修改需求")
        user_demand = st.text_area("你想怎么修改文档？", placeholder="例如：把甲方公司改为『字节跳动』", height=150)
        
        if st.button("✨ 智能分析规则", type="primary"):
            if not uploaded_file:
                st.warning("请先上传文件")
            else:
                with st.spinner("AI 正在分析并提取规则..."):
                    content_bytes = uploaded_file.getvalue()
                    sample_doc = Document(io.BytesIO(content_bytes))
                    # 提取前20个段落作为样本
                    sample_text = "\n".join([p.text for p in sample_doc.paragraphs[:20]])
                    rules = get_deepseek_rules(api_key, base_url, sample_text, user_demand)
                    st.session_state.ai_rules = rules

    with col2:
        st.subheader("3. 执行并下载")
        rules_text = st.text_area("替换规则预览 (可手动编辑)", value=st.session_state.get("ai_rules", ""), height=250)
        
        if st.button("🚀 执行处理", use_container_width=True):
            if not uploaded_file or not rules_text:
                st.error("请确保已上传文件并生成规则")
            else:
                replacements = parse_rules(rules_text)
                with st.spinner("正在克隆正文并物理隔离页脚..."):
                    content_bytes = uploaded_file.getvalue()
                    
                    # 1. 它是我们要保留正确页脚的容器（模板）
                    container_doc = Document(io.BytesIO(content_bytes))
                    # 2. 它是我们要修改文字的内容源
                    source_doc = Document(io.BytesIO(content_bytes))

                    # 3. 在内容源中修改文字
                    safe_replace_text(source_doc, replacements)

                    # 4. 把改好的内容塞进干净的容器
                    clone_content_to_template(source_doc, container_doc)

                    # 5. 保存并提供下载
                    output = io.BytesIO()
                    container_doc.save(output)
                    output.seek(0)
                    
                    st.success("✅ 处理完成！页脚格式已成功保护。")
                    st.download_button(
                        "📥 下载修复后的文档", 
                        data=output, 
                        file_name=f"Fixed_{uploaded_file.name}",
                        use_container_width=True
                    )

if __name__ == "__main__":
    main()
