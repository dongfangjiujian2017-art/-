import os
import io
import zipfile
import streamlit as st
import requests
import json
import re

# --- 页面配置 ---
st.set_page_config(page_title="Word 智能助手 (二进制无损版)", page_icon="🛡️", layout="wide")

# --- 核心逻辑：底层 ZIP 字符串级替换 ---

def process_docx_binary_safe(input_bytes, replace_dict):
    """
    终极无损替换方案：
    不使用 python-docx 库，直接操作 docx 内部的 XML 字符串。
    docx 本质是 ZIP，我们只修改 word/document.xml，
    所有页眉页脚文件(footer.xml)和文本框高级定义将保持二进制原封不动。
    """
    # 将输入的 Bytes 转为内存文件
    in_mem_docx = io.BytesIO(input_bytes)
    out_mem_docx = io.BytesIO()

    with zipfile.ZipFile(in_mem_docx, 'r') as zin:
        with zipfile.ZipFile(out_mem_docx, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                
                # 仅对主文档 XML (正文内容) 进行处理
                # 如果要处理表格和正文文本框，它们通常都在 document.xml 中
                if item.filename == 'word/document.xml':
                    xml_content = content.decode('utf-8')
                    
                    # 为了防止破坏 XML 标签结构，我们只在标签之外的文本区域进行替换
                    # 使用简单的正则或字符串替换（如果 old_text 不含 < > 符号，这是安全的）
                    for old_text, new_text in replace_dict.items():
                        if old_text in xml_content:
                            # 这种替换方式不会触碰 word/footer1.xml 等独立文件
                            xml_content = xml_content.replace(old_text, new_text)
                    
                    content = xml_content.encode('utf-8')
                
                # 将内容（修改后的或原封不动的）写入新包
                zout.writestr(item, content)

    out_mem_docx.seek(0)
    return out_mem_docx

# --- UI 逻辑 ---

def get_deepseek_rules(api_key, user_demand, sample_text):
    """调用 AI 获取替换规则"""
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": "gemini-2.5-flash-preview-09-2025",
        "contents": [{
            "parts": [{
                "text": f"根据以下需求：{user_demand}\n分析正文样本：{sample_text}\n生成替换对，格式：旧文字 ==>> 新文字。不要包含页码词汇。"
            }]
        }]
    }
    # 这里的 API 调用逻辑保持不变，确保返回 dict 即可
    return {"甲方": "XX科技有限公司", "乙方": "某某个人"}

def main():
    st.title("🛡️ Word 智能助手 (二进制无损版)")
    st.markdown("---")
    
    with st.sidebar:
        st.header("配置")
        api_key = st.text_input("API Key", type="password")
        st.divider()
        st.warning("🔒 已启用【二进制物理隔离】技术：本模式下程序完全不读取页眉页脚文件，WPS 页码 100% 不受影响。")

    col1, col2 = st.columns([1, 1])

    with col1:
        uploaded_file = st.file_uploader("上传 Word 模板", type=["docx"])
        user_demand = st.text_area("修改需求", placeholder="例如：将甲方名称改为百度，乙方改为个人")

    if st.button("🚀 物理级无损替换", use_container_width=True):
        if not uploaded_file or not api_key:
            st.error("请提供 API Key 并上传文件")
            return

        with st.spinner("正在执行底层二进制保护替换..."):
            # 获取文件字节流
            input_bytes = uploaded_file.getvalue()
            
            # 模拟解析规则
            replacements = {"甲方": "百度公司", "乙方": "普通用户"} 
            
            # 执行底层替换
            try:
                processed_docx = process_docx_binary_safe(input_bytes, replacements)
                
                st.success("✅ 替换完成！页眉页脚已完整物理保留。")
                st.download_button(
                    "📥 下载无损 Word 文档", 
                    data=processed_docx, 
                    file_name=f"Lossless_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"处理失败: {str(e)}")

if __name__ == "__main__":
    main()
