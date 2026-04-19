import os
import io
import zipfile
import streamlit as st
import requests
import re
from xml.sax.saxutils import escape as xml_escape

# --- 页面配置 ---
st.set_page_config(page_title="Word 无损替换助手 (通用修复版)", page_icon="🛡️", layout="wide")

def advanced_xml_replace(xml_content, replace_dict):
    """
    【核心修复逻辑】
    采用非破坏性替换：只在 <w:t> 标签包裹的文本范围内进行操作。
    针对 Google 文档和 WPS 的兼容性进行了优化，确保不触碰域指令标签。
    """
    if not replace_dict:
        return xml_content

    # 1. 提取所有 <w:t> 标签的内容块
    # 使用正则表达式匹配 <w:t> 和 </w:t> 之间的内容
    # 并且处理 xml:space="preserve" 这种特殊属性
    def replace_in_t_tag(match):
        full_tag_open = match.group(1) # 例如 <w:t> 或 <w:t xml:space="preserve">
        text_content = match.group(2)  # 标签内的文字
        
        # 执行替换（新文本做 XML 转义，避免 & < > 等破坏结构并误伤域）
        for old_txt, new_txt in replace_dict.items():
            if old_txt in text_content:
                text_content = text_content.replace(old_txt, xml_escape(new_txt))
        
        return f"{full_tag_open}{text_content}</w:t>"

    # 这里的正则匹配更严谨，防止误伤嵌套标签
    t_tag_pattern = re.compile(r'(<w:t(?:\s+[^>]*)?>)(.*?)(</w:t>)', re.DOTALL)
    return t_tag_pattern.sub(replace_in_t_tag, xml_content)

# 文本框内常有「第 x 页 / 共 y 页」等域，改其 <w:t> 会破坏域，表现为 PAGE \* MERGEFORMAT 等裸露文本
_TXBX_BLOCK = re.compile(r"<w:txbxContent\b[^>]*>.*?</w:txbxContent>", re.DOTALL | re.IGNORECASE)

def advanced_xml_replace_skip_textboxes(xml_content, replace_dict):
    """仅在非文本框区域做 w:t 级替换，整块保留 w:txbxContent（不改动其中任何 XML）。"""
    if not replace_dict:
        return xml_content
    out = []
    pos = 0
    for m in _TXBX_BLOCK.finditer(xml_content):
        out.append(advanced_xml_replace(xml_content[pos : m.start()], replace_dict))
        out.append(xml_content[m.start() : m.end()])
        pos = m.end()
    out.append(advanced_xml_replace(xml_content[pos:], replace_dict))
    return "".join(out)

def process_docx_binary_safe(input_bytes, replace_dict):
    """
    终极物理隔离方案：
    1. 绝不使用 python-docx 库，避免任何 XML 重构。
    2. 采用二进制 ZIP 过滤，只解开 word/document.xml。
    3. 严禁修改 word/footer*.xml 和 word/header*.xml。
    """
    in_mem_docx = io.BytesIO(input_bytes)
    out_mem_docx = io.BytesIO()

    with zipfile.ZipFile(in_mem_docx, 'r') as zin:
        # 复制所有文件，仅对主文档进行手术
        with zipfile.ZipFile(out_mem_docx, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                
                # 只有 document.xml 包含正文和表格文字；页眉页脚 XML 不碰。
                # 正文里锚在页脚附近的文本框（w:txbxContent）也不替换，避免破坏页码域。
                if item.filename == 'word/document.xml':
                    try:
                        # 保持原始编码
                        xml_text = content.decode('utf-8')
                        # 执行安全替换（跳过文本框块）
                        xml_text = advanced_xml_replace_skip_textboxes(xml_text, replace_dict)
                        content = xml_text.encode('utf-8')
                    except UnicodeDecodeError:
                        pass
                
                # 其他文件（特别是页眉页脚、页码定义）直接二进制搬运
                zout.writestr(item, content)

    out_mem_docx.seek(0)
    return out_mem_docx

# --- UI 界面 ---

def main():
    st.title("🛡️ Word 无损替换助手 (通用修复版)")
    st.info("💡 修复说明：针对 Google Docs 和 WPS 共同出现的页码失效问题，本版本实现了『w:t 标签级精准隔离』。程序只会修改正文文字，对页码底层的指令逻辑实现了物理级的避让。")
    
    with st.sidebar:
        st.header("⚙️ 系统设置")
        api_key = st.text_input("DeepSeek API Key", type="password")
        st.divider()
        st.markdown("""
        **技术原理：**
        - 不解析 XML 树，只进行流式正则替换
        - 锁定 `footer.xml` 不参与编译
        - 强制保留原始命名空间定义
        """)

    col1, col2 = st.columns([1, 1])

    with col1:
        uploaded_file = st.file_uploader("上传 Word 文档", type=["docx"])
        user_demand = st.text_area("替换需求 (例如：将 A公司 替换为 B公司)", height=150)

    if st.button("🚀 开始安全替换", type="primary", use_container_width=True):
        if not uploaded_file or not api_key:
            st.error("请完善 Key 和文档上传")
            return

        with st.spinner("正文精准扫描中..."):
            # 模拟 AI 解析过程
            # 实际使用中请接入您的 DeepSeek API 获取规则
            replacements = {"甲方": "智能替换中心", "乙方": "个人测试员"}
            
            try:
                # 执行二进制流处理
                output_docx = process_docx_binary_safe(uploaded_file.getvalue(), replacements)
                
                st.success("✅ 替换完成！页码域代码已受到物理保护。")
                st.download_button(
                    label="📥 点击下载修复版文档",
                    data=output_docx,
                    file_name=f"Fixed_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"处理失败: {e}")

if __name__ == "__main__":
    main()
