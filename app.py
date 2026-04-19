import os
import io
import json
import time
import streamlit as st
from docx import Document

# #region agent log
_AGENT_LOG_PATH = r"d:\UserData\Documents\cursor\debug-374127.log"

def _agent_log(*, hypothesis_id, location, message, data, run_id="pre-fix"):
    payload = {
        "sessionId": "374127",
        "runId": run_id,
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": data,
        "timestamp": int(time.time() * 1000),
    }
    try:
        with open(_AGENT_LOG_PATH, "a", encoding="utf-8") as _f:
            _f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except OSError:
        pass
# #endregion

# 设置网页标题
st.set_page_config(page_title="Word格式保留替换工具", layout="centered")

def parse_replace_text(text):
    """
    解析用户输入的替换规则。
    支持格式：旧内容 ==>> 新内容
    """
    replace_dict = {}
    lines = text.split('\n')
    for line in lines:
        if "==>>" in line:
            parts = line.split("==>>")
            if len(parts) == 2:
                key = parts[0].strip()
                value = parts[1].strip()
                if key:
                    replace_dict[key] = value
    return replace_dict

def _apply_replace(paragraph, replace_dict):
    """底层替换：遍历运行块并保留格式"""
    for old_text, new_text in replace_dict.items():
        if old_text in paragraph.text:
            runs_hit = [i for i, run in enumerate(paragraph.runs) if old_text in run.text]
            # #region agent log
            if not runs_hit:
                _agent_log(
                    hypothesis_id="H1",
                    location="app.py:_apply_replace",
                    message="old_text in paragraph but no run contains full old_text",
                    data={
                        "old_len": len(old_text),
                        "num_runs": len(paragraph.runs),
                        "para_text_len": len(paragraph.text),
                    },
                )
            # #endregion
            for run in paragraph.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)

def smart_replace(doc, replace_dict):
    """遍历文档所有部分（正文和表格）"""
    # 1. 替换正文
    for p in doc.paragraphs:
        _apply_replace(p, replace_dict)
    
    # 2. 替换表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _apply_replace(p, replace_dict)
    return doc

def main():
    st.title("📄 Word 格式无损替换工具")
    st.markdown("上传模板并输入 AI 生成的替换规则，直接下载处理后的文件。")

    # 1. 文件上传
    uploaded_file = st.file_uploader("第一步：上传你的 Word 模板 (.docx)", type="docx")

    # 2. 输入替换规则
    st.subheader("第二步：输入替换规则")
    example_text = "旧文字A ==>> 新文字A\n旧文字B ==>> 新文字B"
    replace_input = st.text_area(
        "按照格式输入（每行一对）：",
        placeholder=example_text,
        height=200
    )

    if uploaded_file and replace_input:
        if st.button("🚀 开始处理"):
            # 解析规则
            replacements = parse_replace_text(replace_input)
            # #region agent log
            _agent_log(
                hypothesis_id="H2",
                location="app.py:main",
                message="parsed replacement rules",
                data={"rule_count": len(replacements), "key_lens": [len(k) for k in replacements]},
            )
            # #endregion
            
            if not replacements:
                st.error("❌ 未能识别替换规则，请检查格式（旧内容 ==>> 新内容）")
                return

            try:
                # 读取上传的文件到内存
                doc = Document(uploaded_file)
                
                # 执行替换
                with st.spinner('正在处理中...'):
                    processed_doc = smart_replace(doc, replacements)
                
                # 将处理后的文档保存到内存流
                output_stream = io.BytesIO()
                processed_doc.save(output_stream)
                output_stream.seek(0)

                st.success(f"✅ 处理完成！识别到 {len(replacements)} 组替换。")

                # 3. 下载按钮
                st.download_button(
                    label="📥 点击下载新文件",
                    data=output_stream,
                    file_name=f"Processed_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                # #region agent log
                _agent_log(
                    hypothesis_id="H3",
                    location="app.py:main",
                    message="processing exception",
                    data={"exc_type": type(e).__name__, "exc_len": len(str(e))},
                )
                # #endregion
                st.error(f"❌ 处理出错: {str(e)}")

    # 页脚说明
    st.info("💡 提示：为了确保替换成功，请确保模板中的占位词在 Word 里是连续输入的。")

if __name__ == "__main__":
    main()