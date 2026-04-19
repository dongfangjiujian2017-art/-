# --- 核心逻辑：精准 XML 替换（改造版） ---

def _xml_text_replace(element, replace_dict):
    """
    在 XML 元素中深度查找所有 w:t 节点并执行替换。
    """
    if element is None:
        return
    for old_text, new_text in replace_dict.items():
        if not old_text: continue
        for t_node in element.xpath('.//w:t'):
            if t_node.text and old_text in t_node.text:
                t_node.text = t_node.text.replace(old_text, new_text)

def backup_and_process_parts(doc, replace_dict):
    """
    备份页眉页脚。如果原本就没有内容（空段落且无文本框），则标记为无效。
    对原有内容执行替换，空页脚不会被破坏。
    """
    backups = []
    for i, section in enumerate(doc.sections):
        # 检查是否有内容存在（避免强制生成空白页码）
        has_header = any(p.text.strip() for p in section.header.paragraphs) \
                     or len(section.header.tables) > 0 \
                     or section.header._element.xpath('.//w:txbxContent')
        has_footer = any(p.text.strip() for p in section.footer.paragraphs) \
                     or len(section.footer.tables) > 0 \
                     or section.footer._element.xpath('.//w:txbxContent')

        h_xml = copy.deepcopy(section.header._element) if has_header else None
        f_xml = copy.deepcopy(section.footer._element) if has_footer else None
        
        # 在有效的副本中执行替换
        if h_xml is not None:
            _xml_text_replace(h_xml, replace_dict)
        if f_xml is not None:
            _xml_text_replace(f_xml, replace_dict)
        
        backups.append({
            "index": i,
            "header_xml": h_xml,
            "footer_xml": f_xml,
            "header_distance": section.header_distance,
            "footer_distance": section.footer_distance,
            "has_h": has_header,
            "has_f": has_footer
        })
    return backups

def restore_parts(doc, backups):
    """
    还原页眉页脚：
    1. 原本有内容的页眉页脚恢复修改后的 XML；
    2. 原本没有内容的页脚添加一个空段落，避免 Word 自动生成默认页码域。
    """
    for backup in backups:
        section = doc.sections[backup["index"]]
        
        # 还原页眉
        if backup["has_h"] and backup["header_xml"] is not None:
            section.header._element.clear()
            for child in backup["header_xml"]:
                section.header._element.append(child)
        # 还原页脚
        if backup["has_f"] and backup["footer_xml"] is not None:
            section.footer._element.clear()
            for child in backup["footer_xml"]:
                section.footer._element.append(child)
        elif not backup["has_f"]:
            # 原本空页脚，添加一个空段落防止 Word 自动生成页码域
            section.footer._element.clear()
            section.footer.add_paragraph(" ")  # 占位空格
        
        # 恢复页眉页脚距离
        section.header_distance = backup["header_distance"]
        section.footer_distance = backup["footer_distance"]
