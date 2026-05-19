from docx import Document

def read_docx(file_path):
    # 打开文档
    doc = Document(file_path)
    full_text = []

    # 1. 读取所有段落
    print("===== 段落内容 =====")
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            print(f"[段落] {text}")
            full_text.append(text)

    # 2. 读取所有表格
    print("\n===== 表格内容 =====")
    for table in doc.tables:
        for row in table.rows:
            row_cells = [cell.text.strip() for cell in row.cells]
            row_str = "\t".join(row_cells)
            print(f"[表格行] {row_str}")
            full_text.append(row_str)

    # 合并成全文
    return "\n".join(full_text)

if __name__ == "__main__":
    # 你的病历文件路径
    file_path = "samples/patient_01.docx"
    try:
        full_content = read_docx(file_path)
        print("\n===== 全文汇总 =====")
        print(full_content)
    except Exception as e:
        print(f"读取失败：{e}")