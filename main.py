from docx import Document
import re
from datetime import datetime

def read_docx(file_path):
    """读取Word文档，保留你喜欢的输出格式"""
    doc = Document(file_path)
    full_text = []

    print("===== 段落内容 =====")
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            print(f"[段落] {text}")
            full_text.append(text)

    print("\n===== 表格内容 =====")
    for table in doc.tables:
        for row in table.rows:
            row_cells = [cell.text.strip() for cell in row.cells]
            row_str = "\t".join(row_cells)
            print(f"[表格行] {row_str}")
            full_text.append(row_str)

    full_content = "\n".join(full_text)
    print("\n===== 全文汇总 =====")
    print(full_content)

    return full_content, doc

def parse_date(date_str):
    """智能解析日期字符串，支持多种格式，解析失败返回None"""
    if not date_str:
        return None
    # 支持的日期格式：2025-01-20、2025/01/20、2025.01.20、2025年1月20日
    formats = ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None

def extract_patient_info(full_text, doc):
    """
    智能提取患者信息，自动兼容所有3种医院表格格式
    增加日期逻辑校验，避免出院早于入院的错误
    """
    patient_info = {
        "姓名": None,
        "性别": None,
        "年龄": None,
        "住院号": None,
        "科室": None,
        "床号": None,
        "入院日期": None,
        "出院日期": None,
        "诊断": None,
        "医嘱": None
    }

    # 字段名映射表：支持所有常见变体
    field_mapping = {
        "姓名": ["姓名", "患者姓名", "病人姓名"],
        "性别": ["性别"],
        "年龄": ["年龄", "岁数"],
        "住院号": ["住院号", "病案号", "病历号", "住院编号"],
        "科室": ["科室", "入院科室", "所在科室"],
        "床号": ["床号", "床位号"],
        "入院日期": ["入院日期", "入院时间", "住院日期"],
        "出院日期": ["出院日期", "出院时间"]
    }

    # --------------------------
    # 第一步：智能解析所有表格（自动识别3种格式）
    # --------------------------
    for table in doc.tables:
        rows = table.rows
        if len(rows) == 0:
            continue

        # 格式1：两行式（表头在上，值在下）- 最常见
        if len(rows) == 2:
            headers = [cell.text.strip() for cell in rows[0].cells]
            values = [cell.text.strip() for cell in rows[1].cells]
            for i, header in enumerate(headers):
                if i >= len(values):
                    continue
                clean_header = header.replace(" ", "").replace("　", "").rstrip("：:")
                # 匹配字段名变体
                for standard_field, variants in field_mapping.items():
                    if any(variant in clean_header for variant in variants):
                        patient_info[standard_field] = values[i].strip()
                        break

        # 格式2：单列键值对式（两列，项目-内容）
        elif len(rows) >= 3 and len(rows[0].cells) == 2:
            for row in rows:
                cells = [cell.text.strip() for cell in row.cells]
                if len(cells) < 2:
                    continue
                key_cell = cells[0].replace(" ", "").replace("　", "").rstrip("：:")
                value_cell = cells[1].strip()
                # 匹配字段名变体
                for standard_field, variants in field_mapping.items():
                    if any(variant in key_cell for variant in variants):
                        patient_info[standard_field] = value_cell
                        break

        # 格式3：左右键值对式（每行2个字段）
        else:
            for row in rows:
                cells = [cell.text.strip() for cell in row.cells]
                # 每两个单元格为一组（字段-值）
                for i in range(0, len(cells), 2):
                    if i + 1 >= len(cells):
                        break
                    key_cell = cells[i].replace(" ", "").replace("　", "").rstrip("：:")
                    value_cell = cells[i+1].strip()
                    if not key_cell or not value_cell:
                        continue
                    # 匹配字段名变体
                    for standard_field, variants in field_mapping.items():
                        if any(variant in key_cell for variant in variants):
                            patient_info[standard_field] = value_cell
                            break

    # --------------------------
    # 第二步：从段落用正则补全
    # --------------------------
    patterns = {
        "诊断": r"(?:入院|出院|主要)?诊断[：:]\s*(.+?)(?=\n(?:医嘱|出院日期|患者|性别|年龄|住院号|入院日期)[：:]|\Z)",
        "医嘱": r"医嘱[：:]\s*(.+?)(?=\n(?:出院日期|患者|性别|年龄|住院号|入院日期|诊断)[：:]|\Z)",
        "出院日期": r"出院日期[：:]\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}年\d{1,2}月\d{1,2}日)",
        "入院日期": r"入院日期[：:]\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}年\d{1,2}月\d{1,2}日)"
    }

    for field, pattern in patterns.items():
        if patient_info[field] is not None:
            continue
        match = re.search(pattern, full_text, re.DOTALL)
        if match:
            patient_info[field] = match.group(1).strip()

    # --------------------------
    # 第三步：数据清洗 + 逻辑校验（新增！）
    # --------------------------
    # 1. 年龄清洗
    if patient_info["年龄"]:
        age_match = re.search(r"\d+", patient_info["年龄"])
        if age_match:
            patient_info["年龄"] = int(age_match.group())

    # 2. 日期逻辑校验（核心修复！）
    admission_date = parse_date(patient_info["入院日期"])
    discharge_date = parse_date(patient_info["出院日期"])

    # 如果出院日期早于入院日期，判定为无效，置为None
    if admission_date and discharge_date and discharge_date < admission_date:
        print(f"\n⚠️  警告：检测到出院日期({patient_info['出院日期']})早于入院日期({patient_info['入院日期']})，已自动置为未提取到")
        patient_info["出院日期"] = None

    # 3. 未提取到的字段显示为"未提取到"，更友好
    for key in patient_info:
        if patient_info[key] is None:
            patient_info[key] = "未提取到"

    return patient_info

if __name__ == "__main__":
    file_path = "samples/patient_01.docx"
    try:
        full_content, doc = read_docx(file_path)
        print("\n" + "="*50)

        print("\n===== 提取的患者信息 =====")
        patient_info = extract_patient_info(full_content, doc)
        for key, value in patient_info.items():
            print(f"{key}: {value}")

    except Exception as e:
        print(f"处理失败：{e}")