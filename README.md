# README\.md

# 电子病历信息提取器 \(EMR Extractor\)

一个简单易用的 Python 工具，用于从 Word 格式的电子病历中自动提取结构化患者信息，支持批量处理和 Excel 导出。

## ✨ 功能特点

- 读取 Word 文档（\.docx）的段落和表格内容

- 支持医院最常用的**两行式患者基本信息表**

- 自动识别常见字段名变体（如 \&\#34;住院号 / 病案号 / 病历号\&\#34;）

- 智能日期解析与逻辑校验，自动过滤 \&\#34;出院早于入院\&\#34; 的无效数据

- 支持**单个文件处理**和**批量目录处理**

- 提取结果自动导出为 Excel 表格

- 代码结构清晰，易于扩展和修改

## 📋 环境要求

- Python 3\.8 及以上版本

- 依赖包：python\-docx、pandas、openpyxl

## 🚀 快速开始

### 1\. 克隆仓库

```bash
git clone https://github.com/rocoT-T/emr-extractor.git
cd emr-extractor
```

### 2\. 创建并激活虚拟环境

```bash
# Windows
python -m venv venv
.\venv\Scripts\Activate.ps1

# Mac/Linux
python3 -m venv venv
source venv/bin/activate
```

### 3\. 安装依赖

```bash
pip install -r requirements.txt
```

## 📖 使用方法

### 方式一：处理单个文件

```bash
python main.py
# 输入 1 并回车
```

程序会自动处理 `samples/patient\_01\.docx`，并在控制台输出详细的读取过程和提取结果。

### 方式二：批量处理并导出 Excel

1. 将所有需要处理的 Word 病历文件放入 `samples/` 目录

2. 运行程序：

    ```bash
    python main.py
    # 输入 2 并回车
    ```

3. 处理完成后，结果会自动导出到项目根目录的 `output\.xlsx` 文件中

## 📁 项目结构

```Plain Text
emr-extractor/
├── samples/                # 存放待处理的病历文件
│   └── patient_01.docx     # 示例病历文件
├── venv/                   # Python 虚拟环境（已被 .gitignore 忽略）
├── .gitignore              # Git 忽略文件配置
├── main.py                 # 主程序代码
├── requirements.txt        # 依赖包列表
└── README.md               # 项目说明文档
```

## 📊 提取字段说明

|字段名|说明|
|---|---|
|姓名|患者姓名|
|性别|患者性别|
|年龄|患者年龄（自动转为数字）|
|住院号|住院编号 / 病案号 / 病历号|
|科室|入院科室|
|床号|床位号|
|入院日期|入院时间|
|出院日期|出院时间|
|诊断|入院 / 出院 / 主要诊断|
|医嘱|治疗医嘱|

## 📝 更新日志

- `v1\.2\.0` \- 重构代码结构，分离关注点，新增批量处理和 Excel 导出功能

- `v1\.1\.0` \- 增加日期逻辑校验，修复出院日期早于入院日期的问题

- `v1\.0\.0` \- 实现基础功能：读取 Word 文档，提取患者基本信息

## ⚠️ 注意事项

1. **隐私保护**：本工具仅用于学习和研究，请勿上传包含真实患者隐私信息的病历文件到公共仓库

2. **格式支持**：目前仅支持 `\.docx` 格式，不支持旧版 `\.doc` 格式

3. **表格格式**：默认支持最常见的两行式表格（表头在上，值在下），其他格式可自行扩展

## 📄 许可证

MIT License

---

## 📌 下一步计划

- 支持更多表格格式

- 增加 PDF 格式支持

- 优化多行诊断和医嘱的提取

- 增加简单的 Web 界面

需要我帮你把这个 README 直接生成文件，或者调整一下内容侧重点吗？

> （注：文档部分内容可能由 AI 生成）
