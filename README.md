
# Data-Validator

Data-Validator is a Python-based tool that uses Streamlit for validating data in CSV or Excel files. It provides various validators, including custom regex patterns, to ensure data integrity and correctness.

## Features

- Supports CSV and Excel file formats (.csv / .xls / .xlsx).
- **预设方案校验**：上传文件后按表头自动匹配方案（也可手动选择），一键完成多列校验。
- 内置校验规则：必填检测、姓名长度（≤15字）、车牌号、11位手机号、日期时间解析。
- 单个下载文件同时完成：无效单元格标红、无效手机号清空、时间归一化为 yyyy-mm-dd、车牌重复行删除。
- 保留手动模式：逐列选择校验器（Email / EID / Mobile No / License Plate / Date / DateTime / Custom）。

## Installation

To install the required dependencies, run:
```bash
pip install -r requirements.txt
```

## Usage

To run the Data-Validator application, execute the following command:
```bash
streamlit run final.py
```

Windows 下可直接双击 `run_validator.bat`（final.py 数据校验）或 `run_profiler.bat`（app.py 数据画像）。

To run the data profiling app (`app.py`), execute:
```bash
streamlit run app.py
```

## 预设方案（月租车导入数据校验）

上传模板格式的 Excel 文件后，系统按表头自动匹配方案「月租车导入数据校验」，点击「开始校验」并下载校验后文件：

| 规则 | 处理方式 |
|---|---|
| 必填项（项目名称/车辆类型/车主姓名/车位数/一位多车/车牌号/开始时间/结束时间）为空 | 标红，摘要中列出（如 B4、H11）；**车牌为空时删除整行** |
| 车主姓名超过 15 字 | 先复制完整姓名到「车辆备注」列，再截断至 15 字（删除后面部分） |
| 车牌号含字母 O/I | 自动替换为数字 0/1（修复，不报问题） |
| 车牌号省份简称不对或位数不对（普通 7 位 / 新能源 8 位，大小写均可） | 单元格标红，不删除行 |
| 手机号不是 11 位（1[3-9] 开头） | 清空该单元格 |
| 开始/结束时间（支持 yyyy-mm-dd、yy/mm/dd hh:mm:ss 等多种格式） | 统一归一化为 yyyy-mm-dd；无法解析的标红 |
| 车牌号重复 | 删除重复行（保留第一行，空车牌不参与） |

校验摘要以原文件单元格定位显示问题，例如：

```
无效手机号：K5、K8
姓名超15字：D6
车牌无效：第6行
必填项缺失：B4、H11
时间无法解析：I9
```

### 新增/修改方案

方案配置在 `schemes.json` 中，一个方案包含：方案名、期望表头（用于自动匹配）、每列的校验器与失败动作（`mark` 标红 / `clear` 清空）、可选的时间转换（`date_ymd`）、车牌去重配置。新增方案只需在该文件追加一项，无需改代码。

## How It Works

1. **上传文件**：上传 CSV 或 Excel 文件。
2. **自动匹配方案**：系统扫描表头，自动选中匹配的预设方案（可手动切换）。
3. **开始校验**：点击「开始校验」，查看校验摘要（以单元格定位列出各类问题，如 无效手机号：K5、K8）。
4. **下载结果**：下载「修改后的文件」——在**原文件上直接修改**（保留第 1 行说明、表头、列宽等全部格式）：无效车牌/重复车牌整行删除、姓名超15字先复制到备注再截断、无效手机号清空、时间归一化、其余无效单元格标红。

## License

This project is licensed under the MIT License.

## Contributing

Feel free to fork this repository and contribute by submitting pull requests.


