
# Data-Validator

Data-Validator is a Python-based tool that uses Streamlit for validating data in CSV or Excel files. It provides various validators, including custom regex patterns, to ensure data integrity and correctness.

## Features

- Supports CSV and Excel file formats (.csv / .xls / .xlsx).
- **预设方案校验**：上传文件后按表头自动匹配方案（也可手动选择），一键完成多列校验。
- 内置校验规则：必填检测、姓名长度（≤15字）、车牌号、11位手机号、日期时间解析。
- 单个下载文件同时完成：无效单元格标红、无效手机号清空、时间归一化为 yyyy-mm-dd、车牌重复行删除。
- 有标红的行（需人工核对修改）自动集中到文件末尾，正常行保持原序在前，一眼定位问题数据。
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

> **校验密码（可选）**：配置密码后，点击「开始校验」需先通过密码验证才执行；未配置密码时行为不变。
>
> **本地**：在 `.streamlit/secrets.toml`（已被 .gitignore 排除，不会推送到仓库）配置：
> ```toml
> VERIFY_PASSWORD = "你的密码"
> ```
>
> **公网部署启用密码（交给部署方）**：
> 1. 更新服务器代码到最新（含 `secure_password/` 组件目录，GitHub `main` 分支）；
> 2. 在服务器的启动环境（或守护进程/启动脚本中）设置环境变量 `VERIFY_PASSWORD=你的密码`；
>    （也可将上面的 `secrets.toml` 放到服务器项目目录 `.streamlit/` 下，二者等价，环境变量优先级更高）
> 3. 按热更新流程重启服务：杀掉旧 python 进程 → 守护进程自动拉起（或手动启动）；
> 4. 访问公网地址验证：点「开始校验」应弹出密码框，输入正确密码才能执行。

| 规则 | 处理方式 |
|---|---|
| 必填项（项目名称/车辆类型/车主姓名/车位数/一位多车/车牌号/开始时间/结束时间）为空 | 标红，摘要中列出（如 B4、H11）；**车牌为空时删除整行**；**车主姓名为空时用同行车牌号填充** |
| 车主姓名超过 15 字 | 先复制完整姓名到「车辆备注」列，再截断至 15 字（删除后面部分），同时标红 |
| 门牌号 / 车位号超过 20 字 | 先复制完整原值到「车辆备注」列，再截断至 20 字（模板字段限 20 字） |
| 同名不同期 | 仅"一位多车=是"的行参与，按"姓名+手机号"（无手机号仅姓名）分组，组内结束时间不一致时，第二个起的行姓名自动加数字递增（王玉梅→王玉梅1→王玉梅2），避免系统把不同租期误判为一位多车；"一位多车=否"的行不处理；加序号后超 15 字则截断原名保数字 |
| 车牌号含字母 O/I | 自动替换为数字 0/1，**摘要中列出替换位置**（如 H3（鄂AO1234→鄂A01234）） |
| 车牌号省份简称不对或位数不对（普通 7 位 / 新能源 8 位，大小写均可） | 单元格标红，不删除行 |
| 手机号不是 11 位（1[3-9] 开头） | 清空该单元格，原值复制到「车辆备注」列 |
| 开始/结束时间（支持 yyyy-mm-dd、yy/mm/dd hh:mm:ss 等多种格式） | 统一归一化为 yyyy-mm-dd；无法解析的标红 |
| 一个单元格含多个车牌（逗号/顿号/分号/空格分隔） | 拆分为独立车牌逐个校验；下载文件中规范为英文逗号分隔 |
| 重复车牌（多行出现同一车牌，含 O/I 替换后相同） | 按「车位数 = 免费名额」容量优先分配：重复车牌优先保留在还有剩余名额的行；所有出现过的行都满时保留首次出现的行（摘要列出剔除位置与保留位置，如 `第3行（京B22222→保留至第4行）`）；行内重复保留一个 |
| 行车牌数超过车位数 | 仅当该行"一位多车"≠"是"时标红提示"车位不足"（模板允许"一位多车=是 + 车位数1 + 多车牌"，不误报；车位数缺失时只报必填缺失） |
| 分配后整格变空的车牌格 | 按"车牌为空"删除整行 |

校验摘要以原文件单元格定位显示问题，例如：

```
车牌O/I已替换：H3（鄂AO1234→鄂A01234）
无效手机号：K5、K8
姓名缺失已用车牌填充：D6
姓名超15字：D7
车牌无效：第8行
车牌重复：第9行（京B22222→保留至第10行）
车位不足：第11行
车牌为空：第10行
必填项缺失：B4
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


