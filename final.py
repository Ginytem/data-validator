import pandas as pd
import streamlit as st
import re
import io
import json
import os
import datetime
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill


# ============ 内置校验规则 ============

# 全国车牌省份简称（31 个省级行政区）+ 车牌正则：
# 首字为省份简称；第二位为发牌机关代号（A-Z，不含 I/O）；
# 后部 5 位为普通车牌、6 位为新能源车牌（新能源第三位不限于 D/F，
# 各地使用 A/B/C/E/G/H/J/K 等扩展字母），字母均不含 I/O；
# 大小写均有效（校验前统一转大写）。
PLATE_PROVINCES = '京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼'
PLATE_RE = re.compile(r'^[%s][A-HJ-NP-Z][A-HJ-NP-Z0-9]{5,6}$' % PLATE_PROVINCES)

# 11 位手机号（1[3-9] 开头）
MOBILE_RE = re.compile(r'^1[3-9]\d{9}$')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCHEMES_PATH = os.path.join(BASE_DIR, 'schemes.json')


# Normalize cell value to a clean string before regex matching:
# - Excel numeric cells become floats (13812345678.0) -> strip the ".0"
# - strip surrounding whitespace
def _norm(value):
    if pd.isna(value):
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


# ============ 预设方案配置 ============

def load_schemes(path=SCHEMES_PATH):
    """读取 schemes.json 中的预设方案列表。"""
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f).get('schemes', [])
    except Exception:
        return []


def detect_scheme(df_raw, schemes, max_scan=8):
    """按表头自动匹配方案。

    返回 (scheme, header_row_index, hits) 或 None。
    header_row_index 为 0 起始的表头所在行。
    """
    best = None
    for scheme in schemes:
        expected = [str(h).strip() for h in scheme.get('expected_headers', [])]
        if not expected:
            continue
        for r in range(min(len(df_raw), max_scan)):
            row_vals = [str(v).strip() for v in df_raw.iloc[r].tolist()]
            hits = sum(1 for h in expected if h in row_vals)
            if hits >= max(3, int(len(expected) * 0.8)) and (best is None or hits > best[2]):
                best = (scheme, r, hits)
    return best


def find_best_header_row(df_raw, expected, max_scan=10):
    """找不到强匹配时，找表头相似度最高的行。返回 (row_index, hits)。"""
    best_idx, best_hits = None, 0
    for r in range(min(len(df_raw), max_scan)):
        row_vals = [str(v).strip() for v in df_raw.iloc[r].tolist()]
        hits = sum(1 for h in expected if h in row_vals)
        if hits > best_hits:
            best_idx, best_hits = r, hits
    return best_idx, best_hits


def df_with_header(df_raw, header_idx):
    """以 header_idx 行为表头，其下为数据行。"""
    df = df_raw.iloc[header_idx + 1:].copy()
    headers = []
    for i, v in enumerate(df_raw.iloc[header_idx].tolist()):
        headers.append(str(v).strip() if not pd.isna(v) else f'Unnamed:{i}')
    df.columns = headers
    return df.reset_index(drop=True)


# ============ 校验器（按列） ============

def _col_letter(idx1):
    """1 起始列号 -> Excel 列字母（A、B、…、Z、AA…）。"""
    s = ''
    while idx1 > 0:
        idx1, r = divmod(idx1 - 1, 26)
        s = chr(65 + r) + s
    return s


def _check_rule(rule, series):
    """单规则校验，返回 Series（True=有效）。

    非必填语义：plate / mobile / datetime 规则对空值放行，
    空值只交给 required 规则判定，避免同一格重复报两种问题。
    """
    if rule == 'required':
        return series.map(lambda x: _norm(x) != '')
    if rule == 'maxlen15':
        return series.map(lambda x: len(_norm(x)) <= 15)
    if rule == 'plate':
        return series.map(lambda x: _norm(x) == '' or bool(PLATE_RE.match(_norm(x).upper())))
    if rule == 'mobile':
        return series.map(lambda x: _norm(x) == '' or bool(MOBILE_RE.match(_norm(x))))
    if rule == 'datetime':
        empty = series.map(lambda x: _norm(x) == '')
        return empty | _parse_dates(series).notna()
    return pd.Series(True, index=series.index)


# 日期解析：显式格式列表，年月日优先（用户场景为 yy/mm/dd hh:mm:ss 或其它格式，
# pandas 通用解析对两位数年份斜杠格式按日/月/年处理，会解析错，故逐格式尝试）
DATE_FORMATS = [
    '%Y-%m-%d %H:%M:%S',
    '%Y-%m-%d %H:%M',
    '%Y-%m-%d',
    '%Y/%m/%d %H:%M:%S',
    '%Y/%m/%d %H:%M',
    '%Y/%m/%d',
    '%y/%m/%d %H:%M:%S',
    '%y/%m/%d %H:%M',
    '%y/%m/%d',
    '%y-%m-%d %H:%M:%S',
    '%y-%m-%d %H:%M',
    '%y-%m-%d',
    '%Y%m%d',
    '%Y.%m.%d',
]


def _parse_dates(series):
    """按显式格式（年月日优先）解析日期列，解析失败返回 NaT。"""
    parsed = pd.Series(pd.NaT, index=series.index, dtype='datetime64[ns]')
    s = series.astype(object)

    # 1) 已经是日期/时间对象的直接取用
    mask_dt = s.map(lambda x: isinstance(x, (pd.Timestamp, datetime.datetime, np.datetime64)))
    if mask_dt.any():
        parsed.loc[mask_dt] = pd.to_datetime(s[mask_dt], errors='coerce')

    # 2) 字符串按显式格式逐个尝试
    rest = s[~mask_dt].map(lambda x: _norm(x))
    rest = rest[rest.ne('')]
    for fmt in DATE_FORMATS:
        if rest.empty:
            break
        p = pd.to_datetime(rest, errors='coerce', format=fmt)
        ok = p.notna()
        if ok.any():
            parsed.loc[ok.index] = p[ok]
            rest = rest[~ok]

    # 3) 兜底：通用解析（处理未匹配到的其它格式）
    if not rest.empty:
        p = pd.to_datetime(rest, errors='coerce')
        ok = p.notna()
        if ok.any():
            parsed.loc[ok.index] = p[ok]
    return parsed


def run_scheme(df, scheme, header_idx=0):
    """按预设方案执行校验。

    流程：先按车牌去重（删除重复行，保留第一行，空车牌不参与）→
    逐列逐规则校验并按失败动作处理 → 时间列归一化为 yyyy-mm-dd →
    无效车牌整行删除。

    失败动作：
    - mark       ：无效单元格标红（必填缺失 / 时间无法解析）
    - clear      ：清空该单元格（无效手机号）
    - truncate15 ：截断至 15 字（姓名超长）
    - delete_row ：删除整行（无效车牌）

    返回 (result, red_cells, cell_changes, delete_rows, summary)：
    - result: 处理后的 DataFrame（用于界面预览）
    - red_cells: {列名: {原文件行号集合}}，标红单元格位置
    - cell_changes: {(原文件行号, 列名): 新值}，清空/截断/时间归一化写回值
    - delete_rows: 需要整行删除的原文件行号列表
    - summary: {'issues': [{'label','cells'}, ...]}
      issues 里 cells 为原文件中的单元格定位（如 K5）或行号（第5行）
    """
    out = df.copy()
    columns_cfg = scheme.get('columns', {})
    dedup_cfg = scheme.get('dedup') or {}
    first_data_row = header_idx + 2  # 原文件中第一行数据所在的 Excel 行号（1 起始）

    # 记录每一行在原文件中的 Excel 行号，用于摘要定位与修改写回
    out['__orig_row'] = list(range(first_data_row, first_data_row + len(out)))

    issue_items = []  # (label, orig_row, col_letter|None)
    delete_rows = set()
    red_cells = {}
    cell_changes = {}
    cleared = 0
    truncated = 0

    # 1) 车牌去重：空车牌不参与，保留第一行
    dup_col = dedup_cfg.get('column')
    if dup_col and dup_col in out.columns:
        normed = out[dup_col].map(lambda x: _norm(x).upper())
        empty = normed.eq('')
        keep = empty | (~normed.duplicated(keep='first'))
        if (~keep).any():
            for r in out.loc[~keep, '__orig_row'].tolist():
                issue_items.append(('车牌重复', r, None))
                delete_rows.add(r)
            out = out[keep].copy().reset_index(drop=True)

    col_letters = {name: _col_letter(i + 1) for i, name in enumerate(out.columns)}

    # 2) 逐列逐规则校验 + 修复动作
    for col, conf in columns_cfg.items():
        if col not in out.columns:
            continue
        letter = col_letters[col]
        for vconf in conf.get('validators', []):
            rule = vconf.get('rule')
            action = vconf.get('on_invalid', 'mark')
            invalid = ~_check_rule(rule, out[col])
            if not invalid.any():
                continue
            orig_rows = out.loc[invalid, '__orig_row'].tolist()
            if action == 'delete_row':
                delete_rows.update(orig_rows)
                for r in orig_rows:
                    issue_items.append(('车牌无效', r, None))
            elif action == 'clear':
                out.loc[invalid, col] = None
                cleared += len(orig_rows)
                for r in orig_rows:
                    cell_changes[(r, col)] = None
                    issue_items.append(('无效手机号', r, letter))
            elif action == 'truncate15':
                originals = out.loc[invalid, col].map(lambda x: _norm(x))
                new_vals = originals.map(lambda x: x[:15])
                out.loc[invalid, col] = new_vals
                truncated += len(orig_rows)
                for r, v in zip(orig_rows, new_vals.tolist()):
                    cell_changes[(r, col)] = v
                    issue_items.append(('姓名超15字', r, letter))
                # 截断前先把完整姓名复制到备注列，避免丢失重要信息
                copy_to = vconf.get('copy_to')
                if copy_to and copy_to in out.columns:
                    for label, full in zip(out.index[invalid].tolist(), originals.tolist()):
                        r = out.at[label, '__orig_row']
                        prev = _norm(out.at[label, copy_to])
                        new_remark = f'{prev}；原姓名：{full}' if prev else f'原姓名：{full}'
                        out.at[label, copy_to] = new_remark
                        cell_changes[(r, copy_to)] = new_remark
            else:  # mark
                red_cells.setdefault(col, set()).update(orig_rows)
                label = '必填项缺失' if rule == 'required' else ('时间无法解析' if rule == 'datetime' else '无效数据')
                for r in orig_rows:
                    issue_items.append((label, r, letter))

    # 3) 时间归一化：可解析的时间统一为 yyyy-mm-dd
    for col, conf in columns_cfg.items():
        if col not in out.columns or conf.get('transform') != 'date_ymd':
            continue
        parsed = _parse_dates(out[col])
        ok = parsed.notna()
        if ok.any():
            new_vals = parsed[ok].dt.strftime('%Y-%m-%d')
            out[col] = out[col].astype(object)
            out.loc[ok, col] = new_vals
            for r, v in zip(out.loc[ok, '__orig_row'].tolist(), new_vals.tolist()):
                cell_changes[(r, col)] = v

    # 4) 删除整行（无效车牌 / 重复车牌）
    if delete_rows:
        out = out[~out['__orig_row'].isin(delete_rows)].copy().reset_index(drop=True)

    # 5) 汇总：被删除行上的其它问题不再列出
    grouped = {}
    for label, r, letter in issue_items:
        if label not in ('车牌无效', '车牌重复') and r in delete_rows:
            continue
        grouped.setdefault(label, []).append((r, letter))
    issues = []
    for label in ['无效手机号', '姓名超15字', '车牌无效', '车牌重复', '必填项缺失', '时间无法解析']:
        if label not in grouped:
            continue
        rows = sorted(grouped[label])
        if label in ('车牌无效', '车牌重复'):
            cells = [f'第{r}行' for r, _ in rows]
        else:
            cells = [f'{letter}{r}' for r, letter in rows]
        issues.append({'label': label, 'cells': cells})

    result = out.drop(columns=['__orig_row'])
    summary = {
        'issues': issues,
        'cleared': cleared,
        'truncated': truncated,
    }
    return result, red_cells, cell_changes, sorted(delete_rows), summary


# ============ 导出（在原文件上直接修改，保留模板格式） ============

def export_modified_file(data_bytes, sheet_name, col_pos, red_cells, cell_changes, delete_rows):
    """在用户上传的原文件上直接修改：
    - 保留第 1 行说明行、表头、列宽、合并单元格等全部原格式
    - 无效单元格标红（FFC7CE）
    - 无效手机号清空、姓名截断、时间归一化写回原单元格
    - 无效车牌 / 重复车牌所在行整行删除
    """
    wb = openpyxl.load_workbook(io.BytesIO(data_bytes))
    ws = wb[sheet_name]
    red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

    # 1) 写回修复值（清空 / 截断 / 时间归一化）
    for (orig_row, col), val in cell_changes.items():
        if col in col_pos and orig_row not in delete_rows:
            ws.cell(row=orig_row, column=col_pos[col]).value = val

    # 2) 标红无效单元格
    for col, rows in red_cells.items():
        if col not in col_pos:
            continue
        for r in rows:
            if r not in delete_rows:
                ws.cell(row=r, column=col_pos[col]).fill = red

    # 3) 整行删除（从下往上删，避免行号偏移）
    for r in sorted(delete_rows, reverse=True):
        ws.delete_rows(r)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def export_result_file(out, red_cells, sheet_name='Sheet1'):
    """兜底导出（非 xlsx 原文件时使用）：重建表格并标红。"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        out.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.book[sheet_name]
        red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        col_pos = {name: i + 1 for i, name in enumerate(out.columns)}
        for col, labels in red_cells.items():
            if col not in col_pos:
                continue
            for label in labels:
                ws.cell(row=int(label) + 2, column=col_pos[col]).fill = red
    buf.seek(0)
    return buf


# ============ 手动模式（保留原有逐列配置功能） ============

def validate_data(df, validators):
    df = df.copy()

    for column, validator in validators.items():
        if validator == 'Custom':
            custom_pattern = st.session_state.get(f'custom_validator_{column}', '')
            if custom_pattern:
                df[f'Valid_{column}'] = df[column].apply(
                    lambda x: bool(re.match(custom_pattern, _norm(x))))
            else:
                df[f'Valid_{column}'] = True
        elif validator == 'Email':
            df[f'Valid_{column}'] = df[column].apply(
                lambda x: bool(re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', _norm(x))))
        elif validator == 'EID':
            df[f'Valid_{column}'] = df[column].apply(
                lambda x: bool(re.match(r'^[0-9]+-[0-9]+-[0-9]+-\d$', _norm(x))))
        elif validator == 'Mobile No':
            df[f'Valid_{column}'] = df[column].apply(
                lambda x: bool(re.match(r'^1[3-9]\d{9}$', _norm(x))))
        elif validator == 'Plate':
            df[f'Valid_{column}'] = df[column].apply(
                lambda x: bool(PLATE_RE.match(_norm(x).upper())))
        elif validator == 'Date':
            df[f'Valid_{column}'] = pd.to_datetime(df[column], errors='coerce').notnull()
        elif validator == 'DateTime':
            df[f'Valid_{column}'] = pd.to_datetime(df[column], errors='coerce').notnull()
        elif validator is None:
            df[f'Valid_{column}'] = True

    return df


def export_marked_file(df_original, df_validated, selected_columns):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_original.to_excel(writer, sheet_name='校验结果', index=False)
        ws = writer.book['校验结果']
        red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        col_index = {name: i + 1 for i, name in enumerate(df_original.columns)}
        for col in selected_columns:
            invalid = ~df_validated[f'Valid_{col}'].astype(bool)
            for row_pos, flag in enumerate(invalid):
                if flag:
                    ws.cell(row=row_pos + 2, column=col_index[col]).fill = red
    buf.seek(0)
    return buf


VALIDATOR_OPTIONS = {
    'None': None,
    'Email': 'Email',
    'EID': 'EID',
    'Mobile No (11位手机号)': 'Mobile No',
    'License Plate (车牌号)': 'Plate',
    'Date': 'Date',
    'DateTime': 'DateTime',
    'Custom (自定义)': 'Custom',
}


def manual_mode(df, uploaded_file):
    """原有手动逐列校验界面。"""
    st.write('原始数据预览：')
    st.write(df.head(20))

    selected_columns = st.multiselect('选择要校验的列', df.columns)

    custom_validators = {}
    for column in selected_columns:
        custom_pattern = st.text_input(f'为列 "{column}" 输入自定义正则表达式')
        if custom_pattern:
            custom_validators[column] = custom_pattern

    selected_validators = {}
    for column in selected_columns:
        validator = st.selectbox(f'为列 "{column}" 选择校验器', list(VALIDATOR_OPTIONS.keys()))
        selected_validators[column] = VALIDATOR_OPTIONS[validator]
        if validator == 'Custom':
            st.session_state[f'custom_validator_{column}'] = custom_validators.get(column, '')

    if st.button('运行校验（手动模式）'):
        df_validated = validate_data(df[selected_columns], selected_validators)

        validation_results = pd.DataFrame()
        for column, validator in selected_validators.items():
            validation_results.loc['Validation Score', column] = df_validated[f'Valid_{column}'].mean() * 100
            validation_results.loc['Completeness Score', column] = df_validated[column].notnull().mean() * 100

        st.write('校验汇总：')
        st.write(validation_results)

        valid_columns = [col for col in df_validated.columns if col.startswith('Valid_')]
        non_valid_rows = df_validated[~df_validated[valid_columns].all(axis=1)] if valid_columns else df_validated.iloc[0:0]

        if not non_valid_rows.empty:
            st.write('不符合规则的记录：')
            st.write(non_valid_rows)
        else:
            st.write('未发现不符合规则的记录。')

        if selected_validators:
            marked_file = export_marked_file(df, df_validated, list(selected_validators.keys()))
            st.download_button(
                '下载标记后的 Excel（无效单元格标红）',
                marked_file,
                uploaded_file.name.rsplit('.', 1)[0] + '_标记.xlsx',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )


# ============ 方案模式主界面 ============

def main():
    st.set_page_config(page_title='Data Validator')
    st.markdown("""
<style>
[data-testid="stDownloadButton"] button {
    background-color: #52C41A;
    border-color: #52C41A;
    color: #ffffff;
}
[data-testid="stDownloadButton"] button:hover {
    background-color: #3FA015;
    border-color: #3FA015;
    color: #ffffff;
}
</style>
""", unsafe_allow_html=True)
    st.title('Data Validator')

    schemes = load_schemes()
    scheme_names = [s.get('name', f'方案{i + 1}') for i, s in enumerate(schemes)]
    manual_label = '手动配置（不匹配方案）'
    all_options = scheme_names + [manual_label]

    uploaded_file = st.file_uploader('上传 CSV 或 Excel 文件', type=['csv', 'xls', 'xlsx'])

    if uploaded_file is None:
        st.info('请上传文件后开始校验。')
        return

    # 换文件时清除上一次的校验结果（按 文件名+大小 判断是否新文件）
    prev_key = st.session_state.get('scheme_uploaded_key')
    cur_key = (uploaded_file.name, uploaded_file.size)
    if prev_key is not None and prev_key != cur_key:
        for k in ['scheme_out', 'scheme_red', 'scheme_changes', 'scheme_delete',
                  'scheme_sheet', 'scheme_col_pos', 'scheme_raw', 'scheme_header', 'scheme_summary']:
            st.session_state.pop(k, None)
    st.session_state['scheme_uploaded_key'] = cur_key

    data_bytes = uploaded_file.getvalue()
    is_csv = uploaded_file.name.lower().endswith('.csv')

    try:
        if is_csv:
            df_raw = pd.read_csv(io.BytesIO(data_bytes), header=None)
            sheet_names = ['Sheet1']
        else:
            xl = pd.ExcelFile(io.BytesIO(data_bytes))
            sheet_names = xl.sheet_names
            df_raw = xl.parse(sheet_names[0], header=None)
    except Exception as e:
        st.error(f'读取文件失败：{e}')
        return

    if df_raw.empty:
        st.error('文件内容为空。')
        return

    # 自动匹配方案
    detected = detect_scheme(df_raw, schemes)
    default_idx = 0
    if detected and detected[0]['name'] in scheme_names:
        default_idx = scheme_names.index(detected[0]['name'])
        st.success(f'已自动匹配方案：{detected[0]["name"]}')

    chosen = st.selectbox('选择验证方案', all_options, index=default_idx)

    if chosen == manual_label:
        header_idx = detected[1] if detected else 0
        df = df_with_header(df_raw, header_idx)
        manual_mode(df, uploaded_file)
        return

    scheme = next(s for s in schemes if s.get('name') == chosen)

    # 定位表头行：优先用自动检测结果，否则模糊查找，最后退回方案配置的表头行
    if detected and detected[0]['name'] == chosen:
        header_idx = detected[1]
    else:
        expected = [str(h).strip() for h in scheme.get('expected_headers', [])]
        idx, hits = find_best_header_row(df_raw, expected)
        header_idx = idx if (idx is not None and hits >= 3) else int(scheme.get('header_row', 2)) - 1

    df = df_with_header(df_raw, header_idx)

    st.write('原始数据预览：')
    st.write(df.head(20))

    matched_cols = [c for c in scheme.get('columns', {}) if c in df.columns]
    if not matched_cols:
        st.error('表头与所选方案不匹配：未找到方案中配置的任何列。请确认上传的是模板格式文件，或改用手动配置。')
        return

    # 开始校验 / 下载按钮并排：校验前开始校验为高亮主按钮，校验完成后变灰禁用，下载按钮为普通按钮
    has_run = 'scheme_out' in st.session_state
    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        run_clicked = st.button(
            '开始校验',
            type='primary' if not has_run else 'secondary',
            disabled=has_run,
            use_container_width=True,
        )
    with btn_col2:
        if has_run:
            out = st.session_state['scheme_out']
            red_cells = st.session_state['scheme_red']
            cell_changes = st.session_state['scheme_changes']
            delete_rows = st.session_state['scheme_delete']
            sheet = st.session_state['scheme_sheet']
            col_pos = st.session_state['scheme_col_pos']
            raw = st.session_state['scheme_raw']
            header_idx = st.session_state['scheme_header']

            if uploaded_file.name.lower().endswith('.xlsx'):
                # 在原文件上直接修改：保留说明行、表头、列宽等全部格式
                buf = export_modified_file(raw, sheet, col_pos, red_cells, cell_changes, delete_rows)
            else:
                # 非 xlsx（xls/csv）兜底：重建表格导出
                first_data_row = header_idx + 2
                all_orig = list(range(first_data_row, first_data_row + len(df)))
                kept = [r for r in all_orig if r not in set(delete_rows)]
                label_map = {r: i for i, r in enumerate(kept)}
                fallback_red = {}
                for col, rows in red_cells.items():
                    ls = [label_map[r] for r in rows if r in label_map]
                    if ls:
                        fallback_red[col] = set(ls)
                buf = export_result_file(out, fallback_red, sheet_name=sheet)

            # 文件名：项目名称 + 时间戳（项目名称为空时退回原文件名）
            proj_col = '项目名称（必填）'
            proj_name = ''
            if proj_col in out.columns:
                vals = out[proj_col].dropna().astype(str).str.strip()
                if len(vals):
                    proj_name = vals.iloc[0]
            safe_name = re.sub(r'[\\/:*?"<>|]', '_', proj_name).strip()
            base = safe_name or uploaded_file.name.rsplit('.', 1)[0]
            ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            out_name = f'{base}_{ts}.xlsx'
            st.download_button(
                '下载处理好的文件',
                buf,
                out_name,
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
            )
        else:
            st.caption('校验完成后可在此下载处理好的文件')

    if run_clicked:
        out, red_cells, cell_changes, delete_rows, summary = run_scheme(df, scheme, header_idx=header_idx)

        # 列名 -> 列号（用于在原文件上定位单元格）
        col_pos = {}
        for i, v in enumerate(df_raw.iloc[header_idx].tolist()):
            if not pd.isna(v):
                col_pos[str(v).strip()] = i + 1

        st.session_state['scheme_out'] = out
        st.session_state['scheme_red'] = red_cells
        st.session_state['scheme_changes'] = cell_changes
        st.session_state['scheme_delete'] = delete_rows
        st.session_state['scheme_sheet'] = sheet_names[0]
        st.session_state['scheme_col_pos'] = col_pos
        st.session_state['scheme_raw'] = data_bytes
        st.session_state['scheme_header'] = header_idx
        st.session_state['scheme_summary'] = summary
        st.rerun()

    # 校验摘要与结果预览（基于 session_state 渲染，校验后 rerun 依然显示）
    if has_run:
        summary = st.session_state.get('scheme_summary')
        out = st.session_state['scheme_out']
        st.write('校验摘要：')
        if summary and summary['issues']:
            for it in summary['issues']:
                st.write(f"{it['label']}：{'、'.join(it['cells'])}")
        else:
            st.write('未发现不符合规则的数据。')
        st.caption('修改说明：无效车牌/重复车牌已删除整行；姓名超15字已截断（原姓名复制到备注列）；'
                   '无效手机号已清空；标红的单元格请在原表中核对修改。')

        st.write('校验结果预览：')
        st.write(out)


if __name__ == '__main__':
    main()
