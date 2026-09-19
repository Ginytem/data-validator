import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import re
import io
import json
import os
import datetime
import numpy as np
import threading
import openpyxl
import time
import base64
import pyotp
import qrcode
from copy import copy
from openpyxl.styles import PatternFill, Font, Border, Alignment
from openpyxl.utils import get_column_letter


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
USAGE_LOG_PATH = os.path.join(BASE_DIR, 'usage_log.json')
USAGE_COUNT_PATH = os.path.join(BASE_DIR, 'usage_count.json')
TOTP_SECRET_PATH = os.path.join(BASE_DIR, 'totp_secret.txt')
HELP_ATTEMPTS_PATH = os.path.join(BASE_DIR, 'help_auth_attempts.json')
HELP_MAX_WRONG = 3          # 1 分钟内最多 3 次错误
HELP_WINDOW_SEC = 60        # 滑动窗口（秒）
_usage_lock = threading.Lock()

# 安全密码输入组件：普通文本框 + 字符圆点显示，浏览器不识别为密码框、不弹"保存密码"
_PW_COMPONENT = components.declare_component(
    'secure_password_input', path=os.path.join(BASE_DIR, 'secure_password'))


# ============ 使用记录（下载行为日志） ============

def _get_client_ip():
    """获取客户端 IP：优先 Streamlit context，其次 X-Forwarded-For（Cloudflare 隧道场景）。"""
    ip = None
    try:
        ip = getattr(st.context, 'client_ip', None)
    except Exception:
        ip = None
    if not ip or ip in ('127.0.0.1', '::1', 'localhost'):
        try:
            hdrs = getattr(st.context, 'headers', None) or {}
            xff = (hdrs.get('X-Forwarded-For') or hdrs.get('X-Real-IP') or '').strip()
            if xff:
                ip = xff.split(',')[0].strip()
        except Exception:
            pass
    if not ip or ip in ('127.0.0.1', '::1', 'localhost'):
        return '本机'
    return ip


def _load_usage_log():
    try:
        with open(USAGE_LOG_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []


def _save_usage_log(records):
    with _usage_lock:
        tmp = USAGE_LOG_PATH + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        os.replace(tmp, USAGE_LOG_PATH)


def _load_usage_count():
    """累计校验下载次数（初始 18，每次下载 +1；独立于可清空的使用记录）。"""
    try:
        with open(USAGE_COUNT_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return int(data.get('total', 18))
    except Exception:
        return 18


def _save_usage_count(total):
    with _usage_lock:
        tmp = USAGE_COUNT_PATH + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump({'total': total}, f, ensure_ascii=False, indent=2)
        os.replace(tmp, USAGE_COUNT_PATH)


def _append_download_log(filename):
    """记录一次下载：时间、IP、文件名、该 IP 当日下载次数（含本次）。"""
    ip = _get_client_ip()
    today = datetime.date.today().strftime('%Y-%m-%d')
    records = _load_usage_log()
    today_count = sum(1 for r in records if r.get('ip') == ip and r.get('date') == today) + 1
    records.append({
        'time': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'date': today,
        'ip': ip,
        'filename': filename,
        'day_count': today_count,
    })
    _save_usage_log(records)
    # 累计处理次数 +1（独立于可清空的下载记录，与下载行为同步）
    _save_usage_count(_load_usage_count() + 1)
    return ip, today_count


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


# ============ 多车牌支持（一个单元格多个车牌，逗号/顿号/分号/空白分隔） ============

def _split_plates_detailed(value):
    """拆分单元格内的多个车牌（多个车牌用英文逗号隔开，兼容顿号/分号）。

    返回 (plates, dup_in_cell)：
    - plates：去重后（保留首个出现顺序）的车牌列表（已转大写、去首尾空格；不做符号清理）
    - dup_in_cell：该格内重复出现的车牌集合（行内重复，保留一个）
    空值/空串返回 ([], set())。

    注意：此处只做拆分，不自动移除符号/空格——单车牌的特殊符号由
    run_scheme 步骤 0.5 自动清理；多车牌（含分隔符）交给拆分后的逐个校验，
    不正确的标红（用户规范录入多个车牌，带符号视为录入错误）。
    """
    if pd.isna(value):
        return [], set()
    s = str(value).strip()
    if not s:
        return [], set()
    parts = [p.strip().upper() for p in re.split(r'[,，、;；]+', s)]
    plates, dup = [], set()
    for p in parts:
        if not p:
            continue
        if p in plates:
            dup.add(p)
        else:
            plates.append(p)
    return plates, dup


def _capacity(value):
    """车位数 -> 免费名额容量（用于多车牌去重的容量优先分配）。

    - 缺失 / 空串 / 非数字：按 1 处理（保底 1 个名额，避免误删；必填缺失另有标红）
    - 数值：向下取整（0.5 -> 0，1.9 -> 1）；负数按 0
    """
    if pd.isna(value):
        return 1
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return 1
    try:
        f = float(value)
    except (ValueError, TypeError):
        return 1
    if pd.isna(f):
        return 1
    if f < 0:
        return 0
    return int(f)


def _capacity_is_valid(value):
    """车位数是否为有效数字（用于决定是否触发'车位不足'提示；
    缺失/非数字的行只报'必填项缺失'，不重复报'车位不足'）。"""
    if pd.isna(value):
        return False
    if isinstance(value, str) and value.strip() == '':
        return False
    try:
        f = float(value)
    except (ValueError, TypeError):
        return False
    return not pd.isna(f)


def _check_rule(rule, series, vconf=None):
    """单规则校验，返回 Series（True=有效）。

    非必填语义：plate / mobile / datetime 规则对空值放行，
    空值只交给 required 规则判定，避免同一格重复报两种问题。
    """
    if rule == 'required':
        return series.map(lambda x: _norm(x) != '')
    if rule == 'maxlen':
        n = (vconf or {}).get('max_len', 15)
        # 空格是否视为字符由配置决定：
        # 姓名（默认）空格算字符（连续超4个已在预处理压缩为2个）；门牌号/车位号等
        # 配置 ignore_space=true 的列，空格不视为字符（长度按移除空白后计算）
        if (vconf or {}).get('ignore_space'):
            return series.map(lambda x: len(re.sub(r'\s+', '', _norm(x))) <= n)
        return series.map(lambda x: len(_norm(x)) <= n)
    if rule == 'plate':
        def _plate_ok(x):
            if _norm(x) == '':
                return True
            plates, _ = _split_plates_detailed(x)
            if not plates:
                return False
            return all(PLATE_RE.match(p) for p in plates)
        return series.map(_plate_ok)
    if rule == 'mobile':
        # 空格不视为字符：匹配前移除所有空白（"138 1234 5678" 视为有效 11 位号码）
        def _mobile_ok(x):
            s = re.sub(r'\s+', '', _norm(x))
            return s == '' or bool(MOBILE_RE.match(s))
        return series.map(_mobile_ok)
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


def run_scheme(df, scheme, header_idx=0, match_fields=None, enable_multi=False):
    """按预设方案执行校验。

    流程：车牌 O/I 字母自动替换为 0/1 → 多车牌拆分 + 按车位数容量优先去重 →
    逐列逐规则校验并按失败动作处理 → 时间列归一化为 yyyy-mm-dd → 空车牌行删除。

    enable_multi（多位多车功能开关）：
    - False（未勾选，默认）：跳过姓名区分/合并，按原有逻辑做纯数据清洗。
    - True（勾选多位多车）：全部行参与匹配，匹配键可配置（match_fields）。
      匹配到同一车主的组：结束时间相同 → 合并为一行（车牌/门牌/备注汇总，备注≤100字），
      保留行"一位多车"填"是"、车位数写 1，其余行删除；
      结束时间不同 → 第二个起姓名加数字区分（不改写一位多车/车位数）。

    match_fields：多位多车匹配字段（勾选 enable_multi 时必选，可含
    默认/姓名/手机号/门牌号/车位号/身份证号）。
    - 含"默认"：按 姓名+手机号 匹配（手机号为空时退化为仅按姓名）
    - 其他字段：组合键——所选字段全部相同才视为同一车主，任一字段为空的行不参与

    失败动作：
    - mark           ：无效单元格标红（必填缺失 / 时间无法解析 / 车牌省份或位数不对）
    - clear          ：清空该单元格（无效手机号，原值复制到备注列）
    - truncate15     ：截断至 15 字（姓名超长，原姓名复制到备注列）
    - fill_from_plate：姓名缺失时用同行车牌号填充
    - delete_row     ：删除整行（车牌为空）

    多车牌去重（容量优先）：
    - 一格可含多个车牌（逗号/顿号/分号/空白分隔），拆分为独立子车牌逐个校验
    - 行内重复：保留一个，不删行，摘要提示
    - 跨行重复：按"车位数 = 免费名额"做容量优先分配——重复车牌优先保留在
      还有剩余名额的行；所有出现过的行都满时，保留在首次出现的行（宁超不丢，
      该行标红提示车位不足）
    - 分配后被剔空的车牌格视为空，按"车牌为空"删除整行

    返回 (result, red_cells, cell_changes, delete_rows, summary)：
    - result: 处理后的 DataFrame（用于界面预览）
    - red_cells: {列名: {原文件行号集合}}，标红单元格位置
    - cell_changes: {(原文件行号, 列名): 新值}，清空/截断/时间归一化/填充/去重写回值
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

    issue_items = []  # (label, orig_row, col_letter|None|自定义文本)
    oi_fix_items = []  # (orig_row, col_letter, 原值, 新值) —— 车牌 O/I 替换提示
    delete_rows = set()
    red_cells = {}
    cell_changes = {}
    cleared = 0
    truncated = 0

    # 各列原始值备份（用于清空/截断/多值提取时把最原始的完整内容复制到备注列留存，方便后期核对）
    col_originals = {}
    for col in columns_cfg:
        if col in out.columns:
            col_originals[col] = {idx: str(out.at[idx, col]) for idx in out.index}

    # 0) 车牌 O/I 字母自动替换为 0/1（修复），替换后再拆分、去重和校验
    for col, conf in columns_cfg.items():
        if col not in out.columns or conf.get('transform') != 'plate_o_i_fix':
            continue
        vals = out[col].map(lambda x: _norm(x).upper().replace('O', '0').replace('I', '1'))
        changed = out[col].map(lambda x: _norm(x)) != vals
        if changed.any():
            letter = _col_letter(list(out.columns).index(col) + 1)
            for r, old, new in zip(out.loc[changed, '__orig_row'].tolist(),
                                   out.loc[changed, col].tolist(),
                                   vals[changed].tolist()):
                cell_changes[(r, col)] = new
                oi_fix_items.append((r, letter, _norm(old), new))
            out[col] = out[col].astype(object)
            out.loc[changed, col] = vals[changed]

    # 0.5) 车牌符号清理：仅对"单车牌"（不含英文逗号/顿号/分号等分隔符）自动删除特殊符号
    #      （"渝A.52363" → "渝A52363"、"宁B•A2K52" → "宁BA2K52"）；
    #      多车牌（含分隔符）不做符号自动清理，由拆分后的逐个校验判断，不正确的标红
    PLATE_SYMBOL_CLEAN = re.compile(r'[^%sA-Z0-9]' % PLATE_PROVINCES)
    PLATE_SEP_RE = re.compile(r'[,，、;；]')
    for col, conf in columns_cfg.items():
        if col not in out.columns or conf.get('transform') != 'plate_o_i_fix':
            continue

        def _clean_plate_cell(v):
            s = _norm(v).upper()
            if not s or PLATE_SEP_RE.search(s):
                return s, False  # 空值或多车牌（含分隔符）：不自动清理符号
            new = PLATE_SYMBOL_CLEAN.sub('', s)
            if new == s:
                return s, False
            return new, True

        cleaned = out[col].map(_clean_plate_cell)
        changed_mask = cleaned.map(lambda x: x[1])
        if changed_mask.any():
            letter = _col_letter(list(out.columns).index(col) + 1)
            for r, old, new in zip(out.loc[changed_mask, '__orig_row'].tolist(),
                                   out.loc[changed_mask, col].tolist(),
                                   cleaned[changed_mask].map(lambda x: x[0]).tolist()):
                cell_changes[(r, col)] = new
                issue_items.append(('车牌符号已清理', r, letter))
            out[col] = out[col].astype(object)
            out.loc[changed_mask, col] = cleaned[changed_mask].map(lambda x: x[0])

    # 0.6) 姓名/手机号空格处理：
    #      - 姓名：连续超过 4 个空格时压缩为 2 个（其余空格保留，空格是姓名的有效字符），
    #        压缩后字符仍超限的，后续由 maxlen 截断（删减后段 + 复制到备注列）
    #      - 手机号：空格视为无效字符，直接移除
    name_col = next((c for c in columns_cfg if '车主姓名' in c), None)
    phone_col = next((c for c in columns_cfg if '手机号' in c), None)
    if name_col and name_col in out.columns:
        def _compress_name(x):
            s = _norm(x).replace('\u3000', ' ')  # 全角空格转半角
            return re.sub(r' {5,}', '  ', s)     # 连续超过4个（≥5）→ 保留2个

        vals = out[name_col].map(_compress_name)
        changed = out[name_col].map(lambda x: _norm(x)) != vals
        if changed.any():
            letter = _col_letter(list(out.columns).index(name_col) + 1)
            for r, old, new in zip(out.loc[changed, '__orig_row'].tolist(),
                                   out.loc[changed, name_col].tolist(),
                                   vals[changed].tolist()):
                cell_changes[(r, name_col)] = new
                issue_items.append(('姓名空格已压缩', r, letter))
            out[name_col] = out[name_col].astype(object)
            out.loc[changed, name_col] = vals[changed]
    if phone_col and phone_col in out.columns:
        vals = out[phone_col].map(lambda x: re.sub(r'\s+', '', _norm(x)))
        changed = out[phone_col].map(lambda x: _norm(x)) != vals
        if changed.any():
            letter = _col_letter(list(out.columns).index(phone_col) + 1)
            for r, old, new in zip(out.loc[changed, '__orig_row'].tolist(),
                                   out.loc[changed, phone_col].tolist(),
                                   vals[changed].tolist()):
                cell_changes[(r, phone_col)] = new
                issue_items.append(('手机号空格已清理', r, letter))
            out[phone_col] = out[phone_col].astype(object)
            out.loc[changed, phone_col] = vals[changed]

    # 0.7) 手机号多值处理：一个单元格可能出现多个手机号（逗号/顿号/分号/空格分隔），
    #      提取第一个有效手机号（11 位，1[3-9] 开头）保留，完整原值复制到备注列留存；
    #      拆不出有效号码的保持原值，交给后续 mobile 规则校验（清空 + 原值复制备注）
    if phone_col and phone_col in out.columns:
        remark_col = None
        for vconf in columns_cfg.get(phone_col, {}).get('validators', []):
            if vconf.get('copy_to'):
                remark_col = vconf['copy_to']
        letter = _col_letter(list(out.columns).index(phone_col) + 1)

        def _pick_phone(v):
            s = _norm(v)
            if not s:
                return s, '', False
            parts = [p.strip() for p in re.split(r'[,，、;；\s]+', s) if p.strip()]
            for p in parts:
                if MOBILE_RE.match(p):
                    return p, s, True
            return s, '', False

        picked = out[phone_col].map(_pick_phone)
        for idx in out.index:
            new_val, orig_full, is_changed = picked.loc[idx]
            if not is_changed or _norm(out.at[idx, phone_col]) == new_val:
                continue
            r = out.at[idx, '__orig_row']
            out.at[idx, phone_col] = new_val
            cell_changes[(r, phone_col)] = new_val
            if remark_col and remark_col in out.columns:
                out[remark_col] = out[remark_col].astype(object)
                prev = _norm(out.at[idx, remark_col])
                raw_full = col_originals.get(phone_col, {}).get(idx, orig_full)
                new_remark = f'{prev}；原手机号：{raw_full}' if prev else f'原手机号：{raw_full}'
                out.at[idx, remark_col] = new_remark
                cell_changes[(r, remark_col)] = new_remark
            issue_items.append(('手机号多值已保留首个有效', r, letter))

    # 1) 多车牌去重：拆分后按车位数容量优先分配，剔除重复子车牌（不删行），
    #    分配后整格为空的行由第 2 步"车牌为空"删除
    dup_col = dedup_cfg.get('column')
    capacity_col = dedup_cfg.get('capacity_col') or '车位数（必填项）'
    if dup_col and dup_col in out.columns:
        split_map = {}   # idx -> (plates, dup_in_cell)
        plate_rows = {}  # plate -> [idx, ...]（出现过的行）
        for idx in out.index:
            plates, dup_in_cell = _split_plates_detailed(out.at[idx, dup_col])
            split_map[idx] = (plates, dup_in_cell)
            for p in plates:
                plate_rows.setdefault(p, []).append(idx)

        cap_series = out[capacity_col] if capacity_col in out.columns else pd.Series(1, index=out.index)

        def cap_of(idx):
            return _capacity(cap_series.loc[idx])

        assigned = {}  # plate -> idx（最终保留在哪一行）
        row_load = {idx: 0 for idx in out.index}

        # 逐行扫描：重复车牌优先分配给"还有剩余名额"的出现行；全满则留首次出现行
        for idx in out.index:
            plates, _ = split_map[idx]
            for p in plates:
                if p in assigned:
                    continue
                target = None
                for c in plate_rows[p]:
                    if row_load[c] < cap_of(c):
                        target = c
                        break
                if target is None:
                    target = plate_rows[p][0]  # 宁超不丢
                assigned[p] = target
                row_load[target] += 1

        # 重写格子：只保留分配在本行的车牌，其余剔除并记录摘要
        dup_issues = []  # (orig_row, 自定义文本)
        for idx in out.index:
            plates, dup_in_cell = split_map[idx]
            r = out.at[idx, '__orig_row']
            for p in dup_in_cell:
                dup_issues.append((r, f'第{r}行（{p}，行内重复）'))
            kept = [p for p in plates if assigned.get(p) == idx]
            for p in plates:
                if assigned.get(p) != idx:
                    keep_r = out.at[assigned[p], '__orig_row']
                    dup_issues.append((r, f'第{r}行（{p}→保留至第{keep_r}行）'))
            new_val = ','.join(kept) if kept else ''
            if _norm(out.at[idx, dup_col]) != new_val:
                out.at[idx, dup_col] = new_val
                cell_changes[(r, dup_col)] = new_val if new_val else None
        for r, text in dup_issues:
            issue_items.append(('车牌重复', r, text))

        # 超容量提示：车位数有效且该行"一位多车"≠"是"时，保留车牌数 > 容量 → 标红（车位不足）。
        # 模板允许"一位多车=是 + 车位数1 + 多车牌"（一位多车合法，不报车位不足）
        multi_col_name = '一位多车（必填项 是或者否）'
        dup_letter = _col_letter(list(out.columns).index(dup_col) + 1)
        for idx in out.index:
            plates, _ = split_map[idx]
            kept_cnt = sum(1 for p in plates if assigned.get(p) == idx)
            cap = cap_of(idx)
            is_multi_plate = multi_col_name in out.columns and _norm(out.at[idx, multi_col_name]).upper() == '是'
            if kept_cnt > cap and _capacity_is_valid(cap_series.loc[idx]) and not is_multi_plate:
                r = out.at[idx, '__orig_row']
                red_cells.setdefault(dup_col, set()).add(r)
                issue_items.append(('车位不足', r, dup_letter))

    col_letters = {name: _col_letter(i + 1) for i, name in enumerate(out.columns)}

    # 1.5) 多位多车处理（功能开关 enable_multi）：仅勾选"多位多车"时执行。
    #      未勾选 → 跳过本步骤（不做姓名区分、不做合并，按原有逻辑纯数据清洗）。
    #      勾选后：全部行参与匹配（不再受"一位多车=是"限制），匹配键可配置：
    #       - 勾"默认"= 姓名+手机号（手机号为空退化为仅按姓名）
    #       - 勾其他字段（可多选）= 组合键，所选字段全部相同才算同一车主，任一字段为空的行不参与
    #      组内按结束时间再分组：
    #       - 结束时间相同的多行 → 合并为一行（车牌/门牌/备注汇总，备注≤100字），
    #         保留行"一位多车"填"是"、车位数写 1，其余行删除
    #       - 结束时间不同的 → 第二个起姓名加数字区分（不改写一位多车/车位数）
    # 列查找：优先方案 columns 配置，其次在上传文件表头列中模糊匹配
    # （车辆备注/身份证号等可能只出现在表头、未配置校验规则，同样要支持匹配与汇总）
    def _find_col(*keys):
        for k in keys:
            c = next((c for c in columns_cfg if k in c), None)
            if c and c in out.columns:
                return c
        return next((c for c in out.columns if any(k in c for k in keys)), None)

    name_col = _find_col('车主姓名')
    phone_col = _find_col('手机号')
    end_col = _find_col('结束时间')
    plate_col = _find_col('车牌号')
    door_col = _find_col('门牌号')
    remark_col = _find_col('车辆备注', '备注')
    id_col = _find_col('身份证')
    lot_col = _find_col('车位号')
    cap_col = _find_col('车位数')
    multi_col_name = '一位多车（必填项 是或者否）'

    if enable_multi:
        mf = match_fields or ['默认']
        default_mode = '默认' in mf
        field_map = {'姓名': name_col, '手机号': phone_col, '门牌号': door_col,
                     '车位号': lot_col, '身份证号': id_col}
        fields = ['姓名', '手机号'] if default_mode else [f for f in mf if f in field_map and field_map[f]]

        if name_col and end_col and name_col in out.columns and fields:
            def _match_key(row):
                key = []
                for f in fields:
                    col = field_map[f]
                    v = _norm(row.get(col)) if col and col in row.index else ''
                    if default_mode and f == '手机号':
                        if not v:
                            continue  # 默认模式：手机号为空退化为仅按姓名
                    elif not v:
                        return None  # 严格模式：任一字段为空，该行不参与匹配
                    key.append(v)
                return tuple(key)

            groups = {}
            for idx in out.index:
                nm = _norm(out.at[idx, name_col])
                if not nm:
                    continue
                k = _match_key(out.loc[idx])
                if k is None:
                    continue
                groups.setdefault(k, []).append(idx)

            for k, idxs in groups.items():
                if len(idxs) < 2:
                    continue
                # 组内按结束时间细分
                sub = {}
                for i in idxs:
                    e = _norm(out.at[i, end_col])
                    sub.setdefault(e, []).append(i)
                # ① 结束时间相同的子组（>1 行）→ 合并为一行
                for e, sub_idxs in sub.items():
                    if len(sub_idxs) < 2 or not e:
                        continue
                    target = sub_idxs[0]
                    target_r = out.at[target, '__orig_row']
                    # 车牌汇总（去重保序，英文逗号隔开）
                    if plate_col and plate_col in out.columns:
                        cur = _norm(out.at[target, plate_col])
                        merged, seen = [], set()
                        for m in sub_idxs:
                            for sp in re.split(r'[,，、;；\s]+', _norm(out.at[m, plate_col])):
                                if sp and sp not in seen:
                                    seen.add(sp)
                                    merged.append(sp)
                        newp = ','.join(merged)
                        if newp != cur:
                            out.at[target, plate_col] = newp
                            cell_changes[(target_r, plate_col)] = newp
                    # 门牌汇总（去重保序）
                    if door_col and door_col in out.columns:
                        cur = _norm(out.at[target, door_col])
                        merged = [d for d in dict.fromkeys(
                            _norm(out.at[m, door_col]) for m in sub_idxs) if d]
                        newd = '，'.join(merged)
                        if newd != cur:
                            out.at[target, door_col] = newd
                            cell_changes[(target_r, door_col)] = newd
                    # 备注汇总（≤100字）：原备注 + 各被合并行备注（标注来源行号）
                    if remark_col and remark_col in out.columns:
                        cur_remark = _norm(out.at[target, remark_col])
                        parts = [cur_remark] if cur_remark else []
                        for m in sub_idxs[1:]:
                            mr = _norm(out.at[m, remark_col])
                            if mr:
                                parts.append(f'第{out.at[m, "__orig_row"]}行备注：{mr}')
                        new_remark = '；'.join(parts)
                        if len(new_remark) > 100:
                            new_remark = new_remark[:99] + '…'
                        if new_remark != cur_remark:
                            out.at[target, remark_col] = new_remark
                            cell_changes[(target_r, remark_col)] = new_remark
                    # 保留行：一位多车填"是"、车位数写 1（匹配到=同一车主多辆车）
                    if multi_col_name in out.columns and _norm(out.at[target, multi_col_name]) != '是':
                        out.at[target, multi_col_name] = '是'
                        cell_changes[(target_r, multi_col_name)] = '是'
                    if cap_col and cap_col in out.columns and _norm(out.at[target, cap_col]) != '1':
                        out.at[target, cap_col] = 1
                        cell_changes[(target_r, cap_col)] = 1
                    issue_items.append(('多位多车已标注', target_r, f'第{target_r}行（一位多车=是，车位数1）'))
                    # 被合并行删除 + 摘要
                    for m in sub_idxs[1:]:
                        r = out.at[m, '__orig_row']
                        delete_rows.add(r)
                        issue_items.append(('多位多车合并', r, f'第{r}行（车牌合并至第{target_r}行）'))
                # ② 结束时间不同的 → 第二个起姓名加数字（防系统误判多位多车）
                ends = {_norm(out.at[i, end_col]) for i in idxs if _norm(out.at[i, end_col])}
                if len(ends) < 2:
                    continue
                remain = [i for i in idxs if out.at[i, '__orig_row'] not in delete_rows]
                if len(remain) < 2:
                    continue
                base_r = out.at[remain[0], '__orig_row']
                base_nm = _norm(out.at[remain[0], name_col])
                for n, i in enumerate(remain[1:], start=1):
                    cur = _norm(out.at[i, name_col])
                    suffix = str(n)
                    new_name = cur + suffix
                    if len(new_name) > 15:
                        new_name = cur[:15 - len(suffix)] + suffix  # 超 15 字：截断原名保数字
                    out.at[i, name_col] = new_name
                    r = out.at[i, '__orig_row']
                    cell_changes[(r, name_col)] = new_name
                    issue_items.append(('姓名区分', r, f'第{r}行 {base_nm} → {new_name}（与第{base_r}行结束时间不同）'))

    # 2) 逐列逐规则校验 + 修复动作
    for col, conf in columns_cfg.items():
        if col not in out.columns:
            continue
        letter = col_letters[col]
        for vconf in conf.get('validators', []):
            rule = vconf.get('rule')
            action = vconf.get('on_invalid', 'mark')
            invalid = ~_check_rule(rule, out[col], vconf)
            if not invalid.any():
                continue
            orig_rows = out.loc[invalid, '__orig_row'].tolist()
            if action == 'delete_row':
                delete_rows.update(orig_rows)
                dl_label = '车牌为空' if rule == 'required' else '车牌无效'
                for r in orig_rows:
                    issue_items.append((dl_label, r, None))
            elif action == 'clear':
                originals = out.loc[invalid, col].map(lambda x: _norm(x))
                out.loc[invalid, col] = None
                cleared += len(orig_rows)
                for r, v in zip(orig_rows, originals.tolist()):
                    cell_changes[(r, col)] = None
                    issue_items.append(('无效手机号', r, letter))
                # 清空前把原始完整值复制到备注列，避免丢失信息（用未清理的最原始值）
                copy_to = vconf.get('copy_to')
                if copy_to and copy_to in out.columns:
                    out[copy_to] = out[copy_to].astype(object)  # 兼容空备注列（float64 → object）
                    copy_label = vconf.get('copy_label', '原值')
                    for label, full in zip(out.index[invalid].tolist(), originals.tolist()):
                        r = out.at[label, '__orig_row']
                        prev = _norm(out.at[label, copy_to])
                        raw_full = col_originals.get(col, {}).get(label, full)
                        new_remark = f'{prev}；{copy_label}：{raw_full}' if prev else f'{copy_label}：{raw_full}'
                        out.at[label, copy_to] = new_remark
                        cell_changes[(r, copy_to)] = new_remark
            elif action == 'truncate':
                n = vconf.get('max_len', 15)
                originals = out.loc[invalid, col].map(lambda x: _norm(x))
                new_vals = originals.map(lambda x: x[:n])
                out.loc[invalid, col] = new_vals
                truncated += len(orig_rows)
                if vconf.get('red'):  # 姓名超长等要求标红的场景
                    red_cells.setdefault(col, set()).update(orig_rows)
                # 截断前先把原始完整内容复制到备注列，避免丢失重要信息（用未清理的最原始值）
                copy_to = vconf.get('copy_to')
                copy_label = vconf.get('copy_label', '原值')
                if copy_to and copy_to in out.columns:
                    out[copy_to] = out[copy_to].astype(object)  # 兼容空备注列（float64 → object）
                    for label, full in zip(out.index[invalid].tolist(), originals.tolist()):
                        r = out.at[label, '__orig_row']
                        prev = _norm(out.at[label, copy_to])
                        raw_full = col_originals.get(col, {}).get(label, full)
                        new_remark = f'{prev}；{copy_label}：{raw_full}' if prev else f'{copy_label}：{raw_full}'
                        out.at[label, copy_to] = new_remark
                        cell_changes[(r, copy_to)] = new_remark
                # 摘要标签动态化：姓名超15字 / 门牌号超20字 / 车位号超20字
                if '车主姓名' in col:
                    trunc_label = f'姓名超{n}字'
                elif '门牌号' in col:
                    trunc_label = f'门牌号超{n}字'
                elif '车位号' in col:
                    trunc_label = f'车位号超{n}字'
                else:
                    trunc_label = f'{col}超{n}字'
                for r, v in zip(orig_rows, new_vals.tolist()):
                    cell_changes[(r, col)] = v
                    issue_items.append((trunc_label, r, letter))
            elif action == 'fill_from_plate':
                source_col = vconf.get('source_col')
                if source_col and source_col in out.columns:
                    fill_vals = out.loc[invalid, source_col].map(lambda x: _norm(x))
                    can_fill = fill_vals.ne('')
                    fill_positions = fill_vals[can_fill].index
                    if len(fill_positions):
                        new_vals = out.loc[fill_positions, source_col].map(lambda x: _norm(x))
                        out.loc[fill_positions, col] = new_vals
                        for r, v in zip(out.loc[fill_positions, '__orig_row'].tolist(), new_vals.tolist()):
                            cell_changes[(r, col)] = v
                            issue_items.append(('姓名缺失已用车牌填充', r, letter))
                    no_fill_positions = fill_vals[~can_fill].index
                    if len(no_fill_positions):
                        red_cells.setdefault(col, set()).update(out.loc[no_fill_positions, '__orig_row'].tolist())
                        for r in out.loc[no_fill_positions, '__orig_row'].tolist():
                            issue_items.append(('必填项缺失', r, letter))
                else:
                    red_cells.setdefault(col, set()).update(orig_rows)
                    for r in orig_rows:
                        issue_items.append(('必填项缺失', r, letter))
            else:  # mark
                red_cells.setdefault(col, set()).update(orig_rows)
                label = ('必填项缺失' if rule == 'required' else
                         ('时间无法解析' if rule == 'datetime' else
                          ('车牌无效' if rule == 'plate' else '无效数据')))
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

    # 4.5) 问题行集中到末尾：有标红的行（需人工核对修改）排在最后，正常行保持原序在前。
    #      自动修复的行（时间归一化 / O/I 替换 / 去重 / 改名等，摘要已提示）不视为问题行
    problem_rows = set()
    for rows in red_cells.values():
        problem_rows.update(rows)
    if problem_rows:
        normal = out[~out['__orig_row'].isin(problem_rows)]
        problem = out[out['__orig_row'].isin(problem_rows)]
        out = pd.concat([normal, problem], ignore_index=True)

    # 5) 汇总：被删除行上的其它问题不再列出（删除原因本身、O/I 替换、多位多车合并提示除外）
    grouped = {}
    for label, r, letter in issue_items:
        if label not in ('车牌无效', '车牌重复', '车牌为空', '车牌O/I已替换', '多位多车合并') and r in delete_rows:
            continue
        grouped.setdefault(label, []).append((r, letter))

    # 车牌 O/I 替换提示：始终列出（用户可能在源文件自行修改，需要知道哪些车牌被替换）
    for r, letter, old, new in oi_fix_items:
        grouped.setdefault('车牌O/I已替换', []).append((r, letter, old, new))

    issues = []
    order = ['车牌O/I已替换', '车牌符号已清理', '姓名空格已压缩', '手机号空格已清理', '手机号多值已保留首个有效', '无效手机号',
             '姓名缺失已用车牌填充', '姓名超15字', '姓名区分', '多位多车已标注', '多位多车合并', '车牌重复',
             '车牌无效', '车位不足', '车牌为空', '必填项缺失', '时间无法解析']
    # 动态标签（如 门牌号超20字/车位号超20字）跟在白名单之后按序输出
    for label in order + [l for l in grouped if l not in order]:
        if label not in grouped:
            continue
        rows = sorted(grouped[label])
        if label == '车牌O/I已替换':
            cells = [f'{letter}{r}（{old}→{new}）' for r, letter, old, new in rows]
        elif label in ('车牌重复', '姓名区分', '多位多车合并', '多位多车已标注'):
            cells = [text for _, text in rows]
        elif label == '车牌为空':
            cells = [f'第{r}行' for r, _ in rows]
        elif label in ('车牌无效', '车位不足'):
            cells = [f'{letter}{r}' for r, letter in rows]  # 标红位置用单元格定位（如 H28）
        else:
            cells = [f'{letter}{r}' for r, letter in rows]
        issues.append({'label': label, 'cells': cells})

    result = out.drop(columns=['__orig_row'])
    summary = {
        'issues': issues,
        'cleared': cleared,
        'truncated': truncated,
        'order': out['__orig_row'].tolist(),  # 重排后行顺序（原文件行号），导出时按此移动行
    }
    return result, red_cells, cell_changes, sorted(delete_rows), summary


# ============ 导出（在原文件上直接修改，保留模板格式） ============

def export_modified_file(data_bytes, sheet_name, col_pos, red_cells, cell_changes, delete_rows,
                         order=None, header_row=2):
    """在用户上传的原文件上直接修改：
    - 保留第 1 行说明行、表头、列宽、合并单元格等全部原格式
    - 无效单元格标红（FFC7CE）
    - 无效手机号清空、姓名截断、时间归一化写回原单元格
    - 无效车牌 / 重复车牌所在行整行删除
    - order 非空时：按新顺序移动行（问题行集中到末尾，值/样式/行高跟随移动）
    """
    wb = openpyxl.load_workbook(io.BytesIO(data_bytes))
    ws = wb[sheet_name]
    red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

    # 模板样式常量（对齐「月租车辆导入模板」：数据行宋体11、无边框、无填充、
    # 左对齐垂直居中、时间列水平居中；16 列列宽；说明行/表头行行高）
    TEMPLATE_COL_WIDTHS = [38.1, 28.9, 13.0, 35.2, 32.8, 13.0, 31.8, 35.9, 46.8, 42.1,
                           30.7, 30.6, 35.8, 18.3, 19.3, 22.5]
    time_cols = {col_pos[col] for col in col_pos if '时间' in col}

    # 0) 统一数据区样式为模板格式（不依赖上传文件自身样式，保证导出美观一致）：
    #    字体宋体11、无边框、无填充、左对齐垂直居中（时间列水平居中）、列宽按模板
    first_data_row = header_row + 1
    for r in range(first_data_row, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name='宋体', size=11)
            cell.border = Border()
            cell.fill = PatternFill(fill_type=None)
            cell.alignment = Alignment(horizontal='center' if c in time_cols else 'left',
                                       vertical='center', wrap_text=False)
    for i, w in enumerate(TEMPLATE_COL_WIDTHS):
        if i < ws.max_column:
            ws.column_dimensions[get_column_letter(i + 1)].width = w
    if header_row >= 2:
        ws.row_dimensions[1].height = 37
    ws.row_dimensions[header_row].height = 34

    # 数据区样式快照（供被删除/清空的行复用，保持统一边框与格式）
    base_styles = [copy(ws.cell(row=first_data_row, column=c)._style) for c in range(1, ws.max_column + 1)]

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

    # 3) 按新顺序移动数据行（问题行集中到末尾；删除行被排除，原位置被后续行覆盖或清空）
    if order:
        first_data_row = header_row + 1  # 表头下一行是数据起始
        total = len(order) + len(delete_rows)
        last_data_row = first_data_row + total - 1
        # 读取数据区每行的值 + 样式 + 行高
        row_content = {}
        for src in range(first_data_row, last_data_row + 1):
            vals = [ws.cell(row=src, column=c).value for c in range(1, ws.max_column + 1)]
            styles = [copy(ws.cell(row=src, column=c)._style) for c in range(1, ws.max_column + 1)]
            rh = ws.row_dimensions[src].height
            row_content[src] = (vals, styles, rh)
        # 按新顺序写回数据区前部
        for i, orig in enumerate(order):
            target = first_data_row + i
            vals, styles, rh = row_content[orig]
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=target, column=c)
                cell.value = vals[c - 1]
                cell._style = styles[c - 1]
            if rh is not None:
                ws.row_dimensions[target].height = rh
        # 数据区剩余行（被删除行 / 未覆盖位置）清空值并应用统一模板样式（保留边框）
        for leftover in range(first_data_row + len(order), last_data_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=leftover, column=c)
                cell.value = None
                cell._style = copy(base_styles[c - 1])
            ws.row_dimensions[leftover].height = None
    else:
        # 无重排：直接删除整行（从下往上删，避免行号偏移）
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

def main_page():
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
    # 标题 + 累计处理次数（初始 18，每次校验完成并下载 +1）
    usage_total = _load_usage_count()
    st.markdown(
        '<div style="display:flex;align-items:baseline;gap:16px;flex-wrap:wrap;">'
        '<h1 style="margin:0;font-size:2.1rem;">Data Validator</h1>'
        f'<span style="font-size:0.85rem;opacity:0.7;white-space:nowrap;">已累计完成 {usage_total} 次数据处理</span>'
        '</div>',
        unsafe_allow_html=True,
    )

    schemes = load_schemes()
    scheme_names = [s.get('name', f'方案{i + 1}') for i, s in enumerate(schemes)]
    manual_label = '手动配置（不匹配方案）'
    all_options = scheme_names + [manual_label]

    # 上传标签 + 问题反馈入口：问题反馈紧挨"上传 CSV 或 Excel 文件"文字右侧（同一行）
    st.markdown(
        '<div style="display:flex;align-items:center;gap:12px;margin-bottom:4px;">'
        '<span style="font-size:14px;line-height:1.5;">上传 CSV 或 Excel 文件</span>'
        '<a href="https://f.kdocs.cn/g/NiuECq9U/" target="_blank" '
        'style="color:inherit;text-decoration:none;font-size:12px;'
        'background:rgba(128,128,128,0.15);padding:3px 12px;border-radius:12px;'
        'box-shadow:0 1px 4px rgba(0,0,0,0.08);display:inline-block;">问题反馈</a>'
        '</div>',
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader('', type=['csv', 'xls', 'xlsx'], label_visibility='collapsed')

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

    # 多位多车（功能开关）：勾选后启用"多位多车"处理——按所选匹配字段判定同一车主，
    # 匹配到的组：结束时间相同 → 合并为一行（车牌/门牌/备注汇总），保留行"一位多车"填"是"、车位数写1，
    # 其余行删除；结束时间不同 → 第二个起姓名加数字区分。
    # 未勾选 → 跳过姓名区分/合并，按原有逻辑做纯数据清洗。
    match_options = ['默认', '姓名', '手机号', '门牌号', '车位号', '身份证号']
    enable_multi = st.checkbox('多位多车', value=False,
                               help='测试功能，可能会有bug 慎用 如有发现bug 欢迎反馈')
    match_fields = []
    if enable_multi:
        match_fields = st.multiselect('选项', match_options, default=['默认'])
        if not match_fields:
            st.warning('请至少选择一个匹配字段')
    # 配置快照：勾选状态或匹配字段变化时，旧校验结果失效，直接清理（无需 rerun，同一次渲染内生效）
    cfg_sig = (enable_multi, tuple(match_fields))
    if st.session_state.get('scheme_cfg') != cfg_sig:
        for k in ['scheme_out', 'scheme_red', 'scheme_changes', 'scheme_delete',
                  'scheme_sheet', 'scheme_col_pos', 'scheme_raw', 'scheme_header', 'scheme_summary']:
            st.session_state.pop(k, None)
        st.session_state['scheme_cfg'] = cfg_sig

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
                # 在原文件上直接修改：保留说明行、表头、列宽等全部格式；
                # 问题行集中到文件末尾（order 为重排后的原文件行号顺序）
                order = st.session_state.get('scheme_summary', {}).get('order')
                buf = export_modified_file(raw, sheet, col_pos, red_cells, cell_changes, delete_rows,
                                           order=order, header_row=header_idx + 1)
            else:
                # 非 xlsx（xls/csv）兜底：重建表格导出
                order = st.session_state.get('scheme_summary', {}).get('order') or []
                pos_map = {r: i for i, r in enumerate(order)}
                fallback_red = {}
                for col, rows in red_cells.items():
                    ls = [pos_map[r] for r in rows if r in pos_map]
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
                on_click=_append_download_log,
                args=(out_name,),
            )
        else:
            st.caption('校验完成后可在此下载处理好的文件')

    # 校验操作密码（可选）：设置环境变量 VERIFY_PASSWORD（或 .streamlit/secrets.toml 的 VERIFY_PASSWORD）后，
    # 点击"开始校验"需先通过密码验证才执行；未配置密码则保持原行为直接校验
    try:
        auth_password = os.environ.get('VERIFY_PASSWORD') or st.secrets.get('VERIFY_PASSWORD', '')
    except Exception:
        auth_password = os.environ.get('VERIFY_PASSWORD', '')

    if run_clicked:
        if enable_multi and not match_fields:
            st.error('已勾选多位多车，请至少选择一个匹配字段后再开始校验')
        elif auth_password:
            st.session_state['need_auth'] = True
            st.rerun()
        else:
            st.session_state['do_run'] = True
            st.rerun()

    if st.session_state.get('need_auth'):
        with st.container(border=True):
            st.caption('本次校验需要权限验证，请输入校验密码')
            # 安全密码组件：普通输入框 + 圆点显示，浏览器不识别为密码框、不弹保存密码
            reset_token = st.session_state.get('auth_reset_token', 0)
            comp_res = _PW_COMPONENT(reset_token=reset_token, key='auth_pw_comp')
            if comp_res:
                if comp_res.get('action') == 'confirm':
                    if comp_res.get('pw') == auth_password:
                        st.session_state['need_auth'] = False
                        st.session_state['auth_err'] = False
                        st.session_state['do_run'] = True
                        st.rerun()
                    else:
                        st.session_state['auth_err'] = True
                        st.session_state['auth_reset_token'] = reset_token + 1
                        st.rerun()
                elif comp_res.get('action') == 'cancel':
                    st.session_state['need_auth'] = False
                    st.session_state['auth_err'] = False
                    st.rerun()
            if st.session_state.get('auth_err'):
                st.error('密码错误，请重试')

    if st.session_state.pop('do_run', False):
        out, red_cells, cell_changes, delete_rows, summary = run_scheme(df, scheme, header_idx=header_idx,
                                                                        match_fields=match_fields,
                                                                        enable_multi=enable_multi)

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
        # 校验摘要/修改说明：默认不展示（避免使用者看到内部处理逻辑）；
        # 设置环境变量 SHOW_VALIDATION_SUMMARY=1 后显示，便于管理员查看
        show_detail = os.environ.get('SHOW_VALIDATION_SUMMARY', '') == '1'
        if show_detail:
            st.write('校验摘要：')
            if summary and summary['issues']:
                for it in summary['issues']:
                    st.write(f"{it['label']}：{'、'.join(it['cells'])}")
            else:
                st.write('未发现不符合规则的数据。')
            # 修改说明：默认隐藏，小眼睛点击展开
            if 'show_note' not in st.session_state:
                st.session_state['show_note'] = False
            if st.button('修改说明', icon='👁', key='note_toggle'):
                st.session_state['show_note'] = not st.session_state['show_note']
            if st.session_state['show_note']:
                st.caption('修改说明：含字母O/I的车牌已自动替换为0/1（摘要中列出替换位置）；'
                           '无效车牌（省份/位数不对）已标红；一个单元格含多个车牌（逗号分隔）时逐个校验，'
                           '重复车牌按车位数容量优先分配去重（摘要中列出剔除位置与保留位置）；'
                           '非"一位多车"的行车牌数超过车位数时标红提示车位不足；'
                           '无效手机号已清空（原值复制到备注列）；'
                           '姓名超15字已截断（原姓名复制到备注列）；姓名缺失时用同行车牌号填充；'
                           '门牌号/车位号超20字已截断（原值复制到备注列）；'
                           '勾选"多位多车"后：按所选匹配字段判定同一车主（默认=姓名+手机号，手机号空退化为姓名；'
                           '也可选门牌号/车位号/身份证号等组合键，字段为空的行不参与匹配），'
                           '匹配到的组结束时间相同则合并为一行（车牌/门牌/备注汇总，备注≤100字），'
                           '保留行"一位多车"填"是"、车位数写1，其余行删除；'
                           '结束时间不同则第二个起姓名加数字1、2…区分；未勾选"多位多车"时不做姓名区分/合并；'
                           '车牌为空已删除整行；标红的单元格请在原表中核对修改。')

        st.write('校验结果预览：')
        st.write(out)


def _load_totp_secret():
    """读取或首次生成 TOTP 密钥；返回 (secret, is_first)。"""
    try:
        with open(TOTP_SECRET_PATH, 'r', encoding='utf-8') as f:
            s = f.read().strip()
        if s:
            return s, False
    except Exception:
        pass
    s = pyotp.random_base32()
    try:
        with open(TOTP_SECRET_PATH, 'w', encoding='utf-8') as f:
            f.write(s)
    except Exception:
        pass
    return s, True


def _help_attempts_load():
    try:
        with open(HELP_ATTEMPTS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _help_attempts_save(data):
    try:
        with open(HELP_ATTEMPTS_PATH, 'w', encoding='utf-8') as f:
            json.dump(data, f)
    except Exception:
        pass


def _help_check_locked(ip):
    """当前 IP 在滑动窗口内是否已达错误上限；返回 (locked, wait_seconds)。"""
    now = time.time()
    stamps = [float(t) for t in _help_attempts_load().get(ip, []) if now - float(t) < HELP_WINDOW_SEC]
    if len(stamps) >= HELP_MAX_WRONG:
        wait = int(HELP_WINDOW_SEC - (now - min(stamps)))
        return True, max(1, wait)
    return False, 0


def _help_record_wrong(ip):
    data = _help_attempts_load()
    now = time.time()
    stamps = [float(t) for t in data.get(ip, []) if now - float(t) < HELP_WINDOW_SEC]
    # iframe 组件 postMessage 偶发重复触发，1 秒内只记一次
    if stamps and now - stamps[-1] < 1.0:
        return
    stamps.append(now)
    data[ip] = [str(t) for t in stamps]
    _help_attempts_save(data)


def _help_clear_wrong(ip):
    data = _help_attempts_load()
    data[ip] = []
    _help_attempts_save(data)


def help_page():
    """使用说明页：展示校验规则与处理逻辑，管理员自用（访问需密码 258）"""
    st.set_page_config(page_title='使用说明 - Data Validator')

    # cookie 记住登录 3 小时（本地/公网各自独立）
    try:
        raw_cookie = ''
        for k, v in (getattr(st.context, 'headers', {}) or {}).items():
            if k.lower() == 'cookie':
                raw_cookie = v or ''
                break
        if 'help_auth=ok' in raw_cookie:
            st.session_state['help_auth_ok'] = True
    except Exception:
        pass

    secret, is_first = _load_totp_secret()

    if not st.session_state.get('help_auth_ok'):
        st.title('使用说明')
        with st.container(border=True):
            # 首次绑定：显示二维码 + 密钥
            if is_first:
                st.warning('首次使用：请用手机验证器（Google Authenticator / 微软 Authenticator / 微信「二次验证」）扫码绑定')
                otpauth = pyotp.TOTP(secret).provisioning_uri(name='DataValidator', issuer_name='DataValidator')
                st.image(qrcode.make(otpauth), width=200)
                st.caption('扫码失败可手动输入密钥：')
                st.code(secret)
                st.caption('绑定后，把验证器里的 6 位动态码填到下方')

            ip = _get_client_ip()
            locked, wait = _help_check_locked(ip)
            if locked:
                st.error(f"尝试过于频繁，请 {wait} 秒后再试（1 分钟内最多 {HELP_MAX_WRONG} 次错误）。")
                st.stop()
            st.caption('请输入验证码后访问')
            reset_token = st.session_state.get('help_auth_reset', 0)
            comp_res = _PW_COMPONENT(reset_token=reset_token, key='help_pw_comp')
            if comp_res:
                if comp_res.get('action') == 'confirm':
                    code = (comp_res.get('pw') or '').strip()
                    if pyotp.TOTP(secret).verify(code, valid_window=1):
                        _help_clear_wrong(ip)
                        components.html(
                            "<script>document.cookie='help_auth=ok; max-age=10800; path=/; SameSite=Lax'; setTimeout(function(){try{window.parent.location.reload();}catch(e){}}, 200);</script>",
                            height=0,
                        )
                        st.stop()
                    else:
                        _help_record_wrong(ip)
                        st.session_state['help_auth_err'] = True
                        st.session_state['help_auth_reset'] = reset_token + 1
                        st.rerun()
                elif comp_res.get('action') == 'cancel':
                    st.switch_page(st.session_state['_main_page_obj'])
            if st.session_state.get('help_auth_err'):
                st.error('验证码错误，请核对后再试！')
        st.stop()

    st.title('使用说明')
    st.caption('数据校验工具 · 校验规则与处理方式说明')

    st.subheader('一、使用步骤')
    st.markdown("""
1. 上传 CSV 或 Excel 文件（支持 `.xlsx` / `.xls` / `.csv`）；
2. 系统自动识别模板表头并匹配校验方案（也可手动选择方案）；
3. 点击「开始校验」，处理完成后即可下载处理好的文件；
4. 下载文件名以「项目名称 + 时间戳」命名，便于区分版本。
""")

    st.subheader('二、校验规则与处理方式')
    st.markdown("""
| 校验项 | 规则 | 处理方式 |
|---|---|---|
| 必填项 | 项目名称/车辆类型/车主姓名/车位数/一位多车/车牌号/开始时间/结束时间 不能为空 | 缺失时**标红**提示；车牌为空**删除整行**；车主姓名为空时用同行车牌号填充 |
| 车牌号 | 省份简称 + 发牌字母（不含 I/O）+ 后部 5 位（普通）/ 6 位（新能源） | 含字母 O/I 自动替换为数字 0/1；省份简称或位数不对**标红**，不删除 |
| 车牌重复 | 同一车牌在多行/同一格出现 | 按「车位数 = 免费名额」容量优先分配：优先保留在还有剩余名额的行，全满时保留首次出现的行；行内重复保留一个 |
| 多车牌 | 一个单元格可用英文逗号 `,` 隔开多个车牌 | 逐个校验，下载文件中规范为英文逗号分隔 |
| 车主姓名 | 限 15 字；姓名缺失用车牌填充 | 超 15 字先复制完整姓名到「车辆备注」列，再截断至 15 字；同名同手机号且"一位多车=是"的行结束时间不同时，第二个起姓名自动加数字 1、2…区分 |
| 手机号 | 11 位，1[3-9] 开头 | 不符合规则的**清空**该单元格，原值复制到「车辆备注」列 |
| 开始/结束时间 | 支持多种格式（如 yyyy-mm-dd、yy/mm/dd hh:mm:ss） | 统一归一化为 `yyyy-mm-dd`；无法解析的**标红** |
| 门牌号/车位号 | 限 20 字 | 超 20 字先复制完整原值到「车辆备注」列，再截断至 20 字 |
| 车位不足 | 车牌数超过车位数 | 仅当"一位多车"≠"是"时标红提示 |
| 问题行处理 | 需要人工核对修改的行 | 所有标红的行集中移动到文件末尾，正常行保持原序在前 |
""")

    st.subheader('三、下载文件说明')
    st.markdown("""
- 标红的单元格 = 需要人工核对修改的内容（含：无效车牌、必填项缺失、时间无法解析、车位不足）；
- 无效车牌已标红、重复车牌已删除整行、无效手机号已清空、姓名超长已截断（原值均在「车辆备注」列可查）；
- 有标红的行已集中到文件末尾，方便统一处理。
""")

    st.subheader('四、常见问题')
    st.markdown("""
- **一个格子里有多个车牌怎么办？** 用英文逗号 `,` 隔开即可（如 `贵A11111,贵B22222`），系统会逐个校验；
- **车牌含字母 O 或 I？** 会被自动替换为数字 0 和 1；
- **处理后的文件格式会变吗？** 不会，保留原模板的说明行、表头、列宽等全部格式，可直接使用。
""")

    st.subheader('五、使用记录')
    st.caption('记录每次下载处理文件的行为：时间 / IP / 处理后的文件名 / 该 IP 当日下载次数')
    records = _load_usage_log()
    if records:
        rows = [{'time': r.get('time', ''), 'ip': r.get('ip', ''), 'filename': r.get('filename', ''),
                 'day_count': r.get('day_count', 0)} for r in reversed(records)]
        st.dataframe(
            rows,
            use_container_width=True,
            hide_index=True,
            column_config={
                'time': '下载时间',
                'ip': 'IP',
                'filename': '处理后的文件名',
                'day_count': '当日下载次数',
            },
        )
        ip_stat = {}
        for r in records:
            ip_stat[r.get('ip', '')] = ip_stat.get(r.get('ip', ''), 0) + 1
        st.caption(f'共 {len(records)} 条下载记录，涉及 {len(ip_stat)} 个 IP。')
        if st.session_state.get('confirm_clear_usage'):
            c1, c2 = st.columns([1, 3])
            with c1:
                if st.button('确认清空全部记录', type='primary'):
                    _save_usage_log([])
                    st.session_state['confirm_clear_usage'] = False
                    st.rerun()
            with c2:
                if st.button('取消'):
                    st.session_state['confirm_clear_usage'] = False
                    st.rerun()
        else:
            if st.button('清空使用记录'):
                st.session_state['confirm_clear_usage'] = True
                st.rerun()
    else:
        st.caption('暂无使用记录。')

    st.subheader('六、更新记录')
    st.markdown("""
- **2026-09-18**：① 新增"多位多车"功能开关——勾选后需选择匹配字段（默认/姓名/手机号/门牌号/车位号/身份证号，可多选）；勾"默认"=姓名+手机号（手机号空退化为姓名），勾其他字段=组合键（全部相同才匹配，字段为空的行不参与）。匹配到的同一车主：结束时间相同 → 合并为一行（车牌/门牌/备注汇总到第一行，备注≤100字），保留行"一位多车"填"是"、车位数写1，其余行删除；结束时间不同 → 第二个起姓名加数字区分。未勾选"多位多车" → 跳过姓名区分/合并，按原有逻辑纯数据清洗。
- **2026-09-17**：① 车牌符号处理细化——单车牌（无英文逗号）自动删除特殊符号（`渝A.52363`→`渝A52363`、`宁B•A2K52`→`宁BA2K52`）；多车牌（英文逗号隔开）拆分后逐个校验，不正确的标红（不自动删符号）；② 姓名空格——连续超过 4 个空格压缩为 2 个（其余保留），压缩后仍超 15 字走截断+复制备注；③ 手机号空格自动移除；多个手机号保留第一个有效号码，完整原值复制到备注；无效手机号清空并复制原始值到备注（方便核对）；④ 门牌号/车位号长度判断忽略空格；⑤ 导出文件数据区样式统一为模板样式（字体/边框/对齐/底纹一致）。
- **2026-09-16**：`/help` 页面加访问密码 258（管理员自用，复用安全密码组件）。
- **2026-09-15**：新增 `/help` 使用说明页（管理员自用，不展示在前台导航）；校验摘要与修改说明默认隐藏（设置环境变量 `SHOW_VALIDATION_SUMMARY=1` 可显示）；校验密码支持环境变量/`secrets.toml` 配置，密码框改用自研安全组件（普通文本框+圆点显示，浏览器不弹"保存密码"）。
- **2026-09-15**：标红行（需人工核对修改）处理完后集中移动到文件末尾；姓名超 15 字截断后不再标红、不进末尾问题行；"姓名区分加数字"仅对"一位多车=是"的行生效。
- **2026-09-14**：门牌号/车位号超 20 字截断（原值复制到备注列）；车牌重复按车位数容量优先分配去重；问题行摘要以单元格定位展示；下载文件名以"项目名称+时间戳"命名。
""")

    st.caption('本页面为管理员自用说明，前台不展示；主功能请返回「数据校验」页。')


def main():
    # 隐藏侧边栏导航：前台只显示主页面，/help 通过 URL 直接访问（管理员自用）
    main_pg = st.Page(main_page, title='数据校验', url_path='main', default=True)
    help_pg = st.Page(help_page, title='使用说明', url_path='help')
    st.session_state['_main_page_obj'] = main_pg
    st.navigation([main_pg, help_pg], position='hidden').run()


if __name__ == '__main__':
    main()
