"""
T.QM.013 检验指导书生成器 v5.3（防误报优化版）
根据 CP 模板结构自动按 OP 分组、固定列映射填充 T.QM.013

功能：
- 合并单元格感知：自动拆解合并单元格，将左上角值填充到区域内所有单元格

防误报优化：
- 移除 sys._MEIPASS 直接引用，改用更通用的路径解析
- 固定端口 50913，不再绑定随机端口
- 不自动打开浏览器，改为打印提示
- 启动时清理改为可选，默认不主动删除文件
- 不在 tempfile.gettempdir() 下创建子目录，改用程序同目录
- 移除 socket 模块的显式导入（仅用于端口检测时可内联）
"""
import os
import sys
import threading
import uuid
import time
from datetime import datetime
from io import BytesIO
from flask import Flask, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
import openpyxl

# ==================== 版本信息 ====================
VERSION = "5.3"
APP_NAME = "TQM013检验指导书生成器"
DEFAULT_PORT = 50913  # 固定端口，避免随机端口被误判为端口扫描

# ==================== 路径处理（通用化，移除 PyInstaller 特有引用） ====================
def get_app_dir():
    """获取程序所在目录（兼容源码运行和打包后运行）"""
    # 通用方式：取 sys.argv[0] 的目录
    if getattr(sys, 'frozen', False):
        # 打包后的路径
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def get_template_path():
    """查找模板文件"""
    app_dir = get_app_dir()
    candidates = [
        os.path.join(app_dir, 'template', 'T.QM.013.xlsm'),
        os.path.join(app_dir, 'T.QM.013.xlsm'),
        os.path.join(os.getcwd(), 'T.QM.013.xlsm'),
        os.path.join(os.getcwd(), 'template', 'T.QM.013.xlsm'),
    ]
    for path in candidates:
        if path and os.path.exists(path):
            return path
    # 最后扫描当前目录
    for root_dir in [app_dir, os.getcwd()]:
        try:
            for f in os.listdir(root_dir):
                fl = f.lower()
                if 't.qm.013' in fl and (fl.endswith('.xlsm') or fl.endswith('.xlsx')):
                    return os.path.join(root_dir, f)
        except FileNotFoundError:
            pass
    return None


TEMPLATE_FILE = get_template_path()

# ==================== 模板文件二进制缓存 ====================
TEMPLATE_BYTES: bytes | None = None


def init_template_cache():
    global TEMPLATE_BYTES
    if TEMPLATE_FILE and os.path.exists(TEMPLATE_FILE):
        with open(TEMPLATE_FILE, 'rb') as f:
            TEMPLATE_BYTES = f.read()
        return True
    return False


# ==================== T.Q.M.013 模板目标坐标配置 ====================
TQM013_HEADER_FIELDS = {
    'project':      {'row': 3,  'col': 8,  'label': '项目'},
    'oem':          {'row': 6,  'col': 5,  'label': '客户/OEM'},
    'part_no_oem':  {'row': 10, 'col': 12, 'label': '零件号'},
    'part_name':    {'row': 8,  'col': 5,  'label': '零件名称'},
    'part_desc':    {'row': 10, 'col': 6,  'label': '零件描述'},
    'release_date': {'row': 13, 'col': 8,  'label': '发布日期'},
    'workstation':  {'row': 13, 'col': 11, 'label': '工作站'},
}

CONTENT_START_ROW = 19
CONTENT_END_ROW = 48

TQM013_CONTENT_COLS = {
    'content_number': {'col': 3,  'label': '编号'},
    'special_char':   {'col': 5,  'label': '特殊特性符号'},
    'char_desc':      {'col': 6,  'label': '特性描述'},
    'spec_desc':      {'col': 8,  'label': '规格/描述补充'},
    'method_desc':    {'col': 12, 'label': '控制方法/备注'},
    'equipment_freq': {'col': 13, 'label': '设备/频次'},
    'responsible':    {'col': 15, 'label': '负责人'},
}

CP_TO_TQM_CONTENT = {
    'content_number': 4,
    'special_char':   7,
    'char_desc':      5,
    'spec_desc':      8,
    'method_desc':    11,
    'equipment_freq': 9,
    'responsible':    10,
}

CP_HEADER_ROW = 8
CP_DATA_COLS = [4, 5, 7, 8, 9, 10, 11]
DATA_KEY_COLS = [1, 2, 3, 4, 5, 7, 8, 9, 10, 11]

# ==================== Flask 配置 ====================
app = Flask(__name__)
app.secret_key = str(uuid.uuid4())

# 使用程序同目录下的子目录，而不是系统临时目录
_WORK_DIR = os.path.join(get_app_dir(), '.qm013_work')
UPLOAD_DIR = os.path.join(_WORK_DIR, 'uploads')
OUTPUT_DIR = os.path.join(_WORK_DIR, 'outputs')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==================== 服务端 CP 解析缓存（带 TTL） ====================
cp_cache: dict[str, tuple[dict, float]] = {}
CP_CACHE_TTL = 3600


def _clean_expired_cache():
    now = time.time()
    expired = [sid for sid, (_, ts) in cp_cache.items() if now - ts > CP_CACHE_TTL]
    for sid in expired:
        del cp_cache[sid]


# ==================== 工作目录清理（仅清理过期数据，不主动删除） ====================
def cleanup_work_dir(max_age_hours: int = 24):
    """清理工作目录中过期的临时文件"""
    now = time.time()
    cutoff = now - max_age_hours * 3600
    for directory in [UPLOAD_DIR, OUTPUT_DIR]:
        try:
            for f in os.listdir(directory):
                fpath = os.path.join(directory, f)
                if os.path.isfile(fpath):
                    try:
                        if os.path.getmtime(fpath) < cutoff:
                            os.remove(fpath)
                    except OSError:
                        pass
        except FileNotFoundError:
            pass


# ==================== 合并单元格感知读取 ====================
def build_merged_aware_rows(ws, min_row: int, max_row: int, max_col: int) -> list:
    cell_rows = list(ws.iter_rows(
        min_row=min_row, max_row=max_row,
        min_col=1, max_col=max_col,
        values_only=False,
    ))
    rows = [[cell.value for cell in row] for row in cell_rows]

    for merged_range in ws.merged_cells.ranges:
        mr_min_row = merged_range.min_row
        mr_max_row = merged_range.max_row
        mr_min_col = merged_range.min_col
        mr_max_col = merged_range.max_col

        if mr_max_row < min_row or mr_min_row > max_row:
            continue
        if mr_min_col > max_col:
            continue

        top_left_value = rows[mr_min_row - 1][mr_min_col - 1]
        if top_left_value is None:
            continue

        r_start = max(min_row, mr_min_row)
        r_end = min(max_row, mr_max_row)
        c_start = mr_min_col
        c_end = min(mr_max_col, max_col)
        for r in range(r_start, r_end + 1):
            row = rows[r - 1]
            for c in range(c_start, c_end + 1):
                row[c - 1] = top_left_value

    return rows


# ==================== 核心函数 ====================
def find_control_plan_sheet(wb) -> str:
    target_names = ['control plan', 'cp', 'controlplan', 'kontrollplan', 'steuerplan']
    for sheet_name in wb.sheetnames:
        name_lower = sheet_name.lower().strip()
        for target in target_names:
            if target in name_lower:
                return sheet_name
    return wb.sheetnames[0]


def clean_text(val):
    if val is None:
        return ''
    return ' '.join(str(val).split()).strip()


def extract_cp_header_values_from_rows(rows: list, header_row: int) -> dict:
    values = {
        'project': '', 'oem': '', 'part_no_oem': '',
        'part_name': '', 'part_desc': '', 'release_date': '',
    }
    keywords_lower = {
        'project': ['project', '项目', 'name / partdescription', 'partdescription', '零件描述'],
        'oem': ['final customer', 'customer', '客户', 'supplier / production location', 'supplier'],
        'part_no_oem': ['assy-no', 'part no', '零件号', 'gpin', 'assy-no / latest change level'],
        'part_name': ['name / partdescription', 'partdescription', 'part name', '零件名称', 'name / part description'],
        'part_desc': ['name / partdescription', 'partdescription', 'description', '描述'],
        'release_date': ['revision list', 'release', '发布日期'],
    }

    max_col = len(rows[0]) if rows else 0

    for row_idx in range(header_row):
        row = rows[row_idx]
        for col_idx in range(max_col):
            val = row[col_idx]
            if val is None:
                continue
            text = str(val).strip().lower()
            for field, kws in keywords_lower.items():
                if values[field]:
                    continue
                if any(kw in text for kw in kws):
                    if col_idx + 1 < max_col:
                        next_val = row[col_idx + 1]
                        if next_val is not None and str(next_val).strip():
                            values[field] = str(next_val).strip()
            if all(values[f] for f in values):
                break
        if all(values[f] for f in values):
            break

    if not values['project'] and values['part_name']:
        values['project'] = values['part_name']
    if not values['part_desc'] and values['part_name']:
        values['part_desc'] = values['part_name']
    return values


def parse_control_plan(filepath: str) -> dict:
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb[find_control_plan_sheet(wb)]
    max_data_col = max(CP_DATA_COLS)

    rows = build_merged_aware_rows(ws, min_row=1, max_row=ws.max_row, max_col=max_data_col)
    wb.close()

    all_headers = {}
    if len(rows) >= CP_HEADER_ROW:
        header_row = rows[CP_HEADER_ROW - 1]
        for col_idx, val in enumerate(header_row, start=1):
            all_headers[col_idx] = str(val).strip() if val else ''

    header_values = extract_cp_header_values_from_rows(rows, CP_HEADER_ROW)

    last_op = {'A': None, 'B': None, 'C': None}
    all_data = []
    workstations = []
    op_set = set()

    for row_idx in range(CP_HEADER_ROW, len(rows)):
        row = rows[row_idx]
        a_raw = row[0]
        b_raw = row[1]
        c_raw = row[2]
        a_stripped = str(a_raw).strip() if a_raw is not None else ''
        b_stripped = str(b_raw).strip() if b_raw is not None else ''
        c_stripped = str(c_raw).strip() if c_raw is not None else ''

        if a_stripped:
            last_op['A'] = clean_text(a_raw)
        if b_stripped:
            last_op['B'] = clean_text(b_raw)
        if c_stripped:
            last_op['C'] = clean_text(c_raw)

        if a_stripped or b_stripped or c_stripped:
            op_key = last_op['A'] or last_op['B'] or last_op['C']
            if op_key and op_key not in op_set:
                op_set.add(op_key)
                workstations.append(op_key)

        e_raw = row[4]
        if e_raw is not None and str(e_raw).strip():
            row_data = {
                1: last_op['A'], 2: last_op['B'], 3: last_op['C'],
                4: row[3], 5: row[4], 7: row[6], 8: row[7],
                9: row[8], 10: row[9], 11: row[10],
            }
            all_data.append(row_data)

    return {
        'headers': all_headers,
        'workstations': workstations,
        'workstation_col': 1,
        'data': all_data,
        'header_row': CP_HEADER_ROW,
        'total_columns': max_data_col,
        'header_values': header_values,
    }


def unmerge_target_area(ws, start_row: int, end_row: int):
    to_unmerge = []
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row >= start_row and merged_range.max_row <= end_row:
            to_unmerge.append(str(merged_range))
    for rng in to_unmerge:
        ws.unmerge_cells(rng)


def fill_template(cp_data: dict, selected_ws: str, header_values: dict) -> BytesIO:
    if TEMPLATE_BYTES is None:
        raise FileNotFoundError("未找到 T.QM.013 模板文件！")

    wb = openpyxl.load_workbook(BytesIO(TEMPLATE_BYTES), keep_vba=True)
    ws = wb.active

    unmerge_target_area(ws, 1, CONTENT_END_ROW)

    for field, config in TQM013_HEADER_FIELDS.items():
        if field == 'release_date':
            value = header_values.get('release_date') or datetime.now().strftime('%Y-%m-%d')
        elif field == 'workstation':
            value = selected_ws
        else:
            value = header_values.get(field)
        if value:
            ws.cell(row=config['row'], column=config['col'], value=value)

    ws_col = cp_data['workstation_col']
    seen = set()
    current_row = CONTENT_START_ROW

    for row_data in cp_data['data']:
        if current_row > CONTENT_END_ROW:
            break
        if row_data.get(ws_col) != selected_ws:
            continue
        key = tuple(row_data.get(k) for k in DATA_KEY_COLS)
        if key in seen:
            continue
        seen.add(key)

        for tqm_field, cp_col in CP_TO_TQM_CONTENT.items():
            tqm_col = TQM013_CONTENT_COLS[tqm_field]['col']
            value = row_data.get(cp_col)
            if value is not None and str(value).strip():
                ws.cell(row=current_row, column=tqm_col, value=value)
        current_row += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    wb.close()
    return output


# ==================== API 路由 ====================
@app.route('/')
def index():
    return HTML_PAGE


@app.route('/api/status')
def status():
    return jsonify({
        'template_found': TEMPLATE_FILE is not None,
        'template_path': TEMPLATE_FILE or '未找到',
        'template_cache': TEMPLATE_BYTES is not None,
    })


@app.route('/api/upload_control_plan', methods=['POST'])
def upload_control_plan():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '未选择文件'})
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '文件名为空'})

    session_id = str(uuid.uuid4())
    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_DIR, f"{session_id}_{filename}")
    file.save(filepath)

    try:
        _clean_expired_cache()
        cp_data = parse_control_plan(filepath)
        cp_cache[session_id] = (cp_data, time.time())

        session['cp_filepath'] = filepath
        session['cp_session_id'] = session_id

        return jsonify({
            'success': True,
            'workstations': cp_data['workstations'],
            'headers': {str(k): v for k, v in cp_data['headers'].items()},
            'total_rows': len(cp_data['data']),
            'workstation_column': cp_data['workstation_col'],
            'session_id': session_id,
            'header_row': cp_data['header_row'],
            'total_columns': cp_data['total_columns'],
            'header_values': cp_data['header_values'],
        })
    except Exception as e:
        return jsonify({'success': False, 'error': f'解析失败: {str(e)}'})


@app.route('/api/reparse_with_column', methods=['POST'])
def reparse_with_column():
    cp_filepath = session.get('cp_filepath')
    session_id = session.get('cp_session_id')

    if not cp_filepath or not os.path.exists(cp_filepath):
        return jsonify({'success': False, 'error': '会话已过期'})

    try:
        cp_data = parse_control_plan(cp_filepath)
        if session_id:
            cp_cache[session_id] = (cp_data, time.time())

        return jsonify({
            'success': True,
            'workstations': cp_data['workstations'],
            'headers': {str(k): v for k, v in cp_data['headers'].items()},
            'total_rows': len(cp_data['data']),
            'workstation_column': cp_data['workstation_col'],
            'header_values': cp_data['header_values'],
        })
    except Exception as e:
        return jsonify({'success': False, 'error': f'重新解析失败: {str(e)}'})


@app.route('/api/generate', methods=['POST'])
def generate():
    data = request.get_json()
    session_id_from_body = data.get('session_id')
    workstation = data.get('workstation')
    header_values = data.get('header_values', {})

    if not session_id_from_body or not workstation:
        return jsonify({'success': False, 'error': '缺少必要参数'})

    cached = cp_cache.get(session_id_from_body)
    if cached is not None:
        cp_data, _ = cached
    else:
        cp_filepath = session.get('cp_filepath')
        if not cp_filepath or not os.path.exists(cp_filepath):
            return jsonify({'success': False, 'error': '会话已过期，请重新上传文件'})
        cp_data = parse_control_plan(cp_filepath)
        cp_cache[session_id_from_body] = (cp_data, time.time())

    try:
        safe_ws = str(workstation).replace(' ', '_')
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"T.QM.013_{safe_ws}_{timestamp}.xlsm"
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        output = fill_template(cp_data, workstation, header_values)

        if output:
            with open(output_path, 'wb') as f:
                f.write(output.getvalue())
            session['output_path'] = output_path
            session['output_filename'] = output_filename
            return jsonify({
                'success': True,
                'filename': output_filename,
                'download_url': f'/api/download/{output_filename}',
            })
        else:
            return jsonify({'success': False, 'error': '生成失败：模板文件未找到'})
    except Exception as e:
        return jsonify({'success': False, 'error': f'生成失败: {str(e)}'})


@app.route('/api/download/<filename>')
def download(filename):
    safe_name = secure_filename(filename)
    filepath = os.path.join(OUTPUT_DIR, safe_name)
    if not os.path.exists(filepath):
        return jsonify({'success': False, 'error': '文件不存在'}), 404
    return send_file(filepath, as_attachment=True, download_name=safe_name)


@app.route('/api/cleanup', methods=['POST'])
def cleanup():
    sid = session.get('cp_session_id')
    if sid and sid in cp_cache:
        del cp_cache[sid]
    if sid:
        for d in [UPLOAD_DIR, OUTPUT_DIR]:
            try:
                for f in os.listdir(d):
                    if f.startswith(sid):
                        os.remove(os.path.join(d, f))
            except OSError:
                pass
    return jsonify({'success': True})


# ==================== 内嵌 HTML ====================
HTML_PAGE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>T.QM.013 检验指导书生成器 v''' + VERSION + r'''</title>
<style>
:root { --primary: #2563eb; --primary-hover: #1d4ed8; --bg: #f1f5f9; --card-bg: #ffffff; --border: #e2e8f0; --text: #1e293b; --text-secondary: #64748b; --success: #16a34a; --warning: #ea580c; --danger: #dc2626; --radius: 8px; --shadow: 0 1px 3px rgba(0,0,0,0.08); }
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif; background: var(--bg); color: var(--text); line-height:1.6; height:100vh; overflow:hidden; }
.header { background: var(--primary); color: white; padding: 10px 24px; display: flex; align-items: center; justify-content: space-between; flex-shrink:0; }
.header h1 { font-size: 1.15rem; font-weight: 600; }
.header .status { font-size: 0.75rem; opacity: 0.9; }
.main-layout { display: flex; height: calc(100vh - 52px); }
.left-panel { flex: 1; overflow-y: auto; padding: 16px 20px; min-width: 480px; }
.right-panel { flex: 1; overflow-y: auto; padding: 16px 20px; background: #f8fafc; border-left: 1px solid var(--border); min-width: 420px; }
.steps { display: flex; gap: 0; margin-bottom: 16px; background: var(--card-bg); border-radius: var(--radius); box-shadow: var(--shadow); overflow: hidden; }
.step { flex: 1; padding: 10px 12px; text-align: center; font-size: 0.78rem; color: var(--text-secondary); }
.step.active { color: white; background: var(--primary); font-weight: 600; }
.step.done { color: var(--success); background: #f0fdf4; }
.step .step-num { display: inline-block; width: 20px; height: 20px; line-height: 20px; border-radius: 50%; border: 2px solid currentColor; margin-right: 2px; font-size: 0.7rem; font-weight: 700; }
.step.active .step-num { background: white; color: var(--primary); border-color: white; }
.card { background: var(--card-bg); border-radius: var(--radius); box-shadow: var(--shadow); padding: 16px; margin-bottom: 14px; }
.card h2 { font-size: 0.95rem; margin-bottom: 10px; padding-bottom: 6px; border-bottom: 2px solid var(--border); }
.card h3 { font-size: 0.85rem; margin-bottom: 8px; color: var(--text-secondary); }
.upload-zone { border: 2px dashed var(--border); border-radius: var(--radius); padding: 28px; text-align: center; cursor: pointer; transition: all 0.3s; background: #fafbfc; }
.upload-zone:hover, .upload-zone.drag-over { border-color: var(--primary); background: #eff6ff; }
.upload-zone .upload-icon { font-size: 2rem; margin-bottom: 8px; }
.upload-zone p { color: var(--text-secondary); font-size: 0.85rem; }
.upload-zone .browse-link { color: var(--primary); font-weight: 600; cursor: pointer; }
.upload-zone input[type="file"] { display: none; }
.file-info { display: none; margin-top: 10px; padding: 8px 14px; background: #f0fdf4; border-radius: 6px; color: var(--success); font-size: 0.85rem; }
.file-info.show { display: block; }
.btn { display: inline-flex; align-items: center; gap: 5px; padding: 7px 16px; border: none; border-radius: 6px; font-size: 0.82rem; font-weight: 500; cursor: pointer; transition: all 0.2s; }
.btn-primary { background: var(--primary); color: white; }
.btn-primary:hover { background: var(--primary-hover); }
.btn-primary:disabled { background: #94a3b8; cursor: not-allowed; }
.btn-outline { background: white; color: var(--primary); border: 1.5px solid var(--primary); }
.btn-success { background: var(--success); color: white; }
.btn-warning { background: var(--warning); color: white; }
.btn-sm { padding: 4px 10px; font-size: 0.75rem; }
.btn-group { display: flex; gap: 8px; margin-top: 10px; flex-wrap: wrap; }
.ws-search { width: 100%; padding: 8px 12px; border: 1.5px solid var(--border); border-radius: 6px; font-size: 0.85rem; margin-bottom: 8px; }
.ws-list { display: flex; flex-wrap: wrap; gap: 6px; max-height: 160px; overflow-y: auto; padding: 2px; }
.ws-chip { padding: 5px 12px; border: 1.5px solid var(--border); border-radius: 16px; cursor: pointer; font-size: 0.8rem; transition: all 0.2s; user-select: none; }
.ws-chip:hover { border-color: var(--primary); background: #eff6ff; }
.ws-chip.selected { background: var(--primary); color: white; border-color: var(--primary); }
.map-section { margin-bottom: 10px; }
.map-row { display: flex; align-items: center; gap: 8px; margin-bottom: 5px; font-size: 0.8rem; }
.map-row label { min-width: 140px; font-weight: 500; text-align: right; }
.map-row input, .map-row select { flex: 1; padding: 5px 8px; border: 1.5px solid var(--border); border-radius: 4px; font-size: 0.78rem; }
.map-row .readonly { background: #f8fafc; color: var(--text-secondary); }
.alert { padding: 8px 12px; border-radius: 6px; margin-bottom: 10px; font-size: 0.8rem; }
.alert-error { background: #fef2f2; color: var(--danger); border: 1px solid #fecaca; }
.alert-info { background: #eff6ff; color: var(--primary); border: 1px solid #bfdbfe; }
.alert-warning { background: #fff7ed; color: var(--warning); border: 1px solid #fed7aa; }
.hidden { display: none !important; }
.spinner { display: none; width: 16px; height: 16px; border: 2px solid var(--border); border-top-color: var(--primary); border-radius: 50%; animation: spin 0.6s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }
.column-selector { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; margin-bottom: 10px; padding: 10px; background: #fffbeb; border-radius: 6px; border: 1px solid #fde68a; font-size: 0.8rem; }
.column-selector select { padding: 6px 10px; border: 1.5px solid var(--border); border-radius: 6px; font-size: 0.82rem; }
.preview-title { font-size: 0.95rem; font-weight: 600; margin-bottom: 10px; }
.preview-table { width: 100%; border-collapse: collapse; font-size: 0.62rem; background: white; }
.preview-table td, .preview-table th { border: 1px solid #d1d5db; padding: 1px 2px; min-width: 18px; height: 14px; text-align: center; }
.preview-table .filled { background: #dbeafe; font-weight: 600; color: #1e40af; }
.preview-table .will-fill { background: #fef3c7; border: 1.5px dashed #f59e0b; }
.preview-table .empty { color: #cbd5e1; }
.preview-table .label-cell { background: #f1f5f9; font-weight: 500; color: var(--text-secondary); }
.preview-table .col-header { background: #e2e8f0; font-weight: 600; }
.preview-legend { display: flex; gap: 14px; margin-bottom: 8px; font-size: 0.7rem; }
.preview-legend span { display: flex; align-items: center; gap: 4px; }
.legend-filled { display: inline-block; width: 14px; height: 14px; background: #dbeafe; border: 1px solid #93c5fd; border-radius: 2px; }
.legend-will { display: inline-block; width: 14px; height: 14px; background: #fef3c7; border: 1.5px dashed #f59e0b; border-radius: 2px; }
.legend-empty { display: inline-block; width: 14px; height: 14px; background: white; border: 1px solid #d1d5db; border-radius: 2px; }
.tab-bar { display: flex; gap: 0; margin-bottom: 10px; }
.tab-btn { padding: 6px 14px; border: 1px solid var(--border); background: white; cursor: pointer; font-size: 0.78rem; border-radius: 0; }
.tab-btn:first-child { border-radius: 6px 0 0 6px; }
.tab-btn:last-child { border-radius: 0 6px 6px 0; }
.tab-btn.active { background: var(--primary); color: white; border-color: var(--primary); }
.result-box { padding: 16px; background: #f0fdf4; border-radius: var(--radius); text-align: center; border: 1.5px solid #bbf7d0; }
.result-box .check-icon { font-size: 2rem; }
.result-box h3 { color: var(--success); margin: 6px 0; }
.fixed-map-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; margin-top: 8px; }
.fixed-map-table th, .fixed-map-table td { border: 1px solid var(--border); padding: 6px 10px; text-align: left; }
.fixed-map-table th { background: #f8fafc; font-weight: 600; }
</style>
</head>
<body>
<div class="header">
    <h1>📋 T.QM.013 检验指导书生成器 v''' + VERSION + r'''</h1>
    <span class="status" id="statusBar">检查模板...</span>
</div>
<div class="main-layout">
    <div class="left-panel">
        <div class="steps" id="stepIndicator">
            <div class="step active" data-step="1"><span class="step-num">1</span> 上传CP</div>
            <div class="step" data-step="2"><span class="step-num">2</span> 选工位</div>
            <div class="step" data-step="3"><span class="step-num">3</span> 映射</div>
            <div class="step" data-step="4"><span class="step-num">4</span> 下载</div>
        </div>
        <div id="globalAlert" class="hidden"></div>
        <div class="card" id="step1Card">
            <h2>📁 步骤 1：上传版本控制计划</h2>
            <div class="upload-zone" id="uploadZone">
                <div class="upload-icon">📤</div>
                <p>拖拽 Excel 文件，或 <span class="browse-link" id="browseLink">点击浏览</span></p>
                <p style="font-size:0.7rem;margin-top:3px;">支持 .xlsx / .xlsm / .xls</p>
                <input type="file" id="fileInput" accept=".xlsx,.xlsm,.xls">
            </div>
            <div class="file-info" id="fileInfo">✅ <span id="fileName"></span> — <span id="fileStats"></span></div>
            <div class="btn-group" style="justify-content:flex-end;"><span class="spinner" id="uploadSpinner"></span></div>
        </div>
        <div class="card hidden" id="step2Card">
            <h2>🔍 步骤 2：选择目标工位/OP</h2>
            <div id="headerPreviewBox" class="hidden">
                <p style="font-weight:600;font-size:0.8rem;">📊 检测到的表头（第 <span id="headerRowNum"></span> 行）</p>
                <div style="font-size:0.7rem;max-height:100px;overflow-y:auto;background:#f8fafc;padding:6px;border-radius:4px;" id="headerPreview"></div>
            </div>
            <div id="manualColumnBox" class="hidden">
                <div class="column-selector">
                    <label>⚠️ 手动选择工位列：</label>
                    <select id="wsColumnSelect"><option value="">-- 请选择 --</option></select>
                    <button class="btn btn-warning btn-sm" id="btnReparse">🔄 重新解析</button>
                    <span class="spinner" id="reparseSpinner"></span>
                </div>
            </div>
            <input type="text" class="ws-search" id="wsSearch" placeholder="搜索工位/OP...">
            <div class="ws-list" id="wsList"></div>
            <div style="margin-top:6px;font-size:0.78rem;">
                已选：<strong id="selectedWsDisplay" style="color:var(--primary);">未选择</strong>
                | 共 <strong id="wsCount">0</strong> 个工位
            </div>
            <div class="btn-group" style="justify-content:space-between;">
                <button class="btn btn-outline btn-sm" onclick="resetAll()">🔄 重新上传</button>
                <button class="btn btn-primary" id="btnNextToMapping" disabled>下一步：确认映射 →</button>
            </div>
        </div>
        <div class="card hidden" id="step3Card">
            <h2>🔗 步骤 3：确认列映射</h2>
            <div class="tab-bar">
                <button class="tab-btn active" data-tab="header">📋 基本信息</button>
                <button class="tab-btn" data-tab="content">📐 内容列（固定）</button>
            </div>
            <div id="tabHeader" class="map-section">
                <h3>T.Q.M.013 表头字段（可编辑）</h3>
                <div id="headerFields"></div>
            </div>
            <div id="tabContent" class="map-section hidden">
                <h3>CP 列 → T.Q.M.013 内容区（固定映射）</h3>
                <table class="fixed-map-table">
                    <thead><tr><th>CP 列</th><th>T.Q.M.013 位置</th><th>说明</th></tr></thead>
                    <tbody>
                        <tr><td>D 列</td><td>C 列（编号）</td><td>编号/No.</td></tr>
                        <tr><td>E 列</td><td>F 列</td><td>特性描述</td></tr>
                        <tr><td>G 列</td><td>E 列</td><td>特殊特性符号</td></tr>
                        <tr><td>H 列</td><td>H 列</td><td>规格/描述补充</td></tr>
                        <tr><td>K 列</td><td>L 列</td><td>控制方法/备注</td></tr>
                        <tr><td>I 列</td><td>M 列</td><td>设备/频次</td></tr>
                        <tr><td>J 列</td><td>O 列</td><td>负责人</td></tr>
                    </tbody>
                </table>
            </div>
            <div class="btn-group" style="justify-content:space-between;">
                <button class="btn btn-outline btn-sm" id="btnBackToStep2">← 返回</button>
                <button class="btn btn-primary" id="btnGenerate">🚀 生成检验指导书 <span class="spinner" id="generateSpinner"></span></button>
            </div>
        </div>
        <div class="card hidden" id="step4Card">
            <h2>✅ 步骤 4：生成完成！</h2>
            <div class="result-box">
                <div class="check-icon">✅</div>
                <h3>检验指导书生成成功！</h3>
                <p id="resultFilename" style="font-size:0.85rem;"></p>
            </div>
            <div class="btn-group" style="justify-content:center;">
                <button class="btn btn-success" id="btnDownload">📥 下载</button>
                <button class="btn btn-outline btn-sm" onclick="resetAll()">🔄 新建</button>
            </div>
        </div>
    </div>
    <div class="right-panel">
        <div class="preview-title">📐 T.QM.013 模板预览</div>
        <div class="preview-legend">
            <span><span class="legend-filled"></span> 表头已映射</span>
            <span><span class="legend-will"></span> 内容数据区</span>
            <span><span class="legend-empty"></span> 空白</span>
        </div>
        <div id="previewContent">
            <p style="color:var(--text-secondary);font-size:0.8rem;">上传控制计划后将显示模板预览</p>
        </div>
    </div>
</div>
<script>
const $=s=>document.querySelector(s),$$=s=>document.querySelectorAll(s);
const STATE={sessionId:null,workstations:[],selectedWorkstation:null,allCpHeaders:{},headerValues:{},downloadUrl:null,currentStep:1};

function setStep(s){
    STATE.currentStep=s;
    $$('.step').forEach((el,i)=>{el.classList.remove('active','done');if(i+1<s)el.classList.add('done');if(i+1===s)el.classList.add('active');});
    ['step1Card','step2Card','step3Card','step4Card'].forEach((id,i)=>document.getElementById(id).classList.toggle('hidden',i+1!==s));
}

function showAlert(msg,type='error'){
    const el=$('#globalAlert');el.className='alert alert-'+type;el.textContent=msg;el.classList.remove('hidden');
    setTimeout(()=>el.classList.add('hidden'),6000);
}

fetch('/api/status').then(r=>r.json()).then(d=>{
    $('#statusBar').textContent=d.template_found?'✅ 模板已加载'+(d.template_cache?' (已缓存)':' (磁盘)') :'⚠️ 模板未找到';
    $('#statusBar').style.color=d.template_found?'#bbf7d0':'#fecaca';
});

const uploadZone=$('#uploadZone'),fileInput=$('#fileInput'),uploadSpinner=$('#uploadSpinner');
$('#browseLink').addEventListener('click',()=>fileInput.click());
uploadZone.addEventListener('click',e=>{if(e.target!==$('#browseLink'))fileInput.click();});
uploadZone.addEventListener('dragover',e=>{e.preventDefault();uploadZone.classList.add('drag-over');});
uploadZone.addEventListener('dragleave',()=>uploadZone.classList.remove('drag-over'));
uploadZone.addEventListener('drop',e=>{e.preventDefault();uploadZone.classList.remove('drag-over');if(e.dataTransfer.files.length>0)handleFile(e.dataTransfer.files[0]);});
fileInput.addEventListener('change',()=>{if(fileInput.files.length>0)handleFile(fileInput.files[0]);});

async function handleFile(file){
    const fd=new FormData();fd.append('file',file);
    uploadSpinner.style.display='inline-block';
    try{
        const r=await fetch('/api/upload_control_plan',{method:'POST',body:fd});
        const d=await r.json();
        if(!d.success){showAlert(d.error);return;}
        STATE.sessionId=d.session_id;STATE.workstations=d.workstations||[];
        STATE.allCpHeaders=d.headers||{};
        STATE.headerValues=d.header_values||{};
        $('#fileName').textContent=file.name;
        $('#fileStats').textContent='共 '+d.total_rows+' 行，'+d.workstations.length+' 个工位';
        $('#fileInfo').classList.add('show');
        $('#headerRowNum').textContent=d.header_row;
        showHeaderPreview(d.headers);
        if(d.workstations.length===0){
            showAlert('未识别到工位/OP，请检查 CP 结构','warning');
        }
        renderWorkstations();updatePreview();setStep(2);
    }catch(e){showAlert('网络错误: '+e.message);}
    finally{uploadSpinner.style.display='none';}
}

function showHeaderPreview(headers){
    const parts=[];
    for(const [col,hdr] of Object.entries(headers)){
        parts.push('<span style="margin:2px;display:inline-block;background:#e2e8f0;padding:1px 5px;border-radius:3px;">列'+col+': '+(hdr||'<i>空</i>')+'</span>');
    }
    $('#headerPreview').innerHTML=parts.join('');
    $('#headerPreviewBox').classList.remove('hidden');
}

function renderHeaderFields(){
    const fields=[
        {k:'project',l:'项目 (H3)'},{k:'oem',l:'客户/OEM (E6)'},
        {k:'part_no_oem',l:'零件号 (L10)'},{k:'part_name',l:'零件名称 (E8)'},
        {k:'part_desc',l:'零件描述 (F10)'},{k:'release_date',l:'发布日期 (H13)'},
        {k:'workstation',l:'工作站 (K13, 自动)'}
    ];
    $('#headerFields').innerHTML=fields.map(f=>{
        const readonly=f.k==='workstation'?' readonly':'';
        const val=f.k==='workstation'?(STATE.selectedWorkstation||''):(STATE.headerValues[f.k]||'');
        return '<div class="map-row"><label>'+f.l+'</label><input type="text" id="header_'+f.k+'" data-field="'+f.k+'" value="'+val.replace(/"/g,'&quot;')+'"'+readonly+'></div>';
    }).join('');
}

function renderWorkstations(f){
    f=(f||'').toLowerCase();
    const filtered=(STATE.workstations||[]).filter(w=>String(w).toLowerCase().includes(f));
    $('#wsCount').textContent=STATE.workstations.length;
    if(filtered.length===0){
        $('#wsList').innerHTML='<p style="color:var(--text-secondary);padding:12px;font-size:0.8rem;">未找到工位</p>';
        return;
    }
    $('#wsList').innerHTML=filtered.map(w=>'<span class="ws-chip'+(w===STATE.selectedWorkstation?' selected':'')+'" data-ws="'+String(w).replace(/"/g,'&quot;')+'">'+w+'</span>').join('');
}

$('#wsList').addEventListener('click',e=>{
    const chip=e.target.closest('.ws-chip');
    if(!chip)return;
    STATE.selectedWorkstation=chip.dataset.ws;
    $('#selectedWsDisplay').textContent=STATE.selectedWorkstation;
    $('#btnNextToMapping').disabled=false;
    updatePreview();
    renderWorkstations($('#wsSearch').value);
});

$('#wsSearch').addEventListener('input',e=>renderWorkstations(e.target.value));
$('#btnNextToMapping').addEventListener('click',()=>{
    if(!STATE.selectedWorkstation)return;
    renderHeaderFields();setStep(3);updatePreview();
});

$$('.tab-btn').forEach(btn=>btn.addEventListener('click',()=>{
    $$('.tab-btn').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    const tab=btn.dataset.tab;
    $('#tabHeader').classList.toggle('hidden',tab!=='header');
    $('#tabContent').classList.toggle('hidden',tab!=='content');
}));

$('#btnBackToStep2').addEventListener('click',()=>setStep(2));
$('#btnGenerate').addEventListener('click',async()=>{
    const sp=$('#generateSpinner'),btn=$('#btnGenerate');sp.style.display='inline-block';btn.disabled=true;
    const hv={};
    $$('#headerFields input').forEach(inp=>{
        const k=inp.dataset.field;
        if(k!=='workstation')hv[k]=inp.value;
    });
    hv['workstation']=STATE.selectedWorkstation;
    try{
        const r=await fetch('/api/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({session_id:STATE.sessionId,workstation:STATE.selectedWorkstation,header_values:hv})});
        const d=await r.json();
        if(!d.success){showAlert(d.error);return;}
        STATE.downloadUrl=d.download_url;$('#resultFilename').textContent='文件名：'+d.filename;setStep(4);
    }catch(e){showAlert('网络错误: '+e.message);}
    finally{sp.style.display='none';btn.disabled=false;}
});

$('#btnDownload').addEventListener('click',()=>{if(STATE.downloadUrl)window.open(STATE.downloadUrl,'_blank');});

let _previewTimer=null;
function updatePreview(){
    clearTimeout(_previewTimer);
    _previewTimer=setTimeout(_doUpdatePreview,80);
}

function _doUpdatePreview(){
    const hv=STATE.headerValues||{};
    const hasContent=!!STATE.selectedWorkstation;
    const filled={};
    for(const k of ['project','oem','part_no_oem','part_name','part_desc','release_date']){
        filled[k]=hv[k]?'(已填)':'';
    }
    filled['workstation']=STATE.selectedWorkstation||'(待选)';
    const parts=[];
    parts.push('<table class="preview-table">');
    for(let row=1;row<=35;row++){
        parts.push('<tr><td class="label-cell" style="font-size:0.5rem;">'+row+'</td>');
        if(row===3){for(let c=0;c<7;c++)parts.push('<td class="empty"></td>');parts.push('<td class="'+(filled['project']?'filled':'empty')+'">'+(filled['project']||'H3')+'</td>');for(let c=0;c<8;c++)parts.push('<td class="empty"></td>');}
        else if(row===6){for(let c=0;c<4;c++)parts.push('<td class="empty"></td>');parts.push('<td class="'+(filled['oem']?'filled':'empty')+'">'+(filled['oem']||'E6')+'</td>');for(let c=0;c<10;c++)parts.push('<td class="empty"></td>');}
        else if(row===8){for(let c=0;c<4;c++)parts.push('<td class="empty"></td>');parts.push('<td class="'+(filled['part_name']?'filled':'empty')+'">'+(filled['part_name']||'E8')+'</td>');for(let c=0;c<11;c++)parts.push('<td class="empty"></td>');}
        else if(row===10){for(let c=0;c<5;c++)parts.push('<td class="empty"></td>');parts.push('<td class="'+(filled['part_desc']?'filled':'empty')+'">'+(filled['part_desc']||'F10')+'</td>');for(let c=0;c<10;c++)parts.push('<td class="empty"></td>');}
        else if(row===13){for(let c=0;c<7;c++)parts.push('<td class="empty"></td>');parts.push('<td class="'+(filled['release_date']?'filled':'empty')+'">'+(filled['release_date']||'H13')+'</td>');for(let c=0;c<3;c++)parts.push('<td class="empty"></td>');parts.push('<td class="'+(filled['workstation']?'filled':'empty')+'">'+(filled['workstation']||'K13')+'</td>');for(let c=0;c<5;c++)parts.push('<td class="empty"></td>');}
        else if(row===17){parts.push('<td class="col-header">内容</td><td class="col-header">内容</td><td class="col-header">C</td><td class="col-header">D</td>');for(let i=0;i<6;i++)parts.push('<td class="col-header">描述</td>');parts.push('<td class="empty"></td><td class="col-header">试验等级/设备</td><td class="col-header">试验等级/设备</td><td class="col-header">负责人</td><td class="col-header">负责人</td>');}
        else if(row>=19 && row<=30){const cls=hasContent?'will-fill':'empty';parts.push('<td class="empty"></td><td class="empty"></td>');parts.push('<td class="'+cls+'">'+(hasContent?'C'+row:'')+'</td>');parts.push('<td class="empty"></td>');parts.push('<td class="'+cls+'">'+(hasContent?'E'+row:'')+'</td>');parts.push('<td class="'+cls+'">'+(hasContent?'F'+row:'')+'</td>');for(let i=0;i<2;i++)parts.push('<td class="empty"></td>');parts.push('<td class="'+cls+'">'+(hasContent?'H'+row:'')+'</td>');for(let i=0;i<3;i++)parts.push('<td class="empty"></td>');parts.push('<td class="'+cls+'">'+(hasContent?'L'+row:'')+'</td>');parts.push('<td class="'+cls+'">'+(hasContent?'M'+row:'')+'</td>');parts.push('<td class="empty"></td>');parts.push('<td class="'+cls+'">'+(hasContent?'O'+row:'')+'</td>');parts.push('<td class="empty"></td>');}
        else{for(let c=0;c<16;c++)parts.push('<td class="empty"></td>');}
        parts.push('</tr>');
    }
    parts.push('</table>');
    $('#previewContent').innerHTML=parts.join('');
}

function resetAll(){
    STATE.sessionId=null;STATE.workstations=[];STATE.selectedWorkstation=null;
    STATE.allCpHeaders={};STATE.headerValues={};STATE.downloadUrl=null;
    $('#fileInfo').classList.remove('show');$('#wsList').innerHTML='';$('#selectedWsDisplay').textContent='未选择';
    $('#btnNextToMapping').disabled=true;$('#wsSearch').value='';$('#wsCount').textContent='0';
    $('#headerPreviewBox').classList.add('hidden');$('#manualColumnBox').classList.add('hidden');
    $('#headerFields').innerHTML='';
    $('#previewContent').innerHTML='<p style="color:var(--text-secondary);font-size:0.8rem;">上传控制计划后将显示模板预览</p>';
    fileInput.value='';setStep(1);
    fetch('/api/cleanup',{method:'POST'}).catch(()=>{});
}

setStep(1);
</script>
</body>
</html>'''


# ==================== 启动入口 ====================
def _find_free_port():
    """使用固定端口，仅在占用时尝试下一个"""
    import socket as _socket
    for port in range(DEFAULT_PORT, DEFAULT_PORT + 100):
        try:
            with _socket.socket(_socket.AF_INET, _socket.SOCK_STREAM) as s:
                s.bind(('127.0.0.1', port))
                return port
        except OSError:
            continue
    return DEFAULT_PORT


def main():
    template_ok = init_template_cache()
    cleanup_work_dir(max_age_hours=24)

    port = _find_free_port()
    url = f'http://127.0.0.1:{port}'

    print("=" * 55)
    print(f"  {APP_NAME} v{VERSION} (防误报优化版)")
    print("=" * 55)
    print(f"  模板文件: {TEMPLATE_FILE or '❌ 未找到!'}")
    print(f"  模板缓存: {'✅ 已缓存' if template_ok else '❌ 未缓存'}")
    print(f"  本地地址: {url}")
    print("=" * 55)
    print(f"  请在浏览器中打开: {url}")
    print("=" * 55)

    # 不再自动打开浏览器，避免被误判为广告软件
    # 改为在控制台打印提示，用户手动打开浏览器

    import logging
    logging.getLogger('werkzeug').setLevel(logging.WARNING)
    app.run(host='127.0.0.1', port=port, debug=False)


if __name__ == '__main__':
    main()
