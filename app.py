"""
T.QM.013 检验指导书生成器 v2.0
新增：基本信息自动匹配、右侧实时预览、下拉菜单映射
"""
import os
import sys
import socket
import threading
import webbrowser
import uuid
import json
from datetime import datetime
from io import BytesIO

from flask import Flask, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
import openpyxl


# ==================== PyInstaller 路径处理 ====================
def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def get_template_path():
    base = get_base_path()
    candidates = [
        os.path.join(base, 'template', 'T.QM.013.xlsm'),
        os.path.join(base, 'T.QM.013.xlsm'),
        os.path.join(os.path.dirname(sys.executable), 'T.QM.013.xlsm') if getattr(sys, 'frozen', False) else None,
        os.path.join(os.getcwd(), 'T.QM.013.xlsm'),
    ]
    for path in candidates:
        if path and os.path.exists(path):
            return path
    for root_dir in [base, os.getcwd()]:
        try:
            for f in os.listdir(root_dir):
                if 'T.QM.013' in f and (f.endswith('.xlsm') or f.endswith('.xlsx')):
                    return os.path.join(root_dir, f)
        except FileNotFoundError:
            pass
    return None


TEMPLATE_FILE = get_template_path()

# ==================== Flask 配置 ====================
app = Flask(__name__)
app.secret_key = str(uuid.uuid4())

import tempfile
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), 'qm013_uploads')
OUTPUT_DIR = os.path.join(tempfile.gettempdir(), 'qm013_outputs')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==================== T.QM.013 模板完整映射 ====================
# 格式: (row, col, 标签)
TQM013_HEADER_FIELDS = {
    'project':        {'row': 5,  'col': 6,  'label': '项目'},
    'sub_project':    {'row': 5,  'col': 9,  'label': 'Sub-Project'},
    'language':       {'row': 5,  'col': 13, 'label': '语言', 'default': '中文'},
    'oem':            {'row': 7,  'col': 2,  'label': 'OEM'},
    'part_no_oem':    {'row': 7,  'col': 6,  'label': '零件号 OEM'},
    'dir':            {'row': 7,  'col': 9,  'label': 'DIR'},
    'dmba':           {'row': 7,  'col': 13, 'label': 'DmbA', 'default': 'A'},
    'page':           {'row': 7,  'col': 16, 'label': '页'},
    'part_name':      {'row': 8,  'col': 2,  'label': '零件名称'},
    'model_type':     {'row': 8,  'col': 6,  'label': '模型/种类'},
    'bat_material':   {'row': 8,  'col': 9,  'label': 'BAT-物料'},
    'product_group':  {'row': 8,  'col': 13, 'label': '产品组'},
    'instruction':    {'row': 8,  'col': 16, 'label': '指导', 'default': 'ON'},
    'single_part_desc': {'row': 9,  'col': 2,  'label': '单个零件描述'},
    'single_part_no': {'row': 9,  'col': 10, 'label': '单个零件号'},
    'release_date':   {'row': 11, 'col': 2,  'label': '发布日期'},
    'doc_no':         {'row': 11, 'col': 7,  'label': '文件号'},
    'workstation':    {'row': 11, 'col': 10, 'label': '工作站'},
    'name_dept':      {'row': 12, 'col': 2,  'label': '姓名/部门'},
    'plan_date':      {'row': 12, 'col': 7,  'label': '计划日期'},
}

TQM013_CONTENT_COLS = {
    'content_c':      {'col': 3,  'label': '内容 C'},
    'content_d':      {'col': 4,  'label': '内容 D'},
    'description_1':  {'col': 5,  'label': '描述 1'},
    'description_2':  {'col': 6,  'label': '描述 2'},
    'description_3':  {'col': 7,  'label': '描述 3'},
    'description_4':  {'col': 8,  'label': '描述 4'},
    'description_5':  {'col': 9,  'label': '描述 5'},
    'description_6':  {'col': 10, 'label': '描述 6'},
    'test_level_eq':  {'col': 12, 'label': '试验等级/设备'},
    'responsible':    {'col': 14, 'label': '负责人'},
}

CONTENT_START_ROW = 18
CONTENT_END_ROW = 48

# 表头字段的自动匹配关键词
HEADER_MATCH_RULES = {
    'project':        ['项目', 'project', 'projekt', '项目名称', '项目号'],
    'sub_project':    ['sub-project', '子项目', 'sub project', '分项目'],
    'oem':            ['oem', '客户', 'customer', 'kunde', '顾客'],
    'part_no_oem':    ['零件号', 'part no', 'part number', 'teilnummer', 'oem零件号', '客户零件号', '图号', '图纸号', 'drawing'],
    'dir':            ['dir', '方向', 'direction', '位置'],
    'part_name':      ['零件名称', 'part name', '产品名称', 'product name', '名称', '部件名称'],
    'model_type':     ['模型', '种类', 'model', 'type', '型号', '规格型号', '车型'],
    'bat_material':   ['bat', '物料', 'material', '材料', '物料号', '材料号', 'bom'],
    'product_group':  ['产品组', 'product group', '组别', '分组'],
    'single_part_desc': ['单个零件', '零件描述', 'part description', '描述', '说明'],
    'single_part_no': ['单个零件号', '零件编号', 'part id', '编号'],
    'release_date':   ['发布日期', 'release date', '版本日期', '生效日期', 'date'],
    'doc_no':         ['文件号', 'doc no', 'document', '文档号', '编号', '文件编号'],
    'name_dept':      ['姓名', '部门', 'name', 'department', '编制', '创建人', 'author'],
    'plan_date':      ['计划日期', 'plan date', '编制日期', '创建日期'],
}

# 内容列的自动匹配关键词
CONTENT_MATCH_RULES = {
    'content_c':      ['内容', 'content', 'inhalt', '工序', 'process', 'vorgang', 'nr', '序号', '步骤', 'step'],
    'content_d':      ['d', '分类', 'class', 'klasse', '类别', '类型'],
    'description_1':  ['描述', 'description', 'beschreibung', '要求', 'requirement', '说明', '特征', '特性', '检查项目', '检验项目'],
    'description_2':  ['规格', 'spec', 'spezifikation', '值', 'value', '标准值', '公差', '规格值', '尺寸'],
    'description_3':  ['方法', 'method', 'methode', '测量方法', '检验方法', '检测方法'],
    'description_4':  ['备注', 'remark', 'bemerkung', '注释', 'note', '说明'],
    'description_5':  ['参考', 'reference', 'referenz', '文件', 'document', '标准'],
    'description_6':  ['标准', 'standard', 'norm', '规范', '评价', '判定'],
    'test_level_eq':  ['试验', '设备', 'test', 'equipment', 'prüf', 'messmittel', '检测', '测量', '量具', '检具', '仪器', '工具'],
    'responsible':    ['负责人', 'responsible', 'verantwortlich', '责任', '检验人', '执行人', '部门', '频率'],
}


# ==================== 核心解析函数 ====================

def parse_control_plan(filepath, user_ws_col=None):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    # 自动检测表头行
    header_row = 1
    max_cells = 0
    for row_idx in range(1, min(31, ws.max_row or 100)):
        non_empty = 0
        for col_idx in range(1, (ws.max_column or 1) + 1):
            if ws.cell(row=row_idx, column=col_idx).value is not None:
                non_empty += 1
        if non_empty > max_cells:
            max_cells = non_empty
            header_row = row_idx

    all_headers = {}
    for col_idx in range(1, (ws.max_column or 1) + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        all_headers[col_idx] = str(cell.value).strip() if cell.value else ''

    # 自动检测工位列
    workstation_col = None
    if user_ws_col is not None and user_ws_col in all_headers:
        workstation_col = user_ws_col
    else:
        ws_keywords = [
            '工位', 'op', '工作站', 'station', 'arbeitsplatz',
            '工序', 'process', 'vorgang', '操作', 'operation',
            '岗位', 'platz', '工作', 'work', '编号', '序号',
            'station nr', 'station no', 'op-nr', 'op nr',
        ]
        for col_idx, header in all_headers.items():
            header_lower = header.lower()
            if any(kw in header_lower for kw in ws_keywords):
                workstation_col = col_idx
                break

    # 解析数据行
    all_data = []
    workstations = []
    for row_idx in range(header_row + 1, (ws.max_row or header_row + 1) + 1):
        row_data = {}
        has_data = False
        for col_idx in range(1, (ws.max_column or 1) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            row_data[col_idx] = cell.value
            if cell.value is not None and str(cell.value).strip() != '':
                has_data = True
        if has_data:
            all_data.append(row_data)
            if workstation_col and row_data.get(workstation_col) is not None:
                ws_name = str(row_data[workstation_col]).strip()
                if ws_name and ws_name not in workstations:
                    workstations.append(ws_name)

    # ★ 提取第一行数据作为"基本信息"候选（通常 CP 前面几行是产品信息）
    header_info = {}
    if all_data:
        first_row = all_data[0]
        for col_idx, val in first_row.items():
            if val is not None and str(val).strip():
                header_info[col_idx] = str(val).strip()

    wb.close()
    return {
        'headers': all_headers,
        'workstations': workstations,
        'workstation_col': workstation_col,
        'data': all_data,
        'header_row': header_row,
        'total_columns': (ws.max_column or 1),
        'header_info': header_info,
    }


def auto_match_headers(cp_headers):
    """自动匹配 CP 表头 → T.QM.013 表头字段"""
    mapping = {}
    used = set()
    for field, keywords in HEADER_MATCH_RULES.items():
        for col_idx, header in cp_headers.items():
            if col_idx in used:
                continue
            header_lower = header.lower()
            if any(kw in header_lower for kw in keywords):
                mapping[field] = col_idx
                used.add(col_idx)
                break
    return mapping


def auto_match_content(cp_headers):
    """自动匹配 CP 表头 → T.QM.013 内容列"""
    mapping = {}
    used = set()
    for field, keywords in CONTENT_MATCH_RULES.items():
        for col_idx, header in cp_headers.items():
            if col_idx in used:
                continue
            header_lower = header.lower()
            if any(kw in header_lower for kw in keywords):
                mapping[field] = col_idx
                used.add(col_idx)
                break
    return mapping


def fill_template(cp_data, selected_ws, header_mapping, content_mapping):
    """基于 T.QM.013 模板填充数据"""
    if TEMPLATE_FILE is None:
        raise FileNotFoundError("未找到 T.QM.013 模板文件！")

    wb = openpyxl.load_workbook(TEMPLATE_FILE, keep_vba=True)
    ws = wb.active

    # 1. 填充表头默认值
    for field, config in TQM013_HEADER_FIELDS.items():
        if 'default' in config:
            ws.cell(row=config['row'], column=config['col'], value=config['default'])

    # 2. 从 CP 数据中提取表头信息（使用第一行数据）
    if cp_data['data'] and header_mapping:
        first_row = cp_data['data'][0]
        for field, cp_col in header_mapping.items():
            if field in TQM013_HEADER_FIELDS and cp_col in first_row:
                val = first_row[cp_col]
                if val is not None:
                    config = TQM013_HEADER_FIELDS[field]
                    ws.cell(row=config['row'], column=config['col'], value=val)

    # 3. 填充工作站
    ws.cell(row=TQM013_HEADER_FIELDS['workstation']['row'],
            column=TQM013_HEADER_FIELDS['workstation']['col'],
            value=selected_ws)

    # 4. 筛选匹配工位的数据行
    ws_col = cp_data['workstation_col']
    matched = []
    for row_data in cp_data['data']:
        if ws_col and row_data.get(ws_col) is not None:
            cell_val = str(row_data[ws_col]).strip()
            if cell_val == str(selected_ws).strip():
                matched.append(row_data)
            elif str(selected_ws).strip().lower() in cell_val.lower():
                matched.append(row_data)

    # 去重
    seen = set()
    unique_matched = []
    for row in matched:
        key = tuple(str(row.get(k, '')) for k in sorted(row.keys()))
        if key not in seen:
            seen.add(key)
            unique_matched.append(row)

    # 5. 填充内容数据
    current_row = CONTENT_START_ROW
    for row_data in unique_matched:
        if current_row > CONTENT_END_ROW:
            break
        for field, cp_col in content_mapping.items():
            if field in TQM013_CONTENT_COLS:
                col = TQM013_CONTENT_COLS[field]['col']
                value = row_data.get(cp_col)
                if value is not None:
                    ws.cell(row=current_row, column=col, value=value)
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
        cp_data = parse_control_plan(filepath)
        session['cp_filepath'] = filepath
        session['cp_session_id'] = session_id

        # 自动匹配表头和内容列
        header_match = auto_match_headers(cp_data['headers'])
        content_match = auto_match_content(cp_data['headers'])

        return jsonify({
            'success': True,
            'workstations': cp_data['workstations'],
            'headers': {str(k): v for k, v in cp_data['headers'].items()},
            'total_rows': len(cp_data['data']),
            'workstation_column': cp_data['workstation_col'],
            'session_id': session_id,
            'header_row': cp_data['header_row'],
            'total_columns': cp_data['total_columns'],
            'header_match': {k: v for k, v in header_match.items()},
            'content_match': {k: v for k, v in content_match.items()},
            'header_info': {str(k): v for k, v in cp_data['header_info'].items()},
        })
    except Exception as e:
        return jsonify({'success': False, 'error': f'解析失败: {str(e)}'})


@app.route('/api/reparse_with_column', methods=['POST'])
def reparse_with_column():
    data = request.get_json()
    session_id = data.get('session_id')
    ws_col = data.get('workstation_column')

    cp_filepath = session.get('cp_filepath')
    if not cp_filepath or not os.path.exists(cp_filepath):
        return jsonify({'success': False, 'error': '会话已过期'})

    try:
        cp_data = parse_control_plan(cp_filepath, user_ws_col=int(ws_col))
        header_match = auto_match_headers(cp_data['headers'])
        content_match = auto_match_content(cp_data['headers'])

        return jsonify({
            'success': True,
            'workstations': cp_data['workstations'],
            'headers': {str(k): v for k, v in cp_data['headers'].items()},
            'total_rows': len(cp_data['data']),
            'workstation_column': cp_data['workstation_col'],
            'header_match': {k: v for k, v in header_match.items()},
            'content_match': {k: v for k, v in content_match.items()},
        })
    except Exception as e:
        return jsonify({'success': False, 'error': f'重新解析失败: {str(e)}'})


@app.route('/api/generate', methods=['POST'])
def generate():
    data = request.get_json()
    session_id = data.get('session_id')
    workstation = data.get('workstation')
    header_mapping = data.get('header_mapping', {})
    content_mapping = data.get('content_mapping', {})

    if not session_id or not workstation:
        return jsonify({'success': False, 'error': '缺少必要参数'})

    cp_filepath = session.get('cp_filepath')
    if not cp_filepath or not os.path.exists(cp_filepath):
        return jsonify({'success': False, 'error': '会话已过期'})

    try:
        cp_data = parse_control_plan(cp_filepath)
        hm = {k: int(v) for k, v in header_mapping.items() if v}
        cm = {k: int(v) for k, v in content_mapping.items() if v}

        safe_ws = secure_filename(str(workstation))
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"T.QM.013_{safe_ws}_{timestamp}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        output = fill_template(cp_data, workstation, hm, cm)
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
    if sid:
        for d in [UPLOAD_DIR, OUTPUT_DIR]:
            try:
                for f in os.listdir(d):
                    if f.startswith(sid):
                        os.remove(os.path.join(d, f))
            except OSError:
                pass
    return jsonify({'success': True})


# ==================== 内嵌 HTML（两栏布局 + 右侧预览） ====================
HTML_PAGE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>T.QM.013 检验指导书生成器 v2.0</title>
<style>
:root { --primary: #2563eb; --primary-hover: #1d4ed8; --bg: #f1f5f9; --card-bg: #ffffff; --border: #e2e8f0; --text: #1e293b; --text-secondary: #64748b; --success: #16a34a; --warning: #ea580c; --danger: #dc2626; --highlight: #dbeafe; --highlight-border: #93c5fd; --radius: 8px; --shadow: 0 1px 3px rgba(0,0,0,0.08); }
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif; background: var(--bg); color: var(--text); line-height:1.6; height:100vh; overflow:hidden; }
.header { background: var(--primary); color: white; padding: 10px 24px; display: flex; align-items: center; justify-content: space-between; flex-shrink:0; }
.header h1 { font-size: 1.15rem; font-weight: 600; }
.header .status { font-size: 0.75rem; opacity: 0.9; }

.main-layout { display: flex; height: calc(100vh - 52px); }
.left-panel { flex: 1; overflow-y: auto; padding: 16px 20px; min-width: 480px; }
.right-panel { flex: 1; overflow-y: auto; padding: 16px 20px; background: #f8fafc; border-left: 1px solid var(--border); min-width: 420px; }

.steps { display: flex; gap: 0; margin-bottom: 16px; background: var(--card-bg); border-radius: var(--radius); box-shadow: var(--shadow); overflow: hidden; }
.step { flex: 1; padding: 10px 12px; text-align: center; font-size: 0.78rem; color: var(--text-secondary); background: var(--card-bg); transition: all 0.3s; }
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
.btn-outline:hover { background: #eff6ff; }
.btn-success { background: var(--success); color: white; }
.btn-warning { background: var(--warning); color: white; }
.btn-sm { padding: 4px 10px; font-size: 0.75rem; }
.btn-group { display: flex; gap: 8px; margin-top: 10px; flex-wrap: wrap; }

.ws-search { width: 100%; padding: 8px 12px; border: 1.5px solid var(--border); border-radius: 6px; font-size: 0.85rem; margin-bottom: 8px; }
.ws-search:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(37,99,235,0.1); }
.ws-list { display: flex; flex-wrap: wrap; gap: 6px; max-height: 160px; overflow-y: auto; padding: 2px; }
.ws-chip { padding: 5px 12px; border: 1.5px solid var(--border); border-radius: 16px; cursor: pointer; font-size: 0.8rem; transition: all 0.2s; user-select: none; }
.ws-chip:hover { border-color: var(--primary); background: #eff6ff; }
.ws-chip.selected { background: var(--primary); color: white; border-color: var(--primary); }

.map-section { margin-bottom: 10px; }
.map-section h3 { margin-bottom: 6px; font-size: 0.82rem; }
.map-row { display: flex; align-items: center; gap: 8px; margin-bottom: 4px; font-size: 0.8rem; }
.map-row label { min-width: 110px; font-weight: 500; color: var(--text); text-align: right; }
.map-row select { flex: 1; padding: 4px 8px; border: 1.5px solid var(--border); border-radius: 4px; font-size: 0.78rem; }
.map-row .match-badge { font-size: 0.65rem; padding: 2px 6px; border-radius: 10px; min-width: 50px; text-align: center; }
.match-auto { background: #dcfce7; color: #166534; }
.match-manual { background: #fef3c7; color: #92400e; }
.match-none { background: #fee2e2; color: #991b1b; }

.alert { padding: 8px 12px; border-radius: 6px; margin-bottom: 10px; font-size: 0.8rem; }
.alert-error { background: #fef2f2; color: var(--danger); border: 1px solid #fecaca; }
.alert-info { background: #eff6ff; color: var(--primary); border: 1px solid #bfdbfe; }
.alert-warning { background: #fff7ed; color: var(--warning); border: 1px solid #fed7aa; }
.hidden { display: none !important; }

.spinner { display: none; width: 16px; height: 16px; border: 2px solid var(--border); border-top-color: var(--primary); border-radius: 50%; animation: spin 0.6s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }

.column-selector { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; margin-bottom: 10px; padding: 10px; background: #fffbeb; border-radius: 6px; border: 1px solid #fde68a; font-size: 0.8rem; }
.column-selector select { padding: 6px 10px; border: 1.5px solid var(--border); border-radius: 6px; font-size: 0.82rem; }

/* 右侧预览面板 */
.preview-title { font-size: 0.95rem; font-weight: 600; margin-bottom: 10px; display: flex; align-items: center; gap: 8px; }
.preview-hint { font-size: 0.7rem; color: var(--text-secondary); font-weight: 400; }
.preview-table { width: 100%; border-collapse: collapse; font-size: 0.68rem; background: white; }
.preview-table td, .preview-table th { border: 1px solid #d1d5db; padding: 2px 4px; min-width: 24px; height: 18px; text-align: center; }
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
</style>
</head>
<body>

<div class="header">
    <h1>📋 T.QM.013 检验指导书生成器 v2.0</h1>
    <span class="status" id="statusBar">检查模板...</span>
</div>

<div class="main-layout">
    <!-- ====== 左侧：操作面板 ====== -->
    <div class="left-panel" id="leftPanel">

        <div class="steps" id="stepIndicator">
            <div class="step active" data-step="1"><span class="step-num">1</span> 上传CP</div>
            <div class="step" data-step="2"><span class="step-num">2</span> 选工位</div>
            <div class="step" data-step="3"><span class="step-num">3</span> 映射</div>
            <div class="step" data-step="4"><span class="step-num">4</span> 下载</div>
        </div>

        <div id="globalAlert" class="hidden"></div>

        <!-- Step 1 -->
        <div class="card" id="step1Card">
            <h2>📁 步骤 1：上传版本控制计划</h2>
            <div class="upload-zone" id="uploadZone">
                <div class="upload-icon">📤</div>
                <p>拖拽 Excel 文件到此处，或 <span class="browse-link" id="browseLink">点击浏览</span></p>
                <p style="font-size:0.7rem;margin-top:3px;">支持 .xlsx / .xlsm / .xls</p>
                <input type="file" id="fileInput" accept=".xlsx,.xlsm,.xls">
            </div>
            <div class="file-info" id="fileInfo">✅ <span id="fileName"></span> — <span id="fileStats"></span></div>
            <div class="btn-group" style="justify-content:flex-end;"><span class="spinner" id="uploadSpinner"></span></div>
        </div>

        <!-- Step 2 -->
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

        <!-- Step 3 -->
        <div class="card hidden" id="step3Card">
            <h2>🔗 步骤 3：确认列映射关系</h2>

            <div class="tab-bar">
                <button class="tab-btn active" data-tab="header">📋 基本信息映射</button>
                <button class="tab-btn" data-tab="content">📝 内容列映射</button>
            </div>

            <!-- 基本信息映射 -->
            <div id="tabHeader" class="map-section">
                <h3>将 CP 列映射到 T.QM.013 表头字段</h3>
                <div id="headerMappingRows"></div>
            </div>

            <!-- 内容列映射 -->
            <div id="tabContent" class="map-section hidden">
                <h3>将 CP 列映射到 T.QM.013 内容区域</h3>
                <div id="contentMappingRows"></div>
            </div>

            <div class="btn-group" style="justify-content:space-between;">
                <button class="btn btn-outline btn-sm" id="btnBackToStep2">← 返回选工位</button>
                <button class="btn btn-primary" id="btnGenerate">🚀 生成检验指导书 <span class="spinner" id="generateSpinner"></span></button>
            </div>
        </div>

        <!-- Step 4 -->
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

    <!-- ====== 右侧：预览面板 ====== -->
    <div class="right-panel" id="rightPanel">
        <div class="preview-title">
            📐 T.QM.013 模板预览
            <span class="preview-hint">（蓝色=已填充，黄色=内容区，白色=空白）</span>
        </div>
        <div class="preview-legend">
            <span><span class="legend-filled"></span> 表头已映射</span>
            <span><span class="legend-will"></span> 内容数据区</span>
            <span><span class="legend-empty"></span> 空白/未映射</span>
        </div>
        <div id="previewContent">
            <p style="color:var(--text-secondary);font-size:0.8rem;">上传控制计划后将在此显示模板预览</p>
        </div>
    </div>
</div>

<script>
const $=s=>document.querySelector(s),$$=s=>document.querySelectorAll(s);
const STATE={
    sessionId:null,workstations:[],selectedWorkstation:null,
    allCpHeaders:{},headerMatch:{},contentMatch:{},
    headerMapping:{},contentMapping:{},downloadUrl:null,currentStep:1,
    cpData:null
};

// T.QM.013 模板结构定义（用于预览）
const TMPL_HEADER = [
    {r:5, c:1, w:4, t:'操作和试验指导书', cls:'label-cell'},
    {r:5, c:5, w:1, t:'项目:', cls:'label-cell'},
    {r:5, c:6, w:2, t:'', f:'project'},
    {r:5, c:8, w:1, t:'Sub-Project:', cls:'label-cell'},
    {r:5, c:9, w:2, t:'', f:'sub_project'},
    {r:5, c:11, w:1, t:'语言:', cls:'label-cell'},
    {r:5, c:12, w:2, t:'中文', f:'language'},
    {r:7, c:1, w:1, t:'OEM:', cls:'label-cell'},
    {r:7, c:2, w:2, t:'', f:'oem'},
    {r:7, c:5, w:1, t:'零件号:', cls:'label-cell'},
    {r:7, c:6, w:2, t:'', f:'part_no_oem'},
    {r:7, c:8, w:1, t:'DIR:', cls:'label-cell'},
    {r:7, c:9, w:2, t:'', f:'dir'},
    {r:7, c:12, w:1, t:'DmbA:', cls:'label-cell'},
    {r:7, c:13, w:1, t:'A', f:'dmba'},
    {r:7, c:15, w:1, t:'页:', cls:'label-cell'},
    {r:7, c:16, w:1, t:'2', f:'page'},
    {r:8, c:1, w:1, t:'零件名称:', cls:'label-cell'},
    {r:8, c:2, w:2, t:'', f:'part_name'},
    {r:8, c:5, w:1, t:'模型/种类:', cls:'label-cell'},
    {r:8, c:6, w:2, t:'', f:'model_type'},
    {r:8, c:8, w:1, t:'BAT-物料:', cls:'label-cell'},
    {r:8, c:9, w:2, t:'', f:'bat_material'},
    {r:8, c:12, w:1, t:'产品组:', cls:'label-cell'},
    {r:8, c:13, w:2, t:'', f:'product_group'},
    {r:8, c:15, w:1, t:'指导:', cls:'label-cell'},
    {r:8, c:16, w:1, t:'ON', f:'instruction'},
    {r:9, c:1, w:4, t:'单个零件描述', cls:'label-cell'},
    {r:9, c:9, w:1, t:'单个零件号:', cls:'label-cell'},
    {r:9, c:10, w:2, t:'', f:'single_part_no'},
    {r:11, c:1, w:2, t:'发布日期:', cls:'label-cell'},
    {r:11, c:5, w:1, t:'文件号:', cls:'label-cell'},
    {r:11, c:6, w:2, t:'', f:'doc_no'},
    {r:11, c:9, w:1, t:'工作站:', cls:'label-cell'},
    {r:11, c:10, w:2, t:'', f:'workstation'},
    {r:12, c:1, w:2, t:'姓名/部门:', cls:'label-cell'},
    {r:12, c:5, w:1, t:'计划日期:', cls:'label-cell'},
    {r:12, c:6, w:2, t:'', f:'plan_date'},
];
const TMPL_CONTENT_COLS = [
    {c:3, label:'C'},
    {c:4, label:'D'},
    {c:5, label:'描述'},
    {c:6, label:'描述'},
    {c:7, label:'描述'},
    {c:8, label:'描述'},
    {c:9, label:'描述'},
    {c:10, label:'描述'},
    {c:12, label:'试验等级/设备'},
    {c:14, label:'负责人'},
];
const CONTENT_HDR_ROW = 17;
const CONTENT_START = 18;
const CONTENT_END = 30; // 预览只显示前13行内容

function setStep(s){
    STATE.currentStep=s;
    $$('.step').forEach((el,i)=>{el.classList.remove('active','done');if(i+1<s)el.classList.add('done');if(i+1===s)el.classList.add('active');});
    ['step1Card','step2Card','step3Card','step4Card'].forEach((id,i)=>document.getElementById(id).classList.toggle('hidden',i+1!==s));
}
function showAlert(msg,type='error'){
    const el=$('#globalAlert');el.className='alert alert-'+type;el.textContent=msg;el.classList.remove('hidden');
    setTimeout(()=>el.classList.add('hidden'),6000);
}

// 检查模板
fetch('/api/status').then(r=>r.json()).then(d=>{
    $('#statusBar').textContent=d.template_found?'✅ 模板已加载':'⚠️ 模板未找到';
    $('#statusBar').style.color=d.template_found?'#bbf7d0':'#fecaca';
});

// ===== 上传 =====
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
        STATE.allCpHeaders=d.headers||{};STATE.headerMatch=d.header_match||{};
        STATE.contentMatch=d.content_match||{};
        STATE.headerMapping={...STATE.headerMatch};
        STATE.contentMapping={...STATE.contentMatch};

        $('#fileName').textContent=file.name;
        $('#fileStats').textContent='共 '+d.total_rows+' 行，'+d.workstations.length+' 个工位';
        $('#fileInfo').classList.add('show');

        showHeaderPreview(d.headers,d.header_row);

        if(!d.workstation_column||STATE.workstations.length===0){
            showAlert('未自动识别工位列，请手动选择','warning');
            showManualColumnSelector(d.headers);
        }else{$('#manualColumnBox').classList.add('hidden');}

        renderWorkstations();
        updatePreview();
        setStep(2);
    }catch(e){showAlert('网络错误: '+e.message);}
    finally{uploadSpinner.style.display='none';}
}

function showHeaderPreview(headers,headerRow){
    $('#headerRowNum').textContent=headerRow||'?';
    let h='';
    for(const [col,hdr] of Object.entries(headers)){
        h+='<span style="margin:2px;display:inline-block;background:#e2e8f0;padding:1px 5px;border-radius:3px;">列'+col+': '+(hdr||'<i>空</i>')+'</span>';
    }
    $('#headerPreview').innerHTML=h;
    $('#headerPreviewBox').classList.remove('hidden');
}

function showManualColumnSelector(headers){
    const sel=$('#wsColumnSelect');
    sel.innerHTML='<option value="">-- 请选择工位/OP 所在列 --</option>';
    for(const [col,hdr] of Object.entries(headers)){
        sel.innerHTML+='<option value="'+col+'">列 '+col+': '+(hdr||'(空)')+'</option>';
    }
    $('#manualColumnBox').classList.remove('hidden');
}

$('#btnReparse').addEventListener('click',async()=>{
    const wsCol=$('#wsColumnSelect').value;
    if(!wsCol){showAlert('请先选择列','warning');return;}
    const sp=$('#reparseSpinner');sp.style.display='inline-block';
    try{
        const r=await fetch('/api/reparse_with_column',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({session_id:STATE.sessionId,workstation_column:wsCol})});
        const d=await r.json();
        if(!d.success){showAlert(d.error);return;}
        STATE.workstations=d.workstations||[];STATE.allCpHeaders=d.headers||{};
        STATE.headerMatch=d.header_match||{};STATE.contentMatch=d.content_match||{};
        STATE.headerMapping={...STATE.headerMatch};STATE.contentMapping={...STATE.contentMatch};
        STATE.selectedWorkstation=null;$('#selectedWsDisplay').textContent='未选择';$('#btnNextToMapping').disabled=true;
        renderWorkstations();updatePreview();
        showAlert('✅ 重新解析成功！'+STATE.workstations.length+' 个工位','info');
    }catch(e){showAlert('网络错误: '+e.message);}
    finally{sp.style.display='none';}
});

// ===== 工位选择 =====
function renderWorkstations(f){
    f=(f||'').toLowerCase();
    const filtered=(STATE.workstations||[]).filter(w=>String(w).toLowerCase().includes(f));
    $('#wsCount').textContent=STATE.workstations.length;
    if(filtered.length===0){
        $('#wsList').innerHTML='<p style="color:var(--text-secondary);padding:12px;font-size:0.8rem;">'+(STATE.workstations.length===0?'未找到工位，请手动选择工位列':'无匹配')+'</p>';
        return;
    }
    $('#wsList').innerHTML=filtered.map(w=>'<span class="ws-chip'+(w===STATE.selectedWorkstation?' selected':'')+'" data-ws="'+w.replace(/"/g,'&quot;')+'">'+w+'</span>').join('');
    $('#wsList').querySelectorAll('.ws-chip').forEach(c=>c.addEventListener('click',()=>{
        STATE.selectedWorkstation=c.dataset.ws;$('#selectedWsDisplay').textContent=STATE.selectedWorkstation;
        $('#btnNextToMapping').disabled=false;updatePreview();renderWorkstations($('#wsSearch').value);
    }));
}
$('#wsSearch').addEventListener('input',e=>renderWorkstations(e.target.value));
$('#btnNextToMapping').addEventListener('click',()=>{if(!STATE.selectedWorkstation)return;buildMappingUI();setStep(3);});

// ===== 映射界面 =====
function buildMappingUI(){
    const allHeaders=STATE.allCpHeaders;
    const opts=['<option value="">-- 不映射 --</option>'];
    for(const [col,hdr] of Object.entries(allHeaders)){
        opts.push('<option value="'+col+'">列 '+col+': '+(hdr||'(空)')+'</option>');
    }
    const optsHtml=opts.join('');

    // 表头映射
    const hdrLabels={
        project:'项目',sub_project:'Sub-Project',oem:'OEM',part_no_oem:'零件号 OEM',
        dir:'DIR',part_name:'零件名称',model_type:'模型/种类',bat_material:'BAT-物料',
        product_group:'产品组',single_part_desc:'单个零件描述',single_part_no:'单个零件号',
        release_date:'发布日期',doc_no:'文件号',name_dept:'姓名/部门',plan_date:'计划日期'
    };
    let hRows='';
    for(const [f,label] of Object.entries(hdrLabels)){
        const curVal=STATE.headerMapping[f]||'';
        const badge=curVal?'<span class="match-badge match-auto">✅</span>':'<span class="match-badge match-none">—</span>';
        hRows+='<div class="map-row"><label>'+label+'</label><select data-field="'+f+'" data-type="header" class="map-select">'+optsHtml+'</select>'+badge+'</div>';
    }
    $('#headerMappingRows').innerHTML=hRows;

    // 内容映射
    const cntLabels={
        content_c:'内容 C',content_d:'内容 D',description_1:'描述 1',description_2:'描述 2',
        description_3:'描述 3',description_4:'描述 4',description_5:'描述 5',description_6:'描述 6',
        test_level_eq:'试验等级/设备',responsible:'负责人'
    };
    let cRows='';
    for(const [f,label] of Object.entries(cntLabels)){
        const curVal=STATE.contentMapping[f]||'';
        const badge=curVal?'<span class="match-badge match-auto">✅</span>':'<span class="match-badge match-none">—</span>';
        cRows+='<div class="map-row"><label>'+label+'</label><select data-field="'+f+'" data-type="content" class="map-select">'+optsHtml+'</select>'+badge+'</div>';
    }
    $('#contentMappingRows').innerHTML=cRows;

    // 设置 select 当前值
    $$('.map-select').forEach(sel=>{
        const type=sel.dataset.type;
        const field=sel.dataset.field;
        const curMap=type==='header'?STATE.headerMapping:STATE.contentMapping;
        sel.value=curMap[field]||'';
        sel.addEventListener('change',()=>{
            if(type==='header'){
                STATE.headerMapping[field]=sel.value?parseInt(sel.value):null;
                // 更新badge
                const badge=sel.parentElement.querySelector('.match-badge');
                if(badge){badge.className='match-badge '+(sel.value?'match-manual':'match-none');badge.textContent=sel.value?'✏️':'—';}
            }else{
                STATE.contentMapping[field]=sel.value?parseInt(sel.value):null;
                const badge=sel.parentElement.querySelector('.match-badge');
                if(badge){badge.className='match-badge '+(sel.value?'match-manual':'match-none');badge.textContent=sel.value?'✏️':'—';}
            }
            updatePreview();
        });
    });
}

// Tab切换
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
    // 清理 mapping（去掉 null 值）
    const hm={};for(const[k,v]of Object.entries(STATE.headerMapping)){if(v)hm[k]=v;}
    const cm={};for(const[k,v]of Object.entries(STATE.contentMapping)){if(v)cm[k]=v;}
    try{
        const r=await fetch('/api/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({session_id:STATE.sessionId,workstation:STATE.selectedWorkstation,header_mapping:hm,content_mapping:cm})});
        const d=await r.json();
        if(!d.success){showAlert(d.error);return;}
        STATE.downloadUrl=d.download_url;$('#resultFilename').textContent='文件名：'+d.filename;setStep(4);
    }catch(e){showAlert('网络错误: '+e.message);}
    finally{sp.style.display='none';btn.disabled=false;}
});
$('#btnDownload').addEventListener('click',()=>{if(STATE.downloadUrl)window.open(STATE.downloadUrl,'_blank');});

// ===== 右侧预览 =====
function updatePreview(){
    const hm=STATE.headerMapping||{};
    let html='<table class="preview-table">';

    // 生成预览网格（16列，行1-35）
    const filledFields={};
    for(const [f,cpCol] of Object.entries(hm)){
        if(f==='workstation')filledFields[f]=STATE.selectedWorkstation||'(待选)';
        else filledFields[f]='(已映射)';
    }
    // 默认值
    filledFields['language']='中文';filledFields['dmba']='A';filledFields['instruction']='ON';
    if(STATE.selectedWorkstation)filledFields['workstation']=STATE.selectedWorkstation;

    for(let row=1;row<=35;row++){
        html+='<tr><td class="label-cell" style="font-size:0.6rem;">'+row+'</td>';

        if(row===5){
            html+='<td class="label-cell" colspan="4">操作和试验指导书</td><td class="label-cell">项目:</td>';
            html+='<td class="'+(filledFields['project']?'filled':'empty')+'" colspan="2">'+ (filledFields['project']||'')+'</td>';
            html+='<td class="label-cell">Sub-Project:</td><td class="'+(filledFields['sub_project']?'filled':'empty')+'" colspan="2">'+ (filledFields['sub_project']||'')+'</td>';
            html+='<td class="label-cell">语言:</td><td class="filled" colspan="2">中文</td><td></td><td></td><td></td>';
        }else if(row===7){
            html+='<td class="label-cell">OEM:</td><td class="'+(filledFields['oem']?'filled':'empty')+'" colspan="2">'+ (filledFields['oem']||'')+'</td>';
            html+='<td></td><td class="label-cell">零件号:</td><td class="'+(filledFields['part_no_oem']?'filled':'empty')+'" colspan="2">'+ (filledFields['part_no_oem']||'')+'</td>';
            html+='<td class="label-cell">DIR:</td><td class="'+(filledFields['dir']?'filled':'empty')+'" colspan="2">'+ (filledFields['dir']||'')+'</td>';
            html+='<td></td><td class="label-cell">DmbA:</td><td class="filled">A</td><td class="label-cell">页</td><td>2</td>';
        }else if(row===8){
            html+='<td class="label-cell">零件名称:</td><td class="'+(filledFields['part_name']?'filled':'empty')+'" colspan="2">'+ (filledFields['part_name']||'')+'</td>';
            html+='<td></td><td class="label-cell">模型/种类:</td><td class="'+(filledFields['model_type']?'filled':'empty')+'" colspan="2">'+ (filledFields['model_type']||'')+'</td>';
            html+='<td class="label-cell">BAT-物料:</td><td class="'+(filledFields['bat_material']?'filled':'empty')+'" colspan="2">'+ (filledFields['bat_material']||'')+'</td>';
            html+='<td></td><td class="label-cell">产品组:</td><td class="'+(filledFields['product_group']?'filled':'empty')+'" colspan="2">'+ (filledFields['product_group']||'')+'</td>';
            html+='<td class="label-cell">指导:</td><td class="filled">ON</td>';
        }else if(row===9){
            html+='<td class="label-cell" colspan="4">单个零件描述</td>';
            html+='<td></td><td></td><td></td><td></td>';
            html+='<td class="label-cell">单个零件号:</td><td class="'+(filledFields['single_part_no']?'filled':'empty')+'" colspan="2">'+ (filledFields['single_part_no']||'')+'</td>';
            html+='<td></td><td></td><td></td><td></td><td></td>';
        }else if(row===11){
            html+='<td class="label-cell" colspan="2">发布日期:</td><td></td><td></td>';
            html+='<td class="label-cell">文件号:</td><td class="'+(filledFields['doc_no']?'filled':'empty')+'" colspan="2">'+ (filledFields['doc_no']||'')+'</td>';
            html+='<td></td><td class="label-cell">工作站:</td><td class="'+(filledFields['workstation']?'filled':'empty')+'" colspan="2">'+ (filledFields['workstation']||'')+'</td>';
            html+='<td></td><td></td><td></td><td></td>';
        }else if(row===12){
            html+='<td class="label-cell" colspan="2">姓名/部门:</td><td></td><td></td>';
            html+='<td class="label-cell">计划日期:</td><td class="'+(filledFields['plan_date']?'filled':'empty')+'" colspan="3">'+ (filledFields['plan_date']||'')+'</td>';
            html+='<td></td><td></td><td></td><td></td><td></td><td></td><td></td>';
        }else if(row===17){
            // 内容表头
            html+='<td class="col-header">内容</td><td class="col-header">C</td><td class="col-header">D</td>';
            for(let i=0;i<6;i++)html+='<td class="col-header">描述</td>';
            html+='<td></td><td class="col-header">试验等级/设备</td><td></td><td class="col-header">负责人</td><td></td><td></td>';
        }else if(row>=CONTENT_START && row<=CONTENT_END){
            // 内容数据区
            const hasContent=STATE.selectedWorkstation&&Object.keys(STATE.contentMapping||{}).length>0;
            const cls=hasContent?'will-fill':'empty';
            html+='<td class="'+cls+'"></td><td class="'+cls+'"></td><td class="'+cls+'"></td>';
            for(let i=0;i<6;i++)html+='<td class="'+cls+'"></td>';
            html+='<td></td><td class="'+cls+'"></td><td></td><td class="'+cls+'"></td><td></td><td></td>';
        }else if(row>=50){
            html+='<td class="label-cell">生产:</td><td></td><td></td><td></td><td class="label-cell">质量管理:</td><td></td><td></td><td></td><td class="label-cell">替代:</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>';
        }else{
            // 空行
            for(let c=0;c<16;c++)html+='<td class="empty"></td>';
        }
        html+='</tr>';
    }
    html+='</table>';
    $('#previewContent').innerHTML=html;
}

function resetAll(){
    STATE.sessionId=null;STATE.workstations=[];STATE.selectedWorkstation=null;
    STATE.headerMapping={};STATE.contentMapping={};STATE.headerMatch={};STATE.contentMatch={};
    STATE.downloadUrl=null;STATE.allCpHeaders={};
    $('#fileInfo').classList.remove('show');$('#wsList').innerHTML='';$('#selectedWsDisplay').textContent='未选择';
    $('#btnNextToMapping').disabled=true;$('#wsSearch').value='';$('#wsCount').textContent='0';
    $('#headerPreviewBox').classList.add('hidden');$('#manualColumnBox').classList.add('hidden');
    $('#headerMappingRows').innerHTML='';$('#contentMappingRows').innerHTML='';
    $('#previewContent').innerHTML='<p style="color:var(--text-secondary);font-size:0.8rem;">上传控制计划后将在此显示模板预览</p>';
    fileInput.value='';setStep(1);
    fetch('/api/cleanup',{method:'POST'}).catch(()=>{});
}
setStep(1);
</script>
</body>
</html>'''


# ==================== 启动入口 ====================
def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]


def main():
    port = find_free_port()
    url = f'http://127.0.0.1:{port}'

    print("=" * 55)
    print("  T.QM.013 检验指导书生成器 v2.0")
    print("=" * 55)
    print(f"  模板文件: {TEMPLATE_FILE or '❌ 未找到!'}")
    print(f"  本地地址: {url}")
    print("=" * 55)

    threading.Timer(1.5, lambda: webbrowser.open(url)).start()

    import logging
    logging.getLogger('werkzeug').setLevel(logging.WARNING)

    app.run(host='127.0.0.1', port=port, debug=False)


if __name__ == '__main__':
    main()
