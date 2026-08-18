"""
T.QM.013 检验指导书生成器 — 桌面版
PyInstaller 打包后双击即可运行，自动打开浏览器
"""
import os
import sys
import socket
import threading
import webbrowser
import uuid
from datetime import datetime
from io import BytesIO

from flask import Flask, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
import openpyxl


# ==================== PyInstaller 路径处理 ====================
def get_base_path():
    """获取运行时的根目录（兼容 PyInstaller 打包和直接运行）"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后
        return sys._MEIPASS
    else:
        return os.path.dirname(os.path.abspath(__file__))


def get_template_path():
    """查找模板文件"""
    base = get_base_path()
    # 按优先级搜索
    candidates = [
        os.path.join(base, 'template', 'T.QM.013_Template.xlsm'),
        os.path.join(base, 'T.QM.013_Template.xlsm'),
        os.path.join(os.path.dirname(sys.executable), 'T.QM.013_Template.xlsm') if getattr(sys, 'frozen', False) else None,
        os.path.join(os.getcwd(), 'T.QM.013_Template.xlsm'),
    ]
    for path in candidates:
        if path and os.path.exists(path):
            return path

    # 模糊搜索
    for root_dir in [base, os.getcwd()]:
        for f in os.listdir(root_dir):
            if 'T.QM.013' in f and (f.endswith('.xlsm') or f.endswith('.xlsx')):
                return os.path.join(root_dir, f)

    return None


TEMPLATE_FILE = get_template_path()

# ==================== Flask 配置 ====================
app = Flask(__name__)
app.secret_key = str(uuid.uuid4())

# 使用临时目录存储上传和输出文件
import tempfile
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), 'qm013_uploads')
OUTPUT_DIR = os.path.join(tempfile.gettempdir(), 'qm013_outputs')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==================== T.QM.013 模板单元格映射 ====================
TQM013_CELL_MAP = {
    'language':       (5, 13),
    'dmba':           (7, 13),
    'instruction':    (8, 16),
    'workstation':    (11, 10),
    'content_start_row': 18,
    'content_cols': {
        'content_c': 3, 'content_d': 4,
        'description_1': 5, 'description_2': 6, 'description_3': 7,
        'description_4': 8, 'description_5': 9, 'description_6': 10,
        'test_level_eq': 12, 'responsible': 14,
    },
    'content_end_row': 48,
}


def parse_control_plan(filepath):
    """解析控制计划 Excel"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    header_row = 1
    workstation_col = None
    for row in ws.iter_rows(min_row=1, max_row=min(30, ws.max_row or 100)):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip().lower()
                if any(kw in val for kw in ['工位', 'op', '工作站', 'station', 'arbeitsplatz']):
                    header_row = cell.row
                    workstation_col = cell.column
                    break
        if workstation_col:
            break

    headers = {}
    for col_idx in range(1, (ws.max_column or 1) + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        headers[col_idx] = str(cell.value).strip() if cell.value else ''

    all_data = []
    workstations = []
    for row_idx in range(header_row + 1, (ws.max_row or header_row + 1) + 1):
        row_data = {}
        has_data = False
        for col_idx in range(1, (ws.max_column or 1) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            row_data[col_idx] = cell.value
            if cell.value is not None:
                has_data = True
        if has_data:
            all_data.append(row_data)
            if workstation_col and row_data.get(workstation_col):
                ws_name = str(row_data[workstation_col]).strip()
                if ws_name and ws_name not in workstations:
                    workstations.append(ws_name)

    wb.close()
    return {
        'headers': headers,
        'workstations': workstations,
        'workstation_col': workstation_col,
        'data': all_data,
    }


def smart_match_columns(cp_headers):
    rules = {
        'content_c':      ['内容', 'content', 'inhalt', '工序', 'process', 'vorgang', 'nr'],
        'content_d':      ['d', '分类', 'class', 'klasse'],
        'description_1':  ['描述', 'description', 'beschreibung', '要求', 'requirement', '说明', '特征'],
        'description_2':  ['规格', 'spec', 'spezifikation', '值', 'value', '标准值'],
        'description_3':  ['方法', 'method', 'methode', '公差', 'tolerance'],
        'description_4':  ['备注', 'remark', 'bemerkung', '注释', 'note'],
        'description_5':  ['参考', 'reference', 'referenz', '文件', 'document'],
        'description_6':  ['标准', 'standard', 'norm', '规范'],
        'test_level_eq':  ['试验', '设备', 'test', 'equipment', 'prüf', 'messmittel', '检测', '测量', '量具'],
        'responsible':    ['负责人', 'responsible', 'verantwortlich', '责任', '检验人', '执行人', '部门'],
    }
    mapping = {}
    used = set()
    for tqm_col, keywords in rules.items():
        for cp_idx, cp_header in cp_headers.items():
            if cp_idx in used:
                continue
            if any(kw in cp_header.lower() for kw in keywords):
                mapping[tqm_col] = cp_idx
                used.add(cp_idx)
                break
    return mapping


def fill_template(cp_data, selected_ws, column_mapping):
    """基于 T.QM.013 模板填充数据"""
    if TEMPLATE_FILE is None:
        raise FileNotFoundError(
            "未找到 T.QM.013 模板文件！\n"
            "请将模板文件 (.xlsm) 放在与程序相同的目录下。"
        )

    wb = openpyxl.load_workbook(TEMPLATE_FILE, keep_vba=True)
    ws = wb.active
    cm = TQM013_CELL_MAP

    ws.cell(row=cm['language'][0], column=cm['language'][1], value='中文')
    ws.cell(row=cm['dmba'][0], column=cm['dmba'][1], value='A')
    ws.cell(row=cm['instruction'][0], column=cm['instruction'][1], value='ON')
    ws.cell(row=cm['workstation'][0], column=cm['workstation'][1], value=selected_ws)

    ws_col = cp_data['workstation_col']
    matched = []
    for row_data in cp_data['data']:
        if ws_col and row_data.get(ws_col) is not None:
            cell_val = str(row_data[ws_col]).strip()
            if cell_val == str(selected_ws).strip():
                matched.append(row_data)
            elif str(selected_ws).strip().lower() in cell_val.lower():
                matched.append(row_data)

    seen = set()
    unique_matched = []
    for row in matched:
        key = tuple(str(row.get(k, '')) for k in sorted(row.keys()))
        if key not in seen:
            seen.add(key)
            unique_matched.append(row)

    content_cols = cm['content_cols']
    current_row = cm['content_start_row']
    for row_data in unique_matched:
        if current_row > cm['content_end_row']:
            break
        for tqm_col, cp_col_idx in column_mapping.items():
            if tqm_col in content_cols:
                tqm_col_idx = content_cols[tqm_col]
                value = row_data.get(cp_col_idx)
                if value is not None:
                    ws.cell(row=current_row, column=tqm_col_idx, value=value)
        current_row += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    wb.close()
    return output


# ==================== Flask 路由 ====================

@app.route('/')
def index():
    return HTML_PAGE


@app.route('/api/status')
def status():
    """返回模板文件状态"""
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
        return jsonify({
            'success': True,
            'workstations': cp_data['workstations'],
            'headers': {str(k): v for k, v in cp_data['headers'].items()},
            'total_rows': len(cp_data['data']),
            'workstation_column': cp_data['workstation_col'],
            'session_id': session_id,
        })
    except Exception as e:
        return jsonify({'success': False, 'error': f'解析失败: {str(e)}'})


@app.route('/api/get_column_mapping', methods=['POST'])
def get_column_mapping():
    data = request.get_json()
    cp_filepath = session.get('cp_filepath')
    if not cp_filepath or not os.path.exists(cp_filepath):
        return jsonify({'success': False, 'error': '会话已过期'})

    cp_data = parse_control_plan(cp_filepath)
    mapping = smart_match_columns(cp_data['headers'])

    tqm_labels = {
        'content_c': '内容 C', 'content_d': '内容 D',
        'description_1': '描述 1', 'description_2': '描述 2',
        'description_3': '描述 3', 'description_4': '描述 4',
        'description_5': '描述 5', 'description_6': '描述 6',
        'test_level_eq': '试验等级/设备', 'responsible': '负责人',
    }
    mapping_display = {}
    for tqm_col, cp_idx in mapping.items():
        mapping_display[tqm_col] = {
            'tqm_label': tqm_labels.get(tqm_col, tqm_col),
            'cp_column': str(cp_idx),
            'cp_header': cp_data['headers'].get(cp_idx, ''),
        }
    unmapped = [{'tqm_col': k, 'tqm_label': v} for k, v in tqm_labels.items() if k not in mapping]

    return jsonify({
        'success': True,
        'mapping': {k: v for k, v in mapping.items()},
        'mapping_display': mapping_display,
        'unmapped': unmapped,
        'all_cp_headers': {str(k): v for k, v in cp_data['headers'].items()},
    })


@app.route('/api/generate', methods=['POST'])
def generate():
    data = request.get_json()
    session_id = data.get('session_id')
    workstation = data.get('workstation')
    custom_mapping = data.get('mapping', {})

    if not session_id or not workstation:
        return jsonify({'success': False, 'error': '缺少必要参数'})

    cp_filepath = session.get('cp_filepath')
    if not cp_filepath or not os.path.exists(cp_filepath):
        return jsonify({'success': False, 'error': '会话已过期'})

    try:
        cp_data = parse_control_plan(cp_filepath)
        column_mapping = {}
        if custom_mapping:
            column_mapping = {k: int(v) for k, v in custom_mapping.items() if v}
        else:
            column_mapping = smart_match_columns(cp_data['headers'])

        safe_ws = secure_filename(str(workstation))
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"T.QM.013_{safe_ws}_{timestamp}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        fill_template(cp_data, workstation, column_mapping)

        # 直接返回文件内容
        output = fill_template(cp_data, workstation, column_mapping)
        if output:
            # 保存到磁盘供下载
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
        return jsonify({'success': False, 'error': '文件不存在或已过期'}), 404
    return send_file(filepath, as_attachment=True, download_name=safe_name)


@app.route('/api/cleanup', methods=['POST'])
def cleanup():
    sid = session.get('cp_session_id')
    if sid:
        for d in [UPLOAD_DIR, OUTPUT_DIR]:
            for f in os.listdir(d):
                if f.startswith(sid):
                    try:
                        os.remove(os.path.join(d, f))
                    except OSError:
                        pass
    return jsonify({'success': True})


# ==================== 内嵌 HTML（精简版） ====================
HTML_PAGE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>T.QM.013 检验指导书生成器</title>
<style>
:root { --primary: #2563eb; --primary-hover: #1d4ed8; --bg: #f8fafc; --card-bg: #ffffff; --border: #e2e8f0; --text: #1e293b; --text-secondary: #64748b; --success: #16a34a; --warning: #ea580c; --danger: #dc2626; --radius: 8px; --shadow: 0 1px 3px rgba(0,0,0,0.1); }
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif; background: var(--bg); color: var(--text); line-height:1.6; min-height:100vh; }
.header { background: var(--primary); color: white; padding: 14px 32px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); display: flex; align-items: center; justify-content: space-between; }
.header h1 { font-size: 1.3rem; font-weight: 600; }
.header .status { font-size: 0.8rem; opacity: 0.9; }
.container { max-width: 1100px; margin: 0 auto; padding: 24px 16px; }

.steps { display: flex; gap: 0; margin-bottom: 28px; background: var(--card-bg); border-radius: var(--radius); box-shadow: var(--shadow); overflow: hidden; }
.step { flex: 1; padding: 12px 16px; text-align: center; font-size: 0.85rem; color: var(--text-secondary); background: var(--card-bg); transition: all 0.3s; }
.step.active { color: white; background: var(--primary); font-weight: 600; }
.step.done { color: var(--success); background: #f0fdf4; }
.step .step-num { display: inline-block; width: 22px; height: 22px; line-height: 22px; border-radius: 50%; border: 2px solid currentColor; margin-right: 4px; font-size: 0.75rem; font-weight: 700; }
.step.active .step-num { background: white; color: var(--primary); border-color: white; }
.step.done .step-num { background: var(--success); color: white; border-color: var(--success); }

.card { background: var(--card-bg); border-radius: var(--radius); box-shadow: var(--shadow); padding: 24px; margin-bottom: 20px; }
.card h2 { font-size: 1.05rem; margin-bottom: 14px; padding-bottom: 8px; border-bottom: 2px solid var(--border); }
.upload-zone { border: 2px dashed var(--border); border-radius: var(--radius); padding: 36px; text-align: center; cursor: pointer; transition: all 0.3s; background: #fafbfc; }
.upload-zone:hover, .upload-zone.drag-over { border-color: var(--primary); background: #eff6ff; }
.upload-zone .upload-icon { font-size: 2.5rem; margin-bottom: 10px; }
.upload-zone p { color: var(--text-secondary); font-size: 0.9rem; }
.upload-zone .browse-link { color: var(--primary); font-weight: 600; cursor: pointer; }
.upload-zone input[type="file"] { display: none; }
.file-info { display: none; margin-top: 12px; padding: 10px 16px; background: #f0fdf4; border-radius: 6px; color: var(--success); font-size: 0.9rem; align-items: center; gap: 8px; }
.file-info.show { display: flex; }

.btn { display: inline-flex; align-items: center; gap: 6px; padding: 9px 20px; border: none; border-radius: 6px; font-size: 0.9rem; font-weight: 500; cursor: pointer; transition: all 0.2s; }
.btn-primary { background: var(--primary); color: white; }
.btn-primary:hover { background: var(--primary-hover); }
.btn-primary:disabled { background: #94a3b8; cursor: not-allowed; }
.btn-outline { background: white; color: var(--primary); border: 1.5px solid var(--primary); }
.btn-outline:hover { background: #eff6ff; }
.btn-success { background: var(--success); color: white; }
.btn-success:hover { background: #15803d; }
.btn-group { display: flex; gap: 10px; margin-top: 14px; flex-wrap: wrap; }

.ws-search { width: 100%; padding: 9px 14px; border: 1.5px solid var(--border); border-radius: 6px; font-size: 0.9rem; margin-bottom: 10px; }
.ws-search:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(37,99,235,0.1); }
.ws-list { display: flex; flex-wrap: wrap; gap: 8px; max-height: 180px; overflow-y: auto; padding: 4px; }
.ws-chip { padding: 7px 14px; border: 1.5px solid var(--border); border-radius: 20px; cursor: pointer; font-size: 0.85rem; transition: all 0.2s; user-select: none; }
.ws-chip:hover { border-color: var(--primary); background: #eff6ff; }
.ws-chip.selected { background: var(--primary); color: white; border-color: var(--primary); }

.mapping-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
.mapping-table th, .mapping-table td { padding: 8px 12px; border: 1px solid var(--border); text-align: left; }
.mapping-table th { background: #f8fafc; font-weight: 600; color: var(--text-secondary); font-size: 0.8rem; }
.mapping-table select { width: 100%; padding: 5px 8px; border: 1.5px solid var(--border); border-radius: 4px; font-size: 0.8rem; }
.mapping-table .auto-match { color: var(--success); font-size: 0.75rem; font-weight: 600; }
.mapping-table .no-match { color: var(--warning); font-size: 0.75rem; font-weight: 600; }

.result-box { padding: 20px; background: #f0fdf4; border-radius: var(--radius); text-align: center; border: 1.5px solid #bbf7d0; }
.result-box .check-icon { font-size: 2.5rem; }
.result-box h3 { color: var(--success); margin: 8px 0; }

.spinner { display: none; width: 18px; height: 18px; border: 2.5px solid var(--border); border-top-color: var(--primary); border-radius: 50%; animation: spin 0.6s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }
.alert { padding: 10px 14px; border-radius: 6px; margin-bottom: 14px; font-size: 0.85rem; }
.alert-error { background: #fef2f2; color: var(--danger); border: 1px solid #fecaca; }
.alert-info { background: #eff6ff; color: var(--primary); border: 1px solid #bfdbfe; }
.hidden { display: none !important; }
</style>
</head>
<body>

<div class="header">
    <h1>📋 T.QM.013 检验指导书生成器</h1>
    <span class="status" id="statusBar">检查模板...</span>
</div>

<div class="container">
    <div class="steps" id="stepIndicator">
        <div class="step active" data-step="1"><span class="step-num">1</span> 上传控制计划</div>
        <div class="step" data-step="2"><span class="step-num">2</span> 选择工位/OP</div>
        <div class="step" data-step="3"><span class="step-num">3</span> 确认列映射</div>
        <div class="step" data-step="4"><span class="step-num">4</span> 生成 & 下载</div>
    </div>

    <div id="globalAlert" class="hidden"></div>

    <!-- Step 1 -->
    <div class="card" id="step1Card">
        <h2>📁 步骤 1：上传版本控制计划</h2>
        <div class="upload-zone" id="uploadZone">
            <div class="upload-icon">📤</div>
            <p>拖拽 Excel 文件到此处，或 <span class="browse-link" id="browseLink">点击浏览</span></p>
            <p style="font-size:0.75rem;margin-top:4px;">支持 .xlsx / .xlsm / .xls</p>
            <input type="file" id="fileInput" accept=".xlsx,.xlsm,.xls">
        </div>
        <div class="file-info" id="fileInfo">✅ <span id="fileName"></span> — <span id="fileStats"></span></div>
        <div class="btn-group" style="justify-content:flex-end;"><span class="spinner" id="uploadSpinner"></span></div>
    </div>

    <!-- Step 2 -->
    <div class="card hidden" id="step2Card">
        <h2>🔍 步骤 2：选择目标工位/OP</h2>
        <input type="text" class="ws-search" id="wsSearch" placeholder="搜索工位/OP...">
        <div class="ws-list" id="wsList"></div>
        <div style="margin-top:8px;color:var(--text-secondary);font-size:0.8rem;">已选：<strong id="selectedWsDisplay" style="color:var(--primary);">未选择</strong></div>
        <div class="btn-group" style="justify-content:space-between;">
            <button class="btn btn-outline" onclick="resetAll()">🔄 重新上传</button>
            <button class="btn btn-primary" id="btnNextToMapping" disabled>下一步：确认映射 →</button>
        </div>
    </div>

    <!-- Step 3 -->
    <div class="card hidden" id="step3Card">
        <h2>🔗 步骤 3：确认列映射关系</h2>
        <p style="color:var(--text-secondary);font-size:0.8rem;margin-bottom:10px;">系统已自动匹配，如需调整请手动选择。</p>
        <div style="overflow-x:auto;">
            <table class="mapping-table">
                <thead><tr><th>T.QM.013 模板列</th><th>控制计划对应列</th><th>状态</th></tr></thead>
                <tbody id="mappingTableBody"></tbody>
            </table>
        </div>
        <div class="btn-group" style="justify-content:space-between;">
            <button class="btn btn-outline" id="btnBackToStep2">← 返回选择工位</button>
            <button class="btn btn-primary" id="btnGenerate">🚀 生成检验指导书 <span class="spinner" id="generateSpinner"></span></button>
        </div>
    </div>

    <!-- Step 4 -->
    <div class="card hidden" id="step4Card">
        <h2>✅ 步骤 4：生成完成！</h2>
        <div class="result-box" id="resultBox">
            <div class="check-icon">✅</div>
            <h3>检验指导书生成成功！</h3>
            <p id="resultFilename"></p>
        </div>
        <div class="btn-group" style="justify-content:center;">
            <button class="btn btn-success" id="btnDownload">📥 下载检验指导书</button>
            <button class="btn btn-outline" onclick="resetAll()">🔄 生成新的指导书</button>
        </div>
    </div>
</div>

<script>
const $=s=>document.querySelector(s),$$=s=>document.querySelectorAll(s);
const STATE={sessionId:null,workstations:[],selectedWorkstation:null,allCpHeaders:{},currentMapping:{},downloadUrl:null,currentStep:1};

function setStep(s){
    STATE.currentStep=s;
    $$('.step').forEach((el,i)=>{el.classList.remove('active','done');if(i+1<s)el.classList.add('done');if(i+1===s)el.classList.add('active');});
    ['step1Card','step2Card','step3Card','step4Card'].forEach((id,i)=>document.getElementById(id).classList.toggle('hidden',i+1!==s));
}
function showAlert(msg,type='error'){
    const el=$('#globalAlert');el.className='alert alert-'+type;el.textContent=msg;el.classList.remove('hidden');
    setTimeout(()=>el.classList.add('hidden'),6000);
}

// 检查模板状态
fetch('/api/status').then(r=>r.json()).then(d=>{
    $('#statusBar').textContent=d.template_found?'✅ 模板已加载':'⚠️ 模板未找到';
    $('#statusBar').style.color=d.template_found?'#bbf7d0':'#fecaca';
});

// Step 1: 上传
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
        if(!d.success){showAlert(d.error||'上传失败');return;}
        STATE.sessionId=d.session_id;STATE.workstations=d.workstations;STATE.allCpHeaders=d.headers;
        $('#fileName').textContent=file.name;$('#fileStats').textContent='共 '+d.total_rows+' 行，'+d.workstations.length+' 个工位';
        $('#fileInfo').classList.add('show');renderWorkstations();setStep(2);
    }catch(e){showAlert('网络错误: '+e.message);}
    finally{uploadSpinner.style.display='none';}
}

// Step 2: 工位选择
function renderWorkstations(f){
    f=(f||'').toLowerCase();
    const filtered=STATE.workstations.filter(w=>String(w).toLowerCase().includes(f));
    $('#wsList').innerHTML=filtered.map(w=>'<span class="ws-chip'+(w===STATE.selectedWorkstation?' selected':'')+'" data-ws="'+w+'">'+w+'</span>').join('');
    $('#wsList').querySelectorAll('.ws-chip').forEach(c=>c.addEventListener('click',()=>{
        STATE.selectedWorkstation=c.dataset.ws;$('#selectedWsDisplay').textContent=STATE.selectedWorkstation;
        $('#btnNextToMapping').disabled=false;renderWorkstations($('#wsSearch').value);
    }));
}
$('#wsSearch').addEventListener('input',e=>renderWorkstations(e.target.value));
$('#btnNextToMapping').addEventListener('click',async()=>{if(!STATE.selectedWorkstation)return;await loadMapping();setStep(3);});

// Step 3: 列映射
async function loadMapping(){
    try{
        const r=await fetch('/api/get_column_mapping',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({session_id:STATE.sessionId})});
        const d=await r.json();
        if(!d.success){showAlert(d.error);return;}
        STATE.currentMapping=d.mapping;
        const display=d.mapping_display,unmapped=d.unmapped||[],allHeaders=d.all_cp_headers||{};
        let rows='';
        for(const [k,v] of Object.entries(display)){
            rows+='<tr><td><strong>'+v.tqm_label+'</strong></td><td><select data-tqm="'+k+'" class="mapping-select"><option value="">-- 不映射 --</option>';
            for(const [col,hdr] of Object.entries(allHeaders))rows+='<option value="'+col+'"'+(col===v.cp_column?' selected':'')+'>'+col+' — '+(hdr||'(空)')+'</option>';
            rows+='</select></td><td><span class="auto-match">✅ 已匹配</span></td></tr>';
        }
        for(const item of unmapped){
            rows+='<tr><td><strong>'+item.tqm_label+'</strong></td><td><select data-tqm="'+item.tqm_col+'" class="mapping-select"><option value="">-- 不映射 --</option>';
            for(const [col,hdr] of Object.entries(allHeaders))rows+='<option value="'+col+'">'+col+' — '+(hdr||'(空)')+'</option>';
            rows+='</select></td><td><span class="no-match">⚠️ 未匹配</span></td></tr>';
        }
        $('#mappingTableBody').innerHTML=rows;
        $$('.mapping-select').forEach(s=>s.addEventListener('change',()=>{
            STATE.currentMapping={};$$('.mapping-select').forEach(x=>{if(x.value)STATE.currentMapping[x.dataset.tqm]=x.value;});
        }));
    }catch(e){showAlert('网络错误: '+e.message);}
}
$('#btnBackToStep2').addEventListener('click',()=>setStep(2));
$('#btnGenerate').addEventListener('click',async()=>{
    const sp=$('#generateSpinner'),btn=$('#btnGenerate');sp.style.display='inline-block';btn.disabled=true;
    try{
        const r=await fetch('/api/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({session_id:STATE.sessionId,workstation:STATE.selectedWorkstation,mapping:STATE.currentMapping})});
        const d=await r.json();
        if(!d.success){showAlert(d.error);return;}
        STATE.downloadUrl=d.download_url;$('#resultFilename').textContent='文件名：'+d.filename;$('#resultBox').style.display='block';setStep(4);
    }catch(e){showAlert('网络错误: '+e.message);}
    finally{sp.style.display='none';btn.disabled=false;}
});

// Step 4: 下载
$('#btnDownload').addEventListener('click',()=>{if(STATE.downloadUrl)window.open(STATE.downloadUrl,'_blank');});

function resetAll(){
    STATE.sessionId=null;STATE.workstations=[];STATE.selectedWorkstation=null;STATE.currentMapping={};STATE.downloadUrl=null;
    $('#fileInfo').classList.remove('show');$('#wsList').innerHTML='';$('#selectedWsDisplay').textContent='未选择';
    $('#btnNextToMapping').disabled=true;$('#mappingTableBody').innerHTML='';$('#resultBox').style.display='none';
    $('#wsSearch').value='';fileInput.value='';setStep(1);
    fetch('/api/cleanup',{method:'POST'}).catch(()=>{});
}
setStep(1);
</script>
</body>
</html>'''


# ==================== 启动入口 ====================
def find_free_port():
    """找一个可用端口"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]


def main():
    port = find_free_port()
    url = f'http://127.0.0.1:{port}'

    print("=" * 55)
    print("  T.QM.013 检验指导书生成器 v1.0")
    print("=" * 55)
    print(f"  模板文件: {TEMPLATE_FILE or '❌ 未找到!'}")
    print(f"  本地地址: {url}")
    print(f"  按 Ctrl+C 退出")
    print("=" * 55)

    # 延迟打开浏览器（等 Flask 启动完成）
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()

    # 启动 Flask（关闭 debug，隐藏控制台输出）
    import logging
    log = logging.getLogger('werkzeug')
    log.setLevel(logging.WARNING)

    app.run(host='127.0.0.1', port=port, debug=False)


if __name__ == '__main__':
    main()
