"""
T.QM.013 检验指导书自动生成工具 — Streamlit 增强版
"""
import os
import glob
from datetime import datetime
from io import BytesIO

import streamlit as st
import openpyxl

# ==================== 自动查找模板文件 ====================
# 在仓库根目录自动搜索 .xlsm 模板文件
TEMPLATE_FILE = None

# 先尝试固定文件名
FIXED_NAMES = [
    'T.QM.013_Template.xlsm',
    'T.QM.013_Work and Test Instruction_Arbeits und Prüfanweisung_Rev 2026_05 (1).xlsm',
]

for name in FIXED_NAMES:
    if os.path.exists(name):
        TEMPLATE_FILE = name
        break

# 再尝试模糊搜索
if TEMPLATE_FILE is None:
    xlsm_files = glob.glob('*.xlsm') + glob.glob('*.xlsx')
    for f in xlsm_files:
        if 'T.QM.013' in f or 'QM.013' in f:
            TEMPLATE_FILE = f
            break

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


def parse_control_plan(file_bytes):
    """解析上传的控制计划 Excel"""
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
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
            "请确保已将 .xlsm 模板文件上传到 GitHub 仓库根目录。\n"
            f"当前仓库根目录文件列表: {os.listdir('.')}"
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

    # 去重
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


# ==================== Streamlit 界面 ====================
st.set_page_config(page_title="T.QM.013 检验指导书生成器", page_icon="📋", layout="wide")

st.title("📋 T.QM.013 检验指导书生成器")
st.caption("上传版本控制计划 → 选择工位/OP → 确认列映射 → 生成检验指导书")

# 显示模板文件状态
if TEMPLATE_FILE:
    st.sidebar.success(f"✅ 模板: {TEMPLATE_FILE}")
else:
    st.sidebar.error("❌ 未找到模板文件！")
    st.sidebar.info(f"仓库文件: {os.listdir('.')}")

# 初始化 session state
for key, default in {
    'cp_data': None, 'step': 1, 'selected_ws': None,
    'column_mapping': {}, 'output_file': None, 'output_name': ''
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ====== 步骤指示器 ======
steps = ["① 上传控制计划", "② 选择工位/OP", "③ 确认列映射", "④ 生成 & 下载"]
cols = st.columns(4)
for i, (col, label) in enumerate(zip(cols, steps)):
    with col:
        if i + 1 < st.session_state.step:
            st.success(label)
        elif i + 1 == st.session_state.step:
            st.info(f"**{label}**")
        else:
            st.markdown(f"<span style='color:#94a3b8'>{label}</span>", unsafe_allow_html=True)

st.divider()

# ====== Step 1 ======
if st.session_state.step == 1:
    st.subheader("步骤 1：上传版本控制计划 Excel")
    uploaded = st.file_uploader(
        "拖拽或选择 Excel 文件", type=['xlsx', 'xlsm', 'xls'],
        help="上传包含工位/OP 列的版本控制计划"
    )

    if uploaded:
        with st.spinner("正在解析文件..."):
            try:
                cp_data = parse_control_plan(uploaded.getvalue())
                st.session_state.cp_data = cp_data
                st.success(f"✅ 解析成功！共 {len(cp_data['data'])} 行数据，{len(cp_data['workstations'])} 个工位/OP")
                st.info(f"工位列: 第 {cp_data['workstation_col']} 列")
                if st.button("下一步：选择工位 →"):
                    st.session_state.step = 2
                    st.rerun()
            except Exception as e:
                st.error(f"解析失败: {e}")

# ====== Step 2 ======
elif st.session_state.step == 2:
    st.subheader("步骤 2：选择目标工位/OP")

    workstations = st.session_state.cp_data['workstations']
    search = st.text_input("🔍 搜索工位/OP", placeholder="输入关键词过滤...")

    filtered = [w for w in workstations if search.lower() in str(w).lower()] if search else workstations

    cols_per_row = 6
    for i in range(0, len(filtered), cols_per_row):
        row_cols = st.columns(cols_per_row)
        for j, ws_name in enumerate(filtered[i:i+cols_per_row]):
            with row_cols[j]:
                is_selected = (st.session_state.selected_ws == ws_name)
                if st.button(
                    f"✅ {ws_name}" if is_selected else str(ws_name),
                    key=f"ws_{ws_name}",
                    use_container_width=True,
                    type="primary" if is_selected else "secondary"
                ):
                    st.session_state.selected_ws = ws_name
                    st.rerun()

    if st.session_state.selected_ws:
        st.success(f"已选择: **{st.session_state.selected_ws}**")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("← 重新上传"):
            for k in ['cp_data', 'selected_ws', 'column_mapping', 'output_file', 'output_name']:
                st.session_state[k] = None if k != 'column_mapping' else {}
            st.session_state.step = 1
            st.rerun()
    with c2:
        if st.button("下一步：确认映射 →", disabled=(st.session_state.selected_ws is None)):
            st.session_state.step = 3
            st.rerun()

# ====== Step 3 ======
elif st.session_state.step == 3:
    st.subheader(f"步骤 3：确认列映射关系（工位: {st.session_state.selected_ws}）")
    st.caption("系统已自动匹配，如需调整请手动选择。")

    cp_headers = st.session_state.cp_data['headers']
    auto_mapping = smart_match_columns(cp_headers)

    tqm_labels = {
        'content_c': '内容 C', 'content_d': '内容 D',
        'description_1': '描述 1', 'description_2': '描述 2',
        'description_3': '描述 3', 'description_4': '描述 4',
        'description_5': '描述 5', 'description_6': '描述 6',
        'test_level_eq': '试验等级/设备', 'responsible': '负责人',
    }

    cp_options = {}
    for idx, hdr in cp_headers.items():
        label = f"列{idx}: {hdr}" if hdr else f"列{idx} (空)"
        cp_options[label] = idx

    option_labels = ["-- 不映射 --"] + list(cp_options.keys())

    final_mapping = {}
    for tqm_col, label in tqm_labels.items():
        default_idx = auto_mapping.get(tqm_col)
        if default_idx is not None:
            default_hdr = cp_headers.get(default_idx, '')
            default_label = f"列{default_idx}: {default_hdr}" if default_hdr else f"列{default_idx} (空)"
        else:
            default_label = "-- 不映射 --"

        if default_label not in option_labels:
            default_label = "-- 不映射 --"

        selected_label = st.selectbox(
            f"**{label}**  ← 对应",
            option_labels,
            index=option_labels.index(default_label),
            key=f"map_{tqm_col}"
        )
        if selected_label != "-- 不映射 --":
            final_mapping[tqm_col] = cp_options[selected_label]

    st.session_state.column_mapping = final_mapping

    c1, c2 = st.columns(2)
    with c1:
        if st.button("← 返回选择工位"):
            st.session_state.step = 2
            st.rerun()
    with c2:
        if st.button("🚀 生成检验指导书", type="primary"):
            with st.spinner("正在生成..."):
                try:
                    output = fill_template(
                        st.session_state.cp_data,
                        st.session_state.selected_ws,
                        st.session_state.column_mapping
                    )
                    if output:
                        st.session_state.output_file = output
                        st.session_state.output_name = (
                            f"T.QM.013_{st.session_state.selected_ws}_"
                            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        )
                        st.session_state.step = 4
                        st.rerun()
                except Exception as e:
                    st.error(f"生成失败: {e}")
                    st.info(f"模板文件: {TEMPLATE_FILE}")
                    st.info(f"仓库文件列表: {os.listdir('.')}")

# ====== Step 4 ======
elif st.session_state.step == 4:
    st.subheader("✅ 步骤 4：生成完成！")
    st.success(f"检验指导书已生成：**{st.session_state.output_name}**")

    st.download_button(
        label="📥 下载检验指导书",
        data=st.session_state.output_file,
        file_name=st.session_state.output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

    if st.button("🔄 生成新的指导书"):
        for k in ['cp_data', 'selected_ws', 'column_mapping', 'output_file', 'output_name']:
            st.session_state[k] = None if k != 'column_mapping' else {}
        st.session_state.step = 1
        st.rerun()
