from flask import Flask, render_template, request, redirect, url_for, jsonify, session
from datetime import timedelta
import pandas as pd
import os
import difflib
import re
import json
import threading
import traceback
import copy
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

pd.options.mode.chained_assignment = None

app = Flask(__name__)
app.secret_key = 'rbc_lineage_super_secret_key'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(days=7)
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

executor = ThreadPoolExecutor(max_workers=4)
file_lock = threading.RLock()
STATE_FILE = os.path.join(app.config['UPLOAD_FOLDER'], 'auto_resolve_state.json')

def load_state():
    with file_lock:
        if os.path.exists(STATE_FILE):
            try:
                with open(STATE_FILE, 'r') as f: return json.load(f)
            except:
                return {}
        return {}

def save_state(state):
    with file_lock:
        with open(STATE_FILE, 'w') as f: json.dump(state, f, indent=4)

def update_state_for_key(key, data_dict):
    with file_lock:
        state = load_state()
        state[key] = data_dict
        save_state(state)

def delete_state_for_key(key):
    with file_lock:
        state = load_state()
        if key in state:
            del state[key]
            save_state(state)

def flexible_normalize(val):
    val = str(val).strip().upper()
    val = re.sub(r'\d{8}', '', val)
    val = re.sub(r'\d{4}-\d{2}-\d{2}', '', val)
    return re.sub(r'[^A-Z0-9]', '', val)

def get_safe_filename(cde_name):
    return f"{re.sub(r'[^a-zA-Z0-9_\-]', '_', cde_name).strip()}_Consolidation_Workbook.xlsx"

@app.route('/')
def index(): return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def process_upload():
    request.files['target_cde_file'].save(os.path.join(app.config['UPLOAD_FOLDER'], 'reporting_layers.xlsx'))
    if request.files.get('primary_lineage_file').filename: request.files['primary_lineage_file'].save(os.path.join(app.config['UPLOAD_FOLDER'], 'primary_lineage.xlsx'))
    if request.files.get('global_lineage_file').filename: request.files['global_lineage_file'].save(os.path.join(app.config['UPLOAD_FOLDER'], 'global_lineage.xlsx'))
    return redirect(url_for('dashboard'))

@app.route('/dashboard')
def dashboard():
    cde_path = os.path.join(app.config['UPLOAD_FOLDER'], 'reporting_layers.xlsx')
    if not os.path.exists(cde_path): return redirect(url_for('index'))
    
    with file_lock:
        xls = pd.ExcelFile(cde_path)
        df = pd.read_excel(xls, 'Sheet1')
        df_done = pd.read_excel(xls, 'Done') if 'Done' in xls.sheet_names else pd.DataFrame()
        
    done_keys = {f"{str(r.get('CDE name','')).strip()}|{str(r.get('Current app code','')).strip()}|{str(r.get('Current table/file name','')).strip()}|{str(r.get('Current column/field name','')).strip()}" for _, r in df_done.iterrows()}

    if 'Target App Codes' not in df.columns: df['Target App Codes'] = ""
    state_db = load_state()
    
    grouped = df.groupby('CDE name')
    cdes = []
    metrics = {'total_cdes': len(grouped), 'completed_cdes': 0, 'total_instances': len(df), 'completed_instances': 0, 'needs_review': 0, 'processing': 0, 'auto_resolvable': 0}

    for cde_name, group in grouped:
        target_apps = set()
        for apps in group['Target App Codes'].dropna(): target_apps.update([app.strip().upper() for app in str(apps).split(',') if app.strip()])
        instances = []
        cde_completed = True
        
        for _, row in group.iterrows():
            app_code = str(row.get("Current app code", "")).strip()
            table = str(row.get("Current table/file name", "")).strip()
            col = str(row.get("Current column/field name", "")).strip()
            key = f"{cde_name}|{app_code}|{table}|{col}"
            
            if key in done_keys:
                status = "Completed"
                metrics['completed_instances'] += 1
            else:
                cde_completed = False
                job_state = state_db.get(key, {}).get("status", "Pending")
                if job_state == "NEEDS_REVIEW": status, metrics['needs_review'] = "Needs Review", metrics['needs_review'] + 1
                elif job_state == "PROCESSING": status, metrics['processing'] = "Processing", metrics['processing'] + 1
                else: status, metrics['auto_resolvable'] = "Pending", metrics['auto_resolvable'] + 1

            instances.append({"app": app_code, "table": table, "column": col, "status": status, "key": key, "reason": state_db.get(key, {}).get("reason", "")})
            
        if cde_completed: metrics['completed_cdes'] += 1
        cdes.append({"name": cde_name, "target_apps": list(target_apps), "instances": instances, "completed": cde_completed})
        
    return render_template('dashboard.html', metrics=metrics, cdes=cdes)

@app.route('/api/retrace', methods=['POST'])
def api_retrace():
    data = request.json
    cde_name, app_code, table, col = data['cde_name'], data['app'], data['table'], data['column']
    cde_path = os.path.join(app.config['UPLOAD_FOLDER'], 'reporting_layers.xlsx')
    
    with file_lock:
        df_done = pd.read_excel(cde_path, sheet_name='Done')
        df_done = df_done[~((df_done['CDE name'] == cde_name) & (df_done['Current app code'] == app_code) & (df_done['Current table/file name'] == table) & (df_done['Current column/field name'] == col))]
        with pd.ExcelWriter(cde_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_done.to_excel(writer, sheet_name='Done', index=False)
            
    delete_state_for_key(f"{cde_name}|{app_code}|{table}|{col}")
    return jsonify({"status": "success"})

@app.route('/trace/<cde_name>/<app_code>/<table_name>/<column_name>')
def trace(cde_name, app_code, table_name, column_name):
    target_apps = request.args.get('targets', '')
    resume_auto = request.args.get('resume_auto', 'false')
    key = f"{cde_name}|{app_code}|{table_name}|{column_name}"
    
    session.permanent = True
    if 'lineage_stack' not in session: session['lineage_stack'] = []
    
    state = load_state()
    resume_stack = state.get(key, {}).get("stack", [])
    
    return render_template('trace.html', cde_name=cde_name, app_code=app_code, table_name=table_name, column_name=column_name, target_apps=target_apps, resume_stack=json.dumps(resume_stack), resume_auto=resume_auto)

@app.route('/api/undo', methods=['POST'])
def undo():
    if 'lineage_stack' not in session or len(session['lineage_stack']) == 0: return jsonify({"error": "Nothing to undo"}), 400
    session['lineage_stack'].pop()
    session.modified = True
    return jsonify({"success": True, "stack_length": len(session['lineage_stack'])})

def search_dataframe(df, app_code, table, col):
    app_code, table, col = str(app_code).strip().upper(), str(table).strip().upper(), str(col).strip().upper()
    if not app_code or app_code == 'NAN' or not table or table == 'NAN' or not col or col == 'NAN': return pd.DataFrame(), "No Match"

    # FIXED: Bulletproof Pandas string conversion to avoid 'nan' string poisoning
    for c in ['Current app code', 'Current table/file name', 'Current column/field name']:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].fillna("").astype(str).str.strip().str.upper()

    df['clean_app'] = df['Current app code']
    df['clean_table'] = df['Current table/file name']
    df['clean_col'] = df['Current column/field name']

    exact = df[(df['clean_app'] == app_code) & (df['clean_table'] == table) & (df['clean_col'] == col)]
    if not exact.empty: return exact, "Exact Match"

    df['flex_table'] = df['clean_table'].apply(flexible_normalize)
    df['flex_col'] = df['clean_col'].apply(flexible_normalize)
    normalized_table, normalized_col = flexible_normalize(table), flexible_normalize(col)
    
    flex = df[(df['clean_app'] == app_code) & (df['flex_table'] == normalized_table) & (df['flex_col'] == normalized_col)]
    if not flex.empty: return flex, "Flexible Match"

    app_filtered = df[df['clean_app'] == app_code]
    if app_filtered.empty: return pd.DataFrame(), "No Match"

    choices_list = [f"{r['clean_table']} | {r['clean_col']}" for _, r in app_filtered[['clean_table', 'clean_col']].drop_duplicates().iterrows()]
    close_matches = difflib.get_close_matches(f"{table} | {col}", choices_list, n=5, cutoff=0.3)

    if close_matches:
        parts = close_matches[0].split(' | ')
        matched_table, matched_col = parts[0], parts[1] if len(parts) > 1 else ""
        fuzzy = df[(df['clean_app'] == app_code) & (df['clean_table'] == matched_table) & (df['clean_col'] == matched_col)]
        if not fuzzy.empty: return fuzzy, "Fuzzy Match"

    return pd.DataFrame(), "No Match"

def execute_search_api_logic(cde_name, app_code, table, col, target_apps):
    primary_path, global_path = os.path.join(app.config['UPLOAD_FOLDER'], 'primary_lineage.xlsx'), os.path.join(app.config['UPLOAD_FOLDER'], 'global_lineage.xlsx')
    unique_groups = []

    def process_file(file_path, source_label):
        if not os.path.exists(file_path): return
        
        with file_lock:
            df = pd.read_excel(file_path)
            
        matches, match_type = search_dataframe(df, app_code, table, col)
        if matches.empty: return

        matches = matches.copy()
        
        matches['group_id'] = matches['file'].astype(str) + "||" + matches.get('Report name', 'N/A').astype(str) + "||" + matches.get('CDE name', 'N/A').astype(str)
        for gid, group in matches.groupby('group_id'):
            rep_row = group.iloc[0]
            distinct_sources = sorted(group['Source app code'].dropna().astype(str).str.strip().str.upper().unique().tolist())
            if not distinct_sources: continue

            score = 1000000 if cde_name.upper() in str(rep_row.get('CDE name', '')).upper() else 0
            if any(s in [t.upper() for t in target_apps] for s in distinct_sources): score += 500000
            score += {"Exact Match": 100000, "Flexible Match": 50000, "Fuzzy Match": 25000}.get(match_type, 0)

            raw_rows = []
            for _, row in group.iterrows():
                src_app = str(row.get('Source app code', '')).strip().upper()
                if src_app and src_app != 'NAN':
                    raw_rows.append({
                        "next_app": src_app, "next_table": str(row.get('Source table/file name', '')).strip(), "next_col": str(row.get('Source column/field name', '')).strip(),
                        "searched_table": table, "found_table": str(rep_row.get('Current table/file name', '')).strip(), "searched_col": col, "found_col": str(rep_row.get('Current column/field name', '')).strip(),
                        "match_type": match_type, "raw_row_data": row.drop(labels=['clean_app', 'clean_table', 'clean_col', 'flex_table', 'flex_col', 'group_id'], errors='ignore').fillna('').to_dict()
                    })

            if raw_rows:
                unique_groups.append({
                    "source": source_label, "score": score, "match_type": match_type, "file_path": str(rep_row.get('file', 'N/A')), "report_name": str(rep_row.get('Report name', 'N/A')),
                    "cde_found": str(rep_row.get('CDE name', '')), "found_table": str(rep_row.get('Current table/file name', '')), "found_col": str(rep_row.get('Current column/field name', '')),
                    "source_count": len(raw_rows), "distinct_sources": distinct_sources, "raw_rows": raw_rows
                })

    process_file(primary_path, "Primary Report")
    if not unique_groups: process_file(global_path, "Global Lineage")
    unique_groups.sort(key=lambda x: x['score'], reverse=True)
    return unique_groups

@app.route('/api/search', methods=['POST'])
def api_search():
    data = request.json
    results = execute_search_api_logic(data['cde_name'], data['app'], data['table'], data['column'], data['target_apps'])
    return jsonify({"candidates": results})

def write_lineage_to_files(cde_name, stack, is_dead_end):
    with file_lock:
        cde_path = os.path.join(app.config['UPLOAD_FOLDER'], 'reporting_layers.xlsx')
        with pd.ExcelWriter(cde_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df_done = pd.read_excel(cde_path, sheet_name='Done') if 'Done' in pd.ExcelFile(cde_path).sheet_names else pd.DataFrame()
            new_done = pd.DataFrame([{"Report name": "Lineage Engine", "CDE name": cde_name, "Current app code": stack[0]['app'], "Current table/file name": stack[0]['original_table'], "Current column/field name": stack[0]['original_col']}])
            df_done = pd.concat([df_done, new_done], ignore_index=True)
            df_done.to_excel(writer, sheet_name='Done', index=False)

        output_rows = []
        seen_lineages = set()

        for i in range(len(stack)):
            curr = stack[i]
            if i < len(stack) - 1:
                next_node = stack[i+1]
                s_app = next_node['app']
                s_table = next_node.get('reconciled_table', next_node['original_table'])
                s_col = next_node.get('reconciled_col', next_node['original_col'])
            else:
                s_app, s_table, s_col = "", "", ""

            dest = stack[i-1]['app'] if i > 0 else ""
            lineage_key = f"{curr['app']}|{curr.get('reconciled_table', curr['original_table'])}|{curr.get('reconciled_col', curr['original_col'])}->{s_app}|{s_table}|{s_col}"
            if lineage_key in seen_lineages: continue
            seen_lineages.add(lineage_key)

            label = f"↓ {curr['app']}" if i == 0 else f"↑ {dest}   ↓ {curr['app']}"
            output_rows.append({"Label": label})
            
            data_row = curr.get('raw_data', {}).copy()
            current_table_output = curr.get('reconciled_table', curr['original_table'])
            current_col_output = curr.get('reconciled_col', curr['original_col'])

            if s_app and next_node.get('searched_table') and next_node.get('found_table'):
                nst, nsc, nft, nfc = next_node.get('searched_table', ''), next_node.get('searched_col', next_node['original_col']), next_node.get('found_table', ''), next_node.get('found_col', next_node['original_col'])
                s_table = nft if nst == nft else f"{nst}/{nft}"
                s_col = nfc if nsc == nfc else f"{nsc}/{nfc}"

            data_row.update({"CDE name": cde_name, "Destination to (App Code)": dest, "Current app code": curr['app'], "Current table/file name": current_table_output, "Current column/field name": current_col_output, "Source app code": s_app, "Source table/file name": s_table, "Source column/field name": s_col, "QA Notes": curr.get('notes', '')})
            output_rows.append(data_row)

        output_file = os.path.join(app.config['OUTPUT_FOLDER'], get_safe_filename(cde_name))
        df_out = pd.DataFrame(output_rows)
        if os.path.exists(output_file):
            existing_df = pd.read_excel(output_file)
            df_combined = pd.concat([existing_df, df_out], ignore_index=True).drop_duplicates(subset=['Current app code', 'Current table/file name', 'Current column/field name', 'Source app code'], keep='first')
            df_out = df_combined

        cols = df_out.columns.tolist()
        if 'Label' in cols:
            cols.insert(0, cols.pop(cols.index('Label')))
            df_out = df_out[cols]
        df_out.to_excel(output_file, index=False)

@app.route('/api/save_lineage', methods=['POST'])
def save_lineage():
    data = request.json
    write_lineage_to_files(data['cde_name'], data['stack'], data.get('is_dead_end', False))
    if 'lineage_stack' in session:
        session['lineage_stack'] = []
        session.modified = True
    return jsonify({"status": "success", "redirect": url_for('dashboard')})

# --- UNIFIED BACKGROUND JOB ENGINE ---
def process_instance_background(cde_name, target_key, initial_raw_data=None):
    try:
        cde_path = os.path.join(app.config['UPLOAD_FOLDER'], 'reporting_layers.xlsx')
        
        with file_lock:
            xls = pd.ExcelFile(cde_path)
            df = pd.read_excel(xls, 'Sheet1')
        
        target_apps = []
        for _, row in df[df['CDE name'] == cde_name].iterrows():
            target_apps.extend([t.strip().upper() for t in str(row.get("Target App Codes", "")).split(',') if t.strip()])
        
        state = load_state()
        stack = state.get(target_key, {}).get('stack', [])
        
        if not stack:
            _, app_code, table, col = target_key.split('|')
            # FIXED: Passes the actual App Name and metadata to the very first reporting node!
            stack = [{"app": app_code, "original_table": table, "original_col": col, "reconciled_table": table, "reconciled_col": col, "raw_data": initial_raw_data or {}}]
            
        update_state_for_key(target_key, {"status": "PROCESSING", "stack": stack, "timestamp": str(datetime.now())})

        stacks_to_process = [stack]
        completed_paths = []

        while stacks_to_process:
            current_stack = stacks_to_process.pop(0)
            current_node = current_stack[-1]
            
            candidates = execute_search_api_logic(cde_name, current_node['app'], current_node.get('original_table', current_node.get('reconciled_table')), current_node.get('original_col', current_node.get('reconciled_col')), target_apps)

            if len(candidates) == 0:
                completed_paths.append((current_stack, True))
                continue

            best_cand = candidates[0]

            if best_cand['match_type'] == 'Fuzzy Match':
                reason = f"Fuzzy Match detected for '{current_node.get('original_table')}'. Human review required."
                update_state_for_key(target_key, {"status": "NEEDS_REVIEW", "stack": current_stack, "reason": reason, "timestamp": str(datetime.now())})
                return

            for src in best_cand['raw_rows']:
                new_stack = copy.deepcopy(current_stack)
                new_stack[-1]['reconciled_table'] = best_cand['found_table']
                new_stack[-1]['reconciled_col'] = src['found_col'] if src.get('found_col') else current_node.get('original_col', '')
                new_stack[-1]['raw_data'] = src['raw_row_data']

                next_app = src['next_app']
                new_node = {
                    "app": next_app, 
                    "original_table": src['next_table'], 
                    "original_col": src['next_col'], 
                    "reconciled_table": src['next_table'], 
                    "reconciled_col": src['next_col'], 
                    "raw_data": src['raw_row_data']
                }
                new_stack.append(new_node)

                if next_app in target_apps:
                    new_stack[-1]['notes'] = "Target Reached (Auto-Resolved)"
                    completed_paths.append((new_stack, False))
                else:
                    stacks_to_process.append(new_stack)

        for path, is_dead_end in completed_paths:
            write_lineage_to_files(cde_name, path, is_dead_end)

        update_state_for_key(target_key, {"status": "RESOLVED", "timestamp": str(datetime.now())})

    except Exception as e:
        error_msg = f"System Error: {str(e)}"
        print(f"CRASH in background task for {target_key}: {traceback.format_exc()}")
        update_state_for_key(target_key, {"status": "NEEDS_REVIEW", "stack": stack, "reason": error_msg, "timestamp": str(datetime.now())})

@app.route('/api/auto_resolve/start', methods=['POST'])
def start_auto_resolve():
    cde_path = os.path.join(app.config['UPLOAD_FOLDER'], 'reporting_layers.xlsx')
    
    with file_lock:
        xls = pd.ExcelFile(cde_path)
        df = pd.read_excel(xls, 'Sheet1')
        df_done = pd.read_excel(xls, 'Done') if 'Done' in pd.ExcelFile(cde_path).sheet_names else pd.DataFrame()
        
    done_keys = {f"{str(row.get('CDE name','')).strip()}|{str(row.get('Current app code','')).strip()}|{str(row.get('Current table/file name','')).strip()}|{str(row.get('Current column/field name','')).strip()}" for _, row in df_done.iterrows()}
    
    state = load_state()
    for _, row in df.iterrows():
        c_name = str(row.get("CDE name", "")).strip()
        c_app = str(row.get("Current app code", "")).strip()
        c_tab = str(row.get("Current table/file name", "")).strip()
        c_col = str(row.get("Current column/field name", "")).strip()
        
        key = f"{c_name}|{c_app}|{c_tab}|{c_col}"
        
        if key not in done_keys and state.get(key, {}).get("status") != "NEEDS_REVIEW":
            # Extract raw data to pass App Name properly!
            raw_data = row.fillna("").to_dict()
            executor.submit(process_instance_background, c_name, key, raw_data)
            
    return jsonify({"status": "Bulk job started"})

@app.route('/api/auto_resolve/single', methods=['POST'])
def start_single_auto_resolve():
    data = request.json
    target_key = f"{data['cde_name']}|{data['app']}|{data['table']}|{data['column']}"
    
    # We do a quick fetch of the row to give it the App Name
    with file_lock:
        df = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'], 'reporting_layers.xlsx'), 'Sheet1')
        row = df[(df['CDE name'] == data['cde_name']) & (df['Current app code'] == data['app'])].iloc[0]
        raw_data = row.fillna("").to_dict()
        
    executor.submit(process_instance_background, data['cde_name'], target_key, raw_data)
    return jsonify({"status": "Single instance job started"})

@app.route('/api/auto_resolve/resume', methods=['POST'])
def resume_auto_resolve():
    data = request.json
    cde_name, stack = data['cde_name'], data['stack']
    target_key = f"{cde_name}|{stack[0]['app']}|{stack[0]['original_table']}|{stack[0]['original_col']}"
    
    update_state_for_key(target_key, {"status": "PROCESSING", "stack": stack, "timestamp": str(datetime.now())})
    executor.submit(process_instance_background, cde_name, target_key)
    return jsonify({"status": "Resumed in background"})

@app.route('/view_lineage/<path:cde_name>')
def view_lineage(cde_name):
    output_file = os.path.join(app.config['OUTPUT_FOLDER'], get_safe_filename(cde_name))
    if not os.path.exists(output_file): return redirect(url_for('dashboard'))
    return render_template('view_lineage.html', cde_name=cde_name)

@app.route('/api/graph/<path:cde_name>')
def api_graph(cde_name):
    try:
        output_file = os.path.join(app.config['OUTPUT_FOLDER'], get_safe_filename(cde_name))
        if not os.path.exists(output_file): return jsonify({"error": "File not found"})

        df = pd.read_excel(output_file).fillna("")
        
        inst_nodes, inst_edges = {}, {}
        app_nodes, app_edges = {}, {}
        app_meta = {}

        for _, row in df.iterrows():
            if not str(row.get('CDE name', '')).strip(): continue
            
            c_app = str(row.get('Current app code', '')).strip().upper()
            if not c_app or c_app.lower() == 'nan': continue
            
            c_name_val = str(row.get('Current app name', '')).strip()
            c_name = c_name_val if c_name_val.lower() not in ['nan', ''] else c_app
            
            c_tab = str(row.get('Current table/file name', '')).strip()
            c_col = str(row.get('Current column/field name', '')).strip()
            
            s_app = str(row.get('Source app code', '')).strip().upper()
            s_name_val = str(row.get('Source app name', '')).strip()
            s_name = s_name_val if s_name_val.lower() not in ['nan', ''] else s_app
            
            s_tab = str(row.get('Source table/file name', '')).strip()
            s_col = str(row.get('Source column/field name', '')).strip()

            is_manual = 'manual' in str(row.get('Manually entered/Derived', '')).lower() or 'manual' in str(row.get('Created/Sourced', '')).lower()
            is_transformed = 'transform' in str(row.get('Transformed/Passed Through', '')).lower() or str(row.get('Transformation Logic', '')).strip() != ''
            is_filtered = 'filter' in str(row.get('Filtered', '')).lower() or str(row.get('Filtration Logic', '')).strip() != ''

            if c_app not in app_meta:
                app_meta[c_app] = {'code': c_app, 'name': c_name, 'manual': False, 'transformed': False, 'filtered': False, 'instances': []}
            
            app_meta[c_app]['manual'] = app_meta[c_app]['manual'] or is_manual
            app_meta[c_app]['transformed'] = app_meta[c_app]['transformed'] or is_transformed
            app_meta[c_app]['filtered'] = app_meta[c_app]['filtered'] or is_filtered
            
            inst_entry = {'table': c_tab, 'col': c_col}
            if inst_entry not in app_meta[c_app]['instances']:
                app_meta[c_app]['instances'].append(inst_entry)

            # --- INSTANCE GRAPH ---
            c_id = f"{c_app}|{c_tab.upper()}|{c_col.upper()}"
            if c_id not in inst_nodes:
                inst_nodes[c_id] = {"id": c_id, "label": f"🏢 Name: {c_name}\n🔑 Code: {c_app}\n📁 Table: {c_tab}\n📌 Col: {c_col}", "shape": "box", "color": {"background": "#F3F4F6", "border": "#005DAA"}, "font": {"face": "monospace", "align": "left"}}

            if s_app and s_app.lower() != 'nan':
                s_id = f"{s_app}|{s_tab.upper()}|{s_col.upper()}"
                if s_id not in inst_nodes:
                    inst_nodes[s_id] = {"id": s_id, "label": f"🏢 Name: {s_name}\n🔑 Code: {s_app}\n📁 Table: {s_tab}\n📌 Col: {s_col}", "shape": "box", "color": {"background": "#E2E8F0", "border": "#64748B"}, "font": {"face": "monospace", "align": "left"}}
                
                e_id = f"{s_id}->{c_id}"
                inst_edges[e_id] = {"from": s_id, "to": c_id, "arrows": "to", "color": {"color": "#005DAA"}}

                if s_app not in app_meta:
                    app_meta[s_app] = {'code': s_app, 'name': s_name, 'manual': False, 'transformed': False, 'filtered': False, 'instances': []}
                
                e_app_id = f"{s_app}->{c_app}"
                if e_app_id not in app_edges:
                    app_edges[e_app_id] = {"from": s_app, "to": c_app, "arrows": "to", "color": {"color": "#10B981"}, "width": 2}

        src_inst_ids = {e['from'] for e in inst_edges.values()}
        for n_id, n in inst_nodes.items():
            if n_id not in src_inst_ids:
                n['color'], n['borderWidth'] = {"background": "#E0F2FE", "border": "#0284C7"}, 3

        src_app_ids = {e['from'] for e in app_edges.values()}
        for app_key, meta in app_meta.items():
            app_nodes[app_key] = {"id": app_key, "meta": meta, "is_target": app_key not in src_app_ids}

        return jsonify({"instance_graph": {"nodes": list(inst_nodes.values()), "edges": list(inst_edges.values())}, "app_graph": {"nodes": list(app_nodes.values()), "edges": list(app_edges.values())}})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)