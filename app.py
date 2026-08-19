from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    jsonify,
    session,
    flash,
)
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
app.secret_key = "123_lineage_super_secret_key"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=7)

WORKSPACE_DIR = "workspace"
os.makedirs(WORKSPACE_DIR, exist_ok=True)

executor = ThreadPoolExecutor(max_workers=8)
file_lock = threading.RLock()


# --- PROJECT WORKSPACE HELPERS ---
def get_project_dir(project_name, subfolder=""):
    path = os.path.join(WORKSPACE_DIR, project_name, subfolder)
    os.makedirs(path, exist_ok=True)
    return path


def get_state_file(project_name):
    return os.path.join(
        get_project_dir(project_name, "uploads"), "auto_resolve_state.json"
    )


def make_json_serializable(obj):
    if pd.isna(obj):
        return None
    if isinstance(obj, dict):
        return {k: make_json_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, (list, tuple)):
        return [make_json_serializable(item) for item in obj]
    elif hasattr(obj, "isoformat"):
        try:
            return obj.isoformat()
        except:
            return None
    elif isinstance(obj, (int, float, str, bool, type(None))):
        return obj
    else:
        return str(obj)


def clean_raw_data_for_json(raw_dict):
    """Clean raw row data for JSON serialization"""
    if not isinstance(raw_dict, dict):
        return raw_dict

    cleaned = {}
    for k, v in raw_dict.items():
        if pd.isna(v):
            cleaned[k] = None
        elif hasattr(v, "isoformat"):
            try:
                cleaned[k] = v.isoformat()
            except (ValueError, TypeError):
                cleaned[k] = None
        else:
            cleaned[k] = v
    return cleaned


def load_state(project_name):
    with file_lock:
        state_file = get_state_file(project_name)
        if os.path.exists(state_file):
            try:
                with open(state_file, "r") as f:
                    return json.load(f)
            except:
                return {}
        return {}


def save_state(project_name, state):
    with file_lock:
        state = make_json_serializable(state)
        with open(get_state_file(project_name), "w") as f:
            json.dump(state, f, indent=4)


def update_state_for_key(project_name, key, data_dict):
    with file_lock:
        state = load_state(project_name)
        state[key] = data_dict
        save_state(project_name, state)


def delete_state_for_key(project_name, key):
    with file_lock:
        state = load_state(project_name)
        if key in state:
            del state[key]
            save_state(project_name, state)


def track_file_usage(project_name, file_path):
    if not file_path or str(file_path).upper() == "NAN":
        return
    with file_lock:
        state = load_state(project_name)
        if "file_memory" not in state:
            state["file_memory"] = {}
        state["file_memory"][str(file_path)] = (
            state["file_memory"].get(str(file_path), 0) + 1
        )
        save_state(project_name, state)


def get_file_memory(project_name):
    with file_lock:
        return load_state(project_name).get("file_memory", {})


def flexible_normalize(val):
    val = str(val).strip().upper()
    if val == "NAN":
        return ""
    val = re.sub(r"\d{8}", "", val)
    val = re.sub(r"\d{4}-\d{2}-\d{2}", "", val)
    return re.sub(r"[^A-Z0-9]", "", val)


def get_safe_filename(cde_name):
    return f"{re.sub(r'[^a-zA-Z0-9_\-]', '_', cde_name).strip()}_Consolidation_Workbook.xlsx"


# --- ROUTES ---
@app.route("/")
def index():
    projects = [
        d
        for d in os.listdir(WORKSPACE_DIR)
        if os.path.isdir(os.path.join(WORKSPACE_DIR, d))
    ]
    active_project = session.get("active_project")
    return render_template(
        "upload.html", projects=projects, active_project=active_project
    )


@app.route("/set_project", methods=["POST"])
def set_project():
    project_name = request.form.get("project_name", "").strip()
    project_name = re.sub(r"[^a-zA-Z0-9_\- ]", "", project_name)  # Sanitize
    if project_name:
        session["active_project"] = project_name
        get_project_dir(project_name, "uploads")
        get_project_dir(project_name, "outputs")
    return redirect(url_for("index"))


@app.route("/upload", methods=["POST"])
def process_upload():
    project_name = session.get("active_project")
    if not project_name:
        return redirect(url_for("index"))

    upload_dir = get_project_dir(project_name, "uploads")
    if "target_cde_file" in request.files and request.files["target_cde_file"].filename:
        request.files["target_cde_file"].save(
            os.path.join(upload_dir, "reporting_layers.xlsx")
        )
    if (
        "primary_lineage_file" in request.files
        and request.files["primary_lineage_file"].filename
    ):
        request.files["primary_lineage_file"].save(
            os.path.join(upload_dir, "primary_lineage.xlsx")
        )
    if (
        "global_lineage_file" in request.files
        and request.files["global_lineage_file"].filename
    ):
        request.files["global_lineage_file"].save(
            os.path.join(upload_dir, "global_lineage.xlsx")
        )
    return redirect(url_for("dashboard"))


@app.route("/dashboard")
def dashboard():
    project_name = session.get("active_project")
    if not project_name:
        return redirect(url_for("index"))

    cde_path = os.path.join(
        get_project_dir(project_name, "uploads"), "reporting_layers.xlsx"
    )
    if not os.path.exists(cde_path):
        return redirect(url_for("index"))

    with file_lock:
        xls = pd.ExcelFile(cde_path, engine="openpyxl")
        df = pd.read_excel(xls, "Sheet1").fillna("")
        df_done = (
            pd.read_excel(xls, "Done").fillna("")
            if "Done" in xls.sheet_names
            else pd.DataFrame()
        )

    done_keys = {
        f"{str(row.get('CDE name', '')).strip()}|{str(row.get('Current app code', '')).strip()}|{str(row.get('Current table/file name', '')).strip()}|{str(row.get('Current column/field name', '')).strip()}"
        for _, row in df_done.iterrows()
    }

    if "Target App Codes" not in df.columns:
        df["Target App Codes"] = ""
    state_db = load_state(project_name)

    grouped = df.groupby("CDE name")
    cdes = []
    metrics = {
        "total_cdes": len(grouped),
        "completed_cdes": 0,
        "total_instances": len(df),
        "completed_instances": 0,
        "needs_review": 0,
        "processing": 0,
        "auto_resolvable": 0,
    }

    for cde_name, group in grouped:
        target_apps = set()
        for apps in group["Target App Codes"]:
            target_str = str(apps).strip()
            if target_str.lower() not in ["nan", ""]:
                target_apps.update(
                    [
                        app.strip().upper()
                        for app in target_str.split(",")
                        if app.strip()
                    ]
                )

        instances, cde_completed = [], True
        for _, row in group.iterrows():
            app_code = str(row.get("Current app code", "")).strip()
            table = str(row.get("Current table/file name", "")).strip()
            col = str(row.get("Current column/field name", "")).strip()
            key = f"{cde_name}|{app_code}|{table}|{col}"

            if key in done_keys:
                status, metrics["completed_instances"] = (
                    "Completed",
                    metrics["completed_instances"] + 1,
                )
            else:
                cde_completed = False
                job_state = state_db.get(key, {}).get("status", "Pending")
                if job_state == "NEEDS_REVIEW":
                    status, metrics["needs_review"] = (
                        "Needs Review",
                        metrics["needs_review"] + 1,
                    )
                elif job_state == "PROCESSING":
                    status, metrics["processing"] = (
                        "Processing",
                        metrics["processing"] + 1,
                    )
                else:
                    status, metrics["auto_resolvable"] = (
                        "Pending",
                        metrics["auto_resolvable"] + 1,
                    )

            instances.append(
                {
                    "app": app_code,
                    "table": table,
                    "column": col,
                    "status": status,
                    "key": key,
                    "reason": state_db.get(key, {}).get("reason", ""),
                }
            )

        if cde_completed:
            metrics["completed_cdes"] += 1
        cdes.append(
            {
                "name": cde_name,
                "target_apps": list(target_apps),
                "instances": instances,
                "completed": cde_completed,
            }
        )

    return render_template("dashboard.html", metrics=metrics, cdes=cdes)


@app.route("/api/edit_trace", methods=["POST"])
def edit_trace():
    project = session["active_project"]
    data = request.json
    cde_name, app_code, table, col = (
        data["cde_name"],
        data["app"],
        data["table"],
        data["column"],
    )
    key = f"{cde_name}|{app_code}|{table}|{col}"
    cde_path = os.path.join(
        get_project_dir(project, "uploads"), "reporting_layers.xlsx"
    )

    with file_lock:
        df_done = pd.read_excel(cde_path, sheet_name="Done", engine="openpyxl")
        df_done = df_done[
            ~(
                (df_done["CDE name"] == cde_name)
                & (df_done["Current app code"] == app_code)
                & (df_done["Current table/file name"] == table)
                & (df_done["Current column/field name"] == col)
            )
        ]
        with pd.ExcelWriter(
            cde_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            df_done.to_excel(writer, sheet_name="Done", index=False)

    state = load_state(project)
    if key in state:
        state[key]["status"] = "NEEDS_REVIEW"
        state[key]["reason"] = "Editing Trace Configuration"
        save_state(project, state)
    return jsonify({"status": "success"})


@app.route("/api/retrace", methods=["POST"])
def api_retrace():
    project = session["active_project"]
    data = request.json
    cde_name, app_code, table, col = (
        data["cde_name"],
        data["app"],
        data["table"],
        data["column"],
    )
    cde_path = os.path.join(
        get_project_dir(project, "uploads"), "reporting_layers.xlsx"
    )

    with file_lock:
        df_done = pd.read_excel(cde_path, sheet_name="Done", engine="openpyxl")
        df_done = df_done[
            ~(
                (df_done["CDE name"] == cde_name)
                & (df_done["Current app code"] == app_code)
                & (df_done["Current table/file name"] == table)
                & (df_done["Current column/field name"] == col)
            )
        ]
        with pd.ExcelWriter(
            cde_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            df_done.to_excel(writer, sheet_name="Done", index=False)

    delete_state_for_key(project, f"{cde_name}|{app_code}|{table}|{col}")
    return jsonify({"status": "success"})


@app.route("/trace/<cde_name>/<app_code>/<table_name>/<column_name>")
def trace(cde_name, app_code, table_name, column_name):
    project = session["active_project"]
    target_apps = request.args.get("targets", "")
    resume_auto = request.args.get("resume_auto", "false")
    key = f"{cde_name}|{app_code}|{table_name}|{column_name}"

    session.permanent = True
    if "lineage_stack" not in session:
        session["lineage_stack"] = []

    state = load_state(project)
    resume_stack = state.get(key, {}).get("stack", [])

    return render_template(
        "trace.html",
        cde_name=cde_name,
        app_code=app_code,
        table_name=table_name,
        column_name=column_name,
        target_apps=target_apps,
        resume_stack=json.dumps(resume_stack),
        resume_auto=resume_auto,
    )


# --- NEW: SAVE MIDWAY ---
@app.route("/api/save_midway", methods=["POST"])
def save_midway():
    project = session["active_project"]
    data = request.json
    cde_name, stack = data["cde_name"], data["stack"]
    target_key = f"{cde_name}|{stack[0]['app']}|{stack[0]['original_table']}|{stack[0]['original_col']}"

    update_state_for_key(
        project,
        target_key,
        {
            "status": "NEEDS_REVIEW",
            "stack": stack,
            "reason": "Midway Save by User (Paused)",
            "timestamp": str(datetime.now()),
        },
    )
    return jsonify({"status": "success"})


@app.route("/api/undo", methods=["POST"])
def undo():
    if "lineage_stack" not in session or len(session["lineage_stack"]) == 0:
        return jsonify({"error": "Nothing to undo"}), 400
    session["lineage_stack"].pop()
    session.modified = True
    return jsonify({"success": True, "stack_length": len(session["lineage_stack"])})


def search_dataframe(df, app_code, table, col):
    app_code, table, col = (
        str(app_code).strip().upper(),
        str(table).strip().upper(),
        str(col).strip().upper(),
    )
    if (
        not app_code
        or app_code == "NAN"
        or not table
        or table == "NAN"
        or not col
        or col == "NAN"
    ):
        return pd.DataFrame(), "No Match"

    for c in [
        "Current app code",
        "Current table/file name",
        "Current column/field name",
    ]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].astype(str).str.strip().str.upper().replace("NAN", "")

    df["clean_app"] = df["Current app code"]
    df["clean_table"] = df["Current table/file name"]
    df["clean_col"] = df["Current column/field name"]

    exact = df[
        (df["clean_app"] == app_code)
        & (df["clean_table"] == table)
        & (df["clean_col"] == col)
    ]
    if not exact.empty:
        return exact, "Exact Match"

    df["flex_table"] = df["clean_table"].apply(flexible_normalize)
    df["flex_col"] = df["clean_col"].apply(flexible_normalize)
    normalized_table, normalized_col = (
        flexible_normalize(table),
        flexible_normalize(col),
    )

    flex = df[
        (df["clean_app"] == app_code)
        & (df["flex_table"] == normalized_table)
        & (df["flex_col"] == normalized_col)
    ]
    if not flex.empty:
        return flex, "Flexible Match"

    app_filtered = df[df["clean_app"] == app_code]
    if app_filtered.empty:
        return pd.DataFrame(), "No Match"

    choices_list = [
        f"{r['clean_table']} | {r['clean_col']}"
        for _, r in app_filtered[["clean_table", "clean_col"]]
        .drop_duplicates()
        .iterrows()
    ]
    close_matches = difflib.get_close_matches(
        f"{table} | {col}", choices_list, n=5, cutoff=0.3
    )

    if close_matches:
        parts = close_matches[0].split(" | ")
        matched_table, matched_col = parts[0], parts[1] if len(parts) > 1 else ""
        fuzzy = df[
            (df["clean_app"] == app_code)
            & (df["clean_table"] == matched_table)
            & (df["clean_col"] == matched_col)
        ]
        if not fuzzy.empty:
            return fuzzy, "Fuzzy Match"

    return pd.DataFrame(), "No Match"


def execute_search_api_logic(project_name, cde_name, app_code, table, col, target_apps):
    primary_path = os.path.join(
        get_project_dir(project_name, "uploads"), "primary_lineage.xlsx"
    )
    global_path = os.path.join(
        get_project_dir(project_name, "uploads"), "global_lineage.xlsx"
    )
    unique_groups = []
    file_memory = get_file_memory(project_name)

    def process_file(file_path, source_label):
        if not os.path.exists(file_path):
            return
        with file_lock:
            df = pd.read_excel(file_path, engine="openpyxl").fillna("")

        matches, match_type = search_dataframe(df, app_code, table, col)
        if matches.empty:
            return

        matches = matches.copy()
        matches["group_id"] = (
            matches["file"].astype(str)
            + "||"
            + matches.get("Report name", "N/A").astype(str)
            + "||"
            + matches.get("CDE name", "N/A").astype(str)
        )

        for gid, group in matches.groupby("group_id"):
            rep_row = group.iloc[0]
            distinct_sources = sorted(
                group["Source app code"]
                .dropna()
                .astype(str)
                .str.strip()
                .str.upper()
                .unique()
                .tolist()
            )
            if not distinct_sources:
                continue

            file_name_str = str(rep_row.get("file", "N/A"))
            memory_bonus = file_memory.get(file_name_str, 0) * 150000

            score = (
                1000000
                if cde_name.upper() in str(rep_row.get("CDE name", "")).upper()
                else 0
            )
            if any(s in [t.upper() for t in target_apps] for s in distinct_sources):
                score += 500000
            score += {
                "Exact Match": 100000,
                "Flexible Match": 50000,
                "Fuzzy Match": 25000,
            }.get(match_type, 0)
            score += memory_bonus

            raw_rows = []
            for _, row in group.iterrows():
                src_app = str(row.get("Source app code", "")).strip().upper()
                if src_app and src_app != "NAN":
                    raw_rows.append(
                        {
                            "next_app": src_app,
                            "next_table": str(
                                row.get("Source table/file name", "")
                            ).strip(),
                            "next_col": str(
                                row.get("Source column/field name", "")
                            ).strip(),
                            "searched_table": table,
                            "found_table": str(
                                rep_row.get("Current table/file name", "")
                            ).strip(),
                            "searched_col": col,
                            "found_col": str(
                                rep_row.get("Current column/field name", "")
                            ).strip(),
                            "match_type": match_type,
                            "raw_row_data": clean_raw_data_for_json(
                                row.drop(
                                    labels=[
                                        "clean_app",
                                        "clean_table",
                                        "clean_col",
                                        "flex_table",
                                        "flex_col",
                                        "group_id",
                                    ],
                                    errors="ignore",
                                ).to_dict()
                            ),
                        }
                    )

            if raw_rows:
                unique_groups.append(
                    {
                        "source": source_label,
                        "score": score,
                        "match_type": match_type,
                        "file_path": file_name_str,
                        "report_name": str(rep_row.get("Report name", "N/A")),
                        "cde_found": str(rep_row.get("CDE name", "")),
                        "found_table": str(rep_row.get("Current table/file name", "")),
                        "found_col": str(rep_row.get("Current column/field name", "")),
                        "source_count": len(raw_rows),
                        "distinct_sources": distinct_sources,
                        "raw_rows": raw_rows,
                    }
                )

    process_file(primary_path, "Primary Report")
    if not unique_groups:
        process_file(global_path, "Global Lineage")
    unique_groups.sort(key=lambda x: x["score"], reverse=True)
    return unique_groups


@app.route("/api/search", methods=["POST"])
def api_search():
    data = request.json
    results = execute_search_api_logic(
        session["active_project"],
        data["cde_name"],
        data["app"],
        data["table"],
        data["column"],
        data["target_apps"],
    )
    return jsonify({"candidates": results})


def write_lineage_to_files(project_name, cde_name, stack, is_dead_end):
    with file_lock:
        cde_path = os.path.join(
            get_project_dir(project_name, "uploads"), "reporting_layers.xlsx"
        )
        with pd.ExcelWriter(
            cde_path, engine="openpyxl", mode="a", if_sheet_exists="overlay"
        ) as writer:
            df_done = (
                pd.read_excel(cde_path, sheet_name="Done", engine="openpyxl").fillna("")
                if "Done" in pd.ExcelFile(cde_path, engine="openpyxl").sheet_names
                else pd.DataFrame()
            )
            new_done = pd.DataFrame(
                [
                    {
                        "Report name": "Lineage Engine",
                        "CDE name": cde_name,
                        "Current app code": stack[0]["app"],
                        "Current table/file name": stack[0]["original_table"],
                        "Current column/field name": stack[0]["original_col"],
                    }
                ]
            )
            df_done = pd.concat([df_done, new_done], ignore_index=True)
            df_done.to_excel(writer, sheet_name="Done", index=False)

        output_rows = []
        seen_lineages = set()

        for i in range(len(stack)):
            curr = stack[i]
            if i < len(stack) - 1:
                next_node = stack[i + 1]
                s_app = next_node["app"]
                s_table = next_node.get("reconciled_table", next_node["original_table"])
                s_col = next_node.get("reconciled_col", next_node["original_col"])
            else:
                s_app, s_table, s_col = "", "", ""

            dest = stack[i - 1]["app"] if i > 0 else ""
            lineage_key = f"{curr['app']}|{curr.get('reconciled_table', curr['original_table'])}|{curr.get('reconciled_col', curr['original_col'])}->{s_app}|{s_table}|{s_col}"
            if lineage_key in seen_lineages:
                continue
            seen_lineages.add(lineage_key)

            label = f"↓ {curr['app']}" if i == 0 else f"↑ {dest}   ↓ {curr['app']}"
            data_row = curr.get("raw_data", {}).copy()

            c_name = str(data_row.get("Current app name", "")).strip()
            s_name = (
                str(next_node.get("raw_data", {}).get("Source app name", "")).strip()
                if i < len(stack) - 1
                else ""
            )

            data_row.update(
                {
                    "Label": label,
                    "CDE name": cde_name,
                    "Destination to (App Code)": dest,
                    "Current app code": curr["app"],
                    "Current app name": c_name if c_name.lower() != "nan" else "",
                    "Current table/file name": curr.get(
                        "reconciled_table", curr["original_table"]
                    ),
                    "Current column/field name": curr.get(
                        "reconciled_col", curr["original_col"]
                    ),
                    "Source app code": s_app,
                    "Source app name": s_name if s_name.lower() != "nan" else "",
                    "Source table/file name": s_table,
                    "Source column/field name": s_col,
                    "QA Notes": curr.get("notes", ""),
                }
            )
            output_rows.append(data_row)

        output_file = os.path.join(
            get_project_dir(project_name, "outputs"), get_safe_filename(cde_name)
        )
        df_out = pd.DataFrame(output_rows)
        if os.path.exists(output_file):
            existing_df = pd.read_excel(output_file, engine="openpyxl").fillna("")
            df_combined = pd.concat(
                [existing_df, df_out], ignore_index=True
            ).drop_duplicates(
                subset=[
                    "Current app code",
                    "Current table/file name",
                    "Current column/field name",
                    "Source app code",
                ],
                keep="first",
            )
            df_out = df_combined

        cols = df_out.columns.tolist()
        if "Label" in cols:
            cols.insert(0, cols.pop(cols.index("Label")))
            df_out = df_out[cols]
        df_out.to_excel(output_file, index=False)


@app.route("/api/save_lineage", methods=["POST"])
def save_lineage():
    project = session["active_project"]
    data = request.json
    write_lineage_to_files(
        project, data["cde_name"], data["stack"], data.get("is_dead_end", False)
    )

    if len(data["stack"]) > 1 and "file" in data["stack"][-1].get("raw_data", {}):
        track_file_usage(project, data["stack"][-1]["raw_data"]["file"])

    target_key = f"{data['cde_name']}|{data['stack'][0]['app']}|{data['stack'][0]['original_table']}|{data['stack'][0]['original_col']}"
    update_state_for_key(
        project,
        target_key,
        {
            "status": "RESOLVED",
            "stack": data["stack"],
            "timestamp": str(datetime.now()),
        },
    )

    if "lineage_stack" in session:
        session["lineage_stack"] = []
        session.modified = True
    return jsonify({"status": "success", "redirect": url_for("dashboard")})


def process_instance_background(
    project_name, cde_name, target_key, initial_raw_data=None
):
    try:
        cde_path = os.path.join(
            get_project_dir(project_name, "uploads"), "reporting_layers.xlsx"
        )

        with file_lock:
            xls = pd.ExcelFile(cde_path, engine="openpyxl")
            df = pd.read_excel(xls, "Sheet1").fillna("")

        target_apps = []
        for _, row in df[df["CDE name"] == cde_name].iterrows():
            target_str = str(row.get("Target App Codes", ""))
            if target_str.lower() != "nan":
                target_apps.extend(
                    [t.strip().upper() for t in target_str.split(",") if t.strip()]
                )

        with file_lock:
            state = load_state(project_name)
            stack = state.get(target_key, {}).get("stack", [])

        if not stack:
            _, app_code, table, col = target_key.split("|")
            stack = [
                {
                    "app": app_code,
                    "original_table": table,
                    "original_col": col,
                    "reconciled_table": table,
                    "reconciled_col": col,
                    "raw_data": initial_raw_data or {},
                }
            ]

        update_state_for_key(
            project_name,
            target_key,
            {"status": "PROCESSING", "stack": stack, "timestamp": str(datetime.now())},
        )

        stacks_to_process = [stack]
        completed_paths = []

        while stacks_to_process:
            current_stack = stacks_to_process.pop(0)
            current_node = current_stack[-1]

            candidates = execute_search_api_logic(
                project_name,
                cde_name,
                current_node["app"],
                current_node.get(
                    "original_table", current_node.get("reconciled_table")
                ),
                current_node.get("original_col", current_node.get("reconciled_col")),
                target_apps,
            )

            if len(candidates) == 0:
                reason = "End of Discovered Lineage (Dead End). Please confirm or Manually Link."
                update_state_for_key(
                    project_name,
                    target_key,
                    {
                        "status": "NEEDS_REVIEW",
                        "stack": current_stack,
                        "reason": reason,
                        "timestamp": str(datetime.now()),
                    },
                )
                return

            best_cand = candidates[0]

            if best_cand["match_type"] != "Exact Match" or len(candidates) > 1:
                reason = f"{best_cand['match_type']} or Multiple Files detected. Human review required to confirm."
                update_state_for_key(
                    project_name,
                    target_key,
                    {
                        "status": "NEEDS_REVIEW",
                        "stack": current_stack,
                        "reason": reason,
                        "timestamp": str(datetime.now()),
                    },
                )
                return

            track_file_usage(project_name, best_cand.get("file_path"))

            for src in best_cand["raw_rows"]:
                new_stack = copy.deepcopy(current_stack)
                new_stack[-1]["reconciled_table"] = best_cand["found_table"]
                new_stack[-1]["reconciled_col"] = (
                    src["found_col"]
                    if src.get("found_col")
                    else current_node.get("original_col", "")
                )
                new_stack[-1]["raw_data"] = src["raw_row_data"]

                next_app = src["next_app"]
                new_node = {
                    "app": next_app,
                    "original_table": src["next_table"],
                    "original_col": src["next_col"],
                    "reconciled_table": src["next_table"],
                    "reconciled_col": src["next_col"],
                    "raw_data": src["raw_row_data"],
                }
                new_stack.append(new_node)

                if next_app in target_apps:
                    new_stack[-1]["notes"] = "Target Reached (Auto-Resolved)"
                    completed_paths.append((new_stack, False))
                else:
                    stacks_to_process.append(new_stack)

        for path, is_dead_end in completed_paths:
            write_lineage_to_files(project_name, cde_name, path, is_dead_end)

        update_state_for_key(
            project_name,
            target_key,
            {
                "status": "RESOLVED",
                "stack": current_stack,
                "timestamp": str(datetime.now()),
            },
        )

    except Exception as e:
        error_msg = f"System Error: {str(e)}"
        print(f"CRASH in background task for {target_key}: {traceback.format_exc()}")
        update_state_for_key(
            project_name,
            target_key,
            {
                "status": "NEEDS_REVIEW",
                "stack": stack,
                "reason": error_msg,
                "timestamp": str(datetime.now()),
            },
        )


@app.route("/api/auto_resolve/start", methods=["POST"])
def start_auto_resolve():
    project = session["active_project"]
    cde_path = os.path.join(
        get_project_dir(project, "uploads"), "reporting_layers.xlsx"
    )

    with file_lock:
        xls = pd.ExcelFile(cde_path, engine="openpyxl")
        df = pd.read_excel(xls, "Sheet1").fillna("")
        df_done = (
            pd.read_excel(xls, "Done").fillna("")
            if "Done" in xls.sheet_names
            else pd.DataFrame()
        )

    done_keys = {
        f"{str(row.get('CDE name', '')).strip()}|{str(row.get('Current app code', '')).strip()}|{str(row.get('Current table/file name', '')).strip()}|{str(row.get('Current column/field name', '')).strip()}"
        for _, row in df_done.iterrows()
    }

    state = load_state(project)
    for _, row in df.iterrows():
        c_name = str(row.get("CDE name", "")).strip()
        c_app = str(row.get("Current app code", "")).strip()
        c_tab = str(row.get("Current table/file name", "")).strip()
        c_col = str(row.get("Current column/field name", "")).strip()

        key = f"{c_name}|{c_app}|{c_tab}|{c_col}"

        if key not in done_keys and state.get(key, {}).get("status") != "NEEDS_REVIEW":
            raw_data = row.to_dict()
            executor.submit(process_instance_background, project, c_name, key, raw_data)

    return jsonify({"status": "Bulk job started"})


@app.route("/api/auto_resolve/single", methods=["POST"])
def start_single_auto_resolve():
    project = session["active_project"]
    data = request.json
    target_key = f"{data['cde_name']}|{data['app']}|{data['table']}|{data['column']}"

    with file_lock:
        df = pd.read_excel(
            os.path.join(get_project_dir(project, "uploads"), "reporting_layers.xlsx"),
            "Sheet1",
            engine="openpyxl",
        ).fillna("")
        row = df[
            (df["CDE name"] == data["cde_name"])
            & (df["Current app code"] == data["app"])
        ].iloc[0]
        raw_data = row.to_dict()

    executor.submit(
        process_instance_background, project, data["cde_name"], target_key, raw_data
    )
    return jsonify({"status": "Single instance job started"})


@app.route("/api/auto_resolve/resume", methods=["POST"])
def resume_auto_resolve():
    project = session["active_project"]
    data = request.json
    cde_name, stack = data["cde_name"], data["stack"]
    target_key = f"{cde_name}|{stack[0]['app']}|{stack[0]['original_table']}|{stack[0]['original_col']}"

    if len(stack) > 1 and "file" in stack[-1].get("raw_data", {}):
        track_file_usage(project, stack[-1]["raw_data"]["file"])

    update_state_for_key(
        project,
        target_key,
        {"status": "PROCESSING", "stack": stack, "timestamp": str(datetime.now())},
    )
    executor.submit(process_instance_background, project, cde_name, target_key)
    return jsonify({"status": "Resumed in background"})


@app.route("/view_lineage/<path:cde_name>")
def view_lineage(cde_name):
    project = session.get("active_project")
    output_file = os.path.join(
        get_project_dir(project, "outputs"), get_safe_filename(cde_name)
    )
    if not os.path.exists(output_file):
        return redirect(url_for("dashboard"))
    return render_template("view_lineage.html", cde_name=cde_name)


@app.route("/api/graph/<path:cde_name>")
def api_graph(cde_name):
    try:
        project = session.get("active_project")
        output_file = os.path.join(
            get_project_dir(project, "outputs"), get_safe_filename(cde_name)
        )
        if not os.path.exists(output_file):
            return jsonify({"error": "File not found"})

        with file_lock:
            df = pd.read_excel(output_file, engine="openpyxl").fillna("")

        inst_nodes, inst_edges, app_nodes, app_edges, app_meta = {}, {}, {}, {}, {}

        for _, row in df.iterrows():
            if not str(row.get("CDE name", "")).strip():
                continue

            c_app = str(row.get("Current app code", "")).strip().upper()
            if not c_app or c_app.lower() == "nan":
                continue

            c_name_val = str(row.get("Current app name", "")).strip()
            c_name = c_name_val if c_name_val.lower() not in ["nan", ""] else c_app
            c_tab = str(row.get("Current table/file name", "")).strip()
            c_col = str(row.get("Current column/field name", "")).strip()

            s_app = str(row.get("Source app code", "")).strip().upper()
            s_name_val = str(row.get("Source app name", "")).strip()
            s_name = s_name_val if s_name_val.lower() not in ["nan", ""] else s_app
            s_tab = str(row.get("Source table/file name", "")).strip()
            s_col = str(row.get("Source column/field name", "")).strip()

            is_manual = (
                "manual" in str(row.get("Manually entered/Derived", "")).lower()
                or "manual" in str(row.get("Created/Sourced", "")).lower()
            )
            is_transformed = (
                "transform" in str(row.get("Transformed/Passed Through", "")).lower()
                or str(row.get("Transformation Logic", "")).strip() != ""
            )
            is_filtered = (
                "filter" in str(row.get("Filtered", "")).lower()
                or str(row.get("Filtration Logic", "")).strip() != ""
            )

            if c_app not in app_meta:
                app_meta[c_app] = {
                    "code": c_app,
                    "name": c_name,
                    "manual": False,
                    "transformed": False,
                    "filtered": False,
                    "instances": [],
                }
            app_meta[c_app]["manual"] = app_meta[c_app]["manual"] or is_manual
            app_meta[c_app]["transformed"] = (
                app_meta[c_app]["transformed"] or is_transformed
            )
            app_meta[c_app]["filtered"] = app_meta[c_app]["filtered"] or is_filtered

            inst_entry = {"table": c_tab, "col": c_col}
            if inst_entry not in app_meta[c_app]["instances"]:
                app_meta[c_app]["instances"].append(inst_entry)

            c_id = f"{c_app}|{c_tab.upper()}|{c_col.upper()}"
            if c_id not in inst_nodes:
                inst_nodes[c_id] = {
                    "id": c_id,
                    "label": f"🏢 Name: {c_name}\n🔑 Code: {c_app}\n📁 Table: {c_tab}\n📌 Col: {c_col}",
                    "shape": "box",
                    "color": {"background": "#F3F4F6", "border": "#005DAA"},
                    "font": {"face": "monospace", "align": "left"},
                }

            if s_app and s_app.lower() != "nan":
                s_id = f"{s_app}|{s_tab.upper()}|{s_col.upper()}"
                if s_id not in inst_nodes:
                    inst_nodes[s_id] = {
                        "id": s_id,
                        "label": f"🏢 Name: {s_name}\n🔑 Code: {s_app}\n📁 Table: {s_tab}\n📌 Col: {s_col}",
                        "shape": "box",
                        "color": {"background": "#E2E8F0", "border": "#64748B"},
                        "font": {"face": "monospace", "align": "left"},
                    }

                e_id = f"{s_id}->{c_id}"
                inst_edges[e_id] = {
                    "from": s_id,
                    "to": c_id,
                    "arrows": "to",
                    "color": {"color": "#005DAA"},
                }

                if s_app not in app_meta:
                    app_meta[s_app] = {
                        "code": s_app,
                        "name": s_name,
                        "manual": False,
                        "transformed": False,
                        "filtered": False,
                        "instances": [],
                    }

                e_app_id = f"{s_app}->{c_app}"
                if e_app_id not in app_edges:
                    app_edges[e_app_id] = {
                        "from": s_app,
                        "to": c_app,
                        "arrows": "to",
                        "color": {"color": "#10B981"},
                        "width": 2,
                    }

        src_inst_ids = {e["from"] for e in inst_edges.values()}
        for n_id, n in inst_nodes.items():
            if n_id not in src_inst_ids:
                n["color"], n["borderWidth"] = (
                    {"background": "#E0F2FE", "border": "#0284C7"},
                    3,
                )

        src_app_ids = {e["from"] for e in app_edges.values()}
        for app_key, meta in app_meta.items():
            app_nodes[app_key] = {
                "id": app_key,
                "meta": meta,
                "is_target": app_key not in src_app_ids,
            }

        return jsonify(
            {
                "instance_graph": {
                    "nodes": list(inst_nodes.values()),
                    "edges": list(inst_edges.values()),
                },
                "app_graph": {
                    "nodes": list(app_nodes.values()),
                    "edges": list(app_edges.values()),
                },
            }
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- NEW: TOOLS ROUTE ---
@app.route("/tools")
def tools():
    project = session.get("active_project")
    if not project:
        return redirect(url_for("index"))

    out_dir = get_project_dir(project, "outputs")
    files = [f for f in os.listdir(out_dir) if f.endswith(".xlsx")]
    return render_template("tools.html", files=files)


@app.route("/api/tools/<action>", methods=["POST"])
def execute_tool(action):
    # This is a placeholder for where you plug in your future Python scripts!
    # request.form will hold the data from the frontend forms.
    return jsonify(
        {
            "status": "success",
            "message": f"{action.capitalize()} script executed successfully!",
        }
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
