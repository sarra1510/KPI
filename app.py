"""
Flask web interface for the KPI Calculator.
Supports both Scrum (3-sheet: Start, End Sprint, Worklogs) and Kanban (single Worklogs sheet)
formats. The file format is auto-detected on upload.
Provides a drag-and-drop file upload and a KPI dashboard in the browser.
"""

import json
import os
import uuid
from datetime import datetime

import pandas as pd
from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_from_directory,
)

from kpi_calculator import (
    find_sheet_name,
    find_header_row,
    clean_dataframe,
    find_key_column,
    calc_capacity_utilization,
    calc_throughput,
    calc_wip_end_sprint,
    find_no_estimation,
    find_no_tempo,
    calc_resolution_time,
    calc_resolution_time_kanban,
    calc_time_per_project,
    calc_kpi_per_user,
    detect_mode,
    _deduplicate_worklogs,
    _filter_empty_keys,
    SHEET_START,
    SHEET_END,
    SHEET_WORKLOG,
)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "sprint-kpi-web-secret")

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {"xlsx", "xls"}
HISTORY_FILE = os.path.join(UPLOAD_FOLDER, "history.json")
HISTORY_MAX = 5
HOURS_PER_DAY = 8


def load_history():
    """Load upload history from JSON file."""
    if not os.path.isfile(HISTORY_FILE):
        return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


def save_history(entry):
    """Prepend a new entry to the history file, keeping only the last HISTORY_MAX entries."""
    history = load_history()
    history.insert(0, entry)
    history = history[:HISTORY_MAX]
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def load_data_web(filepath):
    """
    Non-interactive version of load_data() for the web interface.
    Auto-detects Kanban (single Worklogs sheet) vs Scrum (3 sheets: Start, End Sprint, Worklogs).
    Raises ValueError instead of calling input() when sheets cannot be detected.

    Returns:
        tuple: (mode, df_start, df_end, df_worklog)
            - mode: "scrum" or "kanban"
            - df_start: Start sheet (Scrum) or empty DataFrame (Kanban)
            - df_end: End Sprint sheet (Scrum) or deduplicated tickets from Worklogs (Kanban)
            - df_worklog: Worklogs sheet (both modes)
    """
    try:
        with pd.ExcelFile(filepath) as xls:
            available_sheets = xls.sheet_names
    except Exception as e:
        raise ValueError(f"Impossible d'ouvrir le fichier Excel : {e}")

    mode = detect_mode(available_sheets)

    sheet_worklog = find_sheet_name(
        available_sheets,
        SHEET_WORKLOG,
        keywords=["worklog", "worklogs", "tempo"],
    )

    if not sheet_worklog:
        raise ValueError(
            f"Feuille Worklogs non détectée. "
            f"Feuilles disponibles : {', '.join(available_sheets)}"
        )

    if mode == "kanban":
        h_wl = find_header_row(filepath, sheet_worklog)
        try:
            df_worklog = pd.read_excel(filepath, sheet_name=sheet_worklog, header=h_wl)
        except Exception as e:
            raise ValueError(f"Erreur de lecture du fichier : {e}")

        df_worklog = clean_dataframe(df_worklog)
        df_worklog = _filter_empty_keys(df_worklog, ["Issue Key", "Issue key", "Key"], "Worklogs")

        df_tickets = _deduplicate_worklogs(df_worklog)
        return mode, pd.DataFrame(), df_tickets, df_worklog

    # Scrum mode: require all 3 sheets
    sheet_start = find_sheet_name(
        available_sheets,
        SHEET_START,
        keywords=["start", "début", "debut", "démarrage", "demarrage"],
        exclude_keywords=["end", "fin"],
    )
    sheet_end = find_sheet_name(
        available_sheets,
        SHEET_END,
        keywords=["end", "fin", "sprint end", "end sprint"],
    )

    missing = []
    if not sheet_start:
        missing.append("Start (début sprint)")
    if not sheet_end:
        missing.append("End Sprint (fin sprint)")

    if missing:
        raise ValueError(
            f"Feuilles non détectées : {', '.join(missing)}. "
            f"Feuilles disponibles : {', '.join(available_sheets)}"
        )

    h_start = find_header_row(filepath, sheet_start)
    h_end = find_header_row(filepath, sheet_end)
    h_wl = find_header_row(filepath, sheet_worklog)

    try:
        df_start = pd.read_excel(filepath, sheet_name=sheet_start, header=h_start)
        df_end = pd.read_excel(filepath, sheet_name=sheet_end, header=h_end)
        df_worklog = pd.read_excel(filepath, sheet_name=sheet_worklog, header=h_wl)
    except Exception as e:
        raise ValueError(f"Erreur de lecture du fichier : {e}")

    df_start = clean_dataframe(df_start)
    df_end = clean_dataframe(df_end)
    df_worklog = clean_dataframe(df_worklog)

    # Remove rows where Key is empty
    df_start = _filter_empty_keys(df_start, ["Key", "Issue Key", "key", "Clé"], "Start")
    df_end = _filter_empty_keys(df_end, ["Key", "Issue Key", "key", "Clé"], "End Sprint")
    df_worklog = _filter_empty_keys(df_worklog, ["Issue Key", "Issue key", "Key"], "Worklogs")

    return mode, df_start, df_end, df_worklog


def find_key_column_web(df, sheet_name):
    """
    Non-interactive version of find_key_column() for the web interface.
    Raises ValueError instead of calling input() when column cannot be detected.
    """
    import re

    known_names = [
        "Key", "key", "KEY",
        "Issue Key", "Issue key", "issue key", "ISSUE KEY",
        "Issue_Key", "Issue_key",
        "Clé", "clé", "Clé du ticket",
        "Ticket Key", "ticket key",
        "Issue ID", "issue id",
    ]
    for col in known_names:
        if col in df.columns:
            return col

    normalized_targets = ["key", "issue key", "issue_key", "clé", "ticket key", "issue id"]
    for col in df.columns:
        if col.strip().lower() in normalized_targets:
            return col

    jira_pattern = re.compile(r"^[A-Z][A-Z0-9]+-\d+$")
    for col in df.columns:
        sample = df[col].dropna().head(10).astype(str)
        if len(sample) == 0:
            continue
        matches = sample.apply(lambda x: bool(jira_pattern.match(x.strip())))
        if matches.sum() >= min(3, len(sample)):
            return col

    raise ValueError(
        f"Colonne clé Jira non détectée dans la feuille '{sheet_name}'. "
        f"Colonnes disponibles : {', '.join(df.columns.tolist())}"
    )


def df_to_records(df):
    """Convert a DataFrame to a list of dicts for template rendering."""
    if df is None or df.empty:
        return []
    return df.where(pd.notnull(df), other="").to_dict(orient="records")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/calculate", methods=["POST"])
def calculate():
    # Validate file upload
    if "file" not in request.files:
        flash("Aucun fichier sélectionné.", "danger")
        return redirect(url_for("index"))

    file = request.files["file"]
    if file.filename == "":
        flash("Aucun fichier sélectionné.", "danger")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash("Format de fichier invalide. Veuillez uploader un fichier .xlsx ou .xls.", "danger")
        return redirect(url_for("index"))

    # Validate capacity (input is now in days; 1 day = 8 hours)
    try:
        capacity_days = float(request.form.get("capacity", "0"))
        if capacity_days <= 0:
            raise ValueError
        capacity_hours = capacity_days * HOURS_PER_DAY
    except (ValueError, TypeError):
        flash("Capacité équipe invalide. Veuillez entrer un nombre positif.", "danger")
        return redirect(url_for("index"))

    # Save uploaded file with a unique name
    unique_id = uuid.uuid4().hex
    original_filename = file.filename
    original_ext = original_filename.rsplit(".", 1)[1].lower()
    saved_filename = f"{unique_id}.{original_ext}"
    filepath = os.path.join(UPLOAD_FOLDER, saved_filename)
    file.save(filepath)

    try:
        mode, df_start, df_end, df_worklog = load_data_web(filepath)

        key_col_end = find_key_column_web(df_end, SHEET_END if mode == "scrum" else SHEET_WORKLOG)

        capacity_util, total_logged = calc_capacity_utilization(df_worklog, capacity_hours)
        total_logged_days = round(total_logged / HOURS_PER_DAY, 2)
        throughput, throughput_details = calc_throughput(df_end, key_col_end)
        wip_count, wip_details = calc_wip_end_sprint(df_end, key_col_end)
        no_est_count, no_est_details = find_no_estimation(df_end, key_col_end)

        if mode == "kanban":
            no_tempo_count, no_tempo_details = 0, pd.DataFrame()
            avg_resolution_days, resolution_details = calc_resolution_time_kanban(
                df_end, df_worklog, key_col_end
            )
            project_totals_df, project_by_priority_df = calc_time_per_project(df_worklog, None, None)
        else:
            key_col_start = find_key_column_web(df_start, SHEET_START)
            no_tempo_count, no_tempo_details = find_no_tempo(df_end, df_worklog, key_col_end)
            avg_resolution_days, resolution_details = calc_resolution_time(
                df_end, df_worklog, key_col_end
            )
            project_totals_df, project_by_priority_df = calc_time_per_project(
                df_worklog, df_end, key_col_end
            )

        user_list, user_kpi_data = calc_kpi_per_user(df_end, df_worklog, key_col_end, mode=mode)

    except ValueError as e:
        flash(str(e), "danger")
        os.remove(filepath)
        return redirect(url_for("index"))
    except Exception as e:
        flash(f"Erreur inattendue lors du calcul : {e}", "danger")
        os.remove(filepath)
        return redirect(url_for("index"))

    # Generate Excel KPI report
    report_filename = f"{unique_id}_KPI_Report.xlsx"
    report_path = os.path.join(UPLOAD_FOLDER, report_filename)

    kpi_data = pd.DataFrame(
        {
            "KPI": [
                "Capacité équipe (j)",
                "Jours logués (j)",
                "Capacity Utilization (%)",
                "Throughput (tickets résolus)",
                "WIP End Sprint",
                "Tickets sans estimation",
                "Tickets sans tempo",
                "Temps moyen de résolution (j)",
            ],
            "Valeur": [
                capacity_days,
                total_logged_days,
                capacity_util,
                throughput,
                wip_count,
                no_est_count,
                no_tempo_count,
                avg_resolution_days if avg_resolution_days is not None else "N/A",
            ],
        }
    )

    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        kpi_data.to_excel(writer, sheet_name="KPI Summary", index=False)
        if not wip_details.empty:
            wip_details.to_excel(writer, sheet_name="WIP End Sprint", index=False)
        if not no_est_details.empty:
            no_est_details.to_excel(writer, sheet_name="Sans Estimation", index=False)
        if not no_tempo_details.empty:
            no_tempo_details.to_excel(writer, sheet_name="Sans Tempo", index=False)
        if not resolution_details.empty:
            resolution_details.to_excel(writer, sheet_name="Temps de résolution", index=False)
        if not project_totals_df.empty:
            project_totals_df.to_excel(writer, sheet_name="Temps par projet", index=False)
        if not project_by_priority_df.empty:
            project_by_priority_df.to_excel(writer, sheet_name="Temps par projet-priorité", index=False)

    kpis = {
        "mode": mode,
        "capacity_days": capacity_days,
        "capacity_hours": capacity_hours,
        "total_logged": total_logged,
        "total_logged_days": total_logged_days,
        "capacity_util": capacity_util,
        "throughput": throughput,
        "wip_count": wip_count,
        "no_est_count": no_est_count,
        "no_tempo_count": no_tempo_count,
        "avg_resolution_days": avg_resolution_days,
    }

    # Save KPI data to JSON for later viewing
    kpi_data_filename = f"{unique_id}_kpi_data.json"
    kpi_data_path = os.path.join(UPLOAD_FOLDER, kpi_data_filename)
    kpi_json_data = {
        "kpis": kpis,
        "wip_rows": df_to_records(wip_details),
        "no_est_rows": df_to_records(no_est_details),
        "no_tempo_rows": df_to_records(no_tempo_details),
        "resolution_rows": df_to_records(resolution_details),
        "project_totals_rows": df_to_records(project_totals_df),
        "project_priority_rows": df_to_records(project_by_priority_df),
        "user_list": user_list,
        "user_kpi_data": user_kpi_data,
    }
    try:
        with open(kpi_data_path, "w", encoding="utf-8") as f:
            json.dump(kpi_json_data, f, ensure_ascii=False, indent=2)
    except OSError:
        pass

    # Save to upload history
    save_history({
        "filename": original_filename,
        "uploaded_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "capacity": capacity_days,
        "report_filename": report_filename,
        "kpi_data_filename": kpi_data_filename,
    })

    return render_template(
        "results.html",
        kpis=kpis,
        wip_rows=df_to_records(wip_details),
        no_est_rows=df_to_records(no_est_details),
        no_tempo_rows=df_to_records(no_tempo_details),
        resolution_rows=df_to_records(resolution_details),
        project_totals_rows=df_to_records(project_totals_df),
        project_priority_rows=df_to_records(project_by_priority_df),
        user_list=user_list,
        user_kpi_data=user_kpi_data,
        report_filename=report_filename,
    )


@app.route("/download/<filename>")
def download(filename):
    # Security: validate filename matches expected pattern (UUID_KPI_Report.xlsx)
    import re
    if not re.fullmatch(r"[0-9a-f]{32}_KPI_Report\.xlsx", filename):
        return "Fichier non autorisé.", 403
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.isfile(filepath):
        return "Fichier introuvable.", 404
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)


@app.route("/reports")
def reports():
    """Display report history page."""
    history = load_history()
    return render_template("reports.html", history=history)


@app.route("/delete-report/<report_filename>", methods=["POST"])
def delete_report(report_filename):
    """Delete a report from history and remove the file."""
    import re
    # Security: validate filename matches expected pattern
    # uuid.uuid4().hex generates 32 hex characters without hyphens
    if not re.fullmatch(r"[0-9a-f]{32}_KPI_Report\.xlsx", report_filename):
        flash("Fichier non autorisé.", "danger")
        return redirect(url_for("reports"))

    # Load history and find entry to delete
    history = load_history()
    new_history = [entry for entry in history if entry.get("report_filename") != report_filename]

    # If no change, the report was not found
    if len(new_history) == len(history):
        flash("Rapport introuvable.", "danger")
        return redirect(url_for("reports"))

    # Save updated history first - don't delete files if this fails
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(new_history, f, ensure_ascii=False, indent=2)
    except OSError:
        flash("Erreur lors de la suppression.", "danger")
        return redirect(url_for("reports"))

    # Delete the report file if it exists
    filepath = os.path.join(UPLOAD_FOLDER, report_filename)
    if os.path.isfile(filepath):
        try:
            os.remove(filepath)
        except OSError:
            # File deletion failed but history is already updated
            # This is acceptable as the main goal (removing from history) succeeded
            pass

    # Also delete the original uploaded file if it exists
    uuid_part = report_filename.replace("_KPI_Report.xlsx", "")
    for ext in ALLOWED_EXTENSIONS:
        original_file = os.path.join(UPLOAD_FOLDER, f"{uuid_part}.{ext}")
        if os.path.isfile(original_file):
            try:
                os.remove(original_file)
            except OSError:
                # Original file deletion failed but this is not critical
                pass

    # Also delete the KPI data JSON file if it exists
    kpi_data_file = os.path.join(UPLOAD_FOLDER, f"{uuid_part}_kpi_data.json")
    if os.path.isfile(kpi_data_file):
        try:
            os.remove(kpi_data_file)
        except OSError:
            pass

    flash("Rapport supprimé avec succès.", "info")
    return redirect(url_for("reports"))


@app.route("/view-report/<report_filename>")
def view_report(report_filename):
    """View a saved KPI report."""
    import re
    # Security: validate filename matches expected pattern
    if not re.fullmatch(r"[0-9a-f]{32}_KPI_Report\.xlsx", report_filename):
        flash("Fichier non autorisé.", "danger")
        return redirect(url_for("reports"))

    # Extract UUID part and load KPI data
    uuid_part = report_filename.replace("_KPI_Report.xlsx", "")
    kpi_data_filename = f"{uuid_part}_kpi_data.json"
    kpi_data_path = os.path.join(UPLOAD_FOLDER, kpi_data_filename)

    if not os.path.isfile(kpi_data_path):
        flash("Données KPI non disponibles pour ce rapport.", "danger")
        return redirect(url_for("reports"))

    try:
        with open(kpi_data_path, "r", encoding="utf-8") as f:
            kpi_json_data = json.load(f)
    except (json.JSONDecodeError, OSError):
        flash("Erreur lors de la lecture des données KPI.", "danger")
        return redirect(url_for("reports"))

    return render_template(
        "results.html",
        kpis=kpi_json_data.get("kpis", {}),
        wip_rows=kpi_json_data.get("wip_rows", []),
        no_est_rows=kpi_json_data.get("no_est_rows", []),
        no_tempo_rows=kpi_json_data.get("no_tempo_rows", []),
        resolution_rows=kpi_json_data.get("resolution_rows", []),
        project_totals_rows=kpi_json_data.get("project_totals_rows", []),
        project_priority_rows=kpi_json_data.get("project_priority_rows", []),
        user_list=kpi_json_data.get("user_list", []),
        user_kpi_data=kpi_json_data.get("user_kpi_data", {}),
        report_filename=report_filename,
    )


@app.route("/current")
def current_calculation():
    """View the most recent KPI calculation if it exists."""
    history = load_history()
    if not history:
        flash("Aucune analyse KPI disponible. Commencez par importer vos données.", "info")
        return redirect(url_for("index"))

    # Get the most recent report
    latest = history[0]
    report_filename = latest.get("report_filename")
    if not report_filename:
        flash("Aucune analyse KPI disponible.", "info")
        return redirect(url_for("index"))

    return redirect(url_for("view_report", report_filename=report_filename))


if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG", "false").lower() == "true")
