"""
KPI Calculator
Supports two modes:
  - Scrum (legacy): reads an Excel file with 3 sheets (Start, End Sprint, Worklogs)
    and computes sprint KPIs.
  - Kanban: reads an Excel file with a single "Worklogs" sheet and derives all KPIs
    from that sheet by deduplicating entries by Issue Key.

Handles Jira Excel exports with metadata header rows.

KPIs:
  - Capacity Utilization (%)
  - Throughput (tickets resolved)
  - Unplanned Tickets (Scrum only: in End but not in Start)
  - WIP End Sprint / WIP (tickets not Done)
  - Support Load (Scrum only: Unplanned / Throughput)
  - Tickets sans estimation
  - Tickets sans tempo (Scrum only; always 0 in Kanban)
  - Temps moyen de résolution (j)
  - Temps par projet
"""

import pandas as pd
import re
import sys
import os


# --- CONFIG ---
SHEET_START = "Start"
SHEET_END = "End Sprint"
SHEET_WORKLOG = "Worklogs"

HOURS_PER_DAY = 8

DONE_STATUSES = [
    "closed",
    "customer pending",
    "released",
    "canceled",
    "done",
]

HEADER_KEYWORDS = [
    "key", "summary", "status", "issue key", "hours",
    "project", "issue type", "priority", "assignee",
    "resolution", "reporter", "created", "resolved",
    "work date", "username", "full name",
]

def find_header_row(filepath, sheet_name, max_scan=20):
    """Scan first rows to find the real header row in Jira exports."""
    try:
        df_raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None, nrows=max_scan)
    except Exception:
        return 0

    for idx, row in df_raw.iterrows():
        values = [
            str(v).strip().lower().replace('\xa0', ' ')
            for v in row if pd.notna(v) and str(v).strip()
        ]
        matches = sum(1 for v in values if v in HEADER_KEYWORDS)
        if matches >= 3:
            return idx

    return 0

def clean_dataframe(df):
    """Clean column names and drop empty rows."""
    df.columns = [str(c).strip().replace('\xa0', ' ') for c in df.columns]
    df = df.dropna(how='all').reset_index(drop=True)
    return df

def find_key_column(df, sheet_name):
    """
    Auto-detect the issue key column with flexible matching.
    1. Known name variants
    2. Normalized match
    3. Regex pattern (PROJ-123)
    4. Interactive fallback
    """
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

    jira_pattern = re.compile(r'^[A-Z][A-Z0-9]+-\d+$')
    for col in df.columns:
        sample = df[col].dropna().head(10).astype(str)
        if len(sample) == 0:
            continue
        matches = sample.apply(lambda x: bool(jira_pattern.match(x.strip())))
        if matches.sum() >= min(3, len(sample)):
            print(f"   🔍 Colonne clé détectée par pattern Jira: '{col}'")
            return col

    print(f"\n⚠️  Colonne clé non détectée dans '{sheet_name}'.")
    print(f"   Colonnes disponibles:")
    cols = list(df.columns)
    for i, c in enumerate(cols):
        sample_val = df[c].dropna().head(1).values
        sample_str = f" (ex: {sample_val[0]})" if len(sample_val) > 0 else ""
        print(f"      [{i}] {c}{sample_str}")

    while True:
        choice = input(f"\n   Entrez le numéro de la colonne contenant les clés Jira: ").strip()
        if choice.isdigit() and 0 <= int(choice) < len(cols):
            selected = cols[int(choice)]
            print(f"   ✅ Colonne clé → '{selected}'")
            return selected
        elif choice in cols:
            print(f"   ✅ Colonne clé → '{choice}'")
            return choice
        else:
            print(f"   ❌ Choix invalide. Réessayez.")

def find_column(df, candidates):
    """Generic flexible column finder."""
    for col in candidates:
        if col in df.columns:
            return col
    normalized_candidates = [c.strip().lower() for c in candidates]
    for col in df.columns:
        if col.strip().lower() in normalized_candidates:
            return col
    return None

def find_sheet_name(available_sheets, expected_name, keywords, exclude_keywords=None):
    """Find a sheet name using flexible matching."""
    if expected_name in available_sheets:
        return expected_name

    normalized = {s: s.strip().lower().replace('\xa0', ' ') for s in available_sheets}

    target = expected_name.lower()
    for original, norm in normalized.items():
        if norm == target:
            return original

    target_stripped = target.replace(" ", "")
    for original, norm in normalized.items():
        if norm.replace(" ", "") == target_stripped:
            return original

    for original, norm in normalized.items():
        if any(kw in norm for kw in keywords):
            if exclude_keywords and any(ex in norm for ex in exclude_keywords):
                continue
            return original

    return None

def detect_mode(available_sheets):
    """Detect whether the file is Kanban (single Worklogs sheet) or Scrum (3 sheets).

    Returns:
        str: "kanban" if only a Worklogs sheet is present (no Start/End Sprint sheets),
             "scrum" if both Start and End Sprint sheets are detected.
    """
    sheet_start = find_sheet_name(
        available_sheets, SHEET_START,
        keywords=["start", "début", "debut", "démarrage", "demarrage"],
        exclude_keywords=["end", "fin"]
    )
    sheet_end = find_sheet_name(
        available_sheets, SHEET_END,
        keywords=["end", "fin", "sprint end", "end sprint"]
    )
    if sheet_start and sheet_end:
        return "scrum"
    return "kanban"

def load_data(filepath):
    """Load sheets with auto-detection of mode (Kanban vs Scrum), header rows and sheet names.

    Returns:
        tuple: (mode, df_start, df_end, df_worklog)
            - mode: "scrum" or "kanban"
            - df_start: Start sheet (Scrum) or empty DataFrame (Kanban)
            - df_end: End Sprint sheet (Scrum) or deduplicated tickets from Worklogs (Kanban)
            - df_worklog: Worklogs sheet (both modes)
    """
    if not os.path.exists(filepath):
        print(f"❌ Fichier introuvable: {filepath}")
        sys.exit(1)

    try:
        with pd.ExcelFile(filepath) as xls:
            available_sheets = xls.sheet_names
    except Exception as e:
        print(f"❌ Impossible d'ouvrir le fichier: {e}")
        sys.exit(1)

    print(f"   📋 Feuilles trouvées: {available_sheets}")

    mode = detect_mode(available_sheets)
    print(f"   🔍 Mode détecté: {mode.upper()}")

    sheet_worklog = find_sheet_name(
        available_sheets, SHEET_WORKLOG,
        keywords=["worklog", "worklogs", "tempo"]
    )

    if mode == "kanban":
        if not sheet_worklog:
            print(f"❌ Feuille Worklogs introuvable. Disponibles: {available_sheets}")
            sys.exit(1)

        h_wl = find_header_row(filepath, sheet_worklog)
        try:
            df_worklog = pd.read_excel(filepath, sheet_name=sheet_worklog, header=h_wl)
        except Exception as e:
            print(f"❌ Erreur de lecture: {e}")
            sys.exit(1)

        df_worklog = clean_dataframe(df_worklog)
        df_worklog = _filter_empty_keys(df_worklog, ["Issue Key", "Issue key", "Key"], "Worklogs")

        df_tickets = _deduplicate_worklogs(df_worklog)
        df_start = pd.DataFrame()
        print(f"   ✅ Worklogs: {len(df_worklog)} entrées")
        print(f"   ✅ Tickets (dédupliqués): {len(df_tickets)}")
        return mode, df_start, df_tickets, df_worklog

    # Scrum mode: load all 3 sheets
    sheet_start = find_sheet_name(
        available_sheets, SHEET_START,
        keywords=["start", "début", "debut", "démarrage", "demarrage"],
        exclude_keywords=["end", "fin"]
    )
    sheet_end = find_sheet_name(
        available_sheets, SHEET_END,
        keywords=["end", "fin", "sprint end", "end sprint"]
    )

    mapping = {
        "Start (début sprint)": sheet_start,
        "End Sprint (fin sprint)": sheet_end,
        "Worklogs": sheet_worklog,
    }
    missing = []
    for label, found in mapping.items():
        if found:
            print(f"   ✅ {label} → '{found}'")
        else:
            print(f"   ❌ {label} → NON TROUVÉE")
            missing.append(label)

    if missing:
        print(f"\n⚠️  Feuilles non détectées. Disponibles:")
        for i, s in enumerate(available_sheets):
            print(f"      [{i}] {s}")

        sheet_vars = {
            "Start (début sprint)": "start",
            "End Sprint (fin sprint)": "end",
            "Worklogs": "worklog",
        }
        selected = {"start": sheet_start, "end": sheet_end, "worklog": sheet_worklog}

        for label in missing:
            var_key = sheet_vars[label]
            while True:
                choice = input(f"\n   Numéro ou nom pour '{label}': ").strip()
                if choice.isdigit() and 0 <= int(choice) < len(available_sheets):
                    sel = available_sheets[int(choice)]
                elif choice in available_sheets:
                    sel = choice
                else:
                    print(f"   ❌ Invalide.")
                    continue
                selected[var_key] = sel
                print(f"   ✅ {label} → '{sel}'")
                break

        sheet_start = selected["start"]
        sheet_end = selected["end"]
        sheet_worklog = selected["worklog"]

    # Auto-detect header rows
    h_start = find_header_row(filepath, sheet_start)
    h_end = find_header_row(filepath, sheet_end)
    h_wl = find_header_row(filepath, sheet_worklog)

    if h_start > 0:
        print(f"   📄 '{sheet_start}': en-tête ligne {h_start}")
    if h_end > 0:
        print(f"   📄 '{sheet_end}': en-tête ligne {h_end}")
    if h_wl > 0:
        print(f"   📄 '{sheet_worklog}': en-tête ligne {h_wl}")

    try:
        df_start = pd.read_excel(filepath, sheet_name=sheet_start, header=h_start)
        df_end = pd.read_excel(filepath, sheet_name=sheet_end, header=h_end)
        df_worklog = pd.read_excel(filepath, sheet_name=sheet_worklog, header=h_wl)
    except Exception as e:
        print(f"❌ Erreur de lecture: {e}")
        sys.exit(1)

    df_start = clean_dataframe(df_start)
    df_end = clean_dataframe(df_end)
    df_worklog = clean_dataframe(df_worklog)

    df_start = _filter_empty_keys(df_start, ["Key", "Issue Key", "key", "Clé"], "Start")
    df_end = _filter_empty_keys(df_end, ["Key", "Issue Key", "key", "Clé"], "End Sprint")
    df_worklog = _filter_empty_keys(df_worklog, ["Issue Key", "Issue key", "Key"], "Worklogs")

    return mode, df_start, df_end, df_worklog


def _filter_empty_keys(df, key_candidates, sheet_label):
    """Remove rows where the key column is empty."""
    for kc in key_candidates:
        if kc in df.columns:
            before = len(df)
            df = df[df[kc].notna() & (df[kc].astype(str).str.strip() != '')].reset_index(drop=True)
            dropped = before - len(df)
            if dropped > 0:
                print(f"   🧹 '{sheet_label}': supprimé {dropped} lignes vides")
            break
    return df


def _deduplicate_worklogs(df_worklog):
    """Deduplicate worklogs by Issue Key, keeping the row with the most recent Work date.

    This produces a ticket-level DataFrame (one row per ticket) from the worklogs sheet,
    using the latest worklog entry's metadata (status, estimates, etc.) for each ticket.
    """
    key_col = find_column(df_worklog, [
        "Issue Key", "Issue key", "issue key", "ISSUE KEY",
        "Key", "key", "KEY",
    ])
    if key_col is None:
        return df_worklog.copy()

    work_date_col = find_column(df_worklog, [
        "Work date", "work date", "Work Date", "WORK DATE", "Date",
    ])

    if work_date_col is not None:
        df_sorted = df_worklog.copy()
        df_sorted[work_date_col] = pd.to_datetime(df_sorted[work_date_col], errors="coerce")
        df_sorted = df_sorted.sort_values(work_date_col, ascending=True, na_position="first")
        df_tickets = df_sorted.drop_duplicates(subset=[key_col], keep="last").reset_index(drop=True)
    else:
        df_tickets = df_worklog.drop_duplicates(subset=[key_col], keep="last").reset_index(drop=True)

    return df_tickets


def calc_capacity_utilization(df_worklog, capacity_hours):
    """Capacity Utilization = Σ Hours worklog / Capacity équipe × 100."""
    hours_col = find_column(df_worklog, ["Hours", "hours", "Time Spent", "Time spent", "Heures", "HOURS"])

    if hours_col is None:
        print(f"⚠️  Colonne 'Hours' introuvable. Colonnes: {list(df_worklog.columns)}")
        return 0.0, 0.0

    total_logged = pd.to_numeric(df_worklog[hours_col], errors="coerce").fillna(0).sum()
    if capacity_hours <= 0:
        return 0.0, total_logged

    utilization = round((total_logged / capacity_hours) * 100, 2)
    return utilization, total_logged

def calc_throughput(df_end, key_col):
    """Throughput = COUNT tickets where Resolved is not null (Scrum) or
    Issue Status in DONE_STATUSES (Kanban fallback).
    """
    resolved_col = find_column(df_end, ["Resolved", "resolved", "Resolution Date", "RESOLVED"])

    if resolved_col is None:
        res_col = find_column(df_end, ["Resolution", "resolution", "RESOLUTION"])
        if res_col:
            resolved = df_end[df_end[res_col].notna() & (df_end[res_col].astype(str).str.strip() != "")]
            return len(resolved), resolved[[key_col]].copy()

        # Kanban fallback: use Issue Status
        status_col = find_column(df_end, [
            "Issue Status", "issue status", "ISSUE STATUS",
            "Status", "status", "STATUS",
        ])
        if status_col:
            resolved = df_end[
                df_end[status_col].astype(str).str.strip().str.lower().isin(DONE_STATUSES)
            ]
            return len(resolved), resolved[[key_col]].copy()

        print(f"⚠️  Colonne 'Resolved' ou 'Issue Status' introuvable dans End Sprint.")
        return 0, pd.DataFrame()

    resolved = df_end[df_end[resolved_col].notna()]
    return len(resolved), resolved[[key_col]].copy()

def calc_unplanned(df_start, df_end, key_col_start, key_col_end):
    """Unplanned = tickets in End Sprint but NOT in Start (Scrum only)."""
    start_keys = set(df_start[key_col_start].dropna().astype(str).str.strip().unique())
    end_keys = set(df_end[key_col_end].dropna().astype(str).str.strip().unique())
    unplanned_keys = end_keys - start_keys
    unplanned_df = df_end[df_end[key_col_end].astype(str).str.strip().isin(unplanned_keys)]

    detail_cols = [key_col_end]
    for col in ["Summary", "Status", "Assignee", "Issue Type", "Priority"]:
        if col in df_end.columns:
            detail_cols.append(col)

    return len(unplanned_keys), unplanned_df[detail_cols].copy()

def calc_wip_end_sprint(df_end, key_col):
    """WIP = tickets where Status NOT IN done statuses.
    Supports both 'Status' (Scrum) and 'Issue Status' (Kanban) column names.
    """
    status_col = find_column(df_end, [
        "Status", "status", "STATUS",
        "Issue Status", "issue status", "ISSUE STATUS",
    ])
    if status_col is None:
        print("⚠️  Colonne 'Status' ou 'Issue Status' introuvable.")
        return 0, pd.DataFrame()

    wip = df_end[
        ~df_end[status_col].astype(str).str.strip().str.lower().isin(DONE_STATUSES)
    ]

    detail_cols = [key_col]
    for col in ["Summary", "Issue summary", status_col, "Assignee", "Issue Type", "Priority"]:
        if col in df_end.columns and col not in detail_cols:
            detail_cols.append(col)

    return len(wip), wip[detail_cols].copy()

def calc_support_load(unplanned_count, throughput):
    """Support Load = Unplanned / Throughput × 100."""
    if throughput == 0:
        return None
    return round((unplanned_count / throughput) * 100, 2)

def find_no_estimation(df_end, key_col):
    """Tickets without Original Estimate.
    Checks both 'Issue Original Estimate' (Kanban) and 'Original Estimate' (Scrum) columns.
    """
    est_col = find_column(df_end, [
        "Issue Original Estimate", "issue original estimate",
        "Original Estimate", "original estimate",
        "Σ Original Estimate", "ORIGINAL ESTIMATE",
    ])

    if est_col is None:
        print("⚠️  Colonne 'Original Estimate' introuvable.")
        return 0, pd.DataFrame()

    no_est = df_end[
        df_end[est_col].isna() | (pd.to_numeric(df_end[est_col], errors="coerce").fillna(0) == 0)
    ]

    detail_cols = [key_col]
    for col in ["Summary", "Issue summary", "Assignee", "Status", "Issue Status"]:
        if col in df_end.columns and col not in detail_cols:
            detail_cols.append(col)

    return len(no_est), no_est[detail_cols].copy()

def find_no_tempo(df_end, df_worklog, key_col_end):
    """Tickets in End Sprint with no worklog entry (Scrum mode).

    In Kanban mode, all tickets come from worklogs by definition, so this always returns 0.
    Callers in Kanban mode should pass the deduplicated tickets DataFrame as df_end and the
    full worklogs as df_worklog; the set difference will be empty.
    """
    wl_key_col = find_column(df_worklog, [
        "Issue Key", "Issue key", "issue key", "ISSUE KEY",
        "Key", "key", "KEY",
    ])
    if wl_key_col is None:
        return 0, pd.DataFrame()

    worklog_keys = set(df_worklog[wl_key_col].dropna().astype(str).str.strip().unique())
    end_keys = set(df_end[key_col_end].dropna().astype(str).str.strip().unique())
    no_tempo_keys = end_keys - worklog_keys

    no_tempo_df = df_end[df_end[key_col_end].astype(str).str.strip().isin(no_tempo_keys)]

    detail_cols = [key_col_end]
    for col in ["Summary", "Issue summary", "Assignee", "Status", "Issue Status"]:
        if col in df_end.columns and col not in detail_cols:
            detail_cols.append(col)

    return len(no_tempo_keys), no_tempo_df[detail_cols].copy()

def calc_resolution_time(df_end, df_worklog, key_col_end):
    """Calculate resolution time (in days) using first Tempo worklog date as start (Scrum mode).

    Resolution Time = Resolved - MIN(Work date) per ticket from Tempo worklogs.

    Returns:
        avg_resolution_days (float | None): average resolution time rounded to 2 decimals.
        resolution_details (DataFrame): detail per resolved ticket.
    """
    resolved_col = find_column(df_end, ["Resolved", "resolved", "Resolution Date", "RESOLVED"])
    if resolved_col is None:
        return None, pd.DataFrame()

    work_date_col = find_column(df_worklog, ["Work date", "work date", "Work Date", "WORK DATE", "Date"])
    if work_date_col is None:
        return None, pd.DataFrame()

    wl_key_col = find_column(df_worklog, ["Issue Key", "Issue key", "issue key", "Key", "key"])
    if wl_key_col is None:
        return None, pd.DataFrame()

    # Convert dates
    df_worklog = df_worklog.copy()
    df_worklog[work_date_col] = pd.to_datetime(df_worklog[work_date_col], errors="coerce")
    df_end = df_end.copy()
    df_end[resolved_col] = pd.to_datetime(df_end[resolved_col], errors="coerce")

    # Get MIN(Work date) per Issue Key from worklogs
    first_work = df_worklog.groupby(wl_key_col)[work_date_col].min().reset_index()
    first_work.columns = ["_join_key", "_first_work_date"]

    # Get resolved tickets from End Sprint
    resolved_df = df_end[df_end[resolved_col].notna()].copy()
    resolved_df["_join_key"] = resolved_df[key_col_end].astype(str).str.strip()
    first_work["_join_key"] = first_work["_join_key"].astype(str).str.strip()

    # Join on Issue Key
    merged = resolved_df.merge(first_work, on="_join_key", how="inner")

    if merged.empty:
        return None, pd.DataFrame()

    # Calculate resolution time in days
    merged["Temps de résolution (j)"] = (
        (merged[resolved_col] - merged["_first_work_date"]).dt.total_seconds() / 86400
    ).round(2)

    # Filter out negative values (data quality issues)
    merged = merged[merged["Temps de résolution (j)"] >= 0]

    if merged.empty:
        return None, pd.DataFrame()

    avg_days = round(merged["Temps de résolution (j)"].mean(), 2)
    if pd.isna(avg_days):
        return None, pd.DataFrame()

    # Derive Projet from Issue Key prefix
    merged["Projet"] = merged[key_col_end].astype(str).str.extract(r"^([A-Z][A-Z0-9]*)-", expand=False)

    # Get Priority column if available
    priority_col = find_column(df_end, ["Priority", "priority", "PRIORITY", "Priorité", "priorité"])

    # Build detail DataFrame with desired column order
    detail_cols = [key_col_end, "Projet"]
    if priority_col is not None and priority_col in merged.columns:
        detail_cols.append(priority_col)
    for col in ["Summary", "Status", "Assignee"]:
        if col in merged.columns:
            detail_cols.append(col)

    detail = merged[detail_cols + ["_first_work_date", resolved_col, "Temps de résolution (j)"]].copy()

    rename_map = {
        "_first_work_date": "Début réel (1er worklog)",
        resolved_col: "Date résolution",
        key_col_end: "Key",
    }
    if priority_col is not None and priority_col in detail.columns and priority_col != "Priorité":
        rename_map[priority_col] = "Priorité"
    detail = detail.rename(columns=rename_map)

    # Format dates for display
    for date_col in ["Début réel (1er worklog)", "Date résolution"]:
        if date_col in detail.columns:
            detail[date_col] = detail[date_col].dt.strftime("%Y-%m-%d").fillna("")

    # Custom priority sort order
    priority_order = {"highest": 0, "high": 1, "medium": 2, "low": 3, "lowest": 4}
    if "Priorité" in detail.columns:
        detail["_prio_sort"] = detail["Priorité"].astype(str).str.lower().map(priority_order).fillna(99)
        detail = detail.sort_values(
            ["Projet", "_prio_sort", "Temps de résolution (j)"],
            ascending=[True, True, False],
        ).reset_index(drop=True)
        detail = detail.drop(columns=["_prio_sort"])
    else:
        detail = detail.sort_values(
            ["Projet", "Temps de résolution (j)"],
            ascending=[True, False],
        ).reset_index(drop=True)

    return avg_days, detail


def calc_resolution_time_kanban(df_tickets, df_worklog, key_col):
    """Calculate resolution time for Kanban mode using work date range per ticket.

    Resolution Time = MAX(Work date) - MIN(Work date) per resolved Issue Key.
    Only resolved tickets (Issue Status in DONE_STATUSES) are considered.

    Returns:
        avg_resolution_days (float | None): average resolution time rounded to 2 decimals.
        resolution_details (DataFrame): detail per resolved ticket.
    """
    work_date_col = find_column(df_worklog, [
        "Work date", "work date", "Work Date", "WORK DATE", "Date",
    ])
    if work_date_col is None:
        return None, pd.DataFrame()

    wl_key_col = find_column(df_worklog, [
        "Issue Key", "Issue key", "issue key", "Key", "key",
    ])
    if wl_key_col is None:
        return None, pd.DataFrame()

    status_col = find_column(df_tickets, [
        "Issue Status", "issue status", "ISSUE STATUS",
        "Status", "status", "STATUS",
    ])
    if status_col is None:
        return None, pd.DataFrame()

    # Get resolved tickets
    resolved_tickets = df_tickets[
        df_tickets[status_col].astype(str).str.strip().str.lower().isin(DONE_STATUSES)
    ]
    if resolved_tickets.empty:
        return None, pd.DataFrame()

    resolved_keys = set(resolved_tickets[key_col].astype(str).str.strip().unique())

    # Filter worklogs for resolved tickets only
    df_wl = df_worklog.copy()
    df_wl[work_date_col] = pd.to_datetime(df_wl[work_date_col], errors="coerce")
    df_wl["_key"] = df_wl[wl_key_col].astype(str).str.strip()
    df_wl = df_wl[df_wl["_key"].isin(resolved_keys)]

    if df_wl.empty:
        return None, pd.DataFrame()

    # Get work date range per issue
    work_range = df_wl.groupby("_key")[work_date_col].agg(["min", "max"]).reset_index()
    work_range.columns = ["Key", "_first_work_date", "_last_work_date"]

    # Calculate resolution time in days
    work_range["Temps de résolution (j)"] = (
        (work_range["_last_work_date"] - work_range["_first_work_date"]).dt.total_seconds() / 86400
    ).round(2)

    # Filter out negatives (data quality issues)
    work_range = work_range[work_range["Temps de résolution (j)"] >= 0]

    if work_range.empty:
        return None, pd.DataFrame()

    avg_days = round(work_range["Temps de résolution (j)"].mean(), 2)
    if pd.isna(avg_days):
        return None, pd.DataFrame()

    # Derive Projet from Issue Key prefix
    work_range["Projet"] = work_range["Key"].str.extract(r"^([A-Z][A-Z0-9]*)-", expand=False)

    detail = work_range[["Key", "Projet", "_first_work_date", "_last_work_date", "Temps de résolution (j)"]].copy()
    detail = detail.rename(columns={
        "_first_work_date": "Début réel (1er worklog)",
        "_last_work_date": "Dernier worklog (résolution)",
    })

    for date_col in ["Début réel (1er worklog)", "Dernier worklog (résolution)"]:
        if date_col in detail.columns:
            detail[date_col] = detail[date_col].dt.strftime("%Y-%m-%d").fillna("")

    detail = detail.sort_values(
        ["Projet", "Temps de résolution (j)"],
        ascending=[True, False],
    ).reset_index(drop=True)

    return avg_days, detail


def calc_time_per_project(df_worklog, df_end=None, key_col_end=None):
    """Break down total logged hours by project (derived from the Issue Key prefix).

    Optionally joins with df_end to include Priority per ticket.

    Returns:
        tuple of (totals_df, by_priority_df):
        - totals_df: DataFrame with columns: Projet, Heures, Jours, % du total
        - by_priority_df: DataFrame with columns: Projet, Priorité, Heures, Jours, % du total
    """
    key_col = find_column(df_worklog, [
        "Issue Key", "Issue key", "issue key", "ISSUE KEY",
        "Key", "key", "KEY",
    ])
    hours_col = find_column(df_worklog, [
        "Hours", "hours", "Time Spent", "Time spent", "Heures", "HOURS",
    ])

    if key_col is None or hours_col is None:
        return pd.DataFrame(), pd.DataFrame()

    df = df_worklog[[key_col, hours_col]].copy()
    df[hours_col] = pd.to_numeric(df[hours_col], errors="coerce").fillna(0)
    df["Projet"] = df[key_col].astype(str).str.extract(r"^([A-Z][A-Z0-9]*)-", expand=False)
    df = df.dropna(subset=["Projet"])

    # 1. Calculate totals per project (no priority breakdown)
    totals = df.groupby("Projet")[hours_col].sum().reset_index()
    totals.columns = ["Projet", "Heures"]
    totals["Heures"] = totals["Heures"].round(2)
    total_hours = totals["Heures"].sum()
    totals["Jours"] = (totals["Heures"] / HOURS_PER_DAY).round(2)
    totals["% du total"] = (
        (totals["Heures"] / total_hours * 100).round(2) if total_hours > 0 else 0
    )
    totals = totals.sort_values("Jours", ascending=False).reset_index(drop=True)

    # 2. Try to join with df_end to get Priority for breakdown
    priority_col = None
    if df_end is not None and key_col_end is not None and not df_end.empty:
        priority_col = find_column(df_end, ["Priority", "priority", "PRIORITY", "Priorité", "priorité"])

    if priority_col is not None:
        prio_df = df_end[[key_col_end, priority_col]].copy()
        prio_df = prio_df.rename(columns={key_col_end: "_join_key", priority_col: "Priorité"})
        prio_df["_join_key"] = prio_df["_join_key"].astype(str).str.strip()
        df["_join_key"] = df[key_col].astype(str).str.strip()
        df_with_prio = df.merge(prio_df, on="_join_key", how="left")
        df_with_prio["Priorité"] = df_with_prio["Priorité"].fillna("Unknown")
        df_with_prio = df_with_prio.drop(columns=["_join_key"])

        by_priority = df_with_prio.groupby(["Projet", "Priorité"])[hours_col].sum().reset_index()
        by_priority.columns = ["Projet", "Priorité", "Heures"]
        by_priority["Heures"] = by_priority["Heures"].round(2)
        by_priority["Jours"] = (by_priority["Heures"] / HOURS_PER_DAY).round(2)
        by_priority["% du total"] = (
            (by_priority["Heures"] / total_hours * 100).round(2) if total_hours > 0 else 0
        )

        priority_order = {
            "blocker": 0, "critical": 1, "high": 2, "medium": 3,
            "low": 4, "minor": 5, "trivial": 6, "unknown": 99
        }
        by_priority["_prio_sort"] = by_priority["Priorité"].astype(str).str.lower().map(priority_order).fillna(99)
        by_priority = by_priority.sort_values(
            ["Projet", "_prio_sort"], ascending=[True, True]
        ).reset_index(drop=True)
        by_priority = by_priority.drop(columns=["_prio_sort"])
    else:
        by_priority = pd.DataFrame()

    return totals, by_priority


def get_capacity_input():
    """Prompt user for team capacity in days (converted to hours internally)."""
    while True:
        try:
            capacity_days = float(input("\n📊 Entrez la capacité de l'équipe pour ce sprint (en jours): "))
            if capacity_days <= 0:
                print("   La capacité doit être > 0.")
                continue
            return capacity_days * HOURS_PER_DAY
        except ValueError:
            print("   Veuillez entrer un nombre valide.")


def main():
    if len(sys.argv) < 2:
        print("Usage: python kpi_calculator.py <fichier_excel.xlsx>")
        print("Exemple: python kpi_calculator.py kanban_worklogs.xlsx")
        sys.exit(1)

    filepath = sys.argv[1]
    print(f"\n📂 Chargement du fichier: {filepath}")

    mode, df_start, df_end, df_worklog = load_data(filepath)

    key_col_end = find_key_column(df_end, SHEET_END if mode == "scrum" else SHEET_WORKLOG)

    if mode == "scrum":
        key_col_start = find_key_column(df_start, SHEET_START)
        print(f"   ✅ Start: {len(df_start)} tickets (clé: '{key_col_start}')")
    print(f"   ✅ {'End Sprint' if mode == 'scrum' else 'Tickets'}: {len(df_end)} tickets (clé: '{key_col_end}')")
    print(f"   ✅ Worklogs: {len(df_worklog)} entrées")

    capacity_hours = get_capacity_input()

    print("\n⏳ Calcul des KPIs...")

    capacity_util, total_logged = calc_capacity_utilization(df_worklog, capacity_hours)
    throughput, throughput_details = calc_throughput(df_end, key_col_end)
    wip_count, wip_details = calc_wip_end_sprint(df_end, key_col_end)
    no_est_count, no_est_details = find_no_estimation(df_end, key_col_end)

    if mode == "kanban":
        no_tempo_count, no_tempo_details = 0, pd.DataFrame()
        avg_resolution_days, resolution_details = calc_resolution_time_kanban(df_end, df_worklog, key_col_end)
    else:
        no_tempo_count, no_tempo_details = find_no_tempo(df_end, df_worklog, key_col_end)
        avg_resolution_days, resolution_details = calc_resolution_time(df_end, df_worklog, key_col_end)

    print("\n" + "=" * 60)
    print(f"            📊 {'KANBAN' if mode == 'kanban' else 'SPRINT'} KPI DASHBOARD")
    print("=" * 60)
    capacity_days = capacity_hours / HOURS_PER_DAY
    total_logged_days = round(total_logged / HOURS_PER_DAY, 2)
    print(f"  📁 Fichier           : {filepath}")
    print(f"  👥 Capacité équipe   : {capacity_days}j")
    print(f"  ⏱️  Jours logués      : {total_logged_days}j")
    print("-" * 60)
    print(f"  🔋 Capacity Utilization : {capacity_util}%")
    print(f"  ✅ Throughput (Resolved): {throughput} tickets")
    print(f"  🔄 WIP                  : {wip_count} tickets")
    print(f"  ⚠️  Sans Estimation      : {no_est_count} tickets")
    if mode == "scrum":
        print(f"  ⚠️  Sans Tempo (Worklog) : {no_tempo_count} tickets")
    if avg_resolution_days is not None:
        print(f"  📅 Temps moyen résol.   : {avg_resolution_days}j")
    print("=" * 60)

    output_file = os.path.splitext(filepath)[0] + "_KPI_Report.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        kpi_rows = [
            "Capacité équipe (j)",
            "Jours logués (j)",
            "Capacity Utilization (%)",
            "Throughput (tickets résolus)",
            "WIP",
            "Tickets sans estimation",
        ]
        kpi_vals = [
            capacity_days,
            total_logged_days,
            capacity_util,
            throughput,
            wip_count,
            no_est_count,
        ]
        if mode == "scrum":
            kpi_rows.append("Tickets sans tempo")
            kpi_vals.append(no_tempo_count)
        if avg_resolution_days is not None:
            kpi_rows.append("Temps moyen de résolution (j)")
            kpi_vals.append(avg_resolution_days)

        kpi_data = pd.DataFrame({"KPI": kpi_rows, "Valeur": kpi_vals})
        kpi_data.to_excel(writer, sheet_name="KPI Summary", index=False)

        if not wip_details.empty:
            wip_details.to_excel(writer, sheet_name="WIP", index=False)
        if not no_est_details.empty:
            no_est_details.to_excel(writer, sheet_name="Sans Estimation", index=False)
        if mode == "scrum" and not no_tempo_details.empty:
            no_tempo_details.to_excel(writer, sheet_name="Sans Tempo", index=False)
        if not resolution_details.empty:
            resolution_details.to_excel(writer, sheet_name="Temps de résolution", index=False)

    print(f"\n✅ Rapport exporté: {output_file}")


if __name__ == "__main__":
    main()
