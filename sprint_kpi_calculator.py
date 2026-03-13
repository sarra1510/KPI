"""
Sprint KPI Calculator
Reads an Excel file with 3 sheets (Start, End Sprint, Worklogs)
and computes sprint KPIs.

Handles Jira Excel exports with metadata header rows.

KPIs:
  - Capacity Utilization (%)
  - Throughput (tickets resolved)
  - Unplanned Tickets (in End but not in Start)
  - WIP End Sprint (tickets not Done)
  - Support Load (Unplanned / Throughput)
  - Tickets sans estimation
  - Tickets sans tempo (no worklog)
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

def load_data(filepath):
    """Load 3 sheets with auto-detection of header rows and flexible sheet names."""
    if not os.path.exists(filepath):
        print(f"❌ Fichier introuvable: {filepath}")
        sys.exit(1)

    try:
        xls = pd.ExcelFile(filepath)
        available_sheets = xls.sheet_names
    except Exception as e:
        print(f"❌ Impossible d'ouvrir le fichier: {e}")
        sys.exit(1)

    print(f"   📋 Feuilles trouvées: {available_sheets}")

    sheet_start = find_sheet_name(
        available_sheets, SHEET_START,
        keywords=["start", "début", "debut", "démarrage", "demarrage"],
        exclude_keywords=["end", "fin"]
    )
    sheet_end = find_sheet_name(
        available_sheets, SHEET_END,
        keywords=["end", "fin", "sprint end", "end sprint"]
    )
    sheet_worklog = find_sheet_name(
        available_sheets, SHEET_WORKLOG,
        keywords=["worklog", "worklogs", "tempo"]
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
        print(f"   📄 '{sheet_start}': en-tête ligne {h_start} (skip {h_start} lignes Jira)")
    if h_end > 0:
        print(f"   📄 '{sheet_end}': en-tête ligne {h_end} (skip {h_end} lignes Jira)")
    if h_wl > 0:
        print(f"   📄 '{sheet_worklog}': en-tête ligne {h_wl} (skip {h_wl} lignes Jira)")

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

    # Remove rows where Key is empty (junk rows in Jira export)
    for name, df, key_candidates in [
        ("Start", df_start, ["Key", "Issue Key", "key", "Clé"]),
        ("End Sprint", df_end, ["Key", "Issue Key", "key", "Clé"]),
        ("Worklogs", df_worklog, ["Issue Key", "Issue key", "Key"]),
    ]:
        for kc in key_candidates:
            if kc in df.columns:
                before = len(df)
                df = df[df[kc].notna() & (df[kc].astype(str).str.strip() != '')].reset_index(drop=True)
                dropped = before - len(df)
                if dropped > 0:
                    print(f"   🧹 '{name}': supprimé {dropped} lignes vides")
                break
        if name == "Start":
            df_start = df
        elif name == "End Sprint":
            df_end = df
        else:
            df_worklog = df

    return df_start, df_end, df_worklog

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
    """Throughput = COUNT tickets where Resolved is not null."""
    resolved_col = find_column(df_end, ["Resolved", "resolved", "Resolution Date", "RESOLVED"])

    if resolved_col is None:
        res_col = find_column(df_end, ["Resolution", "resolution", "RESOLUTION"])
        if res_col:
            resolved = df_end[df_end[res_col].notna() & (df_end[res_col].astype(str).str.strip() != "")]  
            return len(resolved), resolved[[key_col]].copy()

        print(f"⚠️  Colonne 'Resolved' introuvable dans End Sprint.")
        return 0, pd.DataFrame()

    resolved = df_end[df_end[resolved_col].notna()]
    return len(resolved), resolved[[key_col]].copy()

def calc_unplanned(df_start, df_end, key_col_start, key_col_end):
    """Unplanned = tickets in End Sprint but NOT in Start."""
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
    """WIP = tickets where Status NOT IN done statuses."""
    status_col = find_column(df_end, ["Status", "status", "STATUS"])
    if status_col is None:
        print("⚠️  Colonne 'Status' introuvable dans End Sprint.")
        return 0, pd.DataFrame()

    wip = df_end[
        ~df_end[status_col].astype(str).str.strip().str.lower().isin(DONE_STATUSES)
    ]

    detail_cols = [key_col]
    for col in ["Summary", status_col, "Assignee", "Issue Type", "Priority"]:
        if col in df_end.columns and col not in detail_cols:
            detail_cols.append(col)

    return len(wip), wip[detail_cols].copy()

def calc_support_load(unplanned_count, throughput):
    """Support Load = Unplanned / Throughput × 100."""
    if throughput == 0:
        return None
    return round((unplanned_count / throughput) * 100, 2)

def find_no_estimation(df_end, key_col):
    """Tickets without Original Estimate."""
    est_col = find_column(df_end, [
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
    for col in ["Summary", "Assignee", "Status"]:
        if col in df_end.columns:
            detail_cols.append(col)

    return len(no_est), no_est[detail_cols].copy()

def find_no_tempo(df_end, df_worklog, key_col_end):
    """Tickets in End Sprint with no worklog entry."""
    wl_key_col = find_key_column(df_worklog, SHEET_WORKLOG)
    worklog_keys = set(df_worklog[wl_key_col].dropna().astype(str).str.strip().unique())
    end_keys = set(df_end[key_col_end].dropna().astype(str).str.strip().unique())
    no_tempo_keys = end_keys - worklog_keys

    no_tempo_df = df_end[df_end[key_col_end].astype(str).str.strip().isin(no_tempo_keys)]

    detail_cols = [key_col_end]
    for col in ["Summary", "Assignee", "Status"]:
        if col in df_end.columns:
            detail_cols.append(col)

    return len(no_tempo_keys), no_tempo_df[detail_cols].copy()

def calc_resolution_time(df_end, key_col):
    """Calculate resolution time (in days) for resolved tickets.

    Returns:
        avg_resolution_days (float | None): average resolution time rounded to 2 decimals.
        resolution_details (DataFrame): Key, Summary, Status, Created, Resolved, Resolution Time (j).
    """
    created_col = find_column(df_end, ["Created", "created", "Date de création", "CREATED"])
    resolved_col = find_column(df_end, ["Resolved", "resolved", "Resolution Date", "RESOLVED"])
    status_col = find_column(df_end, ["Status", "status", "STATUS"])

    if created_col is None or resolved_col is None:
        return None, pd.DataFrame()

    df = df_end.copy()

    # Consider a ticket resolved if its Resolved column is not null
    # OR if its status is in DONE_STATUSES
    resolved_mask = df[resolved_col].notna()
    if status_col:
        done_mask = df[status_col].astype(str).str.strip().str.lower().isin(DONE_STATUSES)
        resolved_mask = resolved_mask | done_mask

    df_resolved = df[resolved_mask].copy()

    # Keep only rows where both Created and Resolved dates exist
    df_resolved = df_resolved[df_resolved[created_col].notna() & df_resolved[resolved_col].notna()].copy()

    if df_resolved.empty:
        return None, pd.DataFrame()

    df_resolved[created_col] = pd.to_datetime(df_resolved[created_col], errors="coerce")
    df_resolved[resolved_col] = pd.to_datetime(df_resolved[resolved_col], errors="coerce")

    df_resolved = df_resolved.dropna(subset=[created_col, resolved_col])

    if df_resolved.empty:
        return None, pd.DataFrame()

    df_resolved["Resolution Time (j)"] = (
        (df_resolved[resolved_col] - df_resolved[created_col]).dt.total_seconds() / 86400
    ).round(2)

    avg_resolution_days = round(df_resolved["Resolution Time (j)"].mean(), 2)
    if pd.isna(avg_resolution_days):
        return None, pd.DataFrame()

    detail_cols = [key_col]
    for col in ["Summary", status_col, created_col, resolved_col]:
        if col and col in df_resolved.columns and col not in detail_cols:
            detail_cols.append(col)
    detail_cols.append("Resolution Time (j)")

    resolution_details = df_resolved[detail_cols].copy()
    # Rename columns to standard names for display
    rename_map = {}
    if created_col != "Created":
        rename_map[created_col] = "Created"
    if resolved_col != "Resolved":
        rename_map[resolved_col] = "Resolved"
    if status_col and status_col != "Status":
        rename_map[status_col] = "Status"
    if rename_map:
        resolution_details = resolution_details.rename(columns=rename_map)

    return avg_resolution_days, resolution_details


def calc_time_per_project(df_worklog):
    """Break down total logged hours by project (derived from the Issue Key prefix).

    Returns:
        DataFrame with columns: Projet, Heures, Jours, % du total — sorted by Jours descending.
    """
    key_col = find_column(df_worklog, [
        "Issue Key", "Issue key", "issue key", "ISSUE KEY",
        "Key", "key", "KEY",
    ])
    hours_col = find_column(df_worklog, [
        "Hours", "hours", "Time Spent", "Time spent", "Heures", "HOURS",
    ])

    if key_col is None or hours_col is None:
        return pd.DataFrame()

    df = df_worklog[[key_col, hours_col]].copy()
    df[hours_col] = pd.to_numeric(df[hours_col], errors="coerce").fillna(0)
    df["Projet"] = df[key_col].astype(str).str.extract(r"^([A-Z][A-Z0-9]*)-", expand=False)
    df = df.dropna(subset=["Projet"])

    grouped = df.groupby("Projet")[hours_col].sum().reset_index()
    grouped.columns = ["Projet", "Heures"]
    total_hours = grouped["Heures"].sum()
    grouped["Jours"] = (grouped["Heures"] / HOURS_PER_DAY).round(2)
    grouped["% du total"] = (
        (grouped["Heures"] / total_hours * 100).round(1) if total_hours > 0 else 0
    )
    grouped = grouped.sort_values("Jours", ascending=False).reset_index(drop=True)

    return grouped


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
        print("Usage: python sprint_kpi_calculator.py <fichier_excel.xlsx>")
        print("Exemple: python sprint_kpi_calculator.py sprint_data.xlsx")
        sys.exit(1)

    filepath = sys.argv[1]
    print(f"\n📂 Chargement du fichier: {filepath}")

    df_start, df_end, df_worklog = load_data(filepath)

    key_col_start = find_key_column(df_start, SHEET_START)
    key_col_end = find_key_column(df_end, SHEET_END)

    print(f"   ✅ Start: {len(df_start)} tickets (clé: '{key_col_start}')")
    print(f"   ✅ End Sprint: {len(df_end)} tickets (clé: '{key_col_end}')")
    print(f"   ✅ Worklogs: {len(df_worklog)} entrées")

    capacity_hours = get_capacity_input()

    print("\n⏳ Calcul des KPIs...")

    capacity_util, total_logged = calc_capacity_utilization(df_worklog, capacity_hours)
    throughput, throughput_details = calc_throughput(df_end, key_col_end)
    unplanned_count, unplanned_details = calc_unplanned(df_start, df_end, key_col_start, key_col_end)
    wip_count, wip_details = calc_wip_end_sprint(df_end, key_col_end)
    support_load = calc_support_load(unplanned_count, throughput)
    no_est_count, no_est_details = find_no_estimation(df_end, key_col_end)
    no_tempo_count, no_tempo_details = find_no_tempo(df_end, df_worklog, key_col_end)

    print("\n" + "=" * 60)
    print("            📊 SPRINT KPI DASHBOARD")
    print("=" * 60)
    capacity_days = capacity_hours / HOURS_PER_DAY
    total_logged_days = round(total_logged / HOURS_PER_DAY, 2)
    print(f"  📁 Fichier           : {filepath}")
    print(f"  👥 Capacité équipe   : {capacity_days}j")
    print(f"  ⏱️  Jours logués      : {total_logged_days}j")
    print("-" * 60)
    print(f"  🔋 Capacity Utilization : {capacity_util}%")
    print(f"  ✅ Throughput (Resolved): {throughput} tickets")
    print(f"  🆕 Unplanned Tickets    : {unplanned_count} tickets")
    print(f"  🔄 WIP End Sprint       : {wip_count} tickets")
    support_str = f"{support_load}%" if support_load is not None else "N/A (throughput=0)"
    print(f"  🛟 Support Load          : {support_str}")
    print(f"  ⚠️  Sans Estimation      : {no_est_count} tickets")
    print(f"  ⚠️  Sans Tempo (Worklog) : {no_tempo_count} tickets")
    print("=" * 60)

    output_file = os.path.splitext(filepath)[0] + "_KPI_Report.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        kpi_data = pd.DataFrame({
            "KPI": [
                "Capacité équipe (j)",
                "Jours logués (j)",
                "Capacity Utilization (%)",
                "Throughput (tickets résolus)",
                "Unplanned Tickets",
                "WIP End Sprint",
                "Support Load (%)",
                "Tickets sans estimation",
                "Tickets sans tempo",
            ],
            "Valeur": [
                capacity_days,
                total_logged_days,
                capacity_util,
                throughput,
                unplanned_count,
                wip_count,
                support_load if support_load is not None else "N/A",
                no_est_count,
                no_tempo_count,
            ],
        })
        kpi_data.to_excel(writer, sheet_name="KPI Summary", index=False)

        if not unplanned_details.empty:
            unplanned_details.to_excel(writer, sheet_name="Unplanned Tickets", index=False)
        if not wip_details.empty:
            wip_details.to_excel(writer, sheet_name="WIP End Sprint", index=False)
        if not no_est_details.empty:
            no_est_details.to_excel(writer, sheet_name="Sans Estimation", index=False)
        if not no_tempo_details.empty:
            no_tempo_details.to_excel(writer, sheet_name="Sans Tempo", index=False)

    print(f"\n✅ Rapport exporté: {output_file}")
    print("   Onglets: KPI Summary | Unplanned Tickets | WIP End Sprint | Sans Estimation | Sans Tempo")


if __name__ == "__main__":
    main()