"""
Sprint KPI Calculator
Reads an Excel file with 3 sheets (Start, End Sprint, Worklogs)
and computes sprint KPIs.

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
import sys
import os


# --- CONFIG ---
SHEET_START = "Start"
SHEET_END = "End Sprint"
SHEET_WORKLOG = "Worklogs"

DONE_STATUSES = [
    "closed",
    "customer pending",
    "released",
    "canceled",
    "done",
]


def find_key_column(df, sheet_name):
    """Auto-detect the issue key column ('Key' or 'Issue Key')."""
    for col in ["Key", "Issue Key", "Issue key", "key", "issue key"]:
        if col in df.columns:
            return col
    raise ValueError(
        f"Impossible de trouver la colonne 'Key' ou 'Issue Key' dans la feuille '{sheet_name}'.\n"
        f"Colonnes disponibles: {list(df.columns)}"
    )


def load_data(filepath):
    """Load the 3 sheets from the Excel file."""
    if not os.path.exists(filepath):
        print(f"❌ Fichier introuvable: {filepath}")
        sys.exit(1)

    try:
        df_start = pd.read_excel(filepath, sheet_name=SHEET_START)
        df_end = pd.read_excel(filepath, sheet_name=SHEET_END)
        df_worklog = pd.read_excel(filepath, sheet_name=SHEET_WORKLOG)
    except ValueError as e:
        print(f"❌ Erreur de lecture des feuilles: {e}")
        print(f"   Feuilles attendues: '{SHEET_START}', '{SHEET_END}', '{SHEET_WORKLOG}'")
        sys.exit(1)

    return df_start, df_end, df_worklog


def calc_capacity_utilization(df_worklog, capacity_hours):
    """Capacity Utilization = Σ Hours worklog / Capacity équipe × 100."""
    hours_col = None
    for col in ["Hours", "hours", "Time Spent", "Time spent", "Heures"]:
        if col in df_worklog.columns:
            hours_col = col
            break

    if hours_col is None:
        print(f"⚠️  Colonne 'Hours' introuvable dans Worklogs. Colonnes: {list(df_worklog.columns)}")
        return 0.0, 0.0

    total_logged = pd.to_numeric(df_worklog[hours_col], errors="coerce").fillna(0).sum()
    if capacity_hours <= 0:
        return 0.0, total_logged

    utilization = round((total_logged / capacity_hours) * 100, 2)
    return utilization, total_logged


def calc_throughput(df_end, key_col):
    """Throughput = COUNT tickets where Resolved is not null."""
    resolved_col = None
    for col in ["Resolved", "resolved", "Resolution Date"]:
        if col in df_end.columns:
            resolved_col = col
            break

    if resolved_col is None:
        res_col = None
        for col in ["Resolution", "resolution"]:
            if col in df_end.columns:
                res_col = col
                break
        if res_col:
            resolved = df_end[df_end[res_col].notna() & (df_end[res_col] != "")]
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
    if "Status" not in df_end.columns:
        print("⚠️  Colonne 'Status' introuvable dans End Sprint.")
        return 0, pd.DataFrame()

    wip = df_end[
        ~df_end["Status"].astype(str).str.strip().str.lower().isin(DONE_STATUSES)
    ]

    detail_cols = [key_col]
    for col in ["Summary", "Status", "Assignee", "Issue Type", "Priority"]:
        if col in df_end.columns:
            detail_cols.append(col)

    return len(wip), wip[detail_cols].copy()


def calc_support_load(unplanned_count, throughput):
    """Support Load = Unplanned / Throughput × 100."""
    if throughput == 0:
        return None
    return round((unplanned_count / throughput) * 100, 2)


def find_no_estimation(df_end, key_col):
    """Tickets without Original Estimate."""
    est_col = None
    for col in ["Original Estimate", "original estimate", "Σ Original Estimate"]:
        if col in df_end.columns:
            est_col = col
            break

    if est_col is None:
        print("⚠️  Colonne 'Original Estimate' introuvable dans End Sprint.")
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


def get_capacity_input():
    """Prompt user for team capacity in hours."""
    while True:
        try:
            capacity = float(input("\n📊 Entrez la capacité de l'équipe pour ce sprint (en heures): "))
            if capacity <= 0:
                print("   La capacité doit être > 0.")
                continue
            return capacity
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

    print(f"   ✅ Feuille '{SHEET_START}': {len(df_start)} tickets (clé: '{key_col_start}')")
    print(f"   ✅ Feuille '{SHEET_END}': {len(df_end)} tickets (clé: '{key_col_end}')")
    print(f"   ✅ Feuille '{SHEET_WORKLOG}': {len(df_worklog)} entrées")

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
    print(f"  📁 Fichier           : {filepath}")
    print(f"  👥 Capacité équipe   : {capacity_hours}h")
    print(f"  ⏱️  Heures loggées    : {total_logged}h")
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

    output_file = filepath.replace(".xlsx", "_KPI_Report.xlsx").replace(".xls", "_KPI_Report.xlsx")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        kpi_data = pd.DataFrame({
            "KPI": [
                "Capacité équipe (h)",
                "Heures loggées (h)",
                "Capacity Utilization (%)",
                "Throughput (tickets résolus)",
                "Unplanned Tickets",
                "WIP End Sprint",
                "Support Load (%)",
                "Tickets sans estimation",
                "Tickets sans tempo",
            ],
            "Valeur": [
                capacity_hours,
                total_logged,
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
