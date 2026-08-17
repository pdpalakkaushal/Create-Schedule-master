import io
import os
from datetime import date
import pandas as pd
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Call Cycle Reconciler",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# CONSTANTS
# ============================================================

REQUIRED_CLIENT_COLS = [
    "Week Number",
    "Day of Week",
    "User Code",
    "User Name",
    "Location Code",
    "Customer Name",
    "Callage"
]

MASTER_COLUMNS = [
    "Store ID",
    "User ID",
    "Start Date",
    "Frequency",
    "Store Name",
    "End Date",
    "Days of Week",
    "Applicable Weeks",
    "Month Start On",
    "Month End On"
]


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown("""
<style>

.main {
    padding-top: 1rem;
}

.block-container {
    padding-top: 1.5rem;
}

.metric-card {
    padding: 15px;
    border-radius: 10px;
    border: 1px solid #ddd;
    background-color: #ffffff;
}

h1 {
    font-weight: 700;
}

.section-title {
    font-size: 20px;
    font-weight: 700;
    margin-top: 20px;
}

.small-text {
    color: #666;
    font-size: 13px;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# HEADER
# ============================================================

st.title("📅 Call Cycle Reconciler")

st.markdown(
    """
    **Reconcile client call-cycle planning files with an existing master file,
    identify changes, validate data quality, and generate a clean final master CSV.**
    """
)

st.divider()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_cell(value):
    """
    Normalize cell values.

    Example:
    350260774.0 -> 350260774
    """
    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.endswith(".0") and value[:-2].isdigit():
        value = value[:-2]

    return value


def read_uploaded_file(uploaded_file):

    extension = os.path.splitext(uploaded_file.name)[1].lower()

    try:

        if extension in [".csv", ".txt"]:

            df = pd.read_csv(
                uploaded_file,
                dtype=str,
                keep_default_na=False
            )

        elif extension in [".xlsx", ".xlsm"]:

            df = pd.read_excel(
                uploaded_file,
                dtype=str,
                engine="openpyxl"
            )

        elif extension == ".xls":

            df = pd.read_excel(
                uploaded_file,
                dtype=str,
                engine="xlrd"
            )

        else:

            raise ValueError(
                "Unsupported file type. Please upload CSV or Excel file."
            )

        df.columns = [str(c).strip() for c in df.columns]

        return df.fillna("")

    except Exception as e:

        raise ValueError(
            f"Unable to read {uploaded_file.name}: {str(e)}"
        )


def load_client_file(uploaded_file):

    try:

        raw = read_uploaded_file(uploaded_file)

    except Exception as e:

        return None, str(e)

    if raw.empty:

        return None, "File is empty."

    lower_columns = {
        c.lower(): c
        for c in raw.columns
    }

    missing = []

    column_mapping = {}

    for required in REQUIRED_CLIENT_COLS:

        actual = lower_columns.get(required.lower())

        if actual:

            column_mapping[required] = actual

        else:

            missing.append(required)

    if missing:

        return None, (
            "Missing required columns: "
            + ", ".join(missing)
        )

    output = pd.DataFrame({

        required: raw[column_mapping[required]].map(clean_cell)

        for required in REQUIRED_CLIENT_COLS

    })

    output["Week Number"] = pd.to_numeric(
        output["Week Number"],
        errors="coerce"
    )

    output = output.dropna(
        subset=[
            "Week Number",
            "Location Code",
            "User Code",
            "Day of Week"
        ]
    )

    output = output[
        (output["Location Code"] != "") &
        (output["User Code"] != "") &
        (output["Day of Week"] != "")
    ]

    output["Week Number"] = output["Week Number"].astype(int)

    output["source_file"] = uploaded_file.name

    return output, None


def load_master_file(uploaded_file):

    try:

        df = read_uploaded_file(uploaded_file)

    except Exception as e:

        return None, str(e)

    missing = [
        col
        for col in MASTER_COLUMNS
        if col not in df.columns
    ]

    if missing:

        return None, (
            "Invalid master file. Missing columns: "
            + ", ".join(missing)
        )

    df = df[MASTER_COLUMNS].fillna("").astype(str)

    return df, None


def most_common(values):

    values = [
        v for v in values
        if v not in ["", None, "nan"]
    ]

    if not values:

        return ""

    return pd.Series(values).mode().iloc[0]


# ============================================================
# ANALYSIS
# ============================================================

def analyze_client(client_rows, user_code, master_df=None):

    rows = client_rows[
        client_rows["User Code"] == user_code
    ]

    if rows.empty:

        raise ValueError(
            f"No rows found for User Code: {user_code}"
        )

    user_name = rows.iloc[0]["User Name"]

    computed = []

    flags = {
        "day_conflict": [],
        "callage_conflict": [],
        "week_conflict": [],
        "out_of_range": [],
        "cross_file": []
    }

    for location_code, group in rows.groupby(
        "Location Code"
    ):

        entries = group.to_dict("records")

        store_name = entries[0]["Customer Name"]

        # ----------------------------------------------------
        # Invalid week
        # ----------------------------------------------------

        invalid_weeks = [
            e for e in entries
            if not (1 <= e["Week Number"] <= 5)
        ]

        if invalid_weeks:

            flags["out_of_range"].append({
                "Store ID": location_code,
                "Store Name": store_name,
                "Week": invalid_weeks[0]["Week Number"]
            })

        # ----------------------------------------------------
        # Same week conflicts
        # ----------------------------------------------------

        by_week = {}

        for entry in entries:

            by_week.setdefault(
                entry["Week Number"],
                []
            ).append(entry)

        for week, week_entries in by_week.items():

            variants = {
                (
                    e["Day of Week"],
                    e["Callage"]
                )
                for e in week_entries
            }

            if len(variants) > 1:

                flags["week_conflict"].append({
                    "Store ID": location_code,
                    "Store Name": store_name,
                    "Week": week,
                    "Variants": " | ".join(
                        f"{d} / {c}"
                        for d, c in variants
                    )
                })

        # ----------------------------------------------------
        # Cross-file
        # ----------------------------------------------------

        source_files = {
            e["source_file"]
            for e in entries
        }

        if len(source_files) > 1:

            flags["cross_file"].append({
                "Store ID": location_code,
                "Store Name": store_name,
                "Files": ", ".join(
                    sorted(source_files)
                )
            })

        # ----------------------------------------------------
        # Valid weeks
        # ----------------------------------------------------

        weeks_set = sorted({

            e["Week Number"]

            for e in entries

            if 1 <= e["Week Number"] <= 5

        })

        if not weeks_set:

            continue

        # ----------------------------------------------------
        # Earliest-week-wins
        # ----------------------------------------------------

        earliest_week = weeks_set[0]

        earliest_entries = [
            e
            for e in entries
            if e["Week Number"] == earliest_week
        ]

        resolved_day = earliest_entries[0]["Day of Week"]

        resolved_callage = earliest_entries[0]["Callage"]

        # ----------------------------------------------------
        # Day conflict
        # ----------------------------------------------------

        all_days = {
            e["Day of Week"]
            for e in entries
        }

        if len(all_days) > 1:

            flags["day_conflict"].append({
                "Store ID": location_code,
                "Store Name": store_name,
                "Options": " | ".join(
                    f"Week {e['Week Number']}: {e['Day of Week']}"
                    for e in entries
                ),
                "Resolved": resolved_day
            })

        # ----------------------------------------------------
        # Callage conflict
        # ----------------------------------------------------

        all_callage = {
            e["Callage"]
            for e in entries
        }

        if len(all_callage) > 1:

            flags["callage_conflict"].append({
                "Store ID": location_code,
                "Store Name": store_name,
                "Options": " | ".join(
                    f"Week {e['Week Number']}: {e['Callage']}"
                    for e in entries
                ),
                "Resolved": resolved_callage
            })

        # ----------------------------------------------------
        # Applicable weeks
        # ----------------------------------------------------

        bitmask = [
            1 if week in weeks_set else 0
            for week in [1, 2, 3, 4]
        ]

        bitmask.append(
            1 if (
                5 in weeks_set or
                1 in weeks_set
            )
            else 0
        )

        computed.append({

            "Store ID": location_code,

            "Store Name": store_name,

            "Days of Week": resolved_day,

            "Applicable Weeks": ",".join(
                map(str, bitmask)
            ),

            "Weeks": ",".join(
                map(str, weeks_set)
            )

        })

    # ========================================================
    # MASTER COMPARISON
    # ========================================================

    if master_df is not None:

        existing = master_df[
            master_df["User ID"].astype(str)
            == str(user_code)
        ]

    else:

        existing = pd.DataFrame(
            columns=MASTER_COLUMNS
        )

    existing_rows = existing.to_dict("records")

    existing_by_store = {
        str(row["Store ID"]): row
        for row in existing_rows
    }

    computed_by_store = {
        str(row["Store ID"]): row
        for row in computed
    }

    missing_from_master = []

    changed = []

    unchanged = []

    for store_id, current in computed_by_store.items():

        existing_row = existing_by_store.get(store_id)

        if existing_row is None:

            missing_from_master.append(current)

        else:

            if (
                str(existing_row["Days of Week"])
                != str(current["Days of Week"])
                or
                str(existing_row["Applicable Weeks"])
                != str(current["Applicable Weeks"])
            ):

                changed.append({

                    **current,

                    "Old Day":
                        existing_row["Days of Week"],

                    "New Day":
                        current["Days of Week"],

                    "Old Applicable Weeks":
                        existing_row["Applicable Weeks"],

                    "New Applicable Weeks":
                        current["Applicable Weeks"]

                })

            else:

                unchanged.append(current)

    extra_in_master = [

        row

        for row in existing_rows

        if str(row["Store ID"])
        not in computed_by_store

    ]

    return {

        "user_code": user_code,

        "user_name": user_name,

        "computed": computed,

        "existing": existing_rows,

        "missing": missing_from_master,

        "changed": changed,

        "unchanged": unchanged,

        "extra": extra_in_master,

        "flags": flags

    }


# ============================================================
# BUILD OUTPUT
# ============================================================

def build_output_rows(
    analysis,
    start_date="",
    frequency="Weekly",
    end_date="",
    month_start="",
    month_end="",
    keep_extra=True
):

    rows = []

    for store in analysis["computed"]:

        rows.append({

            "Store ID":
                store["Store ID"],

            "User ID":
                analysis["user_code"],

            "Start Date":
                start_date,

            "Frequency":
                frequency,

            "Store Name":
                store["Store Name"],

            "End Date":
                end_date,

            "Days of Week":
                store["Days of Week"],

            "Applicable Weeks":
                store["Applicable Weeks"],

            "Month Start On":
                month_start,

            "Month End On":
                month_end

        })

    if keep_extra:

        rows.extend(
            analysis["extra"]
        )

    return rows


# ============================================================
# EXCEL REPORT
# ============================================================

def create_excel_report(
    final_df,
    summary_df,
    changed_df,
    issues_df
):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        final_df.to_excel(
            writer,
            sheet_name="Final Master",
            index=False
        )

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        changed_df.to_excel(
            writer,
            sheet_name="Changes",
            index=False
        )

        issues_df.to_excel(
            writer,
            sheet_name="Data Quality",
            index=False
        )

        workbook = writer.book

        header_format = workbook.add_format({
            "bold": True,
            "border": 1
        })

        for sheet_name, df in {
            "Final Master": final_df,
            "Summary": summary_df,
            "Changes": changed_df,
            "Data Quality": issues_df
        }.items():

            worksheet = writer.sheets[sheet_name]

            for col_num, column in enumerate(df.columns):

                worksheet.write(
                    0,
                    col_num,
                    column,
                    header_format
                )

                width = max(
                    len(str(column)) + 2,
                    min(
                        40,
                        int(
                            df[column]
                            .astype(str)
                            .str.len()
                            .max()
                        ) + 2
                    )
                    if not df.empty
                    else len(str(column)) + 2
                )

                worksheet.set_column(
                    col_num,
                    col_num,
                    width
                )

    output.seek(0)

    return output.getvalue()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Configuration")

    st.markdown(
        "Upload your call-cycle planning files."
    )

    client_files = st.file_uploader(
        "Client Call Cycle Files",
        type=[
            "csv",
            "xlsx",
            "xls",
            "xlsm"
        ],
        accept_multiple_files=True
    )

    st.divider()

    master_file = st.file_uploader(
        "Existing Master File (Optional)",
        type=[
            "csv",
            "xlsx"
        ],
        accept_multiple_files=False
    )

    st.divider()

    st.subheader("New Client Settings")

    start_date = st.text_input(
        "Start Date",
        placeholder="DD-MM-YYYY"
    )

    frequency = st.selectbox(
        "Frequency",
        [
            "Weekly",
            "Fortnightly",
            "Monthly"
        ]
    )

    end_date = st.text_input(
        "End Date",
        placeholder="DD-MM-YYYY"
    )

    st.divider()

    keep_extra = st.checkbox(
        "Keep stores present in master but missing from client file",
        value=True
    )


# ============================================================
# FILE VALIDATION
# ============================================================

if not client_files:

    st.info(
        "👈 Upload one or more client call-cycle files from the sidebar to begin."
    )

    st.markdown("""
    ### Expected Client Columns

    Your client file should contain:

    - Week Number
    - Day of Week
    - User Code
    - User Name
    - Location Code
    - Customer Name
    - Callage

    ### Expected Master Columns

    - Store ID
    - User ID
    - Start Date
    - Frequency
    - Store Name
    - End Date
    - Days of Week
    - Applicable Weeks
    - Month Start On
    - Month End On
    """)

    st.stop()


# ============================================================
# LOAD CLIENT FILES
# ============================================================

st.subheader("📂 Uploaded Files")

all_client_rows = []

file_status = []

for uploaded_file in client_files:

    data, error = load_client_file(
        uploaded_file
    )

    if error:

        file_status.append({
            "File":
                uploaded_file.name,
            "Status":
                "❌ Error",
            "Rows":
                0,
            "Message":
                error
        })

    else:

        all_client_rows.append(data)

        file_status.append({
            "File":
                uploaded_file.name,
            "Status":
                "✅ Loaded",
            "Rows":
                len(data),
            "Message":
                "Valid"
        })


st.dataframe(
    pd.DataFrame(file_status),
    use_container_width=True,
    hide_index=True
)


if not all_client_rows:

    st.error(
        "No usable client files were found."
    )

    st.stop()


client_rows = pd.concat(
    all_client_rows,
    ignore_index=True
)


# ============================================================
# MASTER
# ============================================================

master_df = None

if master_file:

    master_df, master_error = load_master_file(
        master_file
    )

    if master_error:

        st.error(
            f"Master file error: {master_error}"
        )

        st.stop()

    else:

        st.success(
            f"Master file loaded: {len(master_df):,} rows"
        )


# ============================================================
# CLIENT SELECTION
# ============================================================

clients = (
    client_rows[
        ["User Code", "User Name"]
    ]
    .drop_duplicates()
    .reset_index(drop=True)
)

clients["Display"] = (
    clients["User Name"]
    + " — "
    + clients["User Code"]
)

st.subheader("👤 Select Clients")

selected_display = st.multiselect(
    "Clients to reconcile",
    clients["Display"].tolist(),
    default=clients["Display"].tolist()
)

if not selected_display:

    st.warning(
        "Please select at least one client."
    )

    st.stop()


selected_codes = clients[
    clients["Display"].isin(selected_display)
]["User Code"].tolist()


# ============================================================
# RUN
# ============================================================

run_button = st.button(
    "🚀 Run Reconciliation",
    type="primary",
    use_container_width=True
)


if run_button:

    with st.spinner(
        "Reconciling call-cycle data..."
    ):

        analyses = []

        final_rows = []

        summary_rows = []

        change_rows = []

        issue_rows = []

        for code in selected_codes:

            analysis = analyze_client(
                client_rows,
                code,
                master_df
            )

            analyses.append(analysis)

            existing = analysis["existing"]

            is_new_client = len(existing) == 0

            if is_new_client:

                sd = start_date
                fr = frequency
                ed = end_date
                ms = ""
                me = ""

            else:

                sd = most_common([
                    r["Start Date"]
                    for r in existing
                ])

                fr = most_common([
                    r["Frequency"]
                    for r in existing
                ]) or "Weekly"

                ed = most_common([
                    r["End Date"]
                    for r in existing
                ])

                ms = most_common([
                    r["Month Start On"]
                    for r in existing
                ])

                me = most_common([
                    r["Month End On"]
                    for r in existing
                ])

            output_rows = build_output_rows(
                analysis,
                sd,
                fr,
                ed,
                ms,
                me,
                keep_extra
            )

            final_rows.extend(
                output_rows
            )

            flags = analysis["flags"]

            issue_count = sum(
                len(v)
                for v in flags.values()
            )

            summary_rows.append({

                "User Code":
                    code,

                "User Name":
                    analysis["user_name"],

                "Stores in Client":
                    len(analysis["computed"]),

                "New Stores":
                    len(analysis["missing"]),

                "Changed Stores":
                    len(analysis["changed"]),

                "Stores Removed from Client":
                    len(analysis["extra"]),

                "Unchanged Stores":
                    len(analysis["unchanged"]),

                "Data Quality Issues":
                    issue_count

            })

            for change in analysis["changed"]:

                change_rows.append({

                    "User Code":
                        code,

                    "User Name":
                        analysis["user_name"],

                    "Store ID":
                        change["Store ID"],

                    "Store Name":
                        change["Store Name"],

                    "Old Day":
                        change["Old Day"],

                    "New Day":
                        change["New Day"],

                    "Old Applicable Weeks":
                        change["Old Applicable Weeks"],

                    "New Applicable Weeks":
                        change["New Applicable Weeks"]

                })

            for issue_type, records in flags.items():

                for record in records:

                    issue_rows.append({

                        "User Code":
                            code,

                        "User Name":
                            analysis["user_name"],

                        "Issue Type":
                            issue_type,

                        **record

                    })

    # ========================================================
    # FINAL DATA
    # ========================================================

    final_df = pd.DataFrame(
        final_rows,
        columns=MASTER_COLUMNS
    )

    summary_df = pd.DataFrame(
        summary_rows
    )

    changed_df = pd.DataFrame(
        change_rows
    )

    issues_df = pd.DataFrame(
        issue_rows
    )

    st.session_state["final_df"] = final_df
    st.session_state["summary_df"] = summary_df
    st.session_state["changed_df"] = changed_df
    st.session_state["issues_df"] = issues_df


# ============================================================
# RESULTS
# ============================================================

if "final_df" in st.session_state:

    final_df = st.session_state["final_df"]
    summary_df = st.session_state["summary_df"]
    changed_df = st.session_state["changed_df"]
    issues_df = st.session_state["issues_df"]

    st.divider()

    st.subheader("📊 Reconciliation Summary")

    total_client_stores = int(
        summary_df["Stores in Client"].sum()
    )

    total_new = int(
        summary_df["New Stores"].sum()
    )

    total_changed = int(
        summary_df["Changed Stores"].sum()
    )

    total_removed = int(
        summary_df["Stores Removed from Client"].sum()
    )

    total_unchanged = int(
        summary_df["Unchanged Stores"].sum()
    )

    total_issues = int(
        summary_df["Data Quality Issues"].sum()
    )

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "🏪 Client Stores",
        f"{total_client_stores:,}"
    )

    col2.metric(
        "➕ New Stores",
        f"{total_new:,}"
    )

    col3.metric(
        "🔄 Changed Stores",
        f"{total_changed:,}"
    )

    col4, col5, col6 = st.columns(3)

    col4.metric(
        "➖ Removed",
        f"{total_removed:,}"
    )

    col5.metric(
        "✅ Unchanged",
        f"{total_unchanged:,}"
    )

    col6.metric(
        "⚠️ Data Issues",
        f"{total_issues:,}"
    )

    # ========================================================
    # TABS
    # ========================================================

    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 Summary",
        "🔄 Changes",
        "⚠️ Data Quality",
        "📄 Final Master"
    ])

    with tab1:

        st.dataframe(
            summary_df,
            use_container_width=True,
            hide_index=True
        )

    with tab2:

        if changed_df.empty:

            st.success(
                "No store changes detected."
            )

        else:

            st.dataframe(
                changed_df,
                use_container_width=True,
                hide_index=True
            )

    with tab3:

        if issues_df.empty:

            st.success(
                "No data-quality issues detected."
            )

        else:

            st.dataframe(
                issues_df,
                use_container_width=True,
                hide_index=True
            )

    with tab4:

        st.dataframe(
            final_df.head(500),
            use_container_width=True,
            hide_index=True
        )

        if len(final_df) > 500:

            st.caption(
                f"Showing first 500 of {len(final_df):,} rows."
            )

    # ========================================================
    # DOWNLOADS
    # ========================================================

    st.divider()

    st.subheader("⬇️ Download Results")

    csv_data = final_df.to_csv(
        index=False
    ).encode("utf-8")

    excel_data = create_excel_report(
        final_df,
        summary_df,
        changed_df,
        issues_df
    )

    col1, col2 = st.columns(2)

    with col1:

        st.download_button(
            label="⬇️ Download Final CSV",
            data=csv_data,
            file_name="master_call_cycle_updated.csv",
            mime="text/csv",
            use_container_width=True
        )

    with col2:

        st.download_button(
            label="📊 Download Full Excel Report",
            data=excel_data,
            file_name="call_cycle_reconciliation_report.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

    st.success(
        f"Reconciliation completed successfully. "
        f"{len(final_df):,} rows are ready for download."
    )
