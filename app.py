import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Call Cycle Reconciler",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CONSTANTS
# ============================================================

CLIENT_COLUMNS = [
    "Week Number",
    "Day of Week",
    "User Code",
    "User Name",
    "Location Code",
    "Customer Name",
    "Callage",
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
    "Month End On",
]

DUPLICATE_COLUMNS = [
    "Severity",
    "Duplicate Type",
    "User Code",
    "User Name",
    "Store ID",
    "Store Name",
    "Week Number",
    "Day of Week",
    "Callage",
    "Source File",
    "Duplicate Count",
    "Details",
]


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 3rem;
        }

        .app-title {
            font-size: 2.2rem;
            font-weight: 700;
            margin-bottom: 0.2rem;
        }

        .app-subtitle {
            color: #6b7280;
            font-size: 1rem;
            margin-bottom: 1rem;
        }

        .section-title {
            font-size: 1.25rem;
            font-weight: 700;
            margin-top: 1rem;
            margin-bottom: 0.5rem;
        }

        div[data-testid="stMetric"] {
            border: 1px solid #e5e7eb;
            padding: 12px;
            border-radius: 10px;
        }

        .info-box {
            padding: 12px 15px;
            border-radius: 8px;
            background: #f5f7fa;
            border: 1px solid #e5e7eb;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="app-title">📅 Call Cycle Reconciler</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="app-subtitle">
    Validate client call-cycle files, compare them with the master,
    identify duplicates/conflicts and generate a clean final master.
    </div>
    """,
    unsafe_allow_html=True,
)

st.divider()


# ============================================================
# SESSION STATE
# ============================================================

if "results" not in st.session_state:
    st.session_state.results = None


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_cell(value):
    """Convert cell to clean string and remove Excel .0 from numeric IDs."""

    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.endswith(".0") and value[:-2].isdigit():
        value = value[:-2]

    return value


def read_uploaded_file(uploaded_file):
    """Read CSV/XLS/XLSX/XLSM."""

    extension = os.path.splitext(
        uploaded_file.name
    )[1].lower()

    uploaded_file.seek(0)

    if extension in [".csv", ".txt"]:

        df = pd.read_csv(
            uploaded_file,
            dtype=str,
            keep_default_na=False,
        )

    elif extension in [".xlsx", ".xlsm"]:

        df = pd.read_excel(
            uploaded_file,
            dtype=str,
            engine="openpyxl",
        )

    elif extension == ".xls":

        df = pd.read_excel(
            uploaded_file,
            dtype=str,
            engine="xlrd",
        )

    else:

        raise ValueError(
            f"Unsupported file format: {extension}"
        )

    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    return df.fillna("")


def normalize_client_file(uploaded_file):
    """
    Load and validate one client file.
    """

    try:

        raw = read_uploaded_file(
            uploaded_file
        )

    except Exception as exc:

        return None, f"Could not read file: {exc}"

    if raw.empty:

        return None, "File is empty."

    lower_columns = {
        str(column).lower(): column
        for column in raw.columns
    }

    mapping = {}
    missing = []

    for required_column in CLIENT_COLUMNS:

        actual_column = lower_columns.get(
            required_column.lower()
        )

        if actual_column:

            mapping[required_column] = actual_column

        else:

            missing.append(
                required_column
            )

    if missing:

        return None, (
            "Missing column(s): "
            + ", ".join(missing)
        )

    df = pd.DataFrame(
        {
            required_column:
            raw[mapping[required_column]].map(clean_cell)

            for required_column in CLIENT_COLUMNS
        }
    )

    # --------------------------------------------------------
    # Convert week
    # --------------------------------------------------------

    df["Week Number"] = pd.to_numeric(
        df["Week Number"],
        errors="coerce",
    )

    invalid_week_rows = int(
        df["Week Number"].isna().sum()
    )

    df = df.dropna(
        subset=["Week Number"]
    )

    df["Week Number"] = (
        df["Week Number"]
        .astype(int)
    )

    # --------------------------------------------------------
    # Remove rows with essential blanks
    # --------------------------------------------------------

    essential_columns = [
        "Location Code",
        "User Code",
        "Day of Week",
    ]

    for column in essential_columns:

        df[column] = (
            df[column]
            .astype(str)
            .str.strip()
        )

    df = df[
        (df["Location Code"] != "")
        &
        (df["User Code"] != "")
        &
        (df["Day of Week"] != "")
    ].copy()

    df["source_file"] = uploaded_file.name

    df["_invalid_week_input_count"] = (
        invalid_week_rows
    )

    return df, None


def normalize_master_file(uploaded_file):
    """Load and validate master file."""

    try:

        df = read_uploaded_file(
            uploaded_file
        )

    except Exception as exc:

        return None, f"Could not read master: {exc}"

    missing = [
        column
        for column in MASTER_COLUMNS
        if column not in df.columns
    ]

    if missing:

        return None, (
            "Master file is missing: "
            + ", ".join(missing)
        )

    df = df[
        MASTER_COLUMNS
    ].fillna("").astype(str)

    for column in [
        "Store ID",
        "User ID",
        "Store Name",
    ]:

        df[column] = (
            df[column]
            .map(clean_cell)
        )

    return df, None


def most_common(values):
    """Return most common non-empty value."""

    values = [
        str(value).strip()
        for value in values
        if str(value).strip()
        and str(value).lower() != "nan"
    ]

    if not values:
        return ""

    counts = pd.Series(values).value_counts()

    return counts.index[0]


# ============================================================
# DUPLICATE DETECTION
# ============================================================

def detect_duplicates(client_df):
    """
    Detect duplicate and suspicious records.

    Duplicate categories:

    1. Exact Duplicate
    2. User + Store Duplicate
    3. Store Assigned to Multiple Users
    4. Store + User + Week Duplicate
    5. Cross-File Duplicate
    6. Conflicting Week Duplicate
    """

    reports = []

    def add_report(
        severity,
        duplicate_type,
        row,
        duplicate_count,
        details,
    ):

        reports.append(
            {
                "Severity": severity,
                "Duplicate Type": duplicate_type,
                "User Code": row.get(
                    "User Code", ""
                ),
                "User Name": row.get(
                    "User Name", ""
                ),
                "Store ID": row.get(
                    "Location Code", ""
                ),
                "Store Name": row.get(
                    "Customer Name", ""
                ),
                "Week Number": row.get(
                    "Week Number", ""
                ),
                "Day of Week": row.get(
                    "Day of Week", ""
                ),
                "Callage": row.get(
                    "Callage", ""
                ),
                "Source File": row.get(
                    "source_file", ""
                ),
                "Duplicate Count":
                    duplicate_count,
                "Details": details,
            }
        )

    # --------------------------------------------------------
    # 1. EXACT DUPLICATE
    # --------------------------------------------------------

    exact_keys = [
        "Week Number",
        "Day of Week",
        "User Code",
        "User Name",
        "Location Code",
        "Customer Name",
        "Callage",
    ]

    exact_mask = client_df.duplicated(
        subset=exact_keys,
        keep=False,
    )

    exact_df = client_df[
        exact_mask
    ].copy()

    if not exact_df.empty:

        counts = (
            exact_df
            .groupby(exact_keys, dropna=False)
            .size()
            .reset_index(
                name="Duplicate Count"
            )
        )

        for _, row in exact_df.iterrows():

            key_values = tuple(
                row[column]
                for column in exact_keys
            )

            match = counts[
                counts[exact_keys]
                .apply(
                    tuple,
                    axis=1
                )
                == key_values
            ]

            count = (
                int(
                    match.iloc[0]["Duplicate Count"]
                )
                if not match.empty
                else 2
            )

            add_report(
                "Medium",
                "Exact Duplicate",
                row,
                count,
                "Entire call-cycle record is duplicated.",
            )

    # --------------------------------------------------------
    # 2. USER + STORE DUPLICATE
    # --------------------------------------------------------

    user_store_keys = [
        "User Code",
        "Location Code",
    ]

    user_store_mask = client_df.duplicated(
        subset=user_store_keys,
        keep=False,
    )

    user_store_df = client_df[
        user_store_mask
    ].copy()

    if not user_store_df.empty:

        counts = (
            user_store_df
            .groupby(
                user_store_keys,
                dropna=False
            )
            .size()
            .reset_index(
                name="Duplicate Count"
            )
        )

        for _, row in user_store_df.iterrows():

            match = counts[
                (counts["User Code"]
                 == row["User Code"])
                &
                (counts["Location Code"]
                 == row["Location Code"])
            ]

            count = int(
                match.iloc[0]["Duplicate Count"]
            )

            add_report(
                "High",
                "User + Store Duplicate",
                row,
                count,
                "Same user has the same store more than once.",
            )

    # --------------------------------------------------------
    # 3. STORE ASSIGNED TO MULTIPLE USERS
    # --------------------------------------------------------

    for store_id, group in client_df.groupby(
        "Location Code",
        dropna=False,
    ):

        users = (
            group["User Code"]
            .dropna()
            .astype(str)
            .unique()
        )

        if len(users) <= 1:
            continue

        for _, row in group.iterrows():

            add_report(
                "High",
                "Store Assigned to Multiple Users",
                row,
                len(users),
                (
                    "Store is assigned to multiple users: "
                    + ", ".join(users)
                ),
            )

    # --------------------------------------------------------
    # 4. STORE + USER + WEEK DUPLICATE
    # --------------------------------------------------------

    week_keys = [
        "User Code",
        "Location Code",
        "Week Number",
    ]

    week_mask = client_df.duplicated(
        subset=week_keys,
        keep=False,
    )

    week_df = client_df[
        week_mask
    ].copy()

    if not week_df.empty:

        for _, row in week_df.iterrows():

            group = week_df[
                (week_df["User Code"]
                 == row["User Code"])
                &
                (week_df["Location Code"]
                 == row["Location Code"])
                &
                (week_df["Week Number"]
                 == row["Week Number"])
            ]

            unique_variants = group[
                [
                    "Day of Week",
                    "Callage",
                ]
            ].drop_duplicates()

            if len(unique_variants) > 1:

                duplicate_type = (
                    "Conflicting Week Duplicate"
                )

                severity = "High"

                details = (
                    "Same Store/User/Week has "
                    "different Day or Callage."
                )

            else:

                duplicate_type = (
                    "Store + User + Week Duplicate"
                )

                severity = "Medium"

                details = (
                    "Same Store/User/Week appears "
                    "more than once."
                )

            add_report(
                severity,
                duplicate_type,
                row,
                len(group),
                details,
            )

    # --------------------------------------------------------
    # 5. CROSS-FILE DUPLICATE
    # --------------------------------------------------------

    for (
        user_code,
        store_id,
    ), group in client_df.groupby(
        ["User Code", "Location Code"],
        dropna=False,
    ):

        files = (
            group["source_file"]
            .astype(str)
            .unique()
        )

        if len(files) <= 1:
            continue

        for _, row in group.iterrows():

            add_report(
                "High",
                "Cross-File Duplicate",
                row,
                len(files),
                (
                    "Same User + Store exists "
                    "in multiple uploaded files: "
                    + ", ".join(files)
                ),
            )

    # --------------------------------------------------------
    # REMOVE EXACTLY IDENTICAL REPORT ROWS
    # --------------------------------------------------------

    if not reports:

        return pd.DataFrame(
            columns=DUPLICATE_COLUMNS
        )

    duplicate_df = pd.DataFrame(
        reports
    )

    duplicate_df = (
        duplicate_df
        .drop_duplicates()
        .reset_index(drop=True)
    )

    severity_order = {
        "High": 0,
        "Medium": 1,
        "Low": 2,
    }

    duplicate_df["_sort"] = (
        duplicate_df["Severity"]
        .map(severity_order)
    )

    duplicate_df = (
        duplicate_df
        .sort_values(
            [
                "_sort",
                "User Code",
                "Store ID",
            ]
        )
        .drop(columns="_sort")
        .reset_index(drop=True)
    )

    return duplicate_df


# ============================================================
# CLIENT RECONCILIATION
# ============================================================

def analyze_client(
    client_df,
    user_code,
    master_df,
):
    """Apply original earliest-week-wins reconciliation logic."""

    rows = client_df[
        client_df["User Code"]
        .astype(str)
        == str(user_code)
    ].copy()

    if rows.empty:

        raise ValueError(
            f"No rows found for {user_code}"
        )

    user_name = rows.iloc[0][
        "User Name"
    ]

    computed = []

    flags = {
        "day_conflict": [],
        "callage_conflict": [],
        "week_conflict": [],
        "out_of_range": [],
        "cross_file": [],
    }

    # --------------------------------------------------------
    # STORE-LEVEL PROCESSING
    # --------------------------------------------------------

    for location_code, group in rows.groupby(
        "Location Code",
        sort=False,
    ):

        entries = group.to_dict(
            "records"
        )

        store_name = entries[0][
            "Customer Name"
        ]

        # ----------------------------------------------------
        # Invalid weeks
        # ----------------------------------------------------

        invalid = [
            entry
            for entry in entries
            if not (
                1 <= entry["Week Number"] <= 5
            )
        ]

        if invalid:

            flags["out_of_range"].append(
                {
                    "Store ID":
                        location_code,
                    "Store Name":
                        store_name,
                    "Week":
                        invalid[0][
                            "Week Number"
                        ],
                }
            )

        # ----------------------------------------------------
        # Same-week conflict
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
                    entry["Day of Week"],
                    entry["Callage"],
                )
                for entry in week_entries
            }

            if len(variants) > 1:

                flags[
                    "week_conflict"
                ].append(
                    {
                        "Store ID":
                            location_code,
                        "Store Name":
                            store_name,
                        "Week":
                            week,
                        "Variants":
                            " | ".join(
                                sorted(
                                    f"{day} / {callage}"
                                    for day, callage
                                    in variants
                                )
                            ),
                    }
                )

        # ----------------------------------------------------
        # Cross-file
        # ----------------------------------------------------

        source_files = {
            entry["source_file"]
            for entry in entries
        }

        if len(source_files) > 1:

            flags[
                "cross_file"
            ].append(
                {
                    "Store ID":
                        location_code,
                    "Store Name":
                        store_name,
                    "Files":
                        ", ".join(
                            sorted(source_files)
                        ),
                }
            )

        valid_weeks = sorted(
            {
                entry["Week Number"]
                for entry in entries
                if 1 <= entry["Week Number"] <= 5
            }
        )

        if not valid_weeks:
            continue

        # ----------------------------------------------------
        # EARLIEST WEEK WINS
        # ----------------------------------------------------

        earliest_week = valid_weeks[0]

        earliest_entries = [
            entry
            for entry in entries
            if entry["Week Number"]
            == earliest_week
        ]

        resolved_day = earliest_entries[0][
            "Day of Week"
        ]

        resolved_callage = earliest_entries[0][
            "Callage"
        ]

        # ----------------------------------------------------
        # Day conflict
        # ----------------------------------------------------

        days = {
            entry["Day of Week"]
            for entry in entries
        }

        if len(days) > 1:

            flags[
                "day_conflict"
            ].append(
                {
                    "Store ID":
                        location_code,
                    "Store Name":
                        store_name,
                    "Options":
                        " | ".join(
                            f"Week {entry['Week Number']}: "
                            f"{entry['Day of Week']}"
                            for entry in entries
                        ),
                    "Resolved":
                        resolved_day,
                }
            )

        # ----------------------------------------------------
        # Callage conflict
        # ----------------------------------------------------

        callages = {
            entry["Callage"]
            for entry in entries
        }

        if len(callages) > 1:

            flags[
                "callage_conflict"
            ].append(
                {
                    "Store ID":
                        location_code,
                    "Store Name":
                        store_name,
                    "Options":
                        " | ".join(
                            f"Week {entry['Week Number']}: "
                            f"{entry['Callage']}"
                            for entry in entries
                        ),
                    "Resolved":
                        resolved_callage,
                }
            )

        # ----------------------------------------------------
        # Applicable weeks
        # ----------------------------------------------------

        bitmask = [
            1 if week in valid_weeks else 0
            for week in [1, 2, 3, 4]
        ]

        # Week 5 = Week 1 OR Week 5
        bitmask.append(
            1
            if (
                5 in valid_weeks
                or 1 in valid_weeks
            )
            else 0
        )

        computed.append(
            {
                "Store ID":
                    location_code,

                "Store Name":
                    store_name,

                "Days of Week":
                    resolved_day,

                "Applicable Weeks":
                    ",".join(
                        map(
                            str,
                            bitmask
                        )
                    ),

                "Weeks":
                    ",".join(
                        map(
                            str,
                            valid_weeks
                        )
                    ),
            }
        )

    # ========================================================
    # MASTER COMPARISON
    # ========================================================

    if master_df is not None:

        existing = master_df[
            master_df["User ID"]
            .astype(str)
            == str(user_code)
        ].copy()

    else:

        existing = pd.DataFrame(
            columns=MASTER_COLUMNS
        )

    existing_rows = existing.to_dict(
        "records"
    )

    existing_by_store = {
        str(row["Store ID"]):
            row
        for row in existing_rows
    }

    computed_by_store = {
        str(row["Store ID"]):
            row
        for row in computed
    }

    new_stores = []

    changed_stores = []

    unchanged_stores = []

    for store_id, current in (
        computed_by_store.items()
    ):

        old = existing_by_store.get(
            store_id
        )

        if old is None:

            new_stores.append(
                current
            )

        elif (
            str(old["Days of Week"])
            != str(
                current["Days of Week"]
            )
            or
            str(old["Applicable Weeks"])
            != str(
                current[
                    "Applicable Weeks"
                ]
            )
        ):

            changed_stores.append(
                {
                    **current,
                    "Old Day":
                        old["Days of Week"],
                    "New Day":
                        current[
                            "Days of Week"
                        ],
                    "Old Applicable Weeks":
                        old[
                            "Applicable Weeks"
                        ],
                    "New Applicable Weeks":
                        current[
                            "Applicable Weeks"
                        ],
                }
            )

        else:

            unchanged_stores.append(
                current
            )

    extra_stores = [
        row
        for row in existing_rows
        if str(row["Store ID"])
        not in computed_by_store
    ]

    return {
        "user_code":
            str(user_code),

        "user_name":
            user_name,

        "computed":
            computed,

        "existing":
            existing_rows,

        "new":
            new_stores,

        "changed":
            changed_stores,

        "unchanged":
            unchanged_stores,

        "extra":
            extra_stores,

        "flags":
            flags,
    }


# ============================================================
# OUTPUT BUILDER
# ============================================================

def build_output_rows(
    analysis,
    start_date,
    frequency,
    end_date,
    month_start,
    month_end,
    keep_extra,
):
    """Create final master rows."""

    output = []

    for store in analysis["computed"]:

        output.append(
            {
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
                    store[
                        "Applicable Weeks"
                    ],

                "Month Start On":
                    month_start,

                "Month End On":
                    month_end,
            }
        )

    if keep_extra:

        output.extend(
            analysis["extra"]
        )

    return output


# ============================================================
# EXCEL REPORT
# ============================================================

def create_excel_report(
    final_df,
    summary_df,
    changes_df,
    quality_df,
    duplicates_df,
):
    """Create multi-sheet Excel report."""

    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="xlsxwriter",
    ) as writer:

        sheets = {
            "Final Master":
                final_df,

            "Summary":
                summary_df,

            "Changes":
                changes_df,

            "Data Quality":
                quality_df,

            "Duplicates":
                duplicates_df,
        }

        workbook = writer.book

        header_format = workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "text_wrap": True,
            }
        )

        for sheet_name, df in sheets.items():

            df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
            )

            worksheet = writer.sheets[
                sheet_name
            ]

            # Header
            for col_num, column in enumerate(
                df.columns
            ):

                worksheet.write(
                    0,
                    col_num,
                    column,
                    header_format,
                )

            # Width
            for col_num, column in enumerate(
                df.columns
            ):

                if df.empty:

                    width = len(
                        str(column)
                    ) + 2

                else:

                    max_length = (
                        df[column]
                        .astype(str)
                        .str.len()
                        .max()
                    )

                    width = min(
                        max(
                            len(
                                str(column)
                            ) + 2,
                            int(max_length) + 2,
                        ),
                        45,
                    )

                worksheet.set_column(
                    col_num,
                    col_num,
                    width,
                )

            worksheet.freeze_panes(
                1,
                0,
            )

            worksheet.autofilter(
                0,
                0,
                max(len(df), 1),
                max(len(df.columns) - 1, 0),
            )

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Configuration")

    client_files = st.file_uploader(
        "Upload Client Call Cycle File(s)",
        type=[
            "csv",
            "xlsx",
            "xls",
            "xlsm",
        ],
        accept_multiple_files=True,
    )

    st.divider()

    master_file = st.file_uploader(
        "Upload Existing Master File",
        type=[
            "csv",
            "xlsx",
        ],
        accept_multiple_files=False,
    )

    st.divider()

    st.subheader(
        "New Client Settings"
    )

    start_date = st.text_input(
        "Start Date",
        placeholder="DD-MM-YYYY",
    )

    frequency = st.selectbox(
        "Frequency",
        [
            "Weekly",
            "Fortnightly",
            "Monthly",
        ],
    )

    end_date = st.text_input(
        "End Date",
        placeholder="DD-MM-YYYY",
    )

    st.divider()

    keep_extra = st.checkbox(
        "Keep stores missing from client file",
        value=True,
        help=(
            "If enabled, stores present in the "
            "master but missing from the client "
            "file will be retained."
        ),
    )


# ============================================================
# INITIAL SCREEN
# ============================================================

if not client_files:

    st.info(
        "👈 Upload one or more client call-cycle files "
        "from the sidebar to start."
    )

    st.subheader(
        "Required Client Columns"
    )

    st.dataframe(
        pd.DataFrame(
            {
                "Required Column":
                    CLIENT_COLUMNS
            }
        ),
        hide_index=True,
        use_container_width=True,
    )

    st.subheader(
        "Required Master Columns"
    )

    st.dataframe(
        pd.DataFrame(
            {
                "Required Column":
                    MASTER_COLUMNS
            }
        ),
        hide_index=True,
        use_container_width=True,
    )

    st.stop()


# ============================================================
# LOAD CLIENT FILES
# ============================================================

st.subheader("📂 Uploaded Files")

all_client_data = []

file_status = []

for uploaded_file in client_files:

    df, error = normalize_client_file(
        uploaded_file
    )

    if error:

        file_status.append(
            {
                "File":
                    uploaded_file.name,
                "Status":
                    "❌ Error",
                "Rows":
                    0,
                "Message":
                    error,
            }
        )

    else:

        all_client_data.append(
            df
        )

        file_status.append(
            {
                "File":
                    uploaded_file.name,
                "Status":
                    "✅ Loaded",
                "Rows":
                    len(df),
                "Message":
                    "Valid",
            }
        )


st.dataframe(
    pd.DataFrame(file_status),
    hide_index=True,
    use_container_width=True,
)


if not all_client_data:

    st.error(
        "No valid client files available."
    )

    st.stop()


client_df = pd.concat(
    all_client_data,
    ignore_index=True,
)


# ============================================================
# MASTER LOAD
# ============================================================

master_df = None

if master_file:

    master_df, master_error = (
        normalize_master_file(
            master_file
        )
    )

    if master_error:

        st.error(
            f"❌ Master File Error: {master_error}"
        )

        st.stop()

    st.success(
        f"✅ Master loaded: "
        f"{len(master_df):,} rows"
    )


# ============================================================
# GLOBAL DUPLICATE CHECK
# ============================================================

duplicate_df = detect_duplicates(
    client_df
)


# ============================================================
# CLIENT LIST
# ============================================================

client_list = (
    client_df[
        [
            "User Code",
            "User Name",
        ]
    ]
    .drop_duplicates()
    .reset_index(drop=True)
)

client_list["Display"] = (
    client_list["User Name"].astype(str)
    + " — "
    + client_list["User Code"].astype(str)
)

st.subheader("👤 Client Selection")

selected_clients = st.multiselect(
    "Select client(s)",
    client_list["Display"].tolist(),
    default=client_list["Display"].tolist(),
)


if not selected_clients:

    st.warning(
        "Select at least one client."
    )

    st.stop()


selected_codes = client_list[
    client_list["Display"].isin(
        selected_clients
    )
]["User Code"].tolist()


# ============================================================
# RUN BUTTON
# ============================================================

run_button = st.button(
    "🚀 Run Reconciliation",
    type="primary",
    use_container_width=True,
)


# ============================================================
# PROCESS
# ============================================================

if run_button:

    if (
        any(
            [
                not str(start_date).strip()
                for _ in []
            ]
        )
    ):
        pass

    progress = st.progress(0)

    final_rows = []

    summary_rows = []

    change_rows = []

    quality_rows = []

    for index, user_code in enumerate(
        selected_codes
    ):

        analysis = analyze_client(
            client_df,
            user_code,
            master_df,
        )

        existing = analysis[
            "existing"
        ]

        # ----------------------------------------------------
        # New vs existing client
        # ----------------------------------------------------

        if not existing:

            sd = start_date

            fr = frequency

            ed = end_date

            ms = ""

            me = ""

        else:

            sd = most_common(
                [
                    row["Start Date"]
                    for row in existing
                ]
            )

            fr = (
                most_common(
                    [
                        row["Frequency"]
                        for row in existing
                    ]
                )
                or "Weekly"
            )

            ed = most_common(
                [
                    row["End Date"]
                    for row in existing
                ]
            )

            ms = most_common(
                [
                    row["Month Start On"]
                    for row in existing
                ]
            )

            me = most_common(
                [
                    row["Month End On"]
                    for row in existing
                ]
            )

        output_rows = build_output_rows(
            analysis,
            sd,
            fr,
            ed,
            ms,
            me,
            keep_extra,
        )

        final_rows.extend(
            output_rows
        )

        flags = analysis["flags"]

        issue_count = sum(
            len(value)
            for value in flags.values()
        )

        summary_rows.append(
            {
                "User Code":
                    user_code,

                "User Name":
                    analysis[
                        "user_name"
                    ],

                "Client Stores":
                    len(
                        analysis[
                            "computed"
                        ]
                    ),

                "New Stores":
                    len(
                        analysis["new"]
                    ),

                "Changed Stores":
                    len(
                        analysis[
                            "changed"
                        ]
                    ),

                "Missing from Client":
                    len(
                        analysis["extra"]
                    ),

                "Unchanged Stores":
                    len(
                        analysis[
                            "unchanged"
                        ]
                    ),

                "Data Quality Issues":
                    issue_count,
            }
        )

        # ----------------------------------------------------
        # Changes
        # ----------------------------------------------------

        for change in analysis[
            "changed"
        ]:

            change_rows.append(
                {
                    "User Code":
                        user_code,

                    "User Name":
                        analysis[
                            "user_name"
                        ],

                    "Store ID":
                        change[
                            "Store ID"
                        ],

                    "Store Name":
                        change[
                            "Store Name"
                        ],

                    "Old Day":
                        change[
                            "Old Day"
                        ],

                    "New Day":
                        change[
                            "New Day"
                        ],

                    "Old Applicable Weeks":
                        change[
                            "Old Applicable Weeks"
                        ],

                    "New Applicable Weeks":
                        change[
                            "New Applicable Weeks"
                        ],
                }
            )

        # ----------------------------------------------------
        # Quality
        # ----------------------------------------------------

        for issue_type, records in (
            flags.items()
        ):

            for record in records:

                quality_rows.append(
                    {
                        "User Code":
                            user_code,

                        "User Name":
                            analysis[
                                "user_name"
                            ],

                        "Issue Type":
                            issue_type,

                        **record,
                    }
                )

        progress.progress(
            (index + 1)
            / len(selected_codes)
        )

    # ========================================================
    # BUILD DATAFRAMES
    # ========================================================

    final_df = pd.DataFrame(
        final_rows,
        columns=MASTER_COLUMNS,
    )

    summary_df = pd.DataFrame(
        summary_rows
    )

    changes_df = pd.DataFrame(
        change_rows
    )

    quality_df = pd.DataFrame(
        quality_rows
    )

    # Filter duplicate report for selected users
    selected_duplicate_df = duplicate_df[
        duplicate_df["User Code"]
        .astype(str)
        .isin(
            [
                str(code)
                for code in selected_codes
            ]
        )
    ].copy()

    st.session_state.results = {
        "final":
            final_df,

        "summary":
            summary_df,

        "changes":
            changes_df,

        "quality":
            quality_df,

        "duplicates":
            selected_duplicate_df,
    }

    st.success(
        "✅ Reconciliation completed successfully."
    )


# ============================================================
# RESULTS
# ============================================================

if st.session_state.results:

    results = st.session_state.results

    final_df = results["final"]

    summary_df = results["summary"]

    changes_df = results["changes"]

    quality_df = results["quality"]

    duplicates_df = results[
        "duplicates"
    ]

    st.divider()

    # ========================================================
    # DASHBOARD
    # ========================================================

    st.subheader(
        "📊 Reconciliation Dashboard"
    )

    total_stores = int(
        summary_df[
            "Client Stores"
        ].sum()
    )

    total_new = int(
        summary_df[
            "New Stores"
        ].sum()
    )

    total_changed = int(
        summary_df[
            "Changed Stores"
        ].sum()
    )

    total_missing = int(
        summary_df[
            "Missing from Client"
        ].sum()
    )

    total_unchanged = int(
        summary_df[
            "Unchanged Stores"
        ].sum()
    )

    total_quality = int(
        summary_df[
            "Data Quality Issues"
        ].sum()
    )

    total_duplicates = len(
        duplicates_df
    )

    high_duplicates = int(
        (
            duplicates_df[
                "Severity"
            ] == "High"
        ).sum()
    )

    col1, col2, col3, col4 = st.columns(
        4
    )

    col1.metric(
        "🏪 Client Stores",
        f"{total_stores:,}",
    )

    col2.metric(
        "➕ New Stores",
        f"{total_new:,}",
    )

    col3.metric(
        "🔄 Changed",
        f"{total_changed:,}",
    )

    col4.metric(
        "➖ Missing",
        f"{total_missing:,}",
    )

    col5, col6, col7, col8 = st.columns(
        4
    )

    col5.metric(
        "✅ Unchanged",
        f"{total_unchanged:,}",
    )

    col6.metric(
        "⚠️ Quality Issues",
        f"{total_quality:,}",
    )

    col7.metric(
        "🔴 Duplicate Records",
        f"{total_duplicates:,}",
    )

    col8.metric(
        "🚨 High Severity",
        f"{high_duplicates:,}",
    )

    # ========================================================
    # TABS
    # ========================================================

    (
        tab_summary,
        tab_changes,
        tab_quality,
        tab_duplicates,
        tab_master,
    ) = st.tabs(
        [
            "📋 Summary",
            "🔄 Changes",
            "⚠️ Data Quality",
            "🔴 Duplicates",
            "📄 Final Master",
        ]
    )

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    with tab_summary:

        st.dataframe(
            summary_df,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------
    # CHANGES
    # --------------------------------------------------------

    with tab_changes:

        if changes_df.empty:

            st.success(
                "✅ No changes found."
            )

        else:

            st.dataframe(
                changes_df,
                use_container_width=True,
                hide_index=True,
            )

            changes_csv = (
                changes_df
                .to_csv(index=False)
                .encode("utf-8")
            )

            st.download_button(
                "⬇️ Download Changes Report",
                changes_csv,
                "changes_report.csv",
                "text/csv",
            )

    # --------------------------------------------------------
    # DATA QUALITY
    # --------------------------------------------------------

    with tab_quality:

        if quality_df.empty:

            st.success(
                "✅ No data-quality conflicts found."
            )

        else:

            st.dataframe(
                quality_df,
                use_container_width=True,
                hide_index=True,
            )

            quality_csv = (
                quality_df
                .to_csv(index=False)
                .encode("utf-8")
            )

            st.download_button(
                "⬇️ Download Data Quality Report",
                quality_csv,
                "data_quality_report.csv",
                "text/csv",
            )

    # --------------------------------------------------------
    # DUPLICATES
    # --------------------------------------------------------

    with tab_duplicates:

        if duplicates_df.empty:

            st.success(
                "✅ No duplicate records found."
            )

        else:

            st.warning(
                f"⚠️ {len(duplicates_df):,} "
                "duplicate/suspicious records detected."
            )

            # Duplicate summary
            duplicate_summary = (
                duplicates_df[
                    "Duplicate Type"
                ]
                .value_counts()
                .reset_index()
            )

            duplicate_summary.columns = [
                "Duplicate Type",
                "Records",
            ]

            st.dataframe(
                duplicate_summary,
                use_container_width=True,
                hide_index=True,
            )

            st.markdown(
                "### 🔎 Duplicate Details"
            )

            st.dataframe(
                duplicates_df,
                use_container_width=True,
                hide_index=True,
            )

            duplicate_csv = (
                duplicates_df
                .to_csv(index=False)
                .encode("utf-8")
            )

            st.download_button(
                "⬇️ Download Duplicate Report",
                duplicate_csv,
                "duplicate_report.csv",
                "text/csv",
                use_container_width=True,
            )

    # --------------------------------------------------------
    # FINAL MASTER
    # --------------------------------------------------------

    with tab_master:

        st.caption(
            f"Final master contains "
            f"{len(final_df):,} rows."
        )

        st.dataframe(
            final_df.head(1000),
            use_container_width=True,
            hide_index=True,
        )

        if len(final_df) > 1000:

            st.info(
                "Preview limited to first 1,000 rows. "
                "Download the complete file below."
            )

    # ========================================================
    # DOWNLOAD CENTER
    # ========================================================

    st.divider()

    st.subheader(
        "⬇️ Download Center"
    )

    final_csv = (
        final_df
        .to_csv(index=False)
        .encode("utf-8")
    )

    excel_file = create_excel_report(
        final_df,
        summary_df,
        changes_df,
        quality_df,
        duplicates_df,
    )

    col1, col2 = st.columns(2)

    with col1:

        st.download_button(
            "📄 Download Final Master CSV",
            final_csv,
            "master_call_cycle_updated.csv",
            "text/csv",
            use_container_width=True,
        )

    with col2:

        st.download_button(
            "📊 Download Complete Excel Report",
            excel_file,
            "call_cycle_reconciliation_report.xlsx",
            (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    st.caption(
        "The Excel report contains Final Master, Summary, "
        "Changes, Data Quality and Duplicates sheets."
    )
