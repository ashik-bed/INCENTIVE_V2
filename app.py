"""
SD Incentive Automation System - Streamlit App (single-file)

Save this file as `sd_incentive_app.py` and run:
    pip install streamlit pandas openpyxl
    streamlit run sd_incentive_app.py

Features implemented (matches your spec):
- Upload OLD-OS, NEW-OS, SD-INCENTIVE files (CSV/XLSX)
- Auto-detect common column names (deposit, new account, scheme, realisation date, etc.)
- Preview input files
- Date picker to compute up to a chosen date
- Robust parsing of interest rates (handles percent strings, numeric strings)
- Old-vs-New customer detection using OLD-OS override
- Correct incentive formulas:
    OLD: Deposit * Interest / 12
    NEW: Deposit * Interest * (No_of_Days / 365)
- Handles missing/invalid realisation dates and marks remarks accordingly
- Downloadable Excel output

Notes about interest parsing:
- If input contains a '%' sign, it's interpreted literally ("0.35%" -> 0.0035).
- If value is numeric and > 1, it is treated as a percentage ("35" -> 0.35).
- If numeric between 0 and 1, it's treated as decimal ("0.35" -> 0.35).

"""

import streamlit as st
import pandas as pd
import io
from typing import Optional
from datetime import date as dt_date, datetime as dt_datetime

# --- Page setup ---
st.set_page_config(page_title="SD Incentive Automation System", layout="wide")
st.title("SD Incentive Automation System — Streamlit")
st.markdown("""
This app computes SD incentives for OLD and NEW customers based on the files you upload.
""")

# ---------- Constants & helpers ----------
COMMON_DEPOSIT_COLS = [
    "deposit amount", "depositamount", "deposit", "amount", "outstanding", "outstanding amount", "balance", "bal"
]
COMMON_NEWACC_COLS = [
    "new account number", "newaccountnumber", "newaccno", "account number", "accountno", "accno", "new acc no", "newacc",
    "account"
]
COMMON_SCHEME_COLS = ["scheme code", "schemecode", "scheme", "scheme_name", "schemename", "scheme code"]
COMMON_BRANCH_COLS = ["branch name", "branch", "branchname"]
COMMON_CUSTID_COLS = ["customer id", "customerid", "cust id", "custid", "customer_id"]
COMMON_CUSTNAME_COLS = ["customer name", "customername", "cust name", "custname", "customer_name"]
COMMON_CANVASS_COLS = ["canvassed by", "canvassedby", "canvasser", "employee", "collected by"]
COMMON_OLD_INC_COLS = ["old_incentive", "oldincentive", "old_incent", "oldinc", "old incentive"]
COMMON_REAL_DATE_COLS = [
    "realisation date", "realization date", "realiztation date", "realisationdate", "realisation", "realized date",
    "realisation date", "realisation_date", "realised date", "date of realisation", "realisationdate"
]


def normalize_col_name(c: str) -> str:
    return ''.join(ch.lower() for ch in str(c) if ch.isalnum())


def find_column(df: pd.DataFrame, candidates) -> Optional[str]:
    cols_norm = {normalize_col_name(c): c for c in df.columns}
    # exact normalized match
    for cand in candidates:
        key = normalize_col_name(cand)
        if key in cols_norm:
            return cols_norm[key]
    # partial contains fallback
    for cand in candidates:
        key = normalize_col_name(cand)
        for k, orig in cols_norm.items():
            if key in k:
                return orig
    return None


def read_uploaded_file(uploaded) -> Optional[pd.DataFrame]:
    if uploaded is None:
        return None
    try:
        uploaded.seek(0)
    except Exception:
        pass
    name = getattr(uploaded, 'name', '') or ''
    lower_name = name.lower()
    # Try excel first
    if lower_name.endswith(('.xls', '.xlsx', '.xlsm', '.xlsb')):
        try:
            return pd.read_excel(uploaded, engine='openpyxl')
        except Exception:
            try:
                uploaded.seek(0)
            except Exception:
                pass
            try:
                return pd.read_excel(uploaded)
            except Exception:
                pass
    # General attempt
    try:
        uploaded.seek(0)
        return pd.read_excel(uploaded, engine='openpyxl')
    except Exception:
        pass
    encodings_to_try = ["utf-8", "cp1252", "latin1", "iso-8859-1"]
    for enc in encodings_to_try:
        try:
            uploaded.seek(0)
            return pd.read_csv(uploaded, encoding=enc)
        except Exception:
            continue
    try:
        uploaded.seek(0)
        raw = uploaded.read()
        if isinstance(raw, bytes):
            for enc in encodings_to_try:
                try:
                    text = raw.decode(enc)
                    return pd.read_csv(io.StringIO(text))
                except Exception:
                    continue
    except Exception:
        pass
    st.error("Failed to read file. Please provide a valid Excel (.xls/.xlsx/.xlsm) or CSV file.")
    return None


def parse_interest_value(x) -> float:
    """Parse many interest formats into a decimal fraction.
    Returns 0.0 if cannot parse.
    Rules:
    - If string contains '%': strip and divide by 100.
    - If numeric and > 1 treat as percentage (35 -> 0.35)
    - If numeric between 0 and 1 treat as decimal (0.35 -> 0.35)
    """
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return 0.0
        if isinstance(x, str):
            s = x.strip()
            if s == '':
                return 0.0
            if '%' in s:
                s_clean = s.replace('%', '').replace(',', '')
                return float(s_clean) / 100.0
            # attempt to parse numeric string
            s_clean = s.replace(',', '')
            val = float(s_clean)
            if val > 1:
                return val / 100.0
            return val
        # numeric types
        val = float(x)
        if val > 1:
            return val / 100.0
        return val
    except Exception:
        return 0.0


def load_incentive_map(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    sc_col = find_column(df, COMMON_SCHEME_COLS)
    int_col = None
    for maybe in ['interest', 'rate', 'interest rate', 'value']:
        col = find_column(df, [maybe])
        if col:
            int_col = col
            break
    if sc_col is None:
        sc_col = df.columns[0]
    if int_col is None:
        if len(df.columns) > 1:
            int_col = df.columns[1]
        else:
            raise ValueError("Couldn't find interest column in SD-INCENTIVE file")
    df = df[[sc_col, int_col]].rename(columns={sc_col: 'SchemeCode', int_col: 'InterestRaw'})
    df['SchemeCode'] = df['SchemeCode'].astype(str).str.strip()
    df['Interest'] = df['InterestRaw'].apply(parse_interest_value)
    df['SchemeKey'] = df['SchemeCode'].apply(lambda x: normalize_col_name(x))
    df = df.drop_duplicates(subset=['SchemeKey'], keep='last').set_index('SchemeKey')
    return df


def build_old_incentive_map(df: pd.DataFrame) -> pd.Series:
    df = df.copy()
    newacc_col = find_column(df, COMMON_NEWACC_COLS)
    oldinc_col = find_column(df, COMMON_OLD_INC_COLS)
    # If users used different headers, try reasonable fallbacks
    if newacc_col is None:
        # try first column as account
        newacc_col = df.columns[0]
    if oldinc_col is None:
        # try second column if exists
        if len(df.columns) > 1:
            oldinc_col = df.columns[1]
    if newacc_col is None or oldinc_col is None:
        st.error("OLD-OS must contain: New Account Number and OLD_INCENTIVE (column names flexible).")
        return pd.Series(dtype=float)
    s = pd.Series(df[oldinc_col].values, index=df[newacc_col].astype(str).str.strip().apply(normalize_col_name))
    s = s.apply(parse_interest_value)
    s.name = 'OLD_INTEREST'
    return s


def prepare_os_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Incentive')
    return output.getvalue()

# ---------- UI: File uploads ----------
with st.sidebar:
    st.header("Upload files")
    old_file = st.file_uploader("OLD-OS (New Account Number + OLD_INCENTIVE)", type=["csv", "xlsx", "xls"], key="old")
    new_file = st.file_uploader("NEW-OS (full outstanding)", type=["csv", "xlsx", "xls"], key="new")
    sd_file = st.file_uploader("SD-INCENTIVE (Scheme Code -> Interest)", type=["csv", "xlsx", "xls"], key="sd")
    st.markdown("---")
    st.write("Interest parsing behavior (you can override):")
    interpret_hint = st.selectbox("If numeric values are provided (no % sign), treat as:",
                                  options=["Auto (common heuristics)", "As Decimal (0.35 -> 0.35)", "As Percent (35 -> 0.35, 0.35 -> 0.0035)"], index=0)
    st.markdown("---")
    st.write("Date selection for incentive computation")
    calc_date = st.date_input("Calculate up to (inclusive)", value=dt_date.today())
    st.write("Tip: use a past date for backdated calc or a future date for projection.")

# ensure calc_date is date object
if isinstance(calc_date, dt_datetime):
    calc_date = calc_date.date()

old_df = read_uploaded_file(old_file)
new_df = read_uploaded_file(new_file)
sd_df = read_uploaded_file(sd_file)

old_map = pd.Series(dtype=float)
sd_map = None

if old_df is not None:
    try:
        old_map = build_old_incentive_map(old_df)
    except Exception as e:
        st.error(f"Failed to parse OLD-OS: {e}")
        old_map = pd.Series(dtype=float)

if sd_df is not None:
    try:
        sd_map = load_incentive_map(sd_df)
    except Exception as e:
        st.error(f"Failed to parse SD-INCENTIVE: {e}")
        sd_map = None

# Preview
with st.expander("Preview uploaded files and detected columns (click to expand)"):
    st.subheader("OLD-OS preview")
    if old_df is not None:
        st.dataframe(old_df.head(10))
    else:
        st.write("No OLD-OS uploaded")
    st.subheader("NEW-OS preview")
    if new_df is not None:
        st.dataframe(new_df.head(10))
    else:
        st.write("No NEW-OS uploaded")
    st.subheader("SD-INCENTIVE preview")
    if sd_df is not None:
        st.dataframe(sd_df.head(20))
    else:
        st.write("No SD-INCENTIVE uploaded")

st.markdown("---")
st.write("When ready, press **Compute Combined Incentive** to run the engine.")

if st.button("COMPUTE COMBINED INCENTIVE"):
    if new_df is None:
        st.error("Please upload NEW-OS (full outstanding) first.")
    else:
        new = prepare_os_df(new_df)
        # detect columns
        keys = {
            'deposit': find_column(new, COMMON_DEPOSIT_COLS),
            'newacc': find_column(new, COMMON_NEWACC_COLS),
            'scheme': find_column(new, COMMON_SCHEME_COLS),
            'branch': find_column(new, COMMON_BRANCH_COLS),
            'customer_id': find_column(new, COMMON_CUSTID_COLS),
            'customer_name': find_column(new, COMMON_CUSTNAME_COLS),
            'canvassed_by': find_column(new, COMMON_CANVASS_COLS),
            'realisation_date': find_column(new, COMMON_REAL_DATE_COLS)
        }

        # user-friendly reporting of detected columns
        st.info("Detected columns (from NEW-OS):\n" + '\n'.join([f"{k}: {v if v else 'NOT FOUND'}" for k, v in keys.items()]))

        if keys['newacc'] is None:
            st.error("Could not find 'New Account Number' column in NEW-OS. Please ensure it exists.")
        elif keys['deposit'] is None:
            st.error("Could not find deposit/outstanding column in NEW-OS. Please ensure it exists.")
        else:
            # coerce deposit to numeric robustly
            try:
                new[keys['deposit']] = pd.to_numeric(new[keys['deposit']], errors='coerce').fillna(0)
            except Exception:
                def try_parse_num(x):
                    try:
                        if pd.isna(x):
                            return 0.0
                        s = str(x).replace(',', '').strip()
                        return float(s)
                    except Exception:
                        return 0.0
                new[keys['deposit']] = pd.Series([try_parse_num(x) for x in new[keys['deposit']]], index=new.index)

            # ensure NewAccVal exists for normalization
            new['NewAccVal'] = new[keys['newacc']].astype(str).str.strip()

            # parse realisation dates (dayfirst=True)
            real_col = keys.get('realisation_date')
            if real_col and real_col in new.columns:
                new['_RealisationParsed'] = pd.to_datetime(new[real_col].astype(str).str.strip().replace('nan', ''),
                                                            errors='coerce', dayfirst=True, infer_datetime_format=True)
            else:
                new['_RealisationParsed'] = pd.NaT

            # calculation arrays
            interests = []
            remarks = []
            days_list = []
            incentives = []

            for idx, row in new.iterrows():
                interest_val = 0.0
                acct_key = normalize_col_name(str(row['NewAccVal']))
                is_old = False

                # check OLD override
                if not old_map.empty and acct_key in old_map.index:
                    interest_val = float(old_map.loc[acct_key])
                    is_old = True
                else:
                    # scheme lookup
                    if keys['scheme'] and keys['scheme'] in row.index and sd_map is not None:
                        sk = normalize_col_name(str(row[keys['scheme']]))
                        if sk in sd_map.index:
                            interest_val = float(sd_map.loc[sk, 'Interest'])

                # remark default
                remark = "Old Customer" if is_old else "New Customer"

                # deposit
                try:
                    deposit_val = float(row[keys['deposit']]) if keys['deposit'] in row.index and pd.notnull(row[keys['deposit']]) else 0.0
                except Exception:
                    try:
                        deposit_val = float(str(row.get(keys['deposit'], '')).replace(',', '').strip())
                    except Exception:
                        deposit_val = 0.0

                # handle parsed date
                real_parsed = row.get('_RealisationParsed', pd.NaT)
                if pd.isna(real_parsed):
                    no_of_days = None
                else:
                    try:
                        real_date_only = real_parsed.date() if hasattr(real_parsed, 'date') else pd.to_datetime(real_parsed).date()
                    except Exception:
                        real_date_only = None

                    if real_date_only is None:
                        no_of_days = None
                    else:
                        try:
                            no_of_days = abs((calc_date - real_date_only).days)
                        except Exception:
                            no_of_days = None

                # calculation
                if is_old:
                    incentive_val = deposit_val * interest_val / 12.0
                else:
                    if no_of_days is None:
                        incentive_val = pd.NA
                        remark = "New Customer - Invalid Realisation Date"
                    else:
                        annual_incentive = deposit_val * interest_val
                        incentive_val = annual_incentive * (no_of_days / 365.0)

                interests.append(interest_val)
                remarks.append(remark)
                days_list.append(no_of_days if no_of_days is not None else pd.NA)
                incentives.append(incentive_val if (incentive_val is not None) else pd.NA)

            # attach results
            new['_Interest'] = pd.Series(interests, index=new.index)
            new['_Remark'] = pd.Series(remarks, index=new.index)
            new['_No_of_Days'] = pd.Series(days_list, index=new.index)
            new['_Incentive'] = pd.Series(incentives, index=new.index)

            # build output table
            out = pd.DataFrame(index=new.index)
            out['Branch'] = new[keys['branch']] if keys['branch'] in new.columns else ''
            out['Customer ID'] = new[keys['customer_id']] if keys['customer_id'] in new.columns else ''
            out['New Account Number'] = new[keys['newacc']]
            out['Scheme Code'] = new[keys['scheme']] if keys['scheme'] in new.columns else ''
            out['Customer Name'] = new[keys['customer_name']] if keys['customer_name'] in new.columns else ''
            out['Deposit'] = new[keys['deposit']]
            out['Canvassed By'] = new[keys['canvassed_by']] if keys['canvassed_by'] in new.columns else ''
            if real_col and real_col in new.columns:
                out['Realisation Date'] = new[real_col]
            else:
                out['Realisation Date'] = ''
            out['Interest'] = new['_Interest']
            out['No_of_Days'] = new['_No_of_Days']
            out['Incentive'] = new['_Incentive']
            out['Remark'] = new['_Remark']

            total_rows = out.shape[0]
            invalid_new_count = int(((out['Remark'].str.contains("Invalid Realisation Date", na=False)) & (~out['Remark'].isna())).sum())

            st.success(f"Computed combined incentive for {total_rows} rows (from NEW-OS).")
            if invalid_new_count > 0:
                st.warning(f"{invalid_new_count} rows are New Customers but have invalid/missing Realisation Date — their New Incentive could not be computed (marked empty).")

            st.dataframe(out.head(200))

            # download
            excel_bytes = to_excel_bytes(out.reset_index(drop=True))
            st.download_button(label="Download Combined Incentive Excel", data=excel_bytes,
                               file_name="combined_incentive.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # small summary metrics
            st.markdown("---")
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.metric("Rows processed", total_rows)
            with col_b:
                st.metric("Invalid Realisation (new)", invalid_new_count)
            with col_c:
                st.metric("Total Incentive (sum, numeric)", round(pd.to_numeric(out['Incentive'], errors='coerce').sum(skipna=True), 2))

st.markdown("\n---\nDeveloped to match the SD Incentive Automation System spec. Save this file and run with Streamlit.")
