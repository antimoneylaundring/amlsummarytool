import streamlit as st
import pandas as pd
import time
import re
import datetime
import os
from dotenv import load_dotenv
import psycopg2
from psycopg2 import extras, OperationalError, InterfaceError
from sqlalchemy import create_engine, text

st.set_page_config(page_title="UPI/Bank Import & Check", layout="wide")

# Load environment variables
load_dotenv()

# Database Configuration from .env
DB_URL = os.getenv("DB_URL")

if not DB_URL:
    st.error("❌ DB_URL not found in .env file. Please add your database connection string.")
    st.stop()

# Convert SQLAlchemy format to psycopg2 format if needed
if DB_URL and 'postgresql+psycopg2://' in DB_URL:
    DB_URL = DB_URL.replace('postgresql+psycopg2://', 'postgresql://')

@st.cache_resource
def get_engine():
    return create_engine(
        DB_URL,
        pool_pre_ping=True,
        connect_args={"connect_timeout": 30},
    )

engine = get_engine()

st.markdown("""
    <style>
            /* Headers styling */
            h1 {
                font-weight: 700;
                padding: 0;
                font-size: 2rem;
            }
            /* Main container styling */
            .block-container {
                padding: 3rem 2rem;
            }
    </style>
""", unsafe_allow_html=True)

# Table options and their required columns + conflict column + filter value
TABLE_OPTIONS = {
    "UPI": {
        "table_name": "all_upiiD",
        "required": ["Upi_vpa", "Inserted_date"],
        "conflict_col": "Upi_vpa",
        "filter_value": "UPI"
    },
    "Bank Account": {
        "table_name": "all_bank_acc",
        "required": ["Bank_account_number", "Inserted_date"],
        "conflict_col": "Bank_account_number",
        "filter_value": "Bank Account"
    }
}

def get_db_connection():
    """Create database connection with keepalive settings"""
    return psycopg2.connect(
        DB_URL,
        connect_timeout=10,
        keepalives=1,
        keepalives_idle=30,
        keepalives_interval=10,
        keepalives_count=5
    )

def normalize_colname(name: str) -> str:
    return re.sub(r'[^a-z0-9]', '', str(name).lower())


def map_columns(df_columns):
    mapping = {}
    for col in df_columns:
        mapping[normalize_colname(col)] = col
    return mapping


def find_required_columns(df_cols, required_list):
    mapping = map_columns(df_cols)
    found = {}
    missing = []
    for req in required_list:
        norm_req = normalize_colname(req)
        if norm_req in mapping:
            found[req] = mapping[norm_req]
        else:
            parts = re.split(r'[^a-z0-9]+', req.lower())
            matched = None
            for dfcol in df_cols:
                n = normalize_colname(dfcol)
                if all(p for p in parts if p and p in n):
                    matched = dfcol
                    break
            if matched:
                found[req] = matched
            else:
                missing.append(req)
    return found, missing


def import_with_retries(records,
                        table_name,
                        on_conflict,
                        initial_chunk_size=1000,
                        max_retries=3,
                        backoff_seconds=2):
    """Insert all data in ONE transaction for maximum speed"""
    total = len(records)
    if total == 0:
        return {"inserted": 0, "errors": []}

    conn = None
    
    for attempt in range(max_retries):
        try:
            # Get column names from first record
            columns = list(records[0].keys())
            
            # Prepare ALL values as list of tuples - NO CHUNKING
            values = [tuple(row[col] for col in columns) for row in records]
            
            # Connect to database
            st.info(f"Connecting to database... (Attempt {attempt + 1}/{max_retries})")
            conn = get_db_connection()
            cur = conn.cursor()
            
            # Build INSERT query with ON CONFLICT
            cols_str = ', '.join([f'"{col}"' for col in columns])
            insert_query = f"""
                INSERT INTO "{table_name}" ({cols_str})
                VALUES %s
                ON CONFLICT ("{on_conflict}") DO UPDATE SET
                "Inserted_date" = EXCLUDED."Inserted_date"
            """
            
            # Execute SINGLE batch insert for ALL records
            st.info(f"Inserting {len(values)} records in ONE transaction...")
            start_time = time.time()
            
            # Use larger page_size for better performance
            extras.execute_values(cur, insert_query, values, page_size=5000)
            
            conn.commit()
            elapsed = time.time() - start_time
            
            st.success(f"✅ Inserted {len(values)} records in {elapsed:.2f} seconds!")
            
            cur.close()
            conn.close()
            
            return {"inserted": len(values), "errors": []}
            
        except (OperationalError, InterfaceError) as e:
            msg = str(e)
            if conn:
                try:
                    conn.close()
                except:
                    pass
            
            if attempt < max_retries - 1:
                wait_time = backoff_seconds * (2 ** attempt)
                st.warning(f"Connection error: {msg}")
                st.info(f"Retrying in {wait_time} seconds... (Attempt {attempt + 1}/{max_retries})")
                time.sleep(wait_time)
            else:
                st.error(f"❌ Failed after {max_retries} attempts: {msg}")
                return {"inserted": 0, "errors": [{"error": msg}]}
                
        except Exception as e:
            msg = str(e)
            if conn:
                try:
                    if not conn.closed:
                        conn.rollback()
                    conn.close()
                except:
                    pass
            
            st.error(f"❌ Error during import: {msg}")
            return {"inserted": 0, "errors": [{"error": msg}]}
    
    return {"inserted": 0, "errors": [{"error": "Max retries exceeded"}]}


# ============================================================================
# CHECK FUNCTIONS
# ============================================================================

def check_id_in_db(id_value: str, table_name: str, search_column: str) -> dict:
    """Check if ID exists in database using direct SQL query"""
    conn = None
    try:
        conn = get_db_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        
        query = f'SELECT * FROM "{table_name}" WHERE "{search_column}" = %s LIMIT 1'
        cur.execute(query, (id_value.strip(),))
        
        result = cur.fetchone()
        
        cur.close()
        conn.close()
        
        if result:
            return {
                "exists": True,
                "record": dict(result),
                "error": None
            }
        else:
            return {
                "exists": False,
                "record": None,
                "error": None
            }
    except Exception as e:
        if conn:
            try:
                conn.close()
            except:
                pass
        return {
            "exists": False,
            "record": None,
            "error": str(e)
        }


def check_ids_batch(ids_list: list, table_name: str, search_column: str) -> pd.DataFrame:
    """Check multiple IDs in batch"""
    results = []
    
    # Use batch query for better performance
    conn = None
    try:
        conn = get_db_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        
        # Query all IDs at once
        placeholders = ','.join(['%s'] * len(ids_list))
        query = f'SELECT "{search_column}" FROM "{table_name}" WHERE "{search_column}" IN ({placeholders})'
        cur.execute(query, ids_list)
        
        found_ids = set(row[search_column] for row in cur.fetchall())
        
        cur.close()
        conn.close()
        
        # Build results
        for id_val in ids_list:
            exists = id_val in found_ids
            results.append({
                "ID": id_val,
                "Exists": "✅ Yes" if exists else "❌ No",
                "Status": "Found" if exists else "Not Found"
            })
            
    except Exception as e:
        if conn:
            try:
                conn.close()
            except:
                pass
        # Fallback to individual checks
        st.warning(f"Batch check failed: {e}. Falling back to individual checks...")
        for id_val in ids_list:
            check_result = check_id_in_db(id_val, table_name, search_column)
            results.append({
                "ID": id_val,
                "Exists": "✅ Yes" if check_result["exists"] else "❌ No",
                "Status": check_result["error"] if check_result["error"] else "Found" if check_result["exists"] else "Not Found"
            })
    
    return pd.DataFrame(results)


st.title("Total Database Summary")

try:
    conn = get_db_connection()
    cur = conn.cursor()

    # Total UPI IDs and latest date
    cur.execute('SELECT COUNT(*), MAX("Inserted_date") FROM "all_upiiD"')
    upi_result = cur.fetchone()
    upi_count = upi_result[0]
    upi_latest_date = upi_result[1]

    # Total Bank Accounts and latest date
    cur.execute('SELECT COUNT(*), MAX("Inserted_date") FROM "all_bank_acc"')
    bank_result = cur.fetchone()
    bank_count = bank_result[0]
    bank_latest_date = bank_result[1]

    cur.close()
    conn.close()

    # ===== DISPLAY =====
    col1, col2 = st.columns(2)

    with col1:
        latest_date_str = f" (Latest: {upi_latest_date.strftime('%Y-%m-%d')})" if upi_latest_date else ""
        st.metric("Total UPI IDs", f"{upi_count:,}{latest_date_str}")

    with col2:
        latest_date_str = f" (Latest: {bank_latest_date.strftime('%Y-%m-%d')})" if bank_latest_date else ""
        st.metric("Total Bank Accounts", f"{bank_count:,}{latest_date_str}")

except Exception as e:
    st.error("Failed to fetch data from Nhost DB")
    st.exception(e)

# ============================================================================
# MAIN UI: TWO-COLUMN LAYOUT
# ============================================================================

col1, col2 = st.columns(2, gap="large")

# ============================================================================
# COLUMN 1: IMPORT FUNCTIONALITY
# ============================================================================
with col1:
    st.header("📥 Import IDs")
    st.markdown("**Upload CSV file and import data into database**")
    
    target_label = st.selectbox("Select target to import", list(TABLE_OPTIONS.keys()), key="import_target")
    target_cfg = TABLE_OPTIONS[target_label]
    TABLE_NAME = target_cfg["table_name"]
    REQUIRED_COLS = target_cfg["required"]
    CONFLICT_COL = target_cfg["conflict_col"]
    FILTER_VALUE = target_cfg["filter_value"]
    
    st.markdown(f"**Target table:** `{TABLE_NAME}`")
    
    uploaded_file = st.file_uploader("Upload CSV file", type=["csv", "txt"], key="import_file")
    
    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file, dtype=str)
            st.info(f"📄 Loaded {len(df)} rows with {len(df.columns)} columns")
        except Exception as e:
            st.error(f"Could not read CSV: {e}")
            st.stop()
        
        st.write(f"**Applying filters for {target_label}...**")
        
        mask = (
            (df["Feature_type"].astype(str).str.strip() == "BS Money Laundering") &
            (df["Upi_bank_account_wallet"].astype(str).str.strip().isin([FILTER_VALUE])) &
            (df["Search_for"].astype(str).str.strip().isin(["App", "Web"]))
        )

        filtered_df = df[mask].copy()
        st.write(f"✅ After filtering: {len(filtered_df)} rows (removed {len(df) - len(filtered_df)} rows)")
        
        if len(filtered_df) == 0:
            st.warning("⚠️ No rows matched the filter criteria. Please check your CSV data.")
            st.stop()
        
        # Find required columns
        found_cols_map, missing = find_required_columns(filtered_df.columns.tolist(), REQUIRED_COLS)
        
        if missing:
            st.error(f"❌ CSV must contain columns: {', '.join(missing)}")
            st.error(f"Available columns: {', '.join(filtered_df.columns.tolist())}")
            st.stop()
        
        # Select and rename columns
        actual_cols = [found_cols_map[c] for c in REQUIRED_COLS]
        df_clean = filtered_df[actual_cols].copy()
        rename_map = {found_cols_map[c]: c for c in REQUIRED_COLS}
        df_clean = df_clean.rename(columns=rename_map)
        
        # Validate key column
        key_col = CONFLICT_COL
        if key_col not in df_clean.columns:
            st.error(f"Internal error: expected key column '{key_col}' not found.")
            st.stop()
        
        # ---------- CLEANING ----------
        df_clean = df_clean.dropna(subset=[key_col])
        df_clean[key_col] = df_clean[key_col].astype(str).str.strip()
        
        if key_col.lower() == "upi_vpa":
            df_clean[key_col] = df_clean[key_col].str.lower()
        
        df_clean = df_clean[df_clean[key_col] != ""]
        df_clean = df_clean.drop_duplicates(subset=[key_col])
        
        st.write(f"🧹 After cleaning: {len(df_clean)} unique records")
        
        # --------------------------
        # INSERTED_DATE NORMALIZE & TODAY CHECK
        # --------------------------
        if "Inserted_date" in df_clean.columns:
            parsed_dates = pd.to_datetime(df_clean["Inserted_date"], errors="coerce")
            
            nat_mask = parsed_dates.isna()
            nat_count = int(nat_mask.sum())
            
            if nat_count > 0:
                st.warning(f"⚠️ {nat_count} row(s) have unparseable dates and will be removed.")
                df_clean = df_clean[~nat_mask].copy()
                parsed_dates = parsed_dates[~nat_mask]
            
            # Python date object — matches PostgreSQL date type
            df_clean["Inserted_date"] = parsed_dates.dt.date

            today = datetime.date.today()
            today_mask = df_clean["Inserted_date"] == today
            today_count = int(today_mask.sum())

            if today_count > 0:
                st.warning(f"⚠️ {today_count} row(s) contain today's date ({today.isoformat()}).")
                st.write("If you don't want rows with today's date to be imported, select the checkbox below to remove them and continue.")
                remove_today = st.checkbox("Remove rows with today's date and continue import", value=False, key="remove_today")

                if remove_today:
                    df_clean = df_clean[~today_mask].copy()
                    st.info(f"✅ Removed {today_count} row(s) with today's date. Remaining rows: {len(df_clean)}")
                else:
                    st.error("⛔ Import stopped because file contains rows with today's date. Select the checkbox to remove them and continue.")
                    st.stop()
        # --------------------------
        # END OF DATE CHECK
        # --------------------------

        total_rows = len(df_clean)

        if total_rows == 0:
            st.warning("⚠️ No records to import after all filters and cleaning.")
            st.stop()
        
        st.success(f"✅ Prepared {total_rows:,} unique records to import into `{TABLE_NAME}`")
        
        with st.expander("📋 Preview data (first 10 rows)", expanded=False):
            st.dataframe(df_clean.head(10), use_container_width=True)
        
        CHUNK_SIZE = 10_000
        st.info(f"💡 Import will execute in chunks of {CHUNK_SIZE:,} rows with live progress")
        
        btn = st.button("🚀 Start Import", use_container_width=True, type="primary")
        
        if btn:
            start_time = time.time()

            inserted_total = 0
            skipped_total  = 0

            # --- Live UI elements ---
            progress_bar    = st.progress(0)
            status_text     = st.empty()
            metrics_cols    = st.columns(3)
            inserted_metric = metrics_cols[0].empty()
            skipped_metric  = metrics_cols[1].empty()
            elapsed_metric  = metrics_cols[2].empty()

            try:
                for i in range(0, total_rows, CHUNK_SIZE):
                    chunk = df_clean.iloc[i:i + CHUNK_SIZE]

                    # Build arrays for unnest insert
                    dates = chunk["Inserted_date"].tolist() if "Inserted_date" in chunk.columns else [None] * len(chunk)
                    keys  = chunk[key_col].tolist()

                    with engine.begin() as conn:
                        result = conn.execute(
                            text(f"""
                                INSERT INTO "{TABLE_NAME}" ("Inserted_date", "{key_col}")
                                SELECT d, u
                                FROM unnest(
                                    CAST(:dates AS date[]),
                                    CAST(:keys  AS text[])
                                ) AS t(d, u)
                                ON CONFLICT DO NOTHING
                            """),
                            {
                                "dates": dates,
                                "keys":  keys
                            }
                        )

                        inserted = result.rowcount
                        skipped  = len(chunk) - inserted
                        inserted_total += inserted
                        skipped_total  += skipped

                    # --- Update progress ---
                    processed = min(i + CHUNK_SIZE, total_rows)
                    pct       = processed / total_rows
                    elapsed   = time.time() - start_time

                    progress_bar.progress(pct)
                    status_text.write(
                        f"⏳ Processing chunk {i // CHUNK_SIZE + 1} — "
                        f"{processed:,} / {total_rows:,} rows ({pct*100:.1f}%)"
                    )
                    inserted_metric.metric("✅ Inserted", f"{inserted_total:,}")
                    skipped_metric.metric("⏭️ Skipped (duplicates)", f"{skipped_total:,}")
                    elapsed_metric.metric("⏱️ Elapsed", f"{elapsed:.1f}s")

                # --- Final summary ---
                elapsed = time.time() - start_time
                progress_bar.progress(1.0)
                status_text.success(f"🎉 Import complete in {elapsed:.1f}s!")

                st.divider()
                col_a, col_b, col_c = st.columns(3)
                col_a.metric("✅ Total Inserted", f"{inserted_total:,}")
                col_b.metric("⏭️ Total Skipped", f"{skipped_total:,}")
                col_c.metric("⏱️ Total Time", f"{elapsed:.1f}s")

            except Exception as e:
                st.error(f"❌ Import failed: {e}")

# ============================================================================
# COLUMN 2: CHECK FUNCTIONALITY
# ============================================================================
with col2:
    st.header("🔍 Check IDs")
    st.markdown("**Search for IDs in database**")
    
    check_target = st.selectbox("Select target to check", list(TABLE_OPTIONS.keys()), key="check_target")
    check_cfg = TABLE_OPTIONS[check_target]
    CHECK_TABLE = check_cfg["table_name"]
    SEARCH_COLUMN = check_cfg["conflict_col"]
    
    st.markdown(f"**Searching in:** `{CHECK_TABLE}`")
    
    check_method = st.radio("Check method", ["Single/Multiple IDs", "Batch Upload"], horizontal=True)
    
    if check_method == "Single/Multiple IDs":
        st.markdown("**Enter IDs (one per line or comma-separated)**")
        id_input = st.text_area("Enter IDs to search", 
                                placeholder="user1@upi\nuser2@upi\nuser3@upi\n\nOr: user1@upi, user2@upi, user3@upi",
                                height=100)
        
        if id_input:
            # Parse IDs - handle both newline and comma separated
            if ',' in id_input:
                ids_list = [id.strip() for id in id_input.split(',') if id.strip()]
            else:
                ids_list = [id.strip() for id in id_input.split('\n') if id.strip()]
            
            st.info(f"📊 Found {len(ids_list)} ID(s) to check")
            
            if st.button("🔎 Search All", use_container_width=True, type="primary"):
                with st.spinner("Searching..."):
                    results_df = check_ids_batch(ids_list, CHECK_TABLE, SEARCH_COLUMN)
                    
                    # Show summary
                    col_exists, col_not_exists = st.columns(2)
                    with col_exists:
                        exists_count = (results_df["Exists"] == "✅ Yes").sum()
                        st.metric("Found", exists_count, f"{(exists_count/len(results_df)*100):.1f}%")
                    with col_not_exists:
                        not_exists_count = (results_df["Exists"] == "❌ No").sum()
                        st.metric("Not Found", not_exists_count, f"{(not_exists_count/len(results_df)*100):.1f}%")
                    
                    # Show results table
                    st.dataframe(results_df, use_container_width=True, height=300)
                    
                    # Download option
                    csv = results_df.to_csv(index=False)
                    st.download_button(
                        label="📥 Download Results",
                        data=csv,
                        file_name=f"check_results_{check_target.lower().replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
    
    else:
        batch_file = st.file_uploader("Upload CSV with IDs to check", type=["csv", "txt"], key="check_file")
        
        if batch_file:
            try:
                batch_df = pd.read_csv(batch_file, dtype=str)
                st.info(f"📄 Loaded {len(batch_df)} rows")
            except Exception as e:
                st.error(f"Could not read CSV: {e}")
                st.stop()
            
            found_cols_map, _ = find_required_columns(batch_df.columns.tolist(), [SEARCH_COLUMN])
            
            if SEARCH_COLUMN not in found_cols_map:
                id_column = batch_df.columns[0]
                st.warning(f"⚠️ Column '{SEARCH_COLUMN}' not found. Using first column '{id_column}' as ID column")
            else:
                id_column = found_cols_map[SEARCH_COLUMN]
                st.info(f"✅ Using column '{id_column}' for IDs")
            
            batch_ids = batch_df[id_column].astype(str).str.strip().tolist()
            batch_ids = [id for id in batch_ids if id]  # Remove empty strings
            
            st.info(f"📊 Found {len(batch_ids)} IDs to check")
            
            if st.button("🔎 Check All", use_container_width=True, type="primary"):
                with st.spinner("Checking all IDs..."):
                    results_df = check_ids_batch(batch_ids, CHECK_TABLE, SEARCH_COLUMN)
                    
                    col_exists, col_not_exists = st.columns(2)
                    with col_exists:
                        exists_count = (results_df["Exists"] == "✅ Yes").sum()
                        st.metric("Found", exists_count, f"{(exists_count/len(results_df)*100):.1f}%")
                    with col_not_exists:
                        not_exists_count = (results_df["Exists"] == "❌ No").sum()
                        st.metric("Not Found", not_exists_count, f"{(not_exists_count/len(results_df)*100):.1f}%")
                    
                    st.dataframe(results_df, use_container_width=True, height=400)
                    
                    csv = results_df.to_csv(index=False)
                    st.download_button(
                        label="📥 Download Results",
                        data=csv,
                        file_name=f"check_results_{check_target.lower().replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )