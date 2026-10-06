"""
Laboratory Operations & Inventory Management System (PharmaLab)
Single-file complete application for simple GitHub & Streamlit deployment.
"""

import os
import sqlite3
import json
from datetime import datetime, date
import pandas as pd
import streamlit as st

# ==============================================================================
# 1. DATABASE & PERSISTENCE CONFIGURATION
# ==============================================================================
DATA_DIR = os.environ.get("PERSISTENT_DATA_DIR", os.path.dirname(os.path.abspath(__file__)))
os.makedirs(DATA_DIR, exist_ok=True)

DB_PATH = os.path.join(DATA_DIR, "lab_inventory.db")
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Total chemicals and lab ware list.xlsx")

def get_connection():
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    return conn

def init_db(force_reseed=False):
    """Initializes tables and seeds base records only if empty to prevent data loss on deployment."""
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS inventory (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_id TEXT UNIQUE NOT NULL,
        lab TEXT,
        inventory_group TEXT,
        fr_code TEXT,
        item_name TEXT NOT NULL,
        item_type TEXT,
        unit TEXT,
        expiry TEXT,
        stock_location TEXT,
        current_stock REAL DEFAULT 0.0,
        reorder_level REAL DEFAULT 0.0,
        status TEXT,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS timetable_schedule (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        lab TEXT NOT NULL,
        day_of_week TEXT NOT NULL,
        time_slot TEXT NOT NULL,
        course_name TEXT NOT NULL,
        room TEXT,
        instructor TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS experiment_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        log_id TEXT UNIQUE NOT NULL,
        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        lab TEXT NOT NULL,
        course_name TEXT,
        instructor TEXT,
        experiment_name TEXT NOT NULL,
        apparatus_used TEXT,
        chemicals_used_json TEXT NOT NULL,
        logged_by TEXT,
        observations TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS stock_transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        transaction_type TEXT NOT NULL,
        item_id TEXT NOT NULL,
        item_name TEXT NOT NULL,
        quantity_change REAL NOT NULL,
        balance_after REAL NOT NULL,
        unit TEXT,
        experiment_log_id TEXT,
        reference_reason TEXT,
        user_name TEXT
    )
    """)

    cur.execute("CREATE INDEX IF NOT EXISTS idx_inv_item_id ON inventory(item_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_inv_name ON inventory(item_name);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_sched_lab ON timetable_schedule(lab, day_of_week);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_exp_timestamp ON experiment_logs(timestamp);")

    conn.commit()

    # Seed only if database is completely empty (safeguards existing records across redeployments)
    cur.execute("SELECT COUNT(*) FROM inventory")
    if cur.fetchone()[0] == 0 or force_reseed:
        seed_inventory_from_excel(conn)

    cur.execute("SELECT COUNT(*) FROM timetable_schedule")
    if cur.fetchone()[0] == 0 or force_reseed:
        seed_timetable(conn)

    conn.close()

def seed_inventory_from_excel(conn):
    if not os.path.exists(EXCEL_PATH):
        return

    try:
        df = pd.read_excel(EXCEL_PATH)
        cur = conn.cursor()
        cur.execute("DELETE FROM inventory")

        records = []
        for _, row in df.iterrows():
            item_id = str(row.get('Item ID', '')).strip()
            if not item_id or item_id.lower() == 'nan':
                continue

            lab = str(row.get('Lab', '')).strip() if pd.notna(row.get('Lab')) else ''
            group = str(row.get('Inventory Group', '')).strip() if pd.notna(row.get('Inventory Group')) else ''
            fr_code = str(row.get('FR Code', '')).strip() if pd.notna(row.get('FR Code')) else ''
            name = str(row.get('Item Name', '')).strip() if pd.notna(row.get('Item Name')) else ''
            itype = str(row.get('Type', '')).strip() if pd.notna(row.get('Type')) else ''
            unit = str(row.get('Unit', '')).strip() if pd.notna(row.get('Unit')) else ''
            exp = str(row.get('Expiry', '')).strip() if pd.notna(row.get('Expiry')) else ''
            loc = str(row.get('Stock Location', '')).strip() if pd.notna(row.get('Stock Location')) else ''
            
            try:
                stock = float(row.get('Current Stock', 0.0)) if pd.notna(row.get('Current Stock')) else 0.0
            except Exception:
                stock = 0.0

            try:
                reorder = float(row.get('Reorder Level', 0.0)) if pd.notna(row.get('Reorder Level')) else 0.0
            except Exception:
                reorder = 0.0

            status = str(row.get('Status (as per fresh audit)', 'Available')).strip() if pd.notna(row.get('Status (as per fresh audit)')) else 'Available'
            records.append((item_id, lab, group, fr_code, name, itype, unit, exp, loc, stock, reorder, status))

        cur.executemany("""
        INSERT OR REPLACE INTO inventory 
        (item_id, lab, inventory_group, fr_code, item_name, item_type, unit, expiry, stock_location, current_stock, reorder_level, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, records)
        conn.commit()
    except Exception:
        conn.rollback()

def seed_timetable(conn):
    cur = conn.cursor()
    cur.execute("DELETE FROM timetable_schedule")

    schedule_data = [
        # Lab 135-B
        ("135-B", "Monday", "08:00 - 11:00", "Pharmaceutics 3A", "Room 211C", "Ms. Saman A"),
        ("135-B", "Monday", "11:00 - 14:00", "Practice 2", "Room 315A", "Ms. Muryam AR"),
        ("135-B", "Monday", "14:00 - 17:00", "Pharmaceutics 1A", "Room 121C", "Ms. Muryam"),
        ("135-B", "Monday", "08:00 - 08:50 (Pre-lab)", "Pharmaceutics 2A", "Room 210A", "Dr. Abdul Haleem Khan"),
        ("135-B", "Tuesday", "08:00 - 11:00", "Practice", "Room 319A", "Mr. Omaid K"),
        ("135-B", "Tuesday", "14:00 - 17:00", "Pharmaceutics 3A", "Room 211A", "Ms. Saman A"),
        ("135-B", "Wednesday", "08:00 - 11:00", "Pharmaceutics 3A", "Room 211C", "Ms. Saman A"),
        ("135-B", "Wednesday", "11:00 - 14:00", "Practice 2", "Room 315A", "Ms. Muryam AR"),
        ("135-B", "Wednesday", "14:00 - 17:00", "Pharmaceutics 2A", "Room 210B", "Dr. Abdul Haleem Khan"),
        ("135-B", "Thursday", "11:00 - 14:00", "Practice 2A", "Room 310A", "Mr. Sufyan J"),
        ("135-B", "Thursday", "14:00 - 17:00", "Pharmaceutics 1A", "Room 121B", "Ms. Muryam"),
        ("135-B", "Friday", "08:00 - 11:00", "Pharmaceutics 3A", "Room 211C", "Ms. Saman A"),
        ("135-B", "Friday", "11:00 - 14:00", "Practice 2", "Room 315A", "Ms. Muryam AR"),
        ("135-B", "Friday", "14:00 - 17:00", "Practice 5B", "Room 356B", "Mr. Omaid K"),
        ("135-B", "Friday", "08:00 - 08:50 (Pre-lab)", "Pharmaceutics 2A", "Room 210C", "Dr. Abdul Haleem Khan"),

        # Lab 138-B
        ("138-B", "Monday", "08:00 - 11:00", "Pharmacognosy", "Room 213A", "Mr. Sabi UR"),
        ("138-B", "Monday", "11:00 - 14:00", "Pharmacognosy 2A", "Room 313A", "Mr. Sohaib P"),
        ("138-B", "Monday", "14:00 - 17:00", "Pharm. Chem 3B", "Room 316A", "Mr. Nasir A"),
        ("138-B", "Tuesday", "08:00 - 11:00", "Pharmaceutics 3A", "Room 211C", "Ms. Saman A"),
        ("138-B", "Tuesday", "11:00 - 12:15", "Pharmaceutics 5B", "Room 358A", "Dr. Omer Salman Q"),
        ("138-B", "Tuesday", "14:00 - 17:00", "Pharmacognosy", "Room 213B", "Mr. Sohaib P"),
        ("138-B", "Wednesday", "08:00 - 11:00", "Pharmacognosy", "Room 213A", "Mr. Sabi UR"),
        ("138-B", "Wednesday", "11:00 - 14:00", "Pharmacognosy 2B", "Room 318A", "Mr. Sabi R"),
        ("138-B", "Wednesday", "14:00 - 17:00", "Pharmacognosy", "Room 213C", "Mr. Sohaib P"),
        ("138-B", "Thursday", "08:00 - 11:00", "Pharmaceutics 3A", "Room 211C", "Ms. Saman A"),
        ("138-B", "Thursday", "11:00 - 12:15", "Pharmaceutics 5B", "Room 358A", "Dr. Omer Salman Q"),
        ("138-B", "Thursday", "14:00 - 17:00", "Pharmaceutics 3A", "Room 211C", "Ms. Saman A"),
        ("138-B", "Friday", "08:00 - 11:00", "Pharmacognosy", "Room 213A", "Mr. Sabi UR"),
        ("138-B", "Friday", "11:00 - 14:00", "Pharmacognosy 2B", "Room 318B", "Mr. Sabi R"),
        ("138-B", "Friday", "14:00 - 17:00", "Pharmacognosy", "Room 213A", "Mr. Sabi UR")
    ]

    cur.executemany("""
    INSERT INTO timetable_schedule (lab, day_of_week, time_slot, course_name, room, instructor)
    VALUES (?, ?, ?, ?, ?, ?)
    """, schedule_data)
    conn.commit()

def get_timetable(lab_filter="All"):
    conn = get_connection()
    query = "SELECT * FROM timetable_schedule WHERE 1=1"
    params = []
    if lab_filter and lab_filter != "All":
        query += " AND lab = ?"
        params.append(lab_filter)

    query += " ORDER BY CASE day_of_week WHEN 'Monday' THEN 1 WHEN 'Tuesday' THEN 2 WHEN 'Wednesday' THEN 3 WHEN 'Thursday' THEN 4 WHEN 'Friday' THEN 5 ELSE 6 END, time_slot"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def get_inventory_items(search_query="", lab_filter="All", group_filter="All", status_filter="All", limit=2000):
    conn = get_connection()
    query = "SELECT * FROM inventory WHERE 1=1"
    params = []

    if lab_filter and lab_filter != "All":
        query += " AND lab = ?"
        params.append(lab_filter)

    if group_filter and group_filter != "All":
        query += " AND inventory_group = ?"
        params.append(group_filter)

    if status_filter == "Low Stock":
        query += " AND current_stock <= reorder_level AND current_stock > 0"
    elif status_filter == "Out of Stock":
        query += " AND current_stock <= 0"
    elif status_filter == "Available":
        query += " AND current_stock > 0"

    if search_query:
        query += " AND (item_name LIKE ? OR item_id LIKE ? OR fr_code LIKE ?)"
        q = f"%{search_query.strip()}%"
        params.extend([q, q, q])

    query += " ORDER BY item_name ASC LIMIT ?"
    params.append(limit)
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def deduct_dispense(experiment_info, items_used, user_name="Lab Officer"):
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("BEGIN TRANSACTION")
        log_id = f"EXP-{datetime.now().strftime('%Y%m%d-%H%M%S')}"

        for item in items_used:
            item_id = item['item_id']
            qty = float(item['quantity'])

            cur.execute("SELECT current_stock, item_name, unit FROM inventory WHERE item_id = ?", (item_id,))
            row = cur.fetchone()
            if not row:
                raise ValueError(f"Item ID {item_id} not found.")

            curr_stock = float(row['current_stock'] or 0.0)
            new_stock = max(0.0, curr_stock - qty)

            cur.execute("UPDATE inventory SET current_stock = ?, updated_at = CURRENT_TIMESTAMP WHERE item_id = ?", (new_stock, item_id))
            cur.execute("""
            INSERT INTO stock_transactions 
            (transaction_type, item_id, item_name, quantity_change, balance_after, unit, experiment_log_id, reference_reason, user_name)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, ("DISPENSE / USAGE", item_id, row['item_name'], -qty, new_stock, row['unit'], log_id, experiment_info.get('experiment_name', ''), user_name))

        cur.execute("""
        INSERT INTO experiment_logs 
        (log_id, lab, course_name, instructor, experiment_name, apparatus_used, chemicals_used_json, logged_by, observations)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            log_id,
            experiment_info.get('lab', '135-B'),
            experiment_info.get('course_name', ''),
            experiment_info.get('instructor', ''),
            experiment_info.get('experiment_name', ''),
            experiment_info.get('apparatus_used', ''),
            json.dumps(items_used),
            user_name,
            experiment_info.get('observations', '')
        ))

        conn.commit()
        conn.close()
        return True, log_id
    except Exception as e:
        conn.rollback()
        conn.close()
        return False, str(e)

def restock_item(item_id, qty, reason="Stock Inward", user_name="Lab Incharge"):
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("SELECT current_stock, item_name, unit FROM inventory WHERE item_id = ?", (item_id,))
        row = cur.fetchone()
        if not row:
            return False, f"Item {item_id} not found."

        new_balance = float(row['current_stock'] or 0.0) + float(qty)
        cur.execute("UPDATE inventory SET current_stock = ?, updated_at = CURRENT_TIMESTAMP WHERE item_id = ?", (new_balance, item_id))
        cur.execute("""
        INSERT INTO stock_transactions 
        (transaction_type, item_id, item_name, quantity_change, balance_after, unit, reference_reason, user_name)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, ("RESTOCK", item_id, row['item_name'], qty, new_balance, row['unit'], reason, user_name))
        conn.commit()
        conn.close()
        return True, f"Balance updated to {new_balance} {row['unit']}."
    except Exception as e:
        conn.close()
        return False, str(e)

def get_experiment_logs(lab_filter="All", search_query=""):
    conn = get_connection()
    query = "SELECT * FROM experiment_logs WHERE 1=1"
    params = []
    if lab_filter and lab_filter != "All":
        query += " AND lab = ?"
        params.append(lab_filter)
    if search_query:
        query += " AND (experiment_name LIKE ? OR course_name LIKE ? OR log_id LIKE ?)"
        q = f"%{search_query.strip()}%"
        params.extend([q, q, q])

    query += " ORDER BY timestamp DESC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def get_transactions(limit=300):
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM stock_transactions ORDER BY timestamp DESC LIMIT ?", conn, params=(limit,))
    conn.close()
    return df

# ==============================================================================
# 2. STREAMLIT FRONTEND
# ==============================================================================
st.set_page_config(
    page_title="Lab Inventory & Operations System",
    page_icon="🧪",
    layout="wide"
)

if "db_ready" not in st.session_state:
    init_db()
    st.session_state["db_ready"] = True

if "cart_items" not in st.session_state:
    st.session_state["cart_items"] = []

# --- Header & Metrics ---
st.title("🧪 Laboratory Operations & Inventory System")

inv_all = get_inventory_items(limit=3000)
total_items = len(inv_all)
chems_count = len(inv_all[inv_all['inventory_group'] == 'Chemical'])
hw_count = len(inv_all[inv_all['inventory_group'] == 'Hardware / Labware'])
low_count = len(inv_all[(inv_all['current_stock'] <= inv_all['reorder_level']) & (inv_all['current_stock'] > 0)])

m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Items", f"{total_items:,}")
m2.metric("Chemicals & Reagents", f"{chems_count:,}")
m3.metric("Glassware & Hardware", f"{hw_count:,}")
m4.metric("Low Stock Alerts", f"{low_count:,}")

st.write("---")

# --- Sidebar: Cloud Persistence & Downloads ---
with st.sidebar:
    st.markdown("### 💾 Storage & Offline Backups")
    st.caption("SQLite WAL Persistent Engine")
    
    if os.path.exists(DB_PATH):
        with open(DB_PATH, "rb") as f:
            st.download_button(
                label="⬇️ Download Database (.db)",
                data=f,
                file_name=f"lab_inventory_{date.today().strftime('%Y%m%d')}.db",
                mime="application/x-sqlite3",
                use_container_width=True
            )

    st.markdown("---")
    st.markdown("### 🛠️ Maintenance")
    if st.button("🔄 Reload from Excel Master", help="Reset and reload default stock from Excel"):
        init_db(force_reseed=True)
        st.success("Re-seeded successfully!")
        st.rerun()

# --- Main Tabs ---
tab_sched, tab_dispense, tab_history, tab_inv, tab_audit = st.tabs([
    "📅 Lab Timetable",
    "🧪 Experiment Logging & Dispense Cart",
    "📜 Previous Logged Experiments",
    "📦 Inventory & Restock",
    "📊 Transaction Ledger"
])

# ------------------------------------------------------------------------------
# TAB 1: SIMPLE TIMETABLE
# ------------------------------------------------------------------------------
with tab_sched:
    st.subheader("Department Lab Timetable")
    col_t1, col_t2 = st.columns([1, 3])
    with col_t1:
        selected_lab = st.selectbox("Select Laboratory", ["All", "135-B", "138-B"], index=0)

    sched_df = get_timetable(lab_filter=selected_lab)

    if sched_df.empty:
        st.info("No schedule entries found.")
    else:
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        for day in days:
            day_slots = sched_df[sched_df['day_of_week'] == day]
            if not day_slots.empty:
                st.markdown(f"#### 🗓️ {day}")
                display_table = day_slots[['lab', 'time_slot', 'course_name', 'instructor', 'room']]
                display_table.columns = ['Lab Room', 'Time Slot', 'Course', 'Instructor', 'Assigned Room']
                st.dataframe(display_table, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# TAB 2: EXPERIMENT LOGGING & DISPENSE (POS CART)
# ------------------------------------------------------------------------------
with tab_dispense:
    st.subheader("Log Experiment & Deduct Stock")

    col_form, col_cart = st.columns([1, 1])

    with col_form:
        st.markdown("##### 1. Experiment Details")
        exp_name = st.text_input("Experiment Title *", placeholder="e.g. Synthesis of Aspirin")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            exp_lab = st.selectbox("Lab Room", ["135-B", "138-B", "M001", "M003"])
            exp_course = st.text_input("Course Name", value="Pharmaceutics")
        with col_c2:
            exp_inst = st.text_input("Instructor", value="Faculty")
            exp_user = st.text_input("Logged By", value="Lab Incharge")
        
        exp_app = st.text_area("Apparatus Used", placeholder="e.g. Mortar & pestle, Measuring cylinder 100ml", height=70)
        exp_obs = st.text_area("Observations / Results", placeholder="Practical notes or yield calculation...", height=70)

    with col_cart:
        st.markdown("##### 2. Requisition Slip / Deduct Chemicals")
        search_kw = st.text_input("🔍 Search Chemical or Apparatus", placeholder="Type name or code (e.g., Paraffin, HCl, Ethanol)")

        matching = get_inventory_items(search_query=search_kw, lab_filter=exp_lab, limit=25)
        if not matching.empty:
            item_map = {row['item_id']: f"[{row['item_id']}] {row['item_name']} (In Stock: {row['current_stock']} {row['unit']})" for _, row in matching.iterrows()}
            chosen_id = st.selectbox("Select Item", options=list(item_map.keys()), format_func=lambda x: item_map[x])
            chosen_row = matching[matching['item_id'] == chosen_id].iloc[0]

            avail_stock = float(chosen_row['current_stock'] or 0.0)
            u = chosen_row['unit']

            col_q, col_b = st.columns([2, 1])
            with col_q:
                qty_to_use = st.number_input(f"Quantity to Deduct ({u}):", min_value=0.1, value=10.0 if avail_stock >= 10 else (avail_stock if avail_stock > 0 else 1.0), step=1.0)
            with col_b:
                st.write("<div style='height:28px;'></div>", unsafe_allow_html=True)
                if st.button("➕ Add to Slip", use_container_width=True):
                    st.session_state["cart_items"].append({
                        "item_id": chosen_id,
                        "item_name": chosen_row['item_name'],
                        "quantity": qty_to_use,
                        "unit": u,
                        "available_before": avail_stock,
                        "balance_after": max(0.0, avail_stock - qty_to_use)
                    })
                    st.success(f"Added {qty_to_use} {u} of {chosen_row['item_name']}")
                    st.rerun()
        else:
            st.info("No matching inventory found. Clear the search box to browse.")

    st.write("---")
    st.markdown("##### Active Dispense Slip")
    if st.session_state["cart_items"]:
        cart_df = pd.DataFrame(st.session_state["cart_items"])
        st.dataframe(cart_df[['item_id', 'item_name', 'quantity', 'unit', 'available_before', 'balance_after']], use_container_width=True, hide_index=True)

        col_act1, col_act2 = st.columns([1, 2])
        with col_act1:
            if st.button("🗑️ Clear Slip", use_container_width=True):
                st.session_state["cart_items"] = []
                st.rerun()
        with col_act2:
            if st.button("✅ Confirm Practical & Deduct Stock in Database", type="primary", use_container_width=True):
                if not exp_name.strip():
                    st.error("Please provide an Experiment Title before proceeding.")
                else:
                    success, ref_id = deduct_dispense(
                        experiment_info={
                            "lab": exp_lab,
                            "course_name": exp_course,
                            "instructor": exp_inst,
                            "experiment_name": exp_name,
                            "apparatus_used": exp_app,
                            "observations": exp_obs
                        },
                        items_used=st.session_state["cart_items"],
                        user_name=exp_user
                    )
                    if success:
                        st.success(f"Transaction successful! Reference: `{ref_id}`. Stock quantities deducted.")
                        st.session_state["cart_items"] = []
                        st.rerun()
                    else:
                        st.error(f"Error logging experiment: {ref_id}")
    else:
        st.caption("Requisition slip is currently empty.")

# ------------------------------------------------------------------------------
# TAB 3: PREVIOUS LOGGED EXPERIMENTS (History Viewer)
# ------------------------------------------------------------------------------
with tab_history:
    st.subheader("Historical Logged Experiments")

    col_h1, col_h2 = st.columns([1, 2])
    with col_h1:
        hist_lab = st.selectbox("Filter by Lab", ["All", "135-B", "138-B"], key="hist_lab_filter")
    with col_h2:
        hist_q = st.text_input("Search Logs", placeholder="Search by experiment name, ID, or course...")

    logs = get_experiment_logs(lab_filter=hist_lab, search_query=hist_q)

    if logs.empty:
        st.info("No experiment logs recorded yet. Once an experiment is confirmed in Tab 2, it will appear here.")
    else:
        st.markdown(f"**Total Practical Logs: {len(logs)}**")
        for _, log in logs.iterrows():
            with st.expander(f"🔬 {log['log_id']} — {log['experiment_name']} ({log['lab']} | {log['timestamp'][:16]})"):
                c_info1, c_info2 = st.columns(2)
                with c_info1:
                    st.write(f"**Course:** {log['course_name']}")
                    st.write(f"**Instructor:** {log['instructor']}")
                    st.write(f"**Logged By:** {log['logged_by']}")
                with c_info2:
                    st.write(f"**Apparatus Used:** {log['apparatus_used'] or 'None recorded'}")
                    st.write(f"**Observations:** {log['observations'] or 'No observations'}")

                st.markdown("**Chemicals Deducted:**")
                try:
                    chems = json.loads(log['chemicals_used_json'])
                    disp_chems = pd.DataFrame(chems)[['item_id', 'item_name', 'quantity', 'unit']]
                    disp_chems.columns = ['Item ID', 'Description', 'Quantity Deducted', 'Unit']
                    st.table(disp_chems)
                except Exception:
                    st.text(log['chemicals_used_json'])

# ------------------------------------------------------------------------------
# TAB 4: INVENTORY & RESTOCK
# ------------------------------------------------------------------------------
with tab_inv:
    st.subheader("Central Inventory Stock Room")

    col_f1, col_f2, col_f3, col_f4 = st.columns([2, 1, 1, 1])
    with col_f1:
        inv_search = st.text_input("Search Inventory", placeholder="Search item name, ID, or FR code...")
    with col_f2:
        lab_f = st.selectbox("Room Filter", ["All", "135-B", "138-B", "M001", "M003", "M004", "M005"])
    with col_f3:
        grp_f = st.selectbox("Category", ["All", "Chemical", "Hardware / Labware"])
    with col_f4:
        st_f = st.selectbox("Status", ["All", "Available", "Low Stock", "Out of Stock"])

    filtered_stock = get_inventory_items(search_query=inv_search, lab_filter=lab_f, group_filter=grp_f, status_filter=st_f)

    display_stock = filtered_stock[['item_id', 'lab', 'inventory_group', 'item_name', 'current_stock', 'unit', 'reorder_level', 'stock_location', 'expiry']]
    display_stock.columns = ['Item ID', 'Lab', 'Group', 'Item Name', 'Current Stock', 'Unit', 'Reorder Lvl', 'Location', 'Expiry']
    st.dataframe(display_stock, use_container_width=True, hide_index=True)

    st.write("---")
    st.markdown("##### 📥 Restock / Record Inward Delivery")
    with st.form("restock_form"):
        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            restock_id = st.text_input("Item ID (e.g. CH-135B-0001)")
        with col_r2:
            restock_qty = st.number_input("Quantity Received", min_value=0.5, value=50.0, step=5.0)
        with col_r3:
            restock_reason = st.text_input("Invoice / Delivery Reference", value="Central Store Replenishment")

        if st.form_submit_button("Record Stock Delivery", type="primary"):
            if not restock_id.strip():
                st.error("Please enter an Item ID.")
            else:
                ok, msg = restock_item(restock_id.strip(), restock_qty, reason=restock_reason)
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

# ------------------------------------------------------------------------------
# TAB 5: TRANSACTION LEDGER
# ------------------------------------------------------------------------------
with tab_audit:
    st.subheader("Audit Trail & Balance Movements")
    tx_df = get_transactions()
    if tx_df.empty:
        st.info("No transactions logged yet.")
    else:
        disp_tx = tx_df[['timestamp', 'transaction_type', 'item_id', 'item_name', 'quantity_change', 'balance_after', 'unit', 'user_name', 'reference_reason']]
        disp_tx.columns = ['Timestamp', 'Type', 'Item ID', 'Item Name', 'Change', 'Balance After', 'Unit', 'User', 'Reference']
        st.dataframe(disp_tx, use_container_width=True, hide_index=True)