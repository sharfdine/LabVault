"""
Interactive Lab Automation & Inventory Management Dashboard
PharmD & Chemistry Labs (Lab 135-B & Lab 138-B)
Features:
1. Editable Lab Schedule & Pre-Planning (from Timetable PDF) with Semester & Week support
2. Experiment Logging & Pharmacy-Style Chemical Deduction / POS Dispensing
3. Inventory Management with Real-Time Stock Tracking & Excel Export
4. Categorized Notes, Correspondences & Follow-ups
5. Lab Productivity Tools (Molarity, Dilution & % Solution Calculators, Expiry/Reorder Alerts)
6. Transaction Audit Trail & Experiment History
7. Spacious Semester Storage & Archive Vault (WAL mode, Attachments, Automated Backups, Multi-Sheet Master Archive)
"""

import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date
import database
import cloud_sync

# --- Page Configuration ---
st.set_page_config(
    page_title="PharmaLab OS | Lab Automation & Inventory",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize Database on first load
if "db_initialized" not in st.session_state:
    database.init_db()
    st.session_state["db_initialized"] = True

# Initialize Session State for Experiment Cart / Dispense Requisition
if "cart_items" not in st.session_state:
    st.session_state["cart_items"] = []

if "prefill_experiment" not in st.session_state:
    st.session_state["prefill_experiment"] = None

# --- Light-Coloured Premium UI Styles ---
st.markdown("""
<style>
    /* Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #1e293b;
    }

    /* Overall Background */
    .stApp {
        background-color: #f8fafc;
    }

    /* Top Banner / Hero */
    .hero-banner {
        background: linear-gradient(135deg, #047857 0%, #0d9488 50%, #0284c7 100%);
        color: white;
        padding: 24px 28px;
        border-radius: 16px;
        margin-bottom: 24px;
        box-shadow: 0 10px 25px -5px rgba(13, 148, 136, 0.25);
    }
    
    .hero-title {
        font-size: 26px;
        font-weight: 700;
        letter-spacing: -0.02em;
        margin: 0;
        color: #ffffff;
    }
    
    .hero-subtitle {
        font-size: 14px;
        color: #e6fffa;
        margin-top: 6px;
        font-weight: 400;
    }

    /* Metric Cards */
    .metric-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 16px 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.03);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.06);
    }
    .metric-value {
        font-size: 26px;
        font-weight: 700;
        color: #0f766e;
        line-height: 1.2;
    }
    .metric-label {
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #64748b;
        margin-top: 4px;
    }

    /* Timetable Card */
    .timetable-slot-card {
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-left: 5px solid #0d9488;
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 12px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.03);
        transition: all 0.2s ease;
    }
    .timetable-slot-card:hover {
        border-left-color: #047857;
        box-shadow: 0 6px 12px rgba(13, 148, 136, 0.1);
    }

    /* Status Pills */
    .status-badge {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 9999px;
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 0.02em;
    }
    .status-scheduled { background-color: #e0f2fe; color: #0369a1; }
    .status-prepared { background-color: #fef3c7; color: #b45309; }
    .status-completed { background-color: #d1fae5; color: #047857; }
    .status-cancelled { background-color: #fee2e2; color: #b91c1c; }
    
    .priority-high { background-color: #fee2e2; color: #b91c1c; font-weight: 700; }
    .priority-medium { background-color: #fef3c7; color: #b45309; }
    .priority-low { background-color: #f1f5f9; color: #475569; }

    /* Storage Diagnostic Card */
    .storage-box {
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 14px;
        padding: 18px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.03);
    }

    /* Buttons */
    div.stButton > button {
        border-radius: 8px;
        font-weight: 600;
        font-size: 13px;
        padding: 6px 16px;
        transition: all 0.15s ease-in-out;
    }
    
    /* Primary Action Buttons */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%);
        border: none;
        color: white;
    }
    div.stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #0f766e 0%, #115e59 100%);
        box-shadow: 0 4px 12px rgba(13, 148, 136, 0.35);
    }

    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #e2e8f0;
        padding: 6px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 42px;
        border-radius: 8px;
        font-weight: 600;
        font-size: 13px;
        color: #475569;
        background-color: transparent;
        border: none;
        padding: 0 16px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #0f766e !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    }
</style>
""", unsafe_allow_html=True)

# --- Top Hero Banner ---
st.markdown("""
<div class="hero-banner">
    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
        <div>
            <h1 class="hero-title">🧪 PharmaLab OS — Automated Lab & Semester Inventory System</h1>
            <div class="hero-subtitle">Pharmacy Labs (Lab 135-B & 138-B) | Full Semester Capacity • Pharmacy POS Dispensing • File Attachments • Crash-Safe WAL Storage</div>
        </div>
        <div style="background: rgba(255,255,255,0.18); backdrop-filter: blur(8px); padding: 8px 16px; border-radius: 10px; font-size: 12px; font-weight: 600;">
            💾 Semester Engine Active | WAL Concurrency Ready
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# --- Top Key Metrics Bar ---
inv_df_all = database.get_inventory_items(limit=3000)
storage_diag = database.get_storage_diagnostics()

total_items = storage_diag["total_inventory"]
total_chemicals = len(inv_df_all[inv_df_all['inventory_group'] == 'Chemical'])
total_labware = len(inv_df_all[inv_df_all['inventory_group'] == 'Hardware / Labware'])
low_stock_count = len(inv_df_all[(inv_df_all['current_stock'] <= inv_df_all['reorder_level']) & (inv_df_all['current_stock'] > 0)])
out_of_stock_count = len(inv_df_all[inv_df_all['current_stock'] <= 0])

col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
with col_m1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value">{total_items:,}</div>
        <div class="metric-label">Total Inventory Items</div>
    </div>
    """, unsafe_allow_html=True)
with col_m2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value" style="color: #0284c7;">{total_chemicals:,}</div>
        <div class="metric-label">Chemicals & Reagents</div>
    </div>
    """, unsafe_allow_html=True)
with col_m3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value" style="color: #8b5cf6;">{total_labware:,}</div>
        <div class="metric-label">Glassware & Hardware</div>
    </div>
    """, unsafe_allow_html=True)
with col_m4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value" style="color: #d97706;">{low_stock_count:,}</div>
        <div class="metric-label">Low Stock Alerts</div>
    </div>
    """, unsafe_allow_html=True)
with col_m5:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value" style="color: #10b981;">{storage_diag['free_disk_gb']:.1f} GB</div>
        <div class="metric-label">Free Storage Capacity</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)

# --- Sidebar Filters & Global Settings ---
with st.sidebar:
    st.image("https://img.icons8.com/isometric/100/test-tube.png", width=64)
    st.markdown("### 🎛️ Semester Operations")
    
    # Semester & Week selectors
    active_semester = st.selectbox("Academic Semester", ["Fall 2026", "Spring 2027", "Summer 2027"], index=0)
    active_week = st.selectbox("Current Semester Week", ["All Weeks"] + [f"Week {i}" for i in range(1, 19)], index=0)
    week_filter_val = "All" if active_week == "All Weeks" else active_week.replace("Week ", "")

    selected_lab = st.selectbox("Active Laboratory Room", ["All Labs", "135-B", "138-B"], index=1)
    
    st.markdown("---")
    st.markdown("#### ⚡ Quick Actions")
    if st.button("🔄 Sync & Reseed from Files", use_container_width=True, help="Reload fresh from Excel & PDF"):
        with st.spinner("Re-seeding database from Excel and PDF timetable..."):
            database.init_db(force_reseed=True)
            st.success("Database reloaded successfully!")
            st.rerun()

    # Excel Download of current stock
    if st.button("📥 Export Stock to Excel (.xlsx)", use_container_width=True):
        export_path, count = database.export_inventory_to_excel()
        with open(export_path, "rb") as f:
            st.download_button(
                label=f"⬇️ Download Current Stock ({count} items)",
                data=f,
                file_name=f"Chemical_and_Labware_Stock_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    st.markdown("---")
    st.markdown(f"""
    **Storage Health (WAL Mode):**
    - 📁 DB Size: **{storage_diag['db_size_kb']} KB**
    - 💾 Free Disk: **{storage_diag['free_disk_gb']} GB**
    - 📑 Attachments: **{storage_diag['attachments_count']} files**
    - ⚡ Engine: **SQLite WAL Concurrency**
    """)

# --- Main Tabs Navigation ---
tab_schedule, tab_logger, tab_inventory, tab_notes, tab_tools, tab_audit, tab_storage = st.tabs([
    "📅 Lab Schedule & Pre-Planning",
    "🧪 Experiment Logging & Requisition Cart",
    "📦 Inventory & Stock Room",
    "📝 Notes & Correspondences",
    "🧮 Lab Productivity Tools",
    "📜 Dispense Ledger & Audit Trail",
    "💾 Semester Storage & Archive Vault"
])

# ==============================================================================
# TAB 1: LAB SCHEDULE & PRE-PLANNING (Editable Timetable)
# ==============================================================================
with tab_schedule:
    st.markdown(f"### 📅 Weekly Lab Timetable & Experiment Pre-Planning ({active_semester})")
    st.info("💡 **Pre-Lab Planning Protocol**: View scheduled classes taken directly from `Lab_Timetable_135B_138B_FA26.pdf`. Click **'Edit / Pre-Plan'** on any slot to log experiment titles, required apparatus, and chemicals beforehand. Use **'Send to Requisition Logger'** to auto-fill the dispensing cart!")

    col_sched_f1, col_sched_f2, col_sched_f3 = st.columns([2, 2, 2])
    with col_sched_f1:
        sched_lab = st.selectbox("Filter Timetable by Lab", ["All", "135-B", "138-B"], index=1 if selected_lab == "135-B" else (2 if selected_lab == "138-B" else 0), key="sched_lab_filter")
    with col_sched_f2:
        sched_day = st.selectbox("Day of Week", ["All", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"], index=0, key="sched_day_filter")
    with col_sched_f3:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        with st.expander("➕ Add New Lab Session"):
            with st.form("new_schedule_form"):
                new_s_lab = st.selectbox("Lab", ["135-B", "138-B"])
                new_s_day = st.selectbox("Day", ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])
                new_s_time = st.text_input("Time Slot", value="08:00 - 11:00")
                new_s_course = st.text_input("Course Name", value="Pharmaceutics 3A")
                new_s_room = st.text_input("Room", value="Room 211C")
                new_s_inst = st.text_input("Instructor", value="Ms. Saman A")
                new_s_exp = st.text_input("Planned Experiment Title")
                new_s_app = st.text_input("Apparatus Required")
                new_s_chem = st.text_input("Chemicals Planned")
                new_s_stud = st.number_input("Expected Students", min_value=1, max_value=200, value=35)
                new_s_week = st.number_input("Semester Week (1-18)", min_value=1, max_value=18, value=1)
                new_s_notes = st.text_area("Session Notes")
                if st.form_submit_button("Save New Session", use_container_width=True):
                    database.add_schedule_entry(new_s_lab, new_s_day, new_s_time, new_s_course, new_s_room, new_s_inst, new_s_exp, new_s_app, new_s_chem, new_s_stud, new_s_notes, active_semester, new_s_week)
                    st.success("New lab session scheduled!")
                    st.rerun()

    schedule_df = database.get_schedule(lab_filter=sched_lab, day_filter=sched_day, semester_filter=active_semester, week_filter=week_filter_val)

    if schedule_df.empty:
        st.warning("No schedule slots found matching the selected filters.")
    else:
        # Group by Day of Week
        days_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        present_days = [d for d in days_order if d in schedule_df['day_of_week'].unique()]

        for day in present_days:
            day_slots = schedule_df[schedule_df['day_of_week'] == day]
            st.markdown(f"#### 🗓️ {day} ({len(day_slots)} Sessions)")
            
            for _, slot in day_slots.iterrows():
                slot_id = slot['id']
                status = slot['status'] or 'Scheduled'
                status_class = f"status-{status.lower().replace(' ', '')}"
                slot_week = slot.get('week_number', 1)
                
                with st.container():
                    st.markdown(f"""
                    <div class="timetable-slot-card">
                        <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                            <div>
                                <span style="font-weight: 700; font-size: 16px; color: #0f766e;">Lab {slot['lab']}</span> 
                                <span style="color: #64748b; font-size: 13px; margin-left: 8px;">⏰ {slot['time_slot']}</span>
                                <span style="margin-left: 8px; background: #f1f5f9; color: #475569;" class="status-badge">Week {slot_week}</span>
                                <span style="margin-left: 6px;" class="status-badge {status_class}">{status}</span>
                                <h4 style="margin: 4px 0 2px 0; color: #1e293b; font-size: 17px;">{slot['course_name']}</h4>
                                <div style="color: #475569; font-size: 13px;">🏛️ {slot['room']} &nbsp;|&nbsp; 👨‍🏫 Instructor: <b>{slot['instructor']}</b> &nbsp;|&nbsp; 👥 {slot['expected_students']} Students</div>
                            </div>
                        </div>
                        <div style="background: #f1f5f9; padding: 10px 14px; border-radius: 8px; margin-top: 10px; font-size: 13px;">
                            <div><b>🧪 Planned Experiment:</b> <span style="color: #0f766e; font-weight: 600;">{slot['experiment_name'] if slot['experiment_name'] else 'Not yet planned (Click Edit below)'}</span></div>
                            <div style="margin-top: 3px;"><b>🔬 Apparatus Required:</b> {slot['apparatus_required'] if slot['apparatus_required'] else 'Standard glassware'}</div>
                            <div style="margin-top: 3px;"><b>📦 Chemicals Planned:</b> {slot['chemicals_planned'] if slot['chemicals_planned'] else 'None specified'}</div>
                            {f"<div style='margin-top: 3px; color: #64748b;'><b>📝 Notes:</b> {slot['notes']}</div>" if slot['notes'] else ""}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                    btn_c1, btn_c2, btn_c3 = st.columns([2, 3, 5])
                    with btn_c1:
                        with st.popover(f"✏️ Edit / Pre-Plan Slot #{slot_id}"):
                            st.markdown(f"**Pre-Plan: {slot['course_name']} ({day} {slot['time_slot']})**")
                            edit_exp = st.text_input("Experiment Title", value=slot['experiment_name'] or "", key=f"e_exp_{slot_id}")
                            edit_app = st.text_area("Apparatus Required", value=slot['apparatus_required'] or "", key=f"e_app_{slot_id}", help="e.g. Mortar & pestle, Burettes, Water bath")
                            edit_chem = st.text_area("Chemicals Planned & Estimated Quantities", value=slot['chemicals_planned'] or "", key=f"e_chem_{slot_id}", help="e.g. Liquid Paraffin 200 mL, Acacia 50 gms")
                            edit_stud = st.number_input("Expected Students", min_value=1, max_value=200, value=int(slot['expected_students'] or 35), key=f"e_stud_{slot_id}")
                            edit_week = st.number_input("Week Number", min_value=1, max_value=18, value=int(slot_week), key=f"e_wk_{slot_id}")
                            edit_stat = st.selectbox("Status", ["Scheduled", "Prepared", "In Progress", "Completed", "Cancelled"], index=["Scheduled", "Prepared", "In Progress", "Completed", "Cancelled"].index(status if status in ["Scheduled", "Prepared", "In Progress", "Completed", "Cancelled"] else "Scheduled"), key=f"e_stat_{slot_id}")
                            edit_note = st.text_input("Special Notes / Instructions", value=slot['notes'] or "", key=f"e_note_{slot_id}")
                            
                            if st.button("💾 Save Pre-Plan Details", key=f"save_slot_{slot_id}", use_container_width=True, type="primary"):
                                database.update_schedule_entry(slot_id, edit_exp, edit_app, edit_chem, edit_stud, edit_stat, edit_note, active_semester, edit_week)
                                st.success("Schedule details updated!")
                                st.rerun()

                    with btn_c2:
                        if st.button(f"🚀 Send to Requisition Logger", key=f"send_log_{slot_id}", use_container_width=True, help="Load this planned experiment directly into Tab 2 (Experiment Logger)"):
                            st.session_state["prefill_experiment"] = {
                                "schedule_id": slot_id,
                                "lab": slot['lab'],
                                "course_name": slot['course_name'],
                                "instructor": slot['instructor'],
                                "experiment_name": slot['experiment_name'] or f"{slot['course_name']} Practical",
                                "apparatus_used": slot['apparatus_required'] or "Standard apparatus",
                                "chemicals_text": slot['chemicals_planned'] or "",
                                "week_number": slot_week
                            }
                            st.success(f"Loaded '{slot['course_name']}' into Experiment Logger! Switch to Tab 2 to finalize chemical deductions.")

# ==============================================================================
# TAB 2: EXPERIMENT LOGGING & REQUISITION CART (PHARMACY-GRADE POS DEDUCTION)
# ==============================================================================
with tab_logger:
    st.markdown("### 🧪 Experiment Logging & Real-Time Chemical Requisition")
    st.info("💊 **Pharmacy-Style Dispensing Protocol**: Search chemicals or apparatus from your inventory, enter quantity to deduct, add to the Requisition Slip, and confirm. Quantities are instantly deducted from the inventory database in real time with an immutable transaction ledger.")

    # Check if pre-filled from Tab 1
    prefill = st.session_state.get("prefill_experiment", {})
    if prefill:
        st.success(f"📋 **Active Pre-Filled Session from Schedule**: {prefill.get('course_name')} ({prefill.get('lab')}) — *'{prefill.get('experiment_name')}'* (Week {prefill.get('week_number', 1)})")
        if st.button("Clear Pre-fill & Start Blank Experiment"):
            st.session_state["prefill_experiment"] = None
            st.rerun()

    exp_col1, exp_col2 = st.columns([1, 1])

    with exp_col1:
        st.markdown("#### 1️⃣ Experiment Information")
        with st.container():
            col_l1, col_l2 = st.columns(2)
            with col_l1:
                log_lab = st.selectbox("Lab Room", ["135-B", "138-B", "M001", "M003"], index=0 if prefill.get("lab") != "138-B" else 1, key="log_lab_input")
                log_course = st.text_input("Course Name", value=prefill.get("course_name", "Pharmaceutics 3A"), key="log_course_input")
                log_instructor = st.text_input("Faculty / Instructor", value=prefill.get("instructor", "Ms. Saman A"), key="log_instructor_input")
                log_semester = st.selectbox("Semester", ["Fall 2026", "Spring 2027", "Summer 2027"], index=0, key="log_sem_input")
            with col_l2:
                log_date = st.date_input("Date of Experiment", value=date.today(), key="log_date_input")
                log_user = st.text_input("Logged By (Technician / Officer)", value="Lab Officer", key="log_user_input")
                log_batch = st.text_input("Student Batch / Group", value="Batch FA26 (Section A)", key="log_batch_input")
                log_week_num = st.number_input("Semester Week (1-18)", min_value=1, max_value=18, value=int(prefill.get("week_number", 1)), key="log_wk_input")

            log_exp_name = st.text_input("🔬 Experiment Title", value=prefill.get("experiment_name", ""), placeholder="e.g. Synthesis of Acetylsalicylic Acid", key="log_title_input")
            log_apparatus = st.text_area("🔬 Apparatus & Hardware Used", value=prefill.get("apparatus_used", "Burettes, Pipettes 10ml, Hot plate, Water bath"), height=70, key="log_app_input")
            log_obs = st.text_area("📝 Observations / Results / Yield Notes", placeholder="e.g. Theoretical yield: 12.5g, Practical yield: 10.8g (86.4%). Reaction completed with standard precipitation.", height=70, key="log_obs_input")

    with exp_col2:
        st.markdown("#### 2️⃣ Chemical & Item Dispensing (POS Cart)")
        with st.container():
            st.markdown("<div style='background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 12px; padding: 16px;'>", unsafe_allow_html=True)
            
            # Search Chemical in Inventory
            search_chem_txt = st.text_input("🔍 Quick Search Item (Name, Item ID, or FR Code)", placeholder="Type e.g. Paraffin, Acid, Ethanol, Sodium...", key="chem_cart_search")
            
            candidate_items = database.get_inventory_items(search_query=search_chem_txt, lab_filter=log_lab if log_lab else "All", limit=30)
            
            if not candidate_items.empty:
                item_options = candidate_items['item_id'].tolist()
                item_labels = {
                    row['item_id']: f"[{row['item_id']}] {row['item_name']} | Avail: {row['current_stock']} {row['unit']} | Loc: {row['stock_location']}"
                    for _, row in candidate_items.iterrows()
                }

                selected_item_id = st.selectbox(
                    "Select Item from Inventory:",
                    options=item_options,
                    format_func=lambda x: item_labels.get(x, x),
                    key="cart_item_select"
                )

                selected_item_row = candidate_items[candidate_items['item_id'] == selected_item_id].iloc[0]

                # Show live stock card
                c_stock = float(selected_item_row['current_stock'] or 0.0)
                c_unit = selected_item_row['unit']
                c_name = selected_item_row['item_name']
                c_exp = selected_item_row['expiry'] or "N/A"
                c_loc = selected_item_row['stock_location'] or "Shelf"

                st.markdown(f"""
                <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px; padding: 12px; margin-top: 8px;">
                    <div style="font-weight: 700; color: #0f766e; font-size: 15px;">{c_name}</div>
                    <div style="display: flex; gap: 16px; margin-top: 6px; font-size: 13px;">
                        <span>📦 Current Stock: <b>{c_stock} {c_unit}</b></span>
                        <span>📍 Location: <b>{c_loc}</b></span>
                        <span>⏳ Expiry: <b>{c_exp}</b></span>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                col_qty, col_add = st.columns([2, 1])
                with col_qty:
                    dispense_qty = st.number_input(f"Quantity to Deduct ({c_unit}):", min_value=0.1, max_value=max(100000.0, c_stock), value=10.0 if c_stock >= 10 else (c_stock if c_stock > 0 else 1.0), step=1.0)
                    if dispense_qty > c_stock:
                        st.error(f"⚠️ Requested {dispense_qty} {c_unit} exceeds available stock of {c_stock} {c_unit}!")
                with col_add:
                    st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                    if st.button("➕ Add to Requisition", use_container_width=True, type="secondary"):
                        existing_idx = next((i for i, item in enumerate(st.session_state["cart_items"]) if item["item_id"] == selected_item_id), None)
                        if existing_idx is not None:
                            st.session_state["cart_items"][existing_idx]["quantity"] += dispense_qty
                        else:
                            st.session_state["cart_items"].append({
                                "item_id": selected_item_id,
                                "item_name": c_name,
                                "quantity": dispense_qty,
                                "unit": c_unit,
                                "stock_before": c_stock,
                                "stock_after": max(0.0, c_stock - dispense_qty),
                                "location": c_loc
                            })
                        st.success(f"Added {dispense_qty} {c_unit} of {c_name} to cart!")
                        st.rerun()
            else:
                st.warning("No items found. Clear search to view available inventory.")

            st.markdown("</div>", unsafe_allow_html=True)

    # --- Live Bill of Materials / Requisition Slip Table ---
    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
    st.markdown("#### 3️⃣ Requisition Slip / Bill of Materials (Simulated Deductions)")

    if not st.session_state["cart_items"]:
        st.info("🛒 Your chemical requisition cart is currently empty. Use the search box above to add chemicals & reagents.")
    else:
        cart_df = pd.DataFrame(st.session_state["cart_items"])
        cols_display = cart_df[['item_id', 'item_name', 'quantity', 'unit', 'stock_before', 'stock_after', 'location']]
        cols_display.columns = ['Item ID', 'Chemical / Reagent', 'Deduct Qty', 'Unit', 'Current Stock', 'Remaining Balance', 'Storage Location']
        
        st.dataframe(cols_display, use_container_width=True, hide_index=True)

        col_cart_act1, col_cart_act2, col_cart_act3 = st.columns([2, 2, 4])
        with col_cart_act1:
            if st.button("🗑️ Clear Requisition Cart", use_container_width=True):
                st.session_state["cart_items"] = []
                st.rerun()

        with col_cart_act3:
            if st.button("✅ Confirm Experiment & Deduct Stock in Real-Time", type="primary", use_container_width=True):
                if not log_exp_name.strip():
                    st.error("Please enter an Experiment Title before confirming.")
                else:
                    exp_info = {
                        "lab": log_lab,
                        "course_name": log_course,
                        "instructor": log_instructor,
                        "experiment_name": log_exp_name,
                        "apparatus_used": log_apparatus,
                        "observations": log_obs,
                        "schedule_id": prefill.get("schedule_id") if prefill else None,
                        "semester": log_semester,
                        "week_number": log_week_num
                    }

                    success, log_ref, msg = database.deduct_inventory_for_experiment(
                        experiment_info=exp_info,
                        items_used=st.session_state["cart_items"],
                        user_name=log_user
                    )

                    if success:
                        st.balloons()
                        st.success(f"🎉 **Transaction Complete!** Log Ref: `{log_ref}` (Week {log_week_num}). {len(st.session_state['cart_items'])} items deducted from stock.")
                        st.session_state["cart_items"] = []
                        st.session_state["prefill_experiment"] = None
                        st.rerun()
                    else:
                        st.error(f"Transaction failed: {msg}")

# ==============================================================================
# TAB 3: INVENTORY & STOCK ROOM (Real-Time Search & Audit)
# ==============================================================================
with tab_inventory:
    st.markdown("### 📦 Central Chemical & Labware Inventory")
    st.info("📊 **Real-Time Stock Room**: Live inventory database loaded from `Total chemicals and lab ware list.xlsx`. Filter by Lab, Category, Type, or search instantly by name or FR code. You can also adjust stock or record new shipments.")

    col_inv_f1, col_inv_f2, col_inv_f3, col_inv_f4 = st.columns([2, 2, 2, 3])
    with col_inv_f1:
        inv_lab_filter = st.selectbox("Lab Room", ["All", "135-B", "138-B", "M001", "M003", "M004", "M005"], index=0, key="inv_lab_f")
    with col_inv_f2:
        inv_group_filter = st.selectbox("Inventory Group", ["All", "Chemical", "Hardware / Labware"], index=0, key="inv_grp_f")
    with col_inv_f3:
        inv_status_filter = st.selectbox("Stock Level", ["All", "Available", "Low Stock", "Out of Stock"], index=0, key="inv_stat_f")
    with col_inv_f4:
        inv_search = st.text_input("🔍 Search Name, ID, or Location", placeholder="e.g. Acetone, Burette, CH-135B...", key="inv_srch")

    filtered_inventory = database.get_inventory_items(
        search_query=inv_search,
        lab_filter=inv_lab_filter,
        group_filter=inv_group_filter,
        status_filter=inv_status_filter,
        limit=2000
    )

    st.markdown(f"**Showing {len(filtered_inventory)} items**")

    display_inv_df = filtered_inventory[['item_id', 'lab', 'inventory_group', 'item_name', 'item_type', 'unit', 'current_stock', 'reorder_level', 'stock_location', 'expiry', 'status']]
    display_inv_df.columns = ['Item ID', 'Lab', 'Group', 'Item Name', 'Type', 'Unit', 'Stock', 'Reorder Lvl', 'Location', 'Expiry', 'Audit Status']

    st.dataframe(
        display_inv_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Stock": st.column_config.NumberColumn(format="%.1f"),
            "Reorder Lvl": st.column_config.NumberColumn(format="%.1f")
        }
    )

    # Restock & Adjustment Tool
    st.markdown("---")
    st.markdown("#### 📥 Stock Inward / Restock / Inventory Adjustment")
    with st.expander("➕ Open Stock Adjustment Panel (Pharmacy Stock-In / Recalibration)"):
        with st.form("restock_form"):
            adj_col1, adj_col2, adj_col3 = st.columns(3)
            with adj_col1:
                adj_item_id = st.text_input("Item ID to Adjust / Restock", placeholder="e.g. CH-135B-0001")
            with adj_col2:
                adj_type = st.selectbox("Adjustment Type", ["RESTOCK", "SET_BALANCE"])
                adj_qty = st.number_input("Quantity (Added if RESTOCK, New Balance if SET_BALANCE)", value=100.0, step=10.0)
            with adj_col3:
                adj_reason = st.text_input("Reference Reason / Invoice #", value="Central Store Delivery FA26")
                adj_user = st.text_input("Authorized By", value="Store Incharge")

            if st.form_submit_button("Submit Stock Adjustment", type="primary", use_container_width=True):
                if not adj_item_id.strip():
                    st.error("Please enter an Item ID.")
                else:
                    success, msg = database.restock_or_adjust_item(adj_item_id.strip(), adj_qty, adj_type, adj_reason, adj_user)
                    if success:
                        st.success(f"Stock updated! {msg}")
                        st.rerun()
                    else:
                        st.error(f"Error: {msg}")

# ==============================================================================
# TAB 4: NOTES, CORRESPONDENCES & PRODUCTIVITY TASKS
# ==============================================================================
with tab_notes:
    st.markdown("### 📝 Lab Notes, Correspondences & Follow-ups")
    st.info("📌 **Productivity Organizer**: Manage lab maintenance schedules, safety inspections, vendor communications, and pending department follow-ups.")

    note_col1, note_col2 = st.columns([1, 2])

    with note_col1:
        st.markdown("#### ➕ Create New Note / Action Item")
        with st.form("new_note_form"):
            n_title = st.text_input("Note / Task Title", placeholder="e.g. Follow-up on autoclave pressure gauge repair")
            n_cat = st.selectbox("Category", [
                "Follow-ups",
                "Correspondences",
                "Safety & Compliance",
                "Equipment Maintenance",
                "Procurement & Requisition",
                "General"
            ])
            n_prio = st.selectbox("Priority", ["High", "Medium", "Low"], index=1)
            n_due = st.date_input("Due Date", value=date.today())
            n_tags = st.text_input("Tags (comma-separated)", placeholder="safety, autoclave, urgent")
            n_body = st.text_area("Details & Description", height=100)

            if st.form_submit_button("Add Note / Task", type="primary", use_container_width=True):
                if not n_title.strip():
                    st.error("Please provide a note title.")
                else:
                    database.add_note(n_title, n_cat, n_prio, "Pending", n_due.strftime("%Y-%m-%d"), n_body, n_tags)
                    st.success("Note created successfully!")
                    st.rerun()

    with note_col2:
        st.markdown("#### 📋 Active Notes & Tasks")
        
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            filter_cat = st.selectbox("Filter Category", ["All", "Follow-ups", "Correspondences", "Safety & Compliance", "Equipment Maintenance", "Procurement & Requisition", "General"])
        with f_col2:
            filter_stat = st.selectbox("Filter Status", ["All", "Pending", "In Progress", "Completed"])
        with f_col3:
            filter_search = st.text_input("Search Notes", placeholder="Search text...")

        notes_df = database.get_notes(category=filter_cat, status=filter_stat, search=filter_search)

        if notes_df.empty:
            st.info("No notes found matching current filters.")
        else:
            for _, note in notes_df.iterrows():
                n_id = note['id']
                prio = note['priority']
                status = note['status']
                prio_class = f"priority-{prio.lower()}"
                
                with st.container():
                    st.markdown(f"""
                    <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px; margin-bottom: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.02);">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <span style="font-weight: 700; font-size: 16px; color: #1e293b;">{note['title']}</span>
                            <div>
                                <span class="status-badge {prio_class}">{prio} Priority</span>
                                <span style="margin-left: 6px; background: #e2e8f0; color: #334155;" class="status-badge">{note['category']}</span>
                                <span style="margin-left: 6px;" class="status-badge {'status-completed' if status == 'Completed' else 'status-scheduled'}">{status}</span>
                            </div>
                        </div>
                        <div style="margin-top: 8px; font-size: 14px; color: #475569; line-height: 1.5;">{note['content']}</div>
                        <div style="margin-top: 10px; font-size: 12px; color: #94a3b8; display: flex; justify-content: space-between;">
                            <span>📅 Due: <b>{note['due_date'] if note['due_date'] else 'None'}</b> &nbsp;|&nbsp; 🏷️ {note['tags']}</span>
                            <span>Created: {note['created_at'][:10]}</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                    btn_c1, btn_c2, btn_c3 = st.columns([2, 2, 4])
                    with btn_c1:
                        if status != "Completed":
                            if st.button("✓ Mark as Done", key=f"done_note_{n_id}", use_container_width=True):
                                database.update_note_status(n_id, "Completed")
                                st.rerun()
                        else:
                            if st.button("↩️ Reopen", key=f"reopen_note_{n_id}", use_container_width=True):
                                database.update_note_status(n_id, "Pending")
                                st.rerun()
                    with btn_c2:
                        if st.button("🗑️ Delete", key=f"del_note_{n_id}", use_container_width=True):
                            database.delete_note(n_id)
                            st.rerun()

# ==============================================================================
# TAB 5: LAB PRODUCTIVITY TOOLS & CALCULATORS
# ==============================================================================
with tab_tools:
    st.markdown("### 🧮 Pharmacy & Chemistry Lab Productivity Tools")
    st.info("🔬 Built-in pharmaceutical formulation and laboratory preparation calculators to eliminate manual arithmetic errors.")

    calc_tab1, calc_tab2, calc_tab3, calc_tab4 = st.tabs([
        "⚖️ Molarity & Mass Calculator",
        "🧪 Dilution Calculator (C1V1 = C2V2)",
        "💧 % Percentage Solutions (w/v, v/v)",
        "⚠️ Critical Reorder & Expiry Monitor"
    ])

    # 1. Molarity Calculator
    with calc_tab1:
        st.markdown("#### Molarity to Mass Calculator")
        st.caption("Formula: Mass (g) = Molarity (mol/L) × Volume (L) × Molecular Weight (g/mol)")
        
        m_c1, m_c2, m_c3 = st.columns(3)
        with m_c1:
            req_molarity = st.number_input("Desired Molarity (M)", min_value=0.001, max_value=20.0, value=0.1, step=0.01)
        with m_c2:
            req_vol_ml = st.number_input("Target Volume (mL)", min_value=1.0, max_value=10000.0, value=500.0, step=50.0)
        with m_c3:
            mol_weight = st.number_input("Molecular Weight / MW (g/mol)", min_value=1.0, max_value=2000.0, value=58.44, step=0.1, help="e.g. NaCl = 58.44, NaOH = 40.00, HCl = 36.46")

        calculated_mass = req_molarity * (req_vol_ml / 1000.0) * mol_weight
        st.markdown(f"""
        <div style="background: #e6fffa; border: 1px solid #99f6e4; border-radius: 12px; padding: 18px; margin-top: 14px;">
            <div style="font-size: 13px; font-weight: 600; color: #0f766e; text-transform: uppercase;">Required Solid Mass to Weigh</div>
            <div style="font-size: 32px; font-weight: 700; color: #0d9488; margin-top: 4px;">{calculated_mass:.4f} grams</div>
            <div style="font-size: 13px; color: #334155; margin-top: 4px;">Weigh <b>{calculated_mass:.4f} g</b> of reagent and dissolve in purified water up to <b>{req_vol_ml} mL</b> in a volumetric flask.</div>
        </div>
        """, unsafe_allow_html=True)

    # 2. Dilution Calculator
    with calc_tab2:
        st.markdown("#### Solution Dilution Calculator (C₁ × V₁ = C₂ × V₂)")
        st.caption("Calculate the exact volume of stock concentrated solution required for dilution.")
        
        d_c1, d_c2, d_c3 = st.columns(3)
        with d_c1:
            stock_conc = st.number_input("Stock Concentration (C₁)", min_value=0.01, value=12.0, step=0.5, help="e.g. Concentrated HCl is approx 12M")
        with d_c2:
            final_conc = st.number_input("Desired Final Concentration (C₂)", min_value=0.001, value=1.0, step=0.1)
        with d_c3:
            final_vol = st.number_input("Desired Final Volume (V₂ in mL)", min_value=1.0, value=250.0, step=25.0)

        if final_conc > stock_conc:
            st.error("Final concentration cannot exceed stock concentration.")
        else:
            required_v1 = (final_conc * final_vol) / stock_conc
            water_to_add = final_vol - required_v1
            st.markdown(f"""
            <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 12px; padding: 18px; margin-top: 14px;">
                <div style="font-size: 13px; font-weight: 600; color: #166534; text-transform: uppercase;">Dilution Recipe</div>
                <div style="font-size: 30px; font-weight: 700; color: #15803d; margin-top: 4px;">Pipette {required_v1:.2f} mL of Stock</div>
                <div style="font-size: 13px; color: #334155; margin-top: 4px;">Add <b>{required_v1:.2f} mL</b> of stock solution to a flask containing water, and bring the total volume up to <b>{final_vol} mL</b> (approx. {water_to_add:.2f} mL solvent).</div>
            </div>
            """, unsafe_allow_html=True)

    # 3. % Solutions
    with calc_tab3:
        st.markdown("#### Percentage Weight/Volume (% w/v) & Volume/Volume (% v/v)")
        p_c1, p_c2 = st.columns(2)
        with p_c1:
            pct_val = st.number_input("Desired Percentage (% w/v)", min_value=0.01, value=5.0, step=0.5)
            pct_vol = st.number_input("Volume of Solution (mL)", min_value=1.0, value=100.0, step=10.0, key="pct_sol_vol")
            pct_mass_needed = (pct_val / 100.0) * pct_vol
            st.info(f"👉 Weigh **{pct_mass_needed:.2f} grams** of solute and make up volume to **{pct_vol} mL**.")
        with p_c2:
            st.markdown("""
            **Common Standard Solutions in Labs:**
            - **Lugol's Iodine**: 5% Iodine + 10% KI in H2O
            - **Normal Saline**: 0.9% w/v Sodium Chloride
            - **Chloral Hydrate Cleansing Solution**: 80% w/v in H2O
            - **Phloroglucinol Reagent**: 1% in 90% Alcohol
            """)

    # 4. Critical Reorder & Expiry Monitor
    with calc_tab4:
        st.markdown("#### 🚨 Reorder Alerts & Out-of-Stock List")
        low_items = inv_df_all[(inv_df_all['current_stock'] <= inv_df_all['reorder_level']) | (inv_df_all['current_stock'] <= 0)]
        if low_items.empty:
            st.success("All stock items are currently above reorder thresholds.")
        else:
            st.warning(f"⚠️ {len(low_items)} items are currently at or below their reorder threshold.")
            st.dataframe(
                low_items[['item_id', 'lab', 'item_name', 'current_stock', 'reorder_level', 'unit', 'stock_location']],
                use_container_width=True,
                hide_index=True
            )

# ==============================================================================
# TAB 6: DISPENSE LEDGER & EXPERIMENT AUDIT TRAIL
# ==============================================================================
with tab_audit:
    st.markdown("### 📜 Dispense Ledger & Experiment Audit Trail")
    st.info("🔍 **Pharmacy Audit Trail**: Complete transparency into all chemical deductions, experiment logs, and balance movements across the semester.")

    audit_t1, audit_t2 = st.tabs(["🧪 Completed Experiment Logs", "📊 Stock Transaction Ledger"])

    with audit_t1:
        logs_df = database.get_experiment_logs(lab_filter=selected_lab if selected_lab != "All Labs" else "All", semester_filter=active_semester, week_filter=week_filter_val)
        if logs_df.empty:
            st.info(f"No experiment logs recorded yet for {active_semester} ({active_week}). Complete an experiment in Tab 2 to see it listed here.")
        else:
            for _, exp in logs_df.iterrows():
                wk = exp.get('week_number', 1)
                with st.expander(f"🔬 {exp['log_id']} — {exp['experiment_name']} (Lab {exp['lab']} | Week {wk} | {exp['timestamp'][:16]})"):
                    c_e1, c_e2 = st.columns(2)
                    with c_e1:
                        st.markdown(f"**Course:** {exp['course_name']}")
                        st.markdown(f"**Instructor:** {exp['instructor']}")
                        st.markdown(f"**Logged By:** {exp['logged_by']}")
                        st.markdown(f"**Apparatus Used:** {exp['apparatus_used']}")
                    with c_e2:
                        st.markdown(f"**Semester / Week:** {exp.get('semester', active_semester)} — Week {wk}")
                        st.markdown(f"**Observations & Yield:** {exp['observations']}")
                    
                    st.markdown("**Chemicals Deducted in this Experiment:**")
                    try:
                        chems = json.loads(exp['chemicals_used_json'])
                        st.table(pd.DataFrame(chems)[['item_id', 'item_name', 'quantity', 'unit']])
                    except:
                        st.write(exp['chemicals_used_json'])

    with audit_t2:
        tx_df = database.get_stock_transactions(limit=500)
        if tx_df.empty:
            st.info("No transactions recorded yet.")
        else:
            st.dataframe(
                tx_df[['id', 'timestamp', 'transaction_type', 'item_id', 'item_name', 'quantity_change', 'balance_after', 'unit', 'user_name', 'reference_reason']],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "quantity_change": st.column_config.NumberColumn("Change", format="%.2f"),
                    "balance_after": st.column_config.NumberColumn("Balance After", format="%.2f")
                }
            )

# ==============================================================================
# TAB 7: SEMESTER STORAGE & ARCHIVE VAULT (UNLIMITED STORAGE RECOVERY)
# ==============================================================================
with tab_storage:
    st.markdown("### 💾 Spacious Semester Storage & Archive Vault")
    st.info("🛡️ **Full Semester Storage Architecture**: Configured with crash-safe SQLite **Write-Ahead Logging (WAL)**, local disk attachment repositories, atomic point-in-time database snapshots, and master multi-sheet Excel archiving. Easily holds millions of records with instantaneous query speeds.")

    diag = database.get_storage_diagnostics()

    # Diagnostics Banner
    s_col1, s_col2, s_col3, s_col4 = st.columns(4)
    with s_col1:
        st.markdown(f"""
        <div class="storage-box">
            <div style="font-size: 12px; font-weight: 600; color: #64748b; text-transform: uppercase;">Database Storage Size</div>
            <div style="font-size: 24px; font-weight: 700; color: #0f766e; margin-top: 4px;">{diag['db_size_kb']} KB</div>
            <div style="font-size: 12px; color: #475569; margin-top: 2px;">WAL Mode: <b>{diag['journal_mode'].upper()}</b></div>
        </div>
        """, unsafe_allow_html=True)
    with s_col2:
        st.markdown(f"""
        <div class="storage-box">
            <div style="font-size: 12px; font-weight: 600; color: #64748b; text-transform: uppercase;">Available Host Disk Space</div>
            <div style="font-size: 24px; font-weight: 700; color: #0284c7; margin-top: 4px;">{diag['free_disk_gb']} GB Free</div>
            <div style="font-size: 12px; color: #475569; margin-top: 2px;">Total Drive: <b>{diag['total_disk_gb']} GB</b></div>
        </div>
        """, unsafe_allow_html=True)
    with s_col3:
        st.markdown(f"""
        <div class="storage-box">
            <div style="font-size: 12px; font-weight: 600; color: #64748b; text-transform: uppercase;">Attachments Stored</div>
            <div style="font-size: 24px; font-weight: 700; color: #8b5cf6; margin-top: 4px;">{diag['attachments_count']} Files</div>
            <div style="font-size: 12px; color: #475569; margin-top: 2px;">Attached Size: <b>{diag['attachments_mb']} MB</b></div>
        </div>
        """, unsafe_allow_html=True)
    with s_col4:
        st.markdown(f"""
        <div class="storage-box">
            <div style="font-size: 12px; font-weight: 600; color: #64748b; text-transform: uppercase;">Theoretical DB Limit</div>
            <div style="font-size: 24px; font-weight: 700; color: #10b981; margin-top: 4px;">281 Terabytes</div>
            <div style="font-size: 12px; color: #475569; margin-top: 2px;">Unlimited for Semester Use</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)

    vault_sub1, vault_sub2, vault_sub3 = st.tabs([
        "📸 Automated Snapshots & Backups",
        "📦 Master Multi-Sheet Semester Archive",
        "📎 Lab Manuals & File Attachments Repository"
    ])

    # 1. Snapshots & Backups
    with vault_sub1:
        st.markdown("#### 📸 Semester Database Snapshots & Zero-Downtime Backups")
        st.caption("Create atomic point-in-time copies using SQLite's native backup API. Snapshots can be restored with a single click.")
        
        c_snap1, c_snap2 = st.columns([2, 3])
        with c_snap1:
            snap_tag = st.text_input("Snapshot Tag / Note", value="Week_Audit", placeholder="e.g. Midterm_Week8, Pre_Audit")
            if st.button("📸 Create Instant Snapshot Backup", type="primary", use_container_width=True):
                ok, fname, fpath, fsize = database.create_database_backup(tag_name=snap_tag.strip())
                if ok:
                    st.success(f"Snapshot created: `{fname}` ({fsize} KB)!")
                    st.rerun()
                else:
                    st.error(f"Failed to create snapshot: {fname}")

        with c_snap2:
            st.markdown("**Available Snapshots on Disk:**")
            backups_df = database.list_backups()
            if backups_df.empty:
                st.info("No snapshots created yet. Click the button on the left to create your first backup.")
            else:
                for _, b in backups_df.iterrows():
                    b_col1, b_col2, b_col3 = st.columns([4, 2, 2])
                    with b_col1:
                        st.markdown(f"**`{b['filename']}`**  \n<span style='font-size: 11px; color: #64748b;'>{b['created_at']} | {b['size_kb']} KB</span>", unsafe_allow_html=True)
                    with b_col2:
                        with open(b['path'], "rb") as f_b:
                            st.download_button("⬇️ Download", data=f_b, file_name=b['filename'], key=f"dl_b_{b['filename']}", use_container_width=True)
                    with b_col3:
                        if st.button("🔄 Restore", key=f"rst_b_{b['filename']}", use_container_width=True, help="Roll back database to this snapshot"):
                            res_ok, res_msg = database.restore_database_backup(b['filename'])
                            if res_ok:
                                st.success(res_msg)
                                st.rerun()
                            else:
                                st.error(res_msg)

    # 2. Master Multi-Sheet Semester Archive
    with vault_sub2:
        st.markdown(f"#### 📦 Complete Academic Semester Master Archive ({active_semester})")
        st.caption("Generates a comprehensive multi-sheet Excel workbook containing all semester tables (Inventory Master, Timetable Pre-plan, Experiment Logs, Stock Transactions, Notes & Tasks) formatted for university accreditation, audits, and department archives.")

        if st.button(f"🚀 Generate Master Multi-Sheet Archive for {active_semester}", type="primary"):
            with st.spinner("Compiling full semester tables into master Excel archive..."):
                arch_path, arch_fname = database.export_full_semester_archive(active_semester)
                st.success(f"Master archive generated successfully: `{arch_fname}`")
                with open(arch_path, "rb") as f_arch:
                    st.download_button(
                        label="⬇️ Download Master Excel Workbook (.xlsx)",
                        data=f_arch,
                        file_name=arch_fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

    # 3. Attachments Repository
    with vault_sub3:
        st.markdown("#### 📎 Lab Attachments, Manuals & Protocol Repository")
        st.caption("Upload lab protocol PDFs, MSDS documents, spectra charts, equipment manuals, or experiment photos with up to 2GB storage capacity per file.")

        att_up_col1, att_up_col2 = st.columns([2, 3])
        with att_up_col1:
            uploaded_file = st.file_uploader("Upload Lab Document / Image / Manual", type=["pdf", "png", "jpg", "jpeg", "xlsx", "docx", "csv", "txt"])
            att_type = st.selectbox("Link to Entity", ["EXPERIMENT", "TIMETABLE_SLOT", "NOTE", "GENERAL_MANUAL"])
            att_id = st.text_input("Entity Reference ID / Tag", value="EXP-GENERAL", placeholder="e.g. EXP-20260930-110147 or CH-135B-0001")
            att_user = st.text_input("Uploaded By", value="Lab Incharge")

            if st.button("💾 Save Attachment to Storage", type="primary", use_container_width=True):
                if uploaded_file is None:
                    st.error("Please choose a file to upload.")
                else:
                    bytes_data = uploaded_file.read()
                    file_uuid, path = database.save_attachment(att_type, att_id, bytes_data, uploaded_file.name, att_user)
                    st.success(f"File '{uploaded_file.name}' saved to storage!")
                    st.rerun()

        with att_up_col2:
            st.markdown("**Stored Documents & Attachments:**")
            att_df = database.get_attachments()
            if att_df.empty:
                st.info("No documents or attachments uploaded yet.")
            else:
                for _, att in att_df.iterrows():
                    a_c1, a_c2, a_c3 = st.columns([4, 2, 2])
                    with a_c1:
                        st.markdown(f"**{att['file_name']}**  \n<span style='font-size: 11px; color: #64748b;'>Type: {att['entity_type']} | Ref: {att['entity_id']} | Size: {round(att['file_size_bytes']/1024, 1)} KB</span>", unsafe_allow_html=True)
                    with a_c2:
                        if os.path.exists(att['file_path']):
                            with open(att['file_path'], "rb") as f_att:
                                st.download_button("⬇️ Download", data=f_att, file_name=att['file_name'], key=f"dl_att_{att['id']}", use_container_width=True)
                    with a_c3:
                        if st.button("🗑️ Delete", key=f"del_att_{att['id']}", use_container_width=True):
                            database.delete_attachment(att['id'])
                            st.rerun()
