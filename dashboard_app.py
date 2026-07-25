import io
import os
import sqlite3
import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image
from datetime import datetime

# -------------------------------------------------------------
# 1. PAGE CONFIGURATION & CUSTOM STYLING
# -------------------------------------------------------------
try:
    logo_icon = Image.open("mvvnl.jpeg")
except Exception:
    logo_icon = "⚡"

st.set_page_config(
    page_title="MVVNL Complaint Analytics & Verification",
    page_icon=logo_icon,
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS with LARGER Block-style Tab buttons
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        border-left: 5px solid #1F497D;
    }
    .login-box {
        background-color: #ffffff;
        padding: 30px;
        border-radius: 12px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    }
    
    /* Prominent Block-Form Tab Styling with Extra Large Font Size */
    button[data-baseweb="tab"] {
        background-color: #e2e8f0 !important;
        border-radius: 10px !important;
        padding: 18px 36px !important;
        margin-right: 16px !important;
        font-size: 24px !important;
        font-weight: 800 !important;
        color: #0f172a !important;
        border: 2px solid #cbd5e1 !important;
        box-shadow: 0px 4px 8px rgba(0,0,0,0.1) !important;
        width: 100% !important;
        text-align: center !important;
        transition: all 0.2s ease-in-out !important;
    }
    button[data-baseweb="tab"]:hover {
        background-color: #cbd5e1 !important;
        border-color: #1E3A8A !important;
        cursor: pointer !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #1E3A8A !important;
        color: #ffffff !important;
        border-color: #1E3A8A !important;
        box-shadow: 0px 6px 14px rgba(30, 58, 138, 0.3) !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] p {
        color: #ffffff !important;
        font-size: 24px !important;
        font-weight: 800 !important;
    }
    </style>
""", unsafe_allow_html=True)

DEFAULT_PASSWORD = "123456"
DB_PATH = "feedback_records.db"

# -------------------------------------------------------------
# 2. LOAD AUTHORIZED USER IDs FROM LOGIN.XLSX
# -------------------------------------------------------------
@st.cache_data
def load_authorized_users(login_file_path="LOGIN.xlsx"):
    """Reads LOGIN.xlsx and builds a dictionary mapping USER ID -> Employee Name."""
    try:
        df_login = pd.read_excel(login_file_path)
        df_login.columns = df_login.columns.str.strip()
        
        if 'USER ID' in df_login.columns:
            df_login['USER_ID_STR'] = (
                df_login['USER ID']
                .astype(str)
                .str.strip()
                .str.replace(r'\.0$', '', regex=True)
            )
            
            emp_name_col = 'Employee Name' if 'Employee Name' in df_login.columns else 'USER_ID_STR'
            user_dict = dict(zip(df_login['USER_ID_STR'], df_login[emp_name_col].fillna('Officer')))
            return user_dict
        else:
            st.error("⚠️ 'USER ID' column not found in LOGIN.xlsx!")
            return {}
    except Exception as e:
        st.error(f"⚠️ Error loading LOGIN.xlsx user database: {e}")
        return {}

AUTHORIZED_USERS = load_authorized_users("LOGIN.xlsx")

# -------------------------------------------------------------
# 3. SQLITE MULTI-USER DATABASE ENGINE
# -------------------------------------------------------------
def get_db_connection():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS consumer_feedback (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                complaint_no TEXT UNIQUE,
                consumer_name TEXT,
                consumer_mobile TEXT,
                zone TEXT,
                circle TEXT,
                division TEXT,
                substation TEXT,
                closed_by TEXT,
                feedback_status TEXT,
                feedback_remark TEXT,
                agent_id TEXT
            )
        """)
        conn.commit()

init_db()

def get_completed_complaint_numbers():
    with get_db_connection() as conn:
        df_completed = pd.read_sql_query("SELECT complaint_no FROM consumer_feedback", conn)
    if not df_completed.empty:
        return set(df_completed['complaint_no'].astype(str).str.strip().tolist())
    return set()

def save_feedback_sqlite(record_dict):
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO consumer_feedback (
                    timestamp, complaint_no, consumer_name, consumer_mobile,
                    zone, circle, division, substation, closed_by,
                    feedback_status, feedback_remark, agent_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                record_dict.get('timestamp'), record_dict.get('complaint_no'),
                record_dict.get('consumer_name'), record_dict.get('consumer_mobile'),
                record_dict.get('zone'), record_dict.get('circle'),
                record_dict.get('division'), record_dict.get('substation'),
                record_dict.get('closed_by'), record_dict.get('feedback_status'),
                record_dict.get('feedback_remark'), record_dict.get('agent_id')
            ))
            conn.commit()
            return True
    except sqlite3.IntegrityError:
        st.error("⚠️ Feedback for this complaint number has already been recorded!")
        return False

# -------------------------------------------------------------
# 4. COMPLAINT DATA LOADING & PREPROCESSING
# -------------------------------------------------------------
@st.cache_data
def load_and_process_data(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip().str.upper()

    text_cols = ['COMPLAINT_NO', 'CONSUMER_NAME', 'ZONE', 'CIRCLE', 'DIVISION', 'SUBSTATION', 'STS', 'REMARKS', 'STAFFREMARKS', 'SOURCE', 'COMPLAINT_TYPE']
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    df['STS_CLEAN'] = df['STS'].str.upper() if 'STS' in df.columns else 'UNKNOWN'
    
    mobile_col_candidates = [col for col in df.columns if any(k in col for k in ['MOBILE', 'PHONE', 'CONTACT', 'CONSUMER_NO', 'ACCOUNT_NO'])]
    if mobile_col_candidates:
        df['MOBILE_NO'] = df[mobile_col_candidates[0]].astype(str).str.replace(r'\D', '', regex=True)
    else:
        df['MOBILE_NO'] = 'UNKNOWN'

    df['PENDING_STATUS'] = df['STS_CLEAN'].apply(lambda x: 'Closed' if x == 'CLOSED' else 'Pending')

    for orig_col, dt_col in [('REGISTRATION_DATE', 'REGISTRATION_DT'), ('CLOSINGDATE', 'CLOSING_DT')]:
        if orig_col in df.columns:
            num_series = pd.to_numeric(df[orig_col], errors='coerce')
            if num_series.notna().sum() > (len(df) * 0.5):
                df[dt_col] = pd.to_datetime(num_series, unit='D', origin='1899-12-30', errors='coerce')
            else:
                df[dt_col] = pd.to_datetime(df[orig_col], errors='coerce')
        else:
            df[dt_col] = pd.NaT

    valid_dates = df['CLOSING_DT'].notna() & df['REGISTRATION_DT'].notna()
    df['RESOLUTION_TIME_HRS'] = pd.NA
    df.loc[valid_dates, 'RESOLUTION_TIME_HRS'] = (
        (df.loc[valid_dates, 'CLOSING_DT'] - df.loc[valid_dates, 'REGISTRATION_DT'])
        .dt.total_seconds() / 3600.0
    )
    df['REG_DATE'] = df['REGISTRATION_DT'].dt.date
    return df

try:
    df = load_and_process_data('june_data.xlsx')
except Exception as e:
    st.error(f"Error loading dataset 'june_data.xlsx': {e}")
    st.stop()

def pick_random_complaint():
    completed = get_completed_complaint_numbers()
    pending_pool = df[~df['COMPLAINT_NO'].astype(str).isin(completed)]
    if pending_pool.empty:
        return None
    return pending_pool.sample(n=1).iloc[0].to_dict()

if 'current_complaint' not in st.session_state:
    st.session_state.current_complaint = pick_random_complaint()

# -------------------------------------------------------------
# 5. SIDEBAR FILTERS
# -------------------------------------------------------------
st.sidebar.title("⚡ Dashboard Filters")

zones = ['All'] + sorted([str(z) for z in df['ZONE'].dropna().unique()]) if 'ZONE' in df.columns else ['All']
selected_zone = st.sidebar.selectbox("Select Zone:", zones)

comp_types = ['All'] + sorted([str(ct) for ct in df['COMPLAINT_TYPE'].dropna().unique()]) if 'COMPLAINT_TYPE' in df.columns else ['All']
selected_type = st.sidebar.selectbox("Select Complaint Type:", comp_types)

sts_options = ['All'] + sorted([str(st_val) for st_val in df['STS'].dropna().unique()])
selected_sts = st.sidebar.selectbox("Filter by Status (STS):", sts_options)

filtered_df = df.copy()
if selected_zone != 'All' and 'ZONE' in filtered_df.columns:
    filtered_df = filtered_df[filtered_df['ZONE'] == selected_zone]
if selected_type != 'All' and 'COMPLAINT_TYPE' in filtered_df.columns:
    filtered_df = filtered_df[filtered_df['COMPLAINT_TYPE'] == selected_type]
if selected_sts != 'All':
    filtered_df = filtered_df[filtered_df['STS'] == selected_sts]

# -------------------------------------------------------------
# 6. MAIN HEADER & TABBED NAVIGATION
# -------------------------------------------------------------
header_col1, header_col2 = st.columns([1, 10])
with header_col1:
    try:
        st.image("mvvnl.jpeg", width=70)
    except Exception:
        pass
with header_col2:
    st.title("MVVNL Integrated Feedback & Complaint Analytics")

tab_analytics, tab_feedback = st.tabs(["📊 EXECUTIVE ANALYTICAL DASHBOARD", "🔒 CONSUMER FEEDBACK"])

# =============================================================
# TAB 1: EXECUTIVE ANALYTICAL DASHBOARD
# =============================================================
with tab_analytics:
    # --- TOP KPI METRIC CARDS ---
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Complaints", f"{len(filtered_df):,}")
    with col2:
        valid_res_times = pd.to_numeric(filtered_df['RESOLUTION_TIME_HRS'], errors='coerce')
        avg_hrs = valid_res_times.mean() if not valid_res_times.empty and valid_res_times.notna().sum() > 0 else 0
        st.metric("Avg Closure Time", f"{avg_hrs:.1f} hrs")
    with col3:
        pending_count = (filtered_df['STS_CLEAN'] != 'CLOSED').sum() if not filtered_df.empty else 0
        pending_pct = (pending_count / len(filtered_df) * 100) if len(filtered_df) > 0 else 0
        st.metric("Total Pending", f"{pending_count:,}", delta=f"{pending_pct:.1f}% Pending", delta_color="inverse")
    with col4:
        is_beyond = filtered_df['STS_CLEAN'].str.contains('BEYOND|OVERDUE', regex=True)
        beyond_count = is_beyond.sum() if not filtered_df.empty else 0
        st.metric("Beyond Time Pendency", f"{beyond_count:,} Tickets")

    st.divider()

    # --- ROW 1 VISUALIZATIONS ---
    r1_col1, r1_col2 = st.columns(2)

    # 1. Zone-Wise Comparison of Total Complaints Received
    with r1_col1:
        st.subheader("📈 Zone-Wise Total Complaints Received Trend Over Time")
        if 'ZONE' in filtered_df.columns and 'REG_DATE' in filtered_df.columns and not filtered_df.empty:
            zone_time_df = filtered_df.groupby(['REG_DATE', 'ZONE']).size().reset_index(name='Complaints_Count')
            zone_time_df = zone_time_df.sort_values(by='REG_DATE')
            
            fig_zone_trend = px.line(
                zone_time_df,
                x='REG_DATE',
                y='Complaints_Count',
                color='ZONE',
                markers=True,
                labels={'REG_DATE': 'Date', 'Complaints_Count': 'Total Complaints Received', 'ZONE': 'Zone'}
            )
            fig_zone_trend.update_layout(template="plotly_white", margin=dict(l=20, r=20, t=30, b=20), hovermode="x unified")
            st.plotly_chart(fig_zone_trend, use_container_width=True)

    # 2. Comparative Pendency by Complaint Type
    with r1_col2:
        st.subheader("🎨 Comparative Pendency by Complaint Type")
        if 'COMPLAINT_TYPE' in filtered_df.columns and not filtered_df.empty:
            type_status = filtered_df.groupby(['COMPLAINT_TYPE', 'PENDING_STATUS']).size().reset_index(name='Count')
            
            color_discrete_map = {
                'Closed': '#1F497D',
                'Pending': '#E74C3C',
            }
            
            fig_pendency = px.bar(
                type_status, y='COMPLAINT_TYPE', x='Count', color='PENDING_STATUS',
                barmode='group', orientation='h',
                text='Count',
                color_discrete_map=color_discrete_map,
                labels={'COMPLAINT_TYPE': 'Complaint Type', 'Count': 'Complaints Count', 'PENDING_STATUS': 'Status'}
            )
            fig_pendency.update_traces(textposition='auto')
            fig_pendency.update_layout(template="plotly_white", margin=dict(l=20, r=20, t=30, b=20))
            st.plotly_chart(fig_pendency, use_container_width=True)

    st.divider()

    # --- ROW 2 VISUALIZATIONS ---
    r2_col1, r2_col2 = st.columns(2)

    # 3. Complaint Type-Wise Avg Closure Time
    with r2_col1:
        st.subheader("⏱️ Complaint Type-Wise Avg Closure Time")
        if 'COMPLAINT_TYPE' in filtered_df.columns and not filtered_df.empty:
            df_temp = filtered_df.copy()
            df_temp['RESOLUTION_TIME_HRS'] = pd.to_numeric(df_temp['RESOLUTION_TIME_HRS'], errors='coerce')
            avg_time_df = df_temp.groupby('COMPLAINT_TYPE')['RESOLUTION_TIME_HRS'].mean().reset_index()
            avg_time_df = avg_time_df.dropna().sort_values(by='RESOLUTION_TIME_HRS', ascending=True)

            fig_avg_time = px.bar(
                avg_time_df,
                x='RESOLUTION_TIME_HRS',
                y='COMPLAINT_TYPE',
                orientation='h',
                text=avg_time_df['RESOLUTION_TIME_HRS'].apply(lambda x: f"{x:.1f} hrs"),
                color='RESOLUTION_TIME_HRS',
                color_continuous_scale='Reds',
                labels={'RESOLUTION_TIME_HRS': 'Avg Resolution Time (Hours)', 'COMPLAINT_TYPE': 'Complaint Type'}
            )
            fig_avg_time.update_traces(textposition='outside')
            fig_avg_time.update_layout(template="plotly_white", coloraxis_showscale=False, margin=dict(l=20, r=20, t=30, b=20))
            st.plotly_chart(fig_avg_time, use_container_width=True)

    # 4. Top 10 Substations with Max Complaints
    with r2_col2:
        st.subheader("🏢 Top 10 Substations with Max Complaints")
        if 'SUBSTATION' in filtered_df.columns and not filtered_df.empty:
            sub_df = filtered_df['SUBSTATION'].value_counts().head(10).reset_index()
            sub_df.columns = ['SUBSTATION', 'COUNT']
            sub_df = sub_df.sort_values(by='COUNT', ascending=True)

            fig_sub = px.bar(
                sub_df,
                x='COUNT',
                y='SUBSTATION',
                orientation='h',
                text='COUNT',
                color_discrete_sequence=['#2E86C1'],
                labels={'COUNT': 'Total Complaints', 'SUBSTATION': 'Substation'}
            )
            fig_sub.update_traces(textposition='outside')
            fig_sub.update_layout(template="plotly_white", margin=dict(l=20, r=20, t=30, b=20))
            st.plotly_chart(fig_sub, use_container_width=True)

    st.divider()

    # --- ROW 3 VISUALIZATIONS: REPEATED COMPLAINTS ---
    r3_col1, r3_col2 = st.columns(2)

    # 5. Repeated Consumer Complaints Zone-Wise
    with r3_col1:
        st.subheader("🔁 Repeated Consumer Complaints Zone-Wise")
        if 'MOBILE_NO' in filtered_df.columns and 'ZONE' in filtered_df.columns and not filtered_df.empty:
            repeat_df = filtered_df[filtered_df['MOBILE_NO'] != 'UNKNOWN']
            mobile_counts = repeat_df['MOBILE_NO'].value_counts()
            repeated_mobiles = mobile_counts[mobile_counts > 1].index
            repeat_filtered = repeat_df[repeat_df['MOBILE_NO'].isin(repeated_mobiles)]

            if not repeat_filtered.empty:
                zone_repeat = repeat_filtered.groupby('ZONE')['MOBILE_NO'].nunique().reset_index(name='Repeat_Consumers')
                
                fig_repeat = px.bar(
                    zone_repeat,
                    x='ZONE',
                    y='Repeat_Consumers',
                    text='Repeat_Consumers',
                    color='ZONE',
                    color_discrete_sequence=px.colors.qualitative.Set2,
                    labels={'ZONE': 'Zone', 'Repeat_Consumers': 'Repeat Consumers Count'}
                )
                fig_repeat.update_traces(textposition='outside')
                fig_repeat.update_layout(template="plotly_white", showlegend=False, margin=dict(l=20, r=20, t=30, b=20))
                st.plotly_chart(fig_repeat, use_container_width=True)
            else:
                st.info("✅ No repeated consumer complaints found in current selection.")

    # 6. Repeated Consumer Complaints Type-Wise
    with r3_col2:
        st.subheader("🔁 Repeated Consumer Complaints Type-Wise")
        if 'MOBILE_NO' in filtered_df.columns and 'COMPLAINT_TYPE' in filtered_df.columns and not filtered_df.empty:
            repeat_df = filtered_df[filtered_df['MOBILE_NO'] != 'UNKNOWN']
            mobile_counts = repeat_df['MOBILE_NO'].value_counts()
            repeated_mobiles = mobile_counts[mobile_counts > 1].index
            repeat_filtered = repeat_df[repeat_df['MOBILE_NO'].isin(repeated_mobiles)]

            if not repeat_filtered.empty:
                type_repeat = repeat_filtered.groupby('COMPLAINT_TYPE')['MOBILE_NO'].nunique().reset_index(name='Repeat_Consumers')
                type_repeat = type_repeat.sort_values(by='Repeat_Consumers', ascending=True)

                fig_type_repeat = px.bar(
                    type_repeat,
                    x='Repeat_Consumers',
                    y='COMPLAINT_TYPE',
                    orientation='h',
                    text='Repeat_Consumers',
                    color_discrete_sequence=['#E67E22'],
                    labels={'COMPLAINT_TYPE': 'Complaint Type', 'Repeat_Consumers': 'Repeat Consumers Count'}
                )
                fig_type_repeat.update_traces(textposition='outside')
                fig_type_repeat.update_layout(template="plotly_white", margin=dict(l=20, r=20, t=30, b=20))
                st.plotly_chart(fig_type_repeat, use_container_width=True)
            else:
                st.info("✅ No repeated consumer complaints found in current selection.")

    st.divider()

    # --- ROW 4: COMPLAINT SOURCE BREAKDOWN (DARK VERTICAL BAR CHART WITH VALUE & PERCENTAGE) ---
    st.subheader("📡 Complaint Source Breakdown")
    source_col = [c for c in filtered_df.columns if 'SOURCE' in c or 'CHANNEL' in c or 'MODE' in c]
    if source_col and not filtered_df.empty:
        col_name = source_col[0]
        source_counts = filtered_df[col_name].value_counts().reset_index()
        source_counts.columns = ['Source', 'Count']
        
        # Calculate Percentage
        total_source_count = source_counts['Count'].sum()
        source_counts['Percentage'] = (source_counts['Count'] / total_source_count) * 100
        
        # Combined Label showing both Count Value & Percentage
        source_counts['Label'] = source_counts.apply(
            lambda row: f"{row['Count']:,}<br>({row['Percentage']:.1f}%)", axis=1
        )
        source_counts = source_counts.sort_values(by='Count', ascending=False)

        # Dark Form Vertical Bar Chart
        fig_source_bar = px.bar(
            source_counts,
            x='Source',
            y='Count',
            text='Label',
            color_discrete_sequence=['#0B2545'],  # Dark Navy Solid Color
            labels={'Count': 'Total Complaints', 'Source': 'Channel / Source'}
        )
        fig_source_bar.update_traces(
            textposition='outside',
            textfont_size=13,
            marker_line_color='#020C1B',
            marker_line_width=1.5
        )
        fig_source_bar.update_layout(
            template="plotly_white",
            margin=dict(l=20, r=20, t=40, b=30),
            height=480,
            yaxis=dict(title='Total Complaints', showgrid=True, gridcolor='#E2E8F0'),
            xaxis=dict(title='Complaint Source')
        )
        st.plotly_chart(fig_source_bar, use_container_width=True)
    else:
        st.info("ℹ️ Source column not detected in dataset.")

    st.divider()

    # --- RAW DATA TABLE EXPLORER ---
    st.subheader("📋 Raw Data Explorer")
    st.dataframe(filtered_df.head(100), use_container_width=True)

# =============================================================
# TAB 2: CONSUMER FEEDBACK (RESTRICTED PAGE SECURITY)
# =============================================================
with tab_feedback:
    if "fb_authenticated" not in st.session_state:
        st.session_state.fb_authenticated = False
    if "fb_user_id" not in st.session_state:
        st.session_state.fb_user_id = ""
    if "fb_user_name" not in st.session_state:
        st.session_state.fb_user_name = ""

    # LOGIN FORM
    if not st.session_state.fb_authenticated:
        st.markdown("### 🔒 Authorized Verification Agent Login")
        st.info("Access to the Consumer Feedback module is restricted. Enter your official **User ID** (from `LOGIN.xlsx`) and password.")

        login_col1, login_col2, login_col3 = st.columns([1, 2, 1])
        with login_col2:
            st.markdown("<div class='login-box'>", unsafe_allow_html=True)
            with st.form(key="feedback_login_form"):
                st.subheader("🔑 Agent Authentication")
                user_id_input = st.text_input("User ID (e.g., 10000075)", placeholder="Enter 8-digit User ID")
                password_input = st.text_input("Password", type="password", placeholder="Default: 123456")
                
                login_submit = st.form_submit_button("Authenticate & Access", type="primary", use_container_width=True)

                if login_submit:
                    clean_uid = user_id_input.strip()
                    if clean_uid in AUTHORIZED_USERS:
                        if password_input == DEFAULT_PASSWORD:
                            st.session_state.fb_authenticated = True
                            st.session_state.fb_user_id = clean_uid
                            st.session_state.fb_user_name = AUTHORIZED_USERS[clean_uid]
                            st.success(f"✅ Welcome, {st.session_state.fb_user_name}!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect password! Default password is '123456'.")
                    else:
                        st.error("❌ Invalid User ID. User ID not found in LOGIN.xlsx file.")
            st.markdown("</div>", unsafe_allow_html=True)

    # PROTECTED WORKSPACE
    else:
        sess_col1, sess_col2 = st.columns([3, 1])
        with sess_col1:
            st.success(f"👤 **Authenticated Agent:** `{st.session_state.fb_user_name}` | **ID:** `{st.session_state.fb_user_id}`")
        with sess_col2:
            if st.button("🚪 Logout from Feedback", use_container_width=True):
                st.session_state.fb_authenticated = False
                st.session_state.fb_user_id = ""
                st.session_state.fb_user_name = ""
                st.rerun()

        st.divider()

        completed_set = get_completed_complaint_numbers()
        total_records = len(df)
        completed_count = len(completed_set)
        remaining_count = max(0, total_records - completed_count)

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Complaints in File", f"{total_records:,}")
        m2.metric("Verified Feedback Recorded", f"{completed_count:,}")
        m3.metric("Remaining Pending Pool", f"{remaining_count:,}")

        st.divider()

        current_ticket = st.session_state.current_complaint

        if not current_ticket:
            st.balloons()
            st.success("🎉 All complaints in the dataset have been verified and recorded!")
        else:
            top_col1, top_col2 = st.columns([3, 1])
            with top_col1:
                st.markdown(f"### 🎯 Allotted Ticket: `{current_ticket.get('COMPLAINT_NO', 'N/A')}`")
            with top_col2:
                if st.button("🔀 Skip & Get Next Random Ticket", use_container_width=True):
                    st.session_state.current_complaint = pick_random_complaint()
                    st.rerun()

            with st.container():
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Complaint No", current_ticket.get('COMPLAINT_NO', 'N/A'))
                c2.metric("Consumer Name", current_ticket.get('CONSUMER_NAME', 'N/A'))
                c3.metric("Mobile No", current_ticket.get('MOBILE_NO', 'N/A'))
                c4.metric("Status (STS)", current_ticket.get('STS', 'N/A'))
                
                d1, d2, d3 = st.columns(3)
                d1.write(f"**Zone / Circle:** {current_ticket.get('ZONE', 'N/A')} / {current_ticket.get('CIRCLE', 'N/A')}")
                d2.write(f"**Substation:** {current_ticket.get('SUBSTATION', 'N/A')}")
                d3.write(f"**Closed By:** {current_ticket.get('CLOSEDBY', 'N/A')}")
                
                st.write(f"**Consumer Address:** {current_ticket.get('CONSUMER_ADDRESS', 'N/A')}")
                
                r1, r2 = st.columns(2)
                with r1:
                    st.info(f"🗣️ **Consumer Remarks:**\n{current_ticket.get('REMARKS', 'N/A')}")
                with r2:
                    st.warning(f"🛠️ **Staff Remarks:**\n{current_ticket.get('STAFFREMARKS', 'N/A')}")

            st.divider()

            st.subheader("✍️ Record Feedback Verification")
            with st.form(key="feedback_verification_form", clear_on_submit=True):
                f1, f2 = st.columns(2)
                with f1:
                    feedback_status = st.radio(
                        "Consumer Feedback Status *",
                        options=["Satisfied", "Not Satisfied", "Wrongly Closed"],
                        index=0,
                        horizontal=True
                    )
                with f2:
                    st.text_input("Recording Agent ID & Name:", value=f"{st.session_state.fb_user_id} - {st.session_state.fb_user_name}", disabled=True)

                feedback_remark = st.text_area(
                    "Consumer Feedback Remarks / Call Summary *",
                    placeholder="Enter caller response, status verification notes...",
                    height=100
                )

                submit_button = st.form_submit_button(label="💾 Submit Feedback & Next Ticket", type="primary", use_container_width=True)

            if submit_button:
                if not feedback_remark.strip():
                    st.error("⚠️ Please enter feedback remarks before submitting.")
                else:
                    record = {
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "complaint_no": str(current_ticket.get('COMPLAINT_NO')).strip(),
                        "consumer_name": current_ticket.get('CONSUMER_NAME'),
                        "consumer_mobile": current_ticket.get('MOBILE_NO'),
                        "zone": current_ticket.get('ZONE'),
                        "circle": current_ticket.get('CIRCLE'),
                        "division": current_ticket.get('DIVISION'),
                        "substation": current_ticket.get('SUBSTATION'),
                        "closed_by": current_ticket.get('CLOSEDBY'),
                        "feedback_status": feedback_status,
                        "feedback_remark": feedback_remark,
                        "agent_id": f"{st.session_state.fb_user_id} ({st.session_state.fb_user_name})"
                    }
                    
                    if save_feedback_sqlite(record):
                        st.toast(f"✅ Feedback saved for Complaint #{current_ticket.get('COMPLAINT_NO')}!")
                        st.session_state.current_complaint = pick_random_complaint()
                        st.rerun()

        st.divider()
        st.subheader("📊 Submitted Feedback Audit Log")
        with get_db_connection() as conn:
            fb_df = pd.read_sql_query("SELECT * FROM consumer_feedback ORDER BY id DESC", conn)

        if not fb_df.empty:
            st.dataframe(fb_df, use_container_width=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                fb_df.to_excel(writer, sheet_name='Feedback_Audit', index=False)
                
            st.download_button(
                label="📥 Download SQLite Feedback Audit Report",
                data=buffer.getvalue(),
                file_name=f"MVVNL_Feedback_Audit_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
