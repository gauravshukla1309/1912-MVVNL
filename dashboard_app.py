import io
import os
import sqlite3
import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image
from datetime import datetime

# -------------------------------------------------------------
# 1. PAGE CONFIGURATION WITH MVVNL LOGO & CUSTOM STYLING
# -------------------------------------------------------------
try:
    logo_icon = Image.open("mvvnl.jpeg")
except Exception:
    logo_icon = "⚡"  # Fallback emoji if image file is missing

st.set_page_config(
    page_title="MVVNL Complaint Analytics & Verification",
    page_icon=logo_icon,
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling for Metric Cards, Reports, and Mobile Optimization
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
    .report-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 5px rgba(0,0,0,0.03);
        margin-bottom: 25px;
    }
    .stButton>button {
        border-radius: 6px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

DB_PATH = "feedback_records.db"

# -------------------------------------------------------------
# 2. SQLITE MULTI-USER DATABASE ENGINE
# -------------------------------------------------------------
def get_db_connection():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Initializes the SQLite feedback table if it does not exist."""
    conn = get_db_connection()
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
    conn.close()

init_db()

def get_completed_complaint_numbers():
    """Returns a set of complaint numbers that have already received feedback in SQLite."""
    conn = get_db_connection()
    df_completed = pd.read_sql_query("SELECT complaint_no FROM consumer_feedback", conn)
    conn.close()
    if not df_completed.empty:
        return set(df_completed['complaint_no'].astype(str).str.strip().tolist())
    return set()

def save_feedback_sqlite(record_dict):
    """Inserts a feedback record into the SQLite database atomically."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
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
    finally:
        conn.close()

# -------------------------------------------------------------
# 3. DATA LOADING & PREPROCESSING
# -------------------------------------------------------------
@st.cache_data
def load_and_process_data(file_path):
    df = pd.read_excel(file_path)
    
    # Standardize Column Names (strip whitespace & convert to upper)
    df.columns = df.columns.str.strip().str.upper()

    # Clean text columns
    text_cols = ['COMPLAINT_NO', 'CONSUMER_NAME', 'ZONE', 'CIRCLE', 'DIVISION', 'SUBSTATION', 'STS', 'REMARKS', 'STAFFREMARKS']
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Ensure STS column is uppercase/trimmed for accurate matching
    if 'STS' in df.columns:
        df['STS_CLEAN'] = df['STS'].str.upper()
    else:
        df['STS'] = 'UNKNOWN'
        df['STS_CLEAN'] = 'UNKNOWN'
        
    # Standardize/Clean Mobile Numbers for Repeat Analysis
    mobile_col_candidates = [col for col in df.columns if 'MOBILE' in col or 'PHONE' in col or 'CONTACT' in col]
    if mobile_col_candidates:
        df['MOBILE_NO'] = df[mobile_col_candidates[0]].astype(str).str.replace(r'\D', '', regex=True)
    else:
        df['MOBILE_NO'] = 'UNKNOWN'

    # Create simplified Status category (Pending vs Closed)
    df['PENDING_STATUS'] = df['STS_CLEAN'].apply(lambda x: 'Closed' if x == 'CLOSED' else 'Pending')

    # Robust Date Parsing
    for orig_col, dt_col in [('REGISTRATION_DATE', 'REGISTRATION_DT'), ('CLOSINGDATE', 'CLOSING_DT')]:
        if orig_col in df.columns:
            num_series = pd.to_numeric(df[orig_col], errors='coerce')
            if num_series.notna().sum() > (len(df) * 0.5):
                df[dt_col] = pd.to_datetime(num_series, unit='D', origin='1899-12-30', errors='coerce')
            else:
                df[dt_col] = pd.to_datetime(df[orig_col], errors='coerce')
        else:
            df[dt_col] = pd.NaT
        
    # Calculate Resolution Time in Hours safely
    valid_dates = df['CLOSING_DT'].notna() & df['REGISTRATION_DT'].notna()
    df['RESOLUTION_TIME_HRS'] = pd.NA
    df.loc[valid_dates, 'RESOLUTION_TIME_HRS'] = (
        (df.loc[valid_dates, 'CLOSING_DT'] - df.loc[valid_dates, 'REGISTRATION_DT'])
        .dt.total_seconds() / 3600.0
    )
    
    # Extract date and hour attributes for trend analysis
    df['REG_DATE'] = df['REGISTRATION_DT'].dt.date
    df['REG_HOUR'] = df['REGISTRATION_DT'].dt.hour
    
    return df

# Load Dataset
try:
    df = load_and_process_data('june_data.xlsx')
except Exception as e:
    st.error(f"Error loading 'june_data.xlsx': {e}. Please ensure the file is in the same directory.")
    st.stop()

# -------------------------------------------------------------
# 4. RANDOM ALLOCATION HELPER
# -------------------------------------------------------------
def pick_random_complaint():
    """Picks a random complaint that has not been stored in SQLite yet."""
    completed = get_completed_complaint_numbers()
    pending_pool = df[~df['COMPLAINT_NO'].astype(str).isin(completed)]
    if pending_pool.empty:
        return None
    return pending_pool.sample(n=1).iloc[0].to_dict()

# Initialize session state complaint assignment
if 'current_complaint' not in st.session_state:
    st.session_state.current_complaint = pick_random_complaint()

# -------------------------------------------------------------
# 5. EXCEL REPORT GENERATOR
# -------------------------------------------------------------
@st.cache_data
def generate_zone_excel_report(data_df):
    """Generates an Excel workbook containing multiple zone-wise analysis tabs."""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Zone Performance KPI Summary
        if 'ZONE' in data_df.columns:
            zone_summary = data_df.groupby('ZONE').agg(
                Total_Complaints=('STS_CLEAN', 'count'),
                Closed_Complaints=('STS_CLEAN', lambda x: (x == 'CLOSED').sum()),
                Pending_Complaints=('STS_CLEAN', lambda x: (x != 'CLOSED').sum()),
                Beyond_Time_Pendency=('STS_CLEAN', lambda x: x.str.contains('BEYOND|OVERDUE', regex=True).sum()),
                Avg_Resolution_Time_Hrs=('RESOLUTION_TIME_HRS', lambda x: pd.to_numeric(x, errors='coerce').mean())
            ).reset_index()
            
            zone_summary['Pending_Rate (%)'] = (zone_summary['Pending_Complaints'] / zone_summary['Total_Complaints'] * 100).round(2)
            zone_summary['Avg_Resolution_Time_Hrs'] = zone_summary['Avg_Resolution_Time_Hrs'].round(1)
            zone_summary.to_excel(writer, sheet_name='Zone_KPI_Summary', index=False)
        
        # Sheet 2: Zone vs Complaint Type Breakdown Matrix
        if 'ZONE' in data_df.columns and 'COMPLAINT_TYPE' in data_df.columns:
            zone_type_pivot = pd.crosstab(
                data_df['ZONE'], 
                data_df['COMPLAINT_TYPE'], 
                margins=True, 
                margins_name='Total'
            )
            zone_type_pivot.to_excel(writer, sheet_name='Zone_Complaint_Matrix')

        # Sheet 3: Repeat Consumer Summary Zone-Wise
        valid_mobiles = data_df[~data_df['MOBILE_NO'].isin(['UNKNOWN', '', 'NAN', 'NONE'])]
        if not valid_mobiles.empty and 'ZONE' in valid_mobiles.columns:
            repeat_counts = valid_mobiles.groupby(['MOBILE_NO', 'ZONE']).size().reset_index(name='Tickets')
            repeat_only = repeat_counts[repeat_counts['Tickets'] > 1]
            if not repeat_only.empty:
                repeat_summary = repeat_only.groupby('ZONE').agg(
                    Repeat_Consumers=('MOBILE_NO', 'count'),
                    Total_Repeat_Tickets=('Tickets', 'sum')
                ).reset_index()
                repeat_summary.to_excel(writer, sheet_name='Repeat_Consumers_By_Zone', index=False)

    return output.getvalue()

# -------------------------------------------------------------
# 6. SIDEBAR INTERACTIVE FILTERS & AGENT SESSION
# -------------------------------------------------------------
st.sidebar.title("⚡ Session & Filters")

agent_id = st.sidebar.text_input("OFFICER ID / Name:", value="CallCenter_Agent")
st.sidebar.divider()

st.sidebar.subheader("🔍 Analysis Filters")
zones = ['All'] + sorted([str(z) for z in df['ZONE'].dropna().unique()]) if 'ZONE' in df.columns else ['All']
selected_zone = st.sidebar.selectbox("Select Zone:", zones)

comp_types = ['All'] + sorted([str(ct) for ct in df['COMPLAINT_TYPE'].dropna().unique()]) if 'COMPLAINT_TYPE' in df.columns else ['All']
selected_type = st.sidebar.selectbox("Select Complaint Type:", comp_types)

sources = ['All'] + sorted([str(s) for s in df['COMPLAINT_SOURCE'].dropna().unique()]) if 'COMPLAINT_SOURCE' in df.columns else ['All']
selected_source = st.sidebar.selectbox("Select Complaint Source:", sources)

sts_options = ['All'] + sorted([str(st_val) for st_val in df['STS'].dropna().unique()])
selected_sts = st.sidebar.selectbox("Filter by Status (STS):", sts_options)

# Apply Dynamic Filters
filtered_df = df.copy()
if selected_zone != 'All' and 'ZONE' in filtered_df.columns:
    filtered_df = filtered_df[filtered_df['ZONE'] == selected_zone]
if selected_type != 'All' and 'COMPLAINT_TYPE' in filtered_df.columns:
    filtered_df = filtered_df[filtered_df['COMPLAINT_TYPE'] == selected_type]
if selected_source != 'All' and 'COMPLAINT_SOURCE' in filtered_df.columns:
    filtered_df = filtered_df[filtered_df['COMPLAINT_SOURCE'] == selected_source]
if selected_sts != 'All':
    filtered_df = filtered_df[filtered_df['STS'] == selected_sts]

# -------------------------------------------------------------
# 7. MAIN HEADER & TABBED LAYOUT
# -------------------------------------------------------------
header_col1, header_col2 = st.columns([1, 10])
with header_col1:
    try:
        st.image("mvvnl.jpeg", width=70)
    except Exception:
        pass
with header_col2:
    st.title("MVVNL Integrated Feedback & Complaint Analytics")

st.markdown("Automated random verification feedback system coupled with real-time complaint analytics.")

tab_feedback, tab_analytics = st.tabs(["📲 Mobile Verification Feedback", "📊 Executive Analytics Dashboard"])

# =============================================================
# TAB 1: RANDOM COMPLAINT FEEDBACK VERIFICATION
# =============================================================
with tab_feedback:
    st.subheader("⚡ Automated Verification System")
    
    completed_set = get_completed_complaint_numbers()
    total_records = len(df)
    completed_count = len(completed_set)
    remaining_count = max(0, total_records - completed_count)

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Dataset Complaints", f"{total_records:,}")
    m2.metric("Feedbacks Saved (SQLite)", f"{completed_count:,}")
    m3.metric("Pending Verification Pool", f"{remaining_count:,}")

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

        # Display Allotted Complaint & Consumer Remarks
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

        # Feedback Input Form
        st.subheader("✍️ Record Feedback")
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
                st.text_input("Recording Agent:", value=agent_id, disabled=True)

            feedback_remark = st.text_area(
                "Consumer Feedback Remarks / Comments *",
                placeholder="e.g., Consumer confirmed supply restored properly OR Consumer mentioned transformer is still burnt.",
                height=100
            )

            submit_button = st.form_submit_button(label="💾 Submit Feedback & Load Next Complaint", type="primary", use_container_width=True)

        if submit_button:
            if not feedback_remark.strip():
                st.error("⚠️ Please enter a feedback remark before submitting.")
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
                    "agent_id": agent_id
                }
                
                if save_feedback_sqlite(record):
                    st.toast(f"✅ Feedback for ticket {current_ticket.get('COMPLAINT_NO')} stored in SQLite!")
                    st.session_state.current_complaint = pick_random_complaint()
                    st.rerun()

    # Feedback Database Export Section
    st.divider()
    st.subheader("📊 Submitted Feedback Audit Log")
    conn = get_db_connection()
    fb_df = pd.read_sql_query("SELECT * FROM consumer_feedback ORDER BY id DESC", conn)
    conn.close()

    if not fb_df.empty:
        fb_c1, fb_c2, fb_c3, fb_c4 = st.columns(4)
        fb_c1.metric("Total Submitted", len(fb_df))
        fb_c2.metric("Satisfied", len(fb_df[fb_df['feedback_status'] == 'Satisfied']))
        fb_c3.metric("Not Satisfied", len(fb_df[fb_df['feedback_status'] == 'Not Satisfied']))
        fb_c4.metric("Wrongly Closed ⚠️", len(fb_df[fb_df['feedback_status'] == 'Wrongly Closed']))
        
        st.dataframe(fb_df, use_container_width=True)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            fb_df.to_excel(writer, sheet_name='Feedback_Audit', index=False)
            
        st.download_button(
            label="📥 Download Full SQLite Feedback Log (Excel)",
            data=buffer.getvalue(),
            file_name=f"Transformer_Feedback_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No feedback entries recorded in SQLite database yet.")

# =============================================================
# TAB 2: ANALYTICS DASHBOARD
# =============================================================
with tab_analytics:
    # Top Metrics
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
        st.metric("Total Pending Complaints", f"{pending_count:,}", delta=f"{pending_pct:.1f}% Pending", delta_color="inverse")

    with col4:
        is_beyond_time = filtered_df['STS_CLEAN'].str.contains('BEYOND|OVERDUE', regex=True)
        beyond_count = is_beyond_time.sum() if not filtered_df.empty else 0
        beyond_pct = (beyond_count / len(filtered_df) * 100) if len(filtered_df) > 0 else 0
        st.metric("Beyond Time Pendency", f"{beyond_count:,} Tickets", delta=f"{beyond_pct:.1f}% Share", delta_color="inverse")

    st.write("")

    # Report Downloader Section
    with st.container():
        rep_col1, rep_col2 = st.columns([3, 1])
        with rep_col1:
            st.markdown(
                f"### 📊 Zone-Wise Analytical Executive Report\n"
                f"Download multi-tab Excel summary including KPI metrics, complaint type distribution matrix, and repeat caller analysis for **Zone: {selected_zone}**."
            )
        with rep_col2:
            try:
                excel_data = generate_zone_excel_report(filtered_df)
                st.download_button(
                    label="📥 Download Zone Excel Report",
                    data=excel_data,
                    file_name=f"MVVNL_Zone_Analytical_Report_{selected_zone}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
            except Exception as e:
                st.error(f"Error preparing report: {e}")

    st.divider()

    # Charts Row 1
    row1_col1, row1_col2 = st.columns(2)

    with row1_col1:
        st.subheader("📌 Zone-Wise Complaint Trajectory")
        if 'ZONE' in filtered_df.columns:
            zone_traj = filtered_df.groupby(['REG_DATE', 'ZONE']).size().reset_index(name='Complaint_Count')
            fig_zone = px.line(
                zone_traj, x='REG_DATE', y='Complaint_Count', color='ZONE', markers=True,
                title="Daily Complaint Trajectory Across Zones",
                labels={'REG_DATE': 'Registration Date', 'Complaint_Count': 'Number of Complaints', 'ZONE': 'Zone'}
            )
            fig_zone.update_layout(template="plotly_white", hovermode="x unified")
            st.plotly_chart(fig_zone, use_container_width=True)
        else:
            st.info("Zone information unavailable for trajectory chart.")

    with row1_col2:
        st.subheader("📊 Comparative Pendency by Complaint Type")
        if 'COMPLAINT_TYPE' in filtered_df.columns:
            type_status_df = filtered_df.groupby(['COMPLAINT_TYPE', 'PENDING_STATUS']).size().reset_index(name='Count')
            fig_type_pendency = px.bar(
                type_status_df, y='COMPLAINT_TYPE', x='Count', color='PENDING_STATUS', barmode='group', orientation='h',
                title="Pending vs. Closed Complaints by Type",
                labels={'COMPLAINT_TYPE': 'Complaint Type', 'Count': 'Total Complaints', 'PENDING_STATUS': 'Status'},
                color_discrete_map={'Pending': '#EF553B', 'Closed': '#00CC96'},
                text_auto=True
            )
            fig_type_pendency.update_layout(template="plotly_white", yaxis={'categoryorder': 'total ascending'}, margin=dict(l=150, r=20, t=40, b=20))
            st.plotly_chart(fig_type_pendency, use_container_width=True)

    st.divider()

    # Charts Row 2
    row2_col1, row2_col2 = st.columns(2)

    with row2_col1:
        st.subheader("⏱️ Complaint Type-Wise Avg Closure Time")
        if 'COMPLAINT_TYPE' in filtered_df.columns:
            temp_df = filtered_df.copy()
            temp_df['RESOLUTION_TIME_HRS'] = pd.to_numeric(temp_df['RESOLUTION_TIME_HRS'], errors='coerce')
            avg_type = temp_df.groupby('COMPLAINT_TYPE')['RESOLUTION_TIME_HRS'].mean().reset_index().sort_values('RESOLUTION_TIME_HRS', ascending=True)
            
            fig_type = px.bar(
                avg_type, x='RESOLUTION_TIME_HRS', y='COMPLAINT_TYPE', orientation='h',
                title="Average Closure Time by Type (Hours)",
                labels={'RESOLUTION_TIME_HRS': 'Avg Closure Time (Hrs)', 'COMPLAINT_TYPE': 'Complaint Type'},
                color='RESOLUTION_TIME_HRS', color_continuous_scale="Reds", text_auto='.1f'
            )
            fig_type.update_layout(template="plotly_white", margin=dict(l=150, r=20, t=40, b=20))
            st.plotly_chart(fig_type, use_container_width=True)

    with row2_col2:
        st.subheader("🏢 Top Substations with Max Complaints")
        if 'SUBSTATION' in filtered_df.columns and 'ZONE' in filtered_df.columns:
            sub_df = filtered_df.groupby(['SUBSTATION', 'ZONE']).size().reset_index(name='Complaint_Count').sort_values('Complaint_Count', ascending=False).head(10)
            fig_sub = px.bar(
                sub_df, x='Complaint_Count', y='SUBSTATION', color='ZONE', orientation='h',
                title="Top 10 High-Volume Substations",
                labels={'Complaint_Count': 'Total Complaints', 'SUBSTATION': 'Substation', 'ZONE': 'Zone'},
                text_auto=True
            )
            fig_sub.update_layout(template="plotly_white", yaxis={'categoryorder': 'total ascending'}, margin=dict(l=150, r=20, t=40, b=20))
            st.plotly_chart(fig_sub, use_container_width=True)

    st.divider()

    # Charts Row 3
    row3_col1, row3_col2 = st.columns(2)

    with row3_col1:
        st.subheader("📲 Complaint Source Breakdown")
        if 'COMPLAINT_SOURCE' in filtered_df.columns:
            source_df = filtered_df.groupby('COMPLAINT_SOURCE').size().reset_index(name='Count').sort_values('Count', ascending=False)
            fig_source = px.bar(
                source_df, x='COMPLAINT_SOURCE', y='Count', color='COMPLAINT_SOURCE',
                title="Complaint Distribution by Channel/Source",
                labels={'COMPLAINT_SOURCE': 'Complaint Source', 'Count': 'Total Complaints'},
                text_auto=True
            )
            fig_source.update_layout(template="plotly_white", showlegend=False)
            st.plotly_chart(fig_source, use_container_width=True)

    with row3_col2:
        st.subheader("🔄 Repeated Consumer Complaints Zone-Wise")
        valid_mobile_df = filtered_df[~filtered_df['MOBILE_NO'].isin(['UNKNOWN', '', 'NAN', 'NONE'])]
        
        if not valid_mobile_df.empty and 'ZONE' in valid_mobile_df.columns:
            consumer_counts = valid_mobile_df.groupby(['MOBILE_NO', 'ZONE']).size().reset_index(name='Complaint_Count')
            repeated_df = consumer_counts[consumer_counts['Complaint_Count'] > 1]
            
            zone_repeat_summary = (
                repeated_df.groupby('ZONE')
                .agg(Repeat_Consumers=('MOBILE_NO', 'count'), Total_Repeat_Complaints=('Complaint_Count', 'sum'))
                .reset_index().sort_values('Total_Repeat_Complaints', ascending=False)
            )
            
            fig_repeat = px.bar(
                zone_repeat_summary, x='ZONE', y=['Repeat_Consumers', 'Total_Repeat_Complaints'], barmode='group',
                title="Repeat Consumers (>1 Ticket) and Complaints by Zone",
                labels={'value': 'Count', 'ZONE': 'Zone', 'variable': 'Metric'},
                text_auto=True
            )
            fig_repeat.update_layout(template="plotly_white")
            st.plotly_chart(fig_repeat, use_container_width=True)
        else:
            st.info("No valid mobile number or zone details available for repeat consumer analysis.")

    st.divider()

    # Raw Data Explorer
    st.subheader("📋 Raw Data Explorer")
    columns_to_show = [
        'COMPLAINT_NO', 'ZONE', 'SUBSTATION', 'COMPLAINT_TYPE', 
        'COMPLAINT_SOURCE', 'MOBILE_NO', 'REGISTRATION_DT', 'CLOSING_DT', 
        'RESOLUTION_TIME_HRS', 'STS'
    ]
    available_cols = [col for col in columns_to_show if col in filtered_df.columns]
    st.dataframe(filtered_df[available_cols], use_container_width=True)