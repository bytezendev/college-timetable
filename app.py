import streamlit as st
import pandas as pd
import json
import os
import io
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

# --- PAGE CONFIGURATION (Mobile & Desktop Responsive) ---
st.set_page_config(page_title="College Timetable Portal", page_icon="📅", layout="wide")

# --- CONSTANTS & CONFIGURATION ---
DRIVE_FOLDER_ID = "1z51WeH5-2UiyAxLw4t9ne1i6uRGhauav"
DB_FILE = "college_portal_db.json"
CLASSROOMS = ["Room 31 (12th-A)", "Room 32 (12th-B)", "Room 33 (11th-A)", "Room 34 (11th-B)"]

# --- LOCAL / CLOUD STATE DATABASE (Users, Staff, Remarks) ---
def load_db():
    if not os.path.exists(DB_FILE):
        default_data = {
            "users": {
                "admin": {"password": "123", "role": "Admin", "subject": "Administration"},
                "dev": {"password": "123", "role": "Developer", "subject": "System"},
                "rahul": {"password": "123", "role": "Teacher", "subject": "Physics"},
                "priya": {"password": "123", "role": "Teacher", "subject": "Chemistry"},
                "amit": {"password": "123", "role": "Teacher", "subject": "Maths"},
                "sneha": {"password": "123", "role": "Teacher", "subject": "Biology"},
                "vikram": {"password": "123", "role": "Teacher", "subject": "English"},
                "neha": {"password": "123", "role": "Teacher", "subject": "IT / CS"}
            },
            "remarks": []
        }
        save_db(default_data)
        return default_data
    with open(DB_FILE, "r") as f:
        return json.load(f)

def save_db(data):
    with open(DB_FILE, "w") as f:
        json.dump(data, f, indent=4)

# --- GOOGLE DRIVE API CONNECTION ---
def get_drive_service():
    try:
        # Loads credentials from Streamlit Secrets
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=['https://www.googleapis.com/auth/drive']
        )
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        return None

def save_excel_to_drive(df, date_str):
    service = get_drive_service()
    filename = f"{date_str}.xlsx"
    
    # Convert DataFrame to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Timetable")
    output.seek(0)

    if not service:
        # Fallback local save if Drive secrets aren't configured yet
        df.to_excel(filename, index=False)
        return False, f"Saved locally as {filename} (Configure Drive Secrets for Cloud Sync)"

    # Check if file for this date already exists in the folder (for updating past records)
    query = f"name='{filename}' and '{DRIVE_FOLDER_ID}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])

    media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    if files:
        # Update existing Excel sheet
        file_id = files[0]['id']
        service.files().update(fileId=file_id, media_body=media).execute()
        return True, f"Updated existing record '{filename}' on Google Drive!"
    else:
        # Create new Excel sheet
        file_metadata = {'name': filename, 'parents': [DRIVE_FOLDER_ID]}
        service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return True, f"Saved new record '{filename}' to Google Drive!"

def load_excel_from_drive(date_str):
    service = get_drive_service()
    filename = f"{date_str}.xlsx"
    if not service:
        if os.path.exists(filename):
            return pd.read_excel(filename)
        return None

    query = f"name='{filename}' and '{DRIVE_FOLDER_ID}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])

    if not files:
        return None

    request = service.files().get_media(fileId=files[0]['id'])
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return pd.read_excel(fh)

# --- INITIALIZE SESSION ---
db = load_db()
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""

# ==========================================
# 1. LOGIN & PASSWORD SCREEN
# ==========================================
if not st.session_state.logged_in:
    st.markdown("### 🎓 College Daily Timetable Portal")
    st.caption("Login using your First Name")
    
    with st.form("login_form"):
        first_name = st.text_input("First Name (e.g., admin, rahul)").strip().lower()
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Login", use_container_width=True)
        
        if submit:
            if first_name in db["users"] and db["users"][first_name]["password"] == password:
                st.session_state.logged_in = True
                st.session_state.username = first_name
                st.session_state.role = db["users"][first_name]["role"]
                st.rerun()
            else:
                st.error("Invalid First Name or Password.")
    st.stop()

# ==========================================
# TOP NAVIGATION BAR
# ==========================================
col1, col2 = st.columns([3, 1])
with col1:
    st.markdown(f"#### 👋 Welcome, **{st.session_state.username.capitalize()}** ({st.session_state.role})")
with col2:
    if st.button("Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

# ==========================================
# TABS (Simple, Clean UI for Mobile & Desktop)
# ==========================================
if st.session_state.role in ["Admin", "Developer"]:
    tabs = st.tabs(["📅 View Timetable", "🛠️ Create / Edit Schedule", "👥 Manage Staff", "💬 Remarks"])
else:
    tabs = st.tabs(["📅 Today's Timetable", "💬 Send Remark", "🔑 Change Password"])

# ------------------------------------------
# TAB 1: VIEW TIMETABLE (Everyone)
# ------------------------------------------
with tabs[0]:
    view_date = st.date_input("Select Date to View", value=datetime.today() + timedelta(days=1))
    date_str = view_date.strftime("%Y-%m-%d")
    day_name = view_date.strftime("%A")
    
    st.info(f"📆 **{day_name}, {date_str}** | File: `{date_str}.xlsx`")
    df_view = load_excel_from_drive(date_str)
    
    if df_view is not None:
        st.dataframe(df_view, use_container_width=True, hide_index=True)
        
        # Download button for offline viewing
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_view.to_excel(writer, index=False)
        st.download_button(
            label="📥 Download Excel Sheet",
            data=output.getvalue(),
            file_name=f"{date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning(f"No timetable published yet for {date_str}.")

# ------------------------------------------
# ADMIN / DEVELOPER VIEWS
# ------------------------------------------
if st.session_state.role in ["Admin", "Developer"]:
    
    # TAB 2: CREATE OR UPDATE PAST/FUTURE TIMETABLE
    with tabs[1]:
        st.markdown("##### 1️⃣ Select Date & Shift Mode")
        c1, c2, c3 = st.columns(3)
        with c1:
            sched_date = st.date_input("Schedule Date", value=datetime.today() + timedelta(days=1), key="edit_date")
            sched_date_str = sched_date.strftime("%Y-%m-%d")
            sched_day = sched_date.strftime("%A")
        with c2:
            shift_mode = st.selectbox("College Timing Mode", ["Morning College (8 Lectures)", "Day College (6 Lectures)", "Custom"])
        with c3:
            num_lectures = st.number_input("Number of Lectures", min_value=1, max_value=10, value=8 if "8" in shift_mode else 6)

        st.markdown("##### 2️⃣ Mark Present Teachers for Tomorrow")
        all_teachers = [
            f"{name.capitalize()} ({info['subject']})" 
            for name, info in db["users"].items() 
            if info["role"] == "Teacher"
        ]
        present_teachers = st.multiselect(
            "Select Present Teachers (Unselected = On Leave)",
            options=all_teachers,
            default=all_teachers
        )
        dropdown_options = ["--- Free / Break ---"] + present_teachers

        st.markdown(f"##### 3️⃣ Assign Lectures ({sched_day}, {sched_date_str})")
        
        # Load existing record if editing a past date
        existing_df = load_excel_from_drive(sched_date_str)
        
        default_times_morning = ["07:30-08:15", "08:15-09:00", "09:00-09:45", "10:00-10:45", "10:45-11:30", "11:30-12:15", "12:15-01:00", "01:00-01:45"]
        default_times_day = ["10:30-11:20", "11:20-12:10", "12:10-01:00", "01:30-02:20", "02:20-03:10", "03:10-04:00"]
        base_times = default_times_morning if "Morning" in shift_mode else default_times_day

        rows = []
        for i in range(int(num_lectures)):
            default_t = base_times[i] if i < len(base_times) else f"Lec {i+1}"
            row_data = {"Day": sched_day, "Date": sched_date_str, "Lecture": f"Lecture {i+1}", "Time": default_t}
            
            # Pre-fill if past record exists
            if existing_df is not None and i < len(existing_df):
                row_data["Time"] = existing_df.loc[i, "Time"]
                for room in CLASSROOMS:
                    val = existing_df.loc[i, room] if room in existing_df.columns else "--- Free / Break ---"
                    row_data[room] = val if val in dropdown_options else "--- Free / Break ---"
            else:
                for room in CLASSROOMS:
                    row_data[room] = "--- Free / Break ---"
            rows.append(row_data)

        df_template = pd.DataFrame(rows)

        # Interactive Grid Editor (Works cleanly on Mobile & Desktop)
        column_cfg = {
            "Day": st.column_config.TextColumn(disabled=True),
            "Date": st.column_config.TextColumn(disabled=True),
            "Lecture": st.column_config.TextColumn(disabled=True),
            "Time": st.column_config.TextColumn("Time Slot (Editable)"),
        }
        for room in CLASSROOMS:
            column_cfg[room] = st.column_config.SelectboxColumn(room, options=dropdown_options, required=True)

        edited_df = st.data_editor(df_template, column_config=column_cfg, use_container_width=True, hide_index=True)

        # Clash Detection Check
        clashes = []
        for idx, row in edited_df.iterrows():
            assigned = [row[r] for r in CLASSROOMS if row[r] != "--- Free / Break ---"]
            if len(assigned) != len(set(assigned)):
                clashes.append(f"{row['Lecture']} ({row['Time']})")

        if clashes:
            st.error(f"⚠️ Teacher Clash Detected in: {', '.join(clashes)}. Same teacher is in two rooms!")

        if st.button("💾 Save & Sync to Google Drive", type="primary", use_container_width=True):
            success, msg = save_excel_to_drive(edited_df, sched_date_str)
            if success:
                st.success(msg)
            else:
                st.warning(msg)

    # TAB 3: MANAGE STAFF & ACCOUNTS
    with tabs[2]:
        st.markdown("##### ➕ Add New Member / Teacher")
        with st.form("add_member_form"):
            new_name = st.text_input("First Name (Used for Login)").strip().lower()
            new_pass = st.text_input("Default Password", value="123")
            new_role = st.selectbox("Role", ["Teacher", "Admin", "Developer"])
            new_sub = st.text_input("Subject (e.g., Physics, Maths)")
            if st.form_submit_button("Create Account"):
                if new_name:
                    db["users"][new_name] = {"password": new_pass, "role": new_role, "subject": new_sub}
                    save_db(db)
                    st.success(f"Account created for {new_name.capitalize()}!")
                    st.rerun()

        st.markdown("##### ✏️ Edit or Remove Existing Staff")
        for uname, uinfo in list(db["users"].items()):
            with st.expander(f"{uname.capitalize()} — {uinfo['role']} ({uinfo['subject']})"):
                c1, c2, c3 = st.columns([2, 2, 1])
                updated_name = c1.text_input("First Name", value=uname, key=f"name_{uname}").strip().lower()
                updated_sub = c2.text_input("Subject", value=uinfo["subject"], key=f"sub_{uname}")
                
                if c1.button("Update", key=f"upd_{uname}"):
                    if updated_name != uname:
                        db["users"][updated_name] = db["users"].pop(uname)
                    db["users"][updated_name]["subject"] = updated_sub
                    save_db(db)
                    st.success("Updated!")
                    st.rerun()
                
                if uname not in ["admin", "dev"]:
                    if c3.button("🗑️ Remove", key=f"del_{uname}"):
                        del db["users"][uname]
                        save_db(db)
                        st.warning(f"Removed {uname}")
                        st.rerun()

    # TAB 4: ADMIN REMARKS VIEW
    with tabs[3]:
        st.markdown("##### 📩 Teacher Queries & Remarks")
        if not db["remarks"]:
            st.info("No remarks yet.")
        else:
            for idx, r in enumerate(reversed(db["remarks"])):
                st.markdown(f"**{r['sender'].capitalize()}** ({r['time']}): {r['text']}")
            if st.button("Clear All Remarks"):
                db["remarks"] = []
                save_db(db)
                st.rerun()

# ------------------------------------------
# MEMBER / TEACHER VIEWS
# ------------------------------------------
else:
    # TAB 2: SEND REMARK
    with tabs[1]:
        st.markdown("##### 💬 Send Query / Leave Request to Admin")
        remark_text = st.text_area("Write your message (e.g., 'I will be on leave tomorrow' or 'Need proxy for Lec 3')")
        if st.button("Send Remark", type="primary"):
            if remark_text.strip():
                db["remarks"].append({
                    "sender": st.session_state.username,
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "text": remark_text.strip()
                })
                save_db(db)
                st.success("Message sent to Admin!")

    # TAB 3: CHANGE OWN PASSWORD
    with tabs[2]:
        st.markdown("##### 🔑 Set Your Own Password")
        new_pw = st.text_input("Enter New Password", type="password")
        confirm_pw = st.text_input("Confirm New Password", type="password")
        if st.button("Update Password"):
            if new_pw and new_pw == confirm_pw:
                db["users"][st.session_state.username]["password"] = new_pw
                save_db(db)
                st.success("Password updated successfully!")
            else:
                st.error("Passwords do not match.")