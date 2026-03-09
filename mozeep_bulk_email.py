import streamlit as st
import pandas as pd
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(page_title="Mozeep Bulk Email Tool", layout="wide")
st.title("🚀 Mozeep Auto Send Tool (CV Fail Reply)")
st.caption("Only change name/school/etc. — everything else is automatic")

# ====================== SMTP SETTINGS (Mozeep) ======================
st.sidebar.header("Mozeep Account")
smtp_server = st.sidebar.text_input("SMTP Server", value="mail.mozeep.com")
port = st.sidebar.number_input("Port", value=465)
your_email = st.sidebar.text_input("Your Mozeep Email (full address)")
password = st.sidebar.text_input("Password / App Password", type="password")

# ====================== UPLOAD CONTACT LIST ======================
st.header("1. Upload your contact list")
uploaded_file = st.file_uploader("Upload Excel or CSV (must have columns: email, name, ...)", 
                               type=["csv", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    
    # Add checkbox column
    df["Send?"] = False
    edited_df = st.data_editor(
        df,
        hide_index=True,
        column_config={
            "Send?": st.column_config.CheckboxColumn("Send?", default=False),
            "email": st.column_config.TextColumn("Email", disabled=True),
            "name": st.column_config.TextColumn("Name", disabled=True),
        },
        use_container_width=True
    )
    
    selected = edited_df[edited_df["Send?"] == True].copy()
    st.write(f"✅ Selected: **{len(selected)}** emails")

    # ====================== EMAIL TEMPLATE ======================
    st.header("2. Email Template")
    subject = st.text_input("Subject line", value="Thông báo kết quả ứng tuyển")
    template = st.text_area(
        "Template (use {name}, {school}, {position}... exactly as in your columns)",
        height=300,
        value="""Kính gửi {name},

Chúng tôi đã xem xét hồ sơ ứng tuyển của bạn cho vị trí tại {school}.

Sau khi đánh giá, chúng tôi rất tiếc phải thông báo rằng hồ sơ của bạn chưa phù hợp với yêu cầu ở vòng này.

Chúc bạn thành công trong những cơ hội tiếp theo!

Trân trọng,
Tên của bạn"""
    )

    # ====================== SEND BUTTON ======================
    if st.button("🚀 SEND SELECTED EMAILS", type="primary"):
        if not your_email or not password:
            st.error("Please enter your Mozeep email and password in sidebar")
        elif len(selected) == 0:
            st.warning("No emails selected")
        else:
            try:
                if port == 465:
                    server = smtplib.SMTP_SSL(smtp_server, port)
                else:
                    server = smtplib.SMTP(smtp_server, port)
                    server.starttls()
                server.login(your_email, password)
                
                progress_bar = st.progress(0)
                total = len(selected)
                for idx, (_, row) in enumerate(selected.iterrows()):
                    try:
                        msg = MIMEMultipart()
                        msg['From'] = your_email
                        msg['To'] = row['email']
                        msg['Subject'] = subject
                        
                        body = template.format(**row)
                        msg.attach(MIMEText(body, 'plain'))
                        
                        server.send_message(msg)
                        
                        st.success(f"✅ Sent to {row['name']} ({row['email']})")
                        progress_bar.progress((idx + 1) / total)
                        time.sleep(2)  # avoid rate limit
                    except Exception as e:
                        st.error(f"Failed {row['email']}: {e}")
                
                server.quit()
                st.balloons()
                st.success("🎉 All done!")
                
            except Exception as e:
                st.error(f"SMTP Error: {e}")