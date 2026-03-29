import imaplib
import html
import re
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

import pandas as pd
import streamlit as st

try:
    from streamlit_quill import st_quill

    QUILL_AVAILABLE = True
except ImportError:
    QUILL_AVAILABLE = False


DEFAULT_TEMPLATE_TEXT = """Kính gửi {name},

Chúng tôi đã xem xét hồ sơ ứng tuyển của bạn cho vị trí {position} tại {company}.

Sau khi đánh giá, chúng tôi rất tiếc phải thông báo rằng hồ sơ của bạn chưa phù hợp với yêu cầu ở vòng này.

Chúc bạn thành công trong những cơ hội tiếp theo!"""

DEFAULT_TEMPLATE_HTML = """<p>Kính gửi {name},</p>
<p>Chúng tôi đã xem xét hồ sơ ứng tuyển của bạn cho vị trí <strong>{position}</strong> tại <strong>{company}</strong>.</p>
<p>Sau khi đánh giá, chúng tôi rất tiếc phải thông báo rằng hồ sơ của bạn chưa phù hợp với yêu cầu ở vòng này.</p>
<p>Chúc bạn thành công trong những cơ hội tiếp theo!</p>"""

DEFAULT_SIGNATURE_TEXT = """Thanks and Best regards
---
DUONG DO PHUONG THUY | TALENT ACQUISITION
thuyddp@lighthuman.vn / 0888.652.083
LIGHT HUMAN SOLUTIONS JOIN STOCK COMPANY
Website: lighthuman.vn
Hotline: 028 3842 5188
Address: 213 Chu Van An Street, Binh Thanh District, Ho Chi Minh City."""

DEFAULT_SIGNATURE_HTML = """
<div>---</div>
<div>Thanks and Best regards</div>
<div>---</div>
<div><strong>DUONG DO PHUONG THUY | TALENT ACQUISITION</strong></div>
<div>thuyddp@lighthuman.vn / 0888.652.083</div>
<div><strong>LIGHT HUMAN SOLUTIONS JOIN STOCK COMPANY</strong></div>
<div>Website: <a href="https://lighthuman.vn">lighthuman.vn</a></div>
<div>Hotline: 028 3842 5188</div>
<div>Address: 213 Chu Van An Street, Binh Thanh District, Ho Chi Minh City.</div>"""


def render_template_with_row(template_text, row_dict):
    """Replace {column_name} placeholders only, preserving other braces (e.g. CSS)."""
    pattern = re.compile(r"\{([a-zA-Z0-9_]+)\}")

    def replace_match(match):
        key = match.group(1)
        value = row_dict.get(key)
        if value is None or pd.isna(value):
            return ""
        return str(value)

    return pattern.sub(replace_match, template_text)


def load_uploaded_contacts(uploaded_file):
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except ImportError:
        st.error(
            "Excel support requires the 'openpyxl' package. "
            "Install it with: pip install openpyxl"
        )
        return None
    except Exception as err:
        st.error(f"Could not read uploaded file: {err}")
        return None

    if "email" not in df.columns:
        st.error("Your file must contain an 'email' column.")
        return None

    return df


def render_contact_selector(df):
    editable_df = df.copy()
    editable_df["Send?"] = False

    edited_df = st.data_editor(
        editable_df,
        hide_index=True,
        column_config={
            "Send?": st.column_config.CheckboxColumn("Send?", default=False),
            "email": st.column_config.TextColumn("Email", disabled=True),
            "name": st.column_config.TextColumn("Name", disabled=True),
            "company": st.column_config.TextColumn("Company", disabled=True),
        },
        use_container_width=True,
    )

    selected = edited_df[edited_df["Send?"] == True].copy()
    st.write(f"✅ Selected: **{len(selected)}** emails")

    available_placeholders = ", ".join(
        [f"{{{col}}}" for col in df.columns if col != "Send?"]
    )
    st.caption(f"Available placeholders from your sheet: {available_placeholders}")
    return selected


def render_template_editor():
    st.markdown(
        """
        <style>
        .ql-editor {
            line-height: 1.2 !important;
        }
        .ql-editor p,
        .ql-editor div {
            margin-top: 0 !important;
            margin-bottom: 0 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.header("2. Email Template")
    subject = st.text_input("Subject line", value="Thông báo kết quả ứng tuyển")
    use_rich_text = st.checkbox(
        "Use StreamQuill rich text editor",
        value=True,
        disabled=not QUILL_AVAILABLE,
    )

    if not QUILL_AVAILABLE:
        st.info("Install 'streamlit-quill' to enable rich text editing. Falling back to plain text.")

    attach_signature = st.checkbox("Attach signature", value=True)
    signature_template = ""

    if use_rich_text and QUILL_AVAILABLE:
        st.markdown("**Email body**")
        body_template = st_quill(
            value=DEFAULT_TEMPLATE_HTML,
            html=True,
            preserve_whitespace=False,
            key="body_template_quill",
        )

        if attach_signature:
            st.markdown("**Signature**")
            signature_template = st_quill(
                value=DEFAULT_SIGNATURE_HTML,
                html=True,
                preserve_whitespace=False,
                key="signature_template_quill",
            )
    else:
        body_template = st.text_area(
            "Template (use {name}, {school}, {position}... exactly as in your columns)",
            height=250,
            value=DEFAULT_TEMPLATE_TEXT,
        )

        if attach_signature:
            signature_template = st.text_area(
                "Signature template (supports placeholders like {name})",
                height=180,
                value=DEFAULT_SIGNATURE_TEXT,
            )

    return subject, use_rich_text, attach_signature, body_template, signature_template

def clean_html(html):
    if not html:
        return ""

    html = html.strip()

    # remove ALL empty paragraphs at the end
    html = re.sub(r"(<p><br></p>\s*)+$", "", html)

    return html


def html_to_plain_text(html_content):
    if not html_content:
        return ""

    text = re.sub(r"<br\s*/?>", "\n", html_content, flags=re.IGNORECASE)
    text = re.sub(r"</p>|</div>|</li>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<li[^>]*>", "- ", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)

    # Remove quote markers if any line still starts with '>'.
    lines = [re.sub(r"^\s*>\s?", "", line) for line in text.splitlines()]
    text = "\n".join(lines)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()

def create_message(
    row,
    sender_email,
    subject,
    body_template,
    signature_template,
    attach_signature,
    use_rich_text,
):
    row_dict = row.to_dict()
    body = render_template_with_row(body_template or "", row_dict)
    signature = ""

    if attach_signature:
        signature = render_template_with_row(signature_template or "", row_dict)

    if use_rich_text and QUILL_AVAILABLE:
        final_body_html = clean_html(body)
        if signature.strip():
            final_body_html = f"{final_body_html}{clean_html(signature)}"
        final_body_plain = html_to_plain_text(final_body_html)
    else:
        final_body_plain = body.strip()
        if signature.strip():
            final_body_plain = f"{final_body_plain}\n{signature.strip()}"
        final_body_plain = re.sub(r"\n{2,}", "\n", final_body_plain)
        final_body_html = None

    message = MIMEMultipart("alternative")
    message["From"] = sender_email
    message["To"] = row["email"]
    message["Subject"] = subject
    message["Date"] = formatdate(localtime=True)
    message.attach(MIMEText(final_body_plain, "plain", "utf-8"))
    if final_body_html is not None:
        message.attach(MIMEText(final_body_html, "html", "utf-8"))
    return message, row_dict


def connect_smtp(smtp_server, port, username, password):
    if port == 465:
        server = smtplib.SMTP_SSL(smtp_server, port)
    else:
        server = smtplib.SMTP(smtp_server, port)
        server.starttls()

    server.login(username, password)
    return server


def connect_imap_if_enabled(save_to_sent, imap_server, imap_port, email, password):
    if not save_to_sent:
        return None

    try:
        imap_conn = imaplib.IMAP4_SSL(imap_server, int(imap_port))
        imap_conn.login(email, password)
        return imap_conn
    except Exception as err:
        st.warning(f"IMAP login failed. Emails will still be sent, but not saved to Sent: {err}")
        return None


def append_to_sent(imap_conn, sent_folder, row_email, message):
    if not imap_conn:
        return

    try:
        status, _ = imap_conn.append(
            sent_folder,
            "\\Seen",
            imaplib.Time2Internaldate(time.time()),
            message.as_bytes(),
        )
        if status != "OK":
            st.warning(
                f"Sent but could not append to '{sent_folder}' for {row_email}. "
                "Check Sent folder name (for example: Sent, Sent Items, INBOX.Sent)."
            )
    except Exception as append_err:
        st.warning(f"Sent but failed to save in '{sent_folder}' for {row_email}: {append_err}")


def get_display_name(row_dict, row_email):
    recipient_name = row_dict.get("name")
    if recipient_name is None or pd.isna(recipient_name) or str(recipient_name).strip() == "":
        return row_email
    return recipient_name


def send_bulk_emails(
    selected,
    smtp_server,
    port,
    smtp_username,
    password,
    your_email,
    save_to_sent,
    imap_server,
    imap_port,
    sent_folder,
    subject,
    body_template,
    signature_template,
    attach_signature,
    use_rich_text,
):
    server = None
    imap_conn = None

    try:
        server = connect_smtp(smtp_server, port, smtp_username, password)
        imap_conn = connect_imap_if_enabled(
            save_to_sent,
            imap_server,
            imap_port,
            your_email,
            password,
        )

        progress_bar = st.progress(0)
        total = len(selected)

        for idx, (_, row) in enumerate(selected.iterrows()):
            try:
                msg, row_dict = create_message(
                    row,
                    your_email,
                    subject,
                    body_template,
                    signature_template,
                    attach_signature,
                    use_rich_text,
                )

                server.send_message(msg)
                append_to_sent(imap_conn, sent_folder, row["email"], msg)

                recipient_name = get_display_name(row_dict, row["email"])
                st.success(f"✅ Sent to {recipient_name} ({row['email']})")
                progress_bar.progress((idx + 1) / total)
                time.sleep(2)  # avoid rate limit
            except Exception as err:
                st.error(f"Failed {row['email']}: {err}")

        st.success("🎉 All done!")
    except Exception as err:
        st.error(f"SMTP Error: {err}")
    finally:
        if server:
            server.quit()
        if imap_conn:
            try:
                imap_conn.logout()
            except Exception:
                pass


def render_sidebar_settings():
    st.sidebar.header("Email Account (SendGrid)")
    st.sidebar.info(
        "📧 **SMTP Auth**: Thử email + password trước nha (bỏ trống mục SMTP username). Hong được thì xài SMTP username='apikey' + SendGrid API key (này thử đăng nhập bằng tk cty ở link dưới coi vô lấy đc ko, hong đc thì hỏi mấy anh IT =D). "
        "[Get API key](https://app.sendgrid.com/settings/api_keys)"
    )
    smtp_server = st.sidebar.text_input("SMTP Server", value="smtp.sendgrid.net")
    port = st.sidebar.number_input("SMTP Port", value=587)
    your_email = st.sidebar.text_input("Your Email (full address)")
    smtp_username = st.sidebar.text_input("SMTP Username", value="", placeholder="Bỏ trống nếu dùng email")
    password = st.sidebar.text_input("SMTP Password / API Key", type="password")

    st.sidebar.subheader("Sent Folder Sync (optional)")
    save_to_sent = st.sidebar.checkbox("Save a copy to Sent folder", value=True)
    imap_server = st.sidebar.text_input(
        "IMAP Server (for Sent folder only)",
        value="mail.mozeep.com",
        help="Your mailbox provider for receiving/syncing Sent folder. SendGrid SMTP is send-only."
    )
    imap_port = st.sidebar.number_input("IMAP Port", value=993)
    sent_folder = st.sidebar.text_input("Sent Folder Name", value="Sent")

    # Use email as username if no custom username provided
    if not smtp_username.strip():
        smtp_username = your_email

    return (
        smtp_server,
        port,
        smtp_username,
        password,
        your_email,
        save_to_sent,
        imap_server,
        imap_port,
        sent_folder,
    )


def main():
    st.set_page_config(page_title="Mozeep Bulk Email Tool", layout="wide")
    st.title("PHƯƠNG THÙY AUTO SEND CV TOOL (PHƯƠNG THÙY = 😈)")
    st.caption("Only change name/school/etc. — everything else is automatic")

    (
        smtp_server,
        port,
        smtp_username,
        password,
        your_email,
        save_to_sent,
        imap_server,
        imap_port,
        sent_folder,
    ) = render_sidebar_settings()

    st.header("1. Upload your contact list")
    uploaded_file = st.file_uploader(
        "Upload Excel or CSV (must have columns: email, name, ...)",
        type=["csv", "xlsx"],
    )

    if not uploaded_file:
        return

    df = load_uploaded_contacts(uploaded_file)
    if df is None:
        return

    selected = render_contact_selector(df)
    (
        subject,
        use_rich_text,
        attach_signature,
        body_template,
        signature_template,
    ) = render_template_editor()

    if st.button("🚀 SEND SELECTED EMAILS", type="primary"):
        if not your_email or not password:
            st.error("Please enter your email and SendGrid API key in sidebar")
            return

        if len(selected) == 0:
            st.warning("No emails selected")
            return

        send_bulk_emails(
            selected,
            smtp_server,
            port,
            smtp_username,
            password,
            your_email,
            save_to_sent,
            imap_server,
            imap_port,
            sent_folder,
            subject,
            body_template,
            signature_template,
            attach_signature,
            use_rich_text,
        )


if __name__ == "__main__":
    main()