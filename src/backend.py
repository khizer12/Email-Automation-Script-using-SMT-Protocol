import csv
import re
import os
import json
import time
import random
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders

# Constants
JSON_EXTENSION = '.json'
TEMPLATE_DIR = "templates"
os.makedirs(TEMPLATE_DIR, exist_ok=True)

# --- New: timeout and connector helper
DEFAULT_TIMEOUT = 12  # seconds; adjust as needed for slow servers


def _connect_smtp(smtp_config):
    """
    Create and return an smtplib connection with timeout + correct SSL/TLS handling.
    - If 'ssl' in smtp_config or port == 465 => use SMTP_SSL
    - Else use SMTP and optionally starttls() when 'tls' is True
    """
    host = smtp_config.get('server', '')
    port = int(smtp_config.get('port', 0) or 0)
    use_ssl = bool(smtp_config.get('ssl', port == 465))
    use_tls = bool(smtp_config.get('tls', not use_ssl))

    if use_ssl:
        server = smtplib.SMTP_SSL(host, port, timeout=DEFAULT_TIMEOUT)
        server.ehlo()
    else:
        server = smtplib.SMTP(host, port, timeout=DEFAULT_TIMEOUT)
        server.ehlo()
        if use_tls:
            # starttls may raise; caller handles exceptions
            server.starttls()
            server.ehlo()
    return server


# ------------------- CSV & Email Validation -------------------
def load_emails_from_csv(file_path):
    """Load emails from CSV file."""
    emails = []
    try:
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            for row in reader:
                for item in row:
                    item = item.strip()
                    if is_valid_email(item):
                        emails.append(item)
    except Exception as e:
        print(f"Error loading CSV: {e}")
    return emails


def load_emails_from_txt(file_path):
    """Load emails from TXT file."""
    emails = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                email = line.strip()
                if is_valid_email(email):
                    emails.append(email)
    except Exception as e:
        print(f"Error loading TXT: {e}")
    return emails


def load_emails(file_path):
    """Load emails from CSV/TXT and validate format."""
    if not file_path or not os.path.exists(file_path):
        return []

    if file_path.lower().endswith('.csv'):
        emails = load_emails_from_csv(file_path)
    else:
        emails = load_emails_from_txt(file_path)

    return list(set(emails))  # Remove duplicates


def is_valid_email(email):
    """Simple regex email validation."""
    if not email or not isinstance(email, str):
        return False
    pattern = r"[^@]+@[^@]+\.[^@]+"
    return re.match(pattern, email) is not None


# ------------------- SMTP Email Sending -------------------
def create_email_message(smtp_config, recipient, subject, body, attachments=None):
    """Create a MIME email message with optional attachments."""
    # Create multipart message
    msg = MIMEMultipart("related")
    msg['From'] = smtp_config['email']
    msg['To'] = recipient
    msg['Subject'] = subject

    # Create the alternative part for HTML and plain text
    alt = MIMEMultipart("alternative")
    msg.attach(alt)

    # Add plain text version
    alt.attach(MIMEText("This is an HTML email. Please view in HTML capable client.", "plain"))

    # Add HTML version
    alt.attach(MIMEText(body, "html"))

    # Attach files
    if attachments:
        for file_path in attachments:
            try:
                if os.path.exists(file_path):
                    with open(file_path, 'rb') as f:
                        part = MIMEBase('application', "octet-stream")
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
                        msg.attach(part)
            except Exception as e:
                print(f"Attachment error: {e}")

    return msg


def send_email(smtp_config, recipient, subject, body, attachments=None):
    """Send a single email.

    Uses _connect_smtp() to ensure timeout and proper SSL/TLS handling.
    Returns (True, None) on success or (False, error_message) on failure.
    """
    msg = create_email_message(smtp_config, recipient, subject, body, attachments)

    server = None
    try:
        server = _connect_smtp(smtp_config)
        # login if credentials provided
        if smtp_config.get('email') and smtp_config.get('password'):
            server.login(smtp_config['email'], smtp_config['password'])
        server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)
    finally:
        try:
            if server:
                server.quit()
        except Exception:
            pass


# ------------------- Bulk Sending -------------------
def bulk_send(smtp_config, email_list, subject, body, attachments=None, delay_range=(2, 5), retry_failed=True):
    """Send emails in bulk with delay and retry logic."""
    logs = []
    for recipient in email_list:
        success, error = send_email(smtp_config, recipient, subject, body, attachments)
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        status = "Sent" if success else f"Failed: {error}"
        logs.append({'timestamp': timestamp, 'recipient': recipient, 'status': status})

        # Throttle
        time.sleep(random.uniform(*delay_range))

        # Retry if failed
        if retry_failed and not success:
            time.sleep(random.uniform(1, 3))
            success, error = send_email(smtp_config, recipient, subject, body, attachments)
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            status = "Sent (Retry)" if success else f"Failed (Retry): {error}"
            logs.append({'timestamp': timestamp, 'recipient': recipient, 'status': status})
    return logs


# ------------------- Inline Image Helper -------------------
def inline_image(file_path, width=400):
    """Return HTML tag for inline newsletter image."""
    # For GUI preview, we use file:// URL
    if os.path.exists(file_path):
        return f'<img src="file://{os.path.abspath(file_path)}" width="{width}"><br>'
    return ""


# ------------------- Template Management -------------------
def save_template(name, subject, body, attachments=None):
    """Save template locally as JSON."""
    # Ensure we're saving in the templates directory
    if not name.endswith(JSON_EXTENSION):
        name += JSON_EXTENSION

    # Extract just the filename for storage
    filename = os.path.basename(name)
    filepath = os.path.join(TEMPLATE_DIR, filename)

    data = {
        "name": filename,
        "subject": subject,
        "body": body,
        "attachments": attachments or []
    }
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving template: {e}")
        return False


def load_templates():
    """Load all saved templates."""
    templates = []
    try:
        for f in os.listdir(TEMPLATE_DIR):
            if f.endswith(JSON_EXTENSION):
                path = os.path.join(TEMPLATE_DIR, f)
                with open(path, 'r', encoding='utf-8') as file:
                    template_data = json.load(file)
                    templates.append(template_data)
    except Exception as e:
        print(f"Error loading templates: {e}")
    return templates


def delete_template(name):
    """Delete a saved template."""
    # Ensure we're looking in the templates directory
    if not name.endswith(JSON_EXTENSION):
        name += JSON_EXTENSION

    path = os.path.join(TEMPLATE_DIR, name)
    if os.path.exists(path):
        os.remove(path)
        return True
    return False


def get_template_by_name(name):
    """Get a specific template by name."""
    templates = load_templates()
    for template in templates:
        if template["name"] == name:
            return template
    return None


# ------------------- Utility Functions -------------------
def validate_smtp_config(smtp_config):
    """Validate SMTP configuration by attempting to connect.

    Returns (True, message) or (False, error_message).
    This function uses _connect_smtp() which applies a timeout and handles SSL/TLS.
    """
    server = None
    try:
        server = _connect_smtp(smtp_config)
        if smtp_config.get('email') and smtp_config.get('password'):
            server.login(smtp_config['email'], smtp_config['password'])
        try:
            server.noop()
        except Exception:
            pass
        return True, "SMTP configuration is valid"
    except Exception as e:
        return False, f"SMTP validation failed: {str(e)}"
    finally:
        try:
            if server:
                server.quit()
        except Exception:
            pass


def export_logs_to_csv(logs, filename):
    """Export logs to a CSV file."""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['timestamp', 'recipient', 'status'])
            writer.writeheader()
            for log in logs:
                writer.writerow(log)
        return True
    except Exception as e:
        print(f"Error exporting logs: {e}")
        return False


def clean_email_list(email_list):
    """Clean and validate a list of email addresses."""
    cleaned_emails = []
    for email in email_list:
        if isinstance(email, str) and is_valid_email(email.strip()):
            cleaned_emails.append(email.strip())
    return list(set(cleaned_emails))  # Remove duplicates


def count_emails_in_file(file_path):
    """Count valid emails in a file without loading them all."""
    if not file_path or not os.path.exists(file_path):
        return 0

    try:
        if file_path.lower().endswith('.csv'):
            return _count_emails_csv(file_path)
        else:
            return _count_emails_txt(file_path)
    except Exception as e:
        print(f"Error counting emails: {e}")
        return 0


def _count_emails_csv(file_path):
    """Count emails in CSV file."""
    count = 0
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        for row in reader:
            for item in row:
                if is_valid_email(item.strip()):
                    count += 1
    return count


def _count_emails_txt(file_path):
    """Count emails in text file."""
    count = 0
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            if is_valid_email(line.strip()):
                count += 1
    return count
