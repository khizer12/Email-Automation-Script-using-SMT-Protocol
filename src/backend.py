import csv
import re
import os
import json
import time
import random
import smtplib
import socket
import ssl
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders

# Constants
JSON_EXTENSION = '.json'
TEMPLATE_DIR = "templates"
os.makedirs(TEMPLATE_DIR, exist_ok=True)
DEFAULT_TIMEOUT = 15  # Increased timeout for slow connections


def _connect_smtp(smtp_config):
    """
    Create and return an smtplib connection with timeout + correct SSL/TLS handling.
    """
    host = smtp_config.get('server', '')
    port = int(smtp_config.get('port', 0) or 0)
    use_ssl = bool(smtp_config.get('ssl', port == 465))
    use_tls = bool(smtp_config.get('tls', port != 465 and port != 25))

    # Basic sanity checks to catch common misconfigurations early (gives clearer GUI feedback)
    if port == 465 and not use_ssl:
        raise Exception("Port 465 requires SSL. Enable 'Use SSL' or switch to port 587 for STARTTLS.")
    if port == 587 and use_ssl:
        raise Exception("Port 587 typically uses STARTTLS (TLS checkbox) not direct SSL. Disable 'Use SSL' and enable 'Use TLS'.")
    if port == 25 and (use_ssl or use_tls):
        # Not always wrong, but often confusing for users. Provide guidance instead of blocking.
        pass

    # Test if we can even reach the server first
    try:
        # Test basic connectivity
        test_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        test_socket.settimeout(DEFAULT_TIMEOUT)
        test_socket.connect((host, port))
        test_socket.close()
    except (socket.timeout, socket.gaierror, ConnectionRefusedError, OSError) as e:
        raise Exception(f"Cannot reach SMTP server {host}:{port}. Error: {str(e)}")

    try:
        if use_ssl:
            # For SSL connections
            context = ssl.create_default_context()
            server = smtplib.SMTP_SSL(host, port, timeout=DEFAULT_TIMEOUT, context=context)
        else:
            # For non-SSL connections
            server = smtplib.SMTP(host, port, timeout=DEFAULT_TIMEOUT)
        # Optional: debug level (set to 0 in production to reduce noise)
        server.set_debuglevel(0)

        # Identify ourselves to the SMTP server
        server.ehlo()

        # Start TLS if requested and server supports it
        if use_tls and server.has_extn('STARTTLS'):
            server.starttls()
            server.ehlo()  # Re-identify ourselves after TLS negotiation

        return server
    except smtplib.SMTPConnectError as e:
        raise Exception(f"Connection to SMTP server failed: {str(e)}")
    except smtplib.SMTPException as e:
        raise Exception(f"SMTP error: {str(e)}")
    except Exception as e:
        raise Exception(f"Failed to connect to SMTP server: {str(e)}")


def load_emails_from_csv(file_path):
    """Load emails from CSV file with improved parsing."""
    emails = []
    try:
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            # Try different delimiters
            for delimiter in [',', ';', '\t']:
                try:
                    f.seek(0)  # Reset file pointer
                    reader = csv.reader(f, delimiter=delimiter)
                    for row in reader:
                        for item in row:
                            item = item.strip()
                            if is_valid_email(item):
                                emails.append(item)
                    break  # Success with this delimiter
                except csv.Error:
                    continue
    except (IOError, OSError, csv.Error) as e:
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
    except (IOError, OSError) as e:
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
    """Improved regex email validation."""
    if not email or not isinstance(email, str):
        return False
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None


def create_email_message(smtp_config, recipient, subject, body, attachments=None):
    """Create a MIME email message with optional attachments."""
    msg = MIMEMultipart("related")
    msg['From'] = smtp_config['email']
    msg['To'] = recipient
    msg['Subject'] = subject

    alt = MIMEMultipart("alternative")
    msg.attach(alt)

    alt.attach(MIMEText("This is an HTML email. Please view in HTML capable client.", "plain"))
    alt.attach(MIMEText(body, "html"))

    if attachments:
        for file_path in attachments:
            try:
                if os.path.exists(file_path):
                    with open(file_path, 'rb') as f:
                        part = MIMEBase('application', "octet-stream")
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        filename = os.path.basename(file_path)
                        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                        msg.attach(part)
            except (IOError, OSError) as e:
                print(f"Attachment error: {e}")
    return msg


def send_email(smtp_config, recipient, subject, body, attachments=None):
    """Send a single email."""
    msg = create_email_message(smtp_config, recipient, subject, body, attachments)

    server = None
    try:
        server = _connect_smtp(smtp_config)
        if smtp_config.get('email') and smtp_config.get('password'):
            server.login(smtp_config['email'], smtp_config['password'])
        server.send_message(msg)
        return True, None
    except (smtplib.SMTPException, OSError, Exception) as e:
        return False, str(e)
    finally:
        try:
            if server:
                server.quit()
        except smtplib.SMTPException:
            pass


def bulk_send(smtp_config, email_list, subject, body, attachments=None, delay_range=(2, 5), retry_failed=True):
    """Send emails in bulk with delay and retry logic."""
    logs = []
    total = len(email_list)
    for i, recipient in enumerate(email_list):
        success, error = send_email(smtp_config, recipient, subject, body, attachments)
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        status = "Sent" if success else f"Failed: {error}"
        logs.append({'timestamp': timestamp, 'recipient': recipient, 'status': status})

        if i < total - 1:  # Don't sleep after the last email
            time.sleep(random.uniform(*delay_range))

        if retry_failed and not success:
            time.sleep(random.uniform(1, 3))
            success, error = send_email(smtp_config, recipient, subject, body, attachments)
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            status = "Sent (Retry)" if success else f"Failed (Retry): {error}"
            logs.append({'timestamp': timestamp, 'recipient': recipient, 'status': status})
    return logs


def inline_image(file_path, width=400):
    """Return HTML tag for inline newsletter image."""
    if os.path.exists(file_path):
        return f'<img src="file://{os.path.abspath(file_path)}" width="{width}"><br>'
    return ""


def save_template(name, subject, body, attachments=None):
    """Save template locally as JSON."""
    if not name.endswith(JSON_EXTENSION):
        name += JSON_EXTENSION

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
    except (IOError, OSError) as e:
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
    except (IOError, OSError, json.JSONDecodeError) as e:
        print(f"Error loading templates: {e}")
    return templates


def delete_template(name):
    """Delete a saved template."""
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


def validate_smtp_config(smtp_config):
    """Validate SMTP configuration by attempting to connect."""
    server = None
    try:
        server = _connect_smtp(smtp_config)
        if smtp_config.get('email') and smtp_config.get('password'):
            server.login(smtp_config['email'], smtp_config['password'])
        try:
            server.noop()
        except smtplib.SMTPException:
            pass
        return True, "SMTP configuration is valid"
    except (smtplib.SMTPException, OSError, Exception) as e:
        return False, f"SMTP validation failed: {str(e)}"
    finally:
        try:
            if server:
                server.quit()
        except smtplib.SMTPException:
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
    except (IOError, OSError) as e:
        print(f"Error exporting logs: {e}")
        return False


def clean_email_list(email_list):
    """Clean and validate a list of email addresses."""
    cleaned_emails = []
    for email in email_list:
        if isinstance(email, str) and is_valid_email(email.strip()):
            cleaned_emails.append(email.strip())
    return list(set(cleaned_emails))


def count_emails_in_file(file_path):
    """Count valid emails in a file without loading them all."""
    if not file_path or not os.path.exists(file_path):
        return 0

    try:
        if file_path.lower().endswith('.csv'):
            return _count_emails_csv(file_path)
        else:
            return _count_emails_txt(file_path)
    except (IOError, OSError) as e:
        print(f"Error counting emails: {e}")
        return 0


def _count_emails_csv(file_path):
    """Count emails in CSV file with improved parsing."""
    count = 0
    try:
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            for delimiter in [',', ';', '\t']:
                try:
                    f.seek(0)  # Reset file pointer
                    reader = csv.reader(f, delimiter=delimiter)
                    for row in reader:
                        for item in row:
                            if is_valid_email(item.strip()):
                                count += 1
                    break
                except csv.Error:
                    continue
    except (IOError, OSError) as e:
        print(f"Error counting CSV emails: {e}")
    return count


def _count_emails_txt(file_path):
    """Count emails in text file."""
    count = 0
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                if is_valid_email(line.strip()):
                    count += 1
    except (IOError, OSError) as e:
        print(f"Error counting TXT emails: {e}")
    return count
