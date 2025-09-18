import time
import sys
import os
import random
import traceback
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QTextEdit, QListWidget, QFileDialog, QVBoxLayout, QHBoxLayout,
    QGridLayout, QTabWidget, QMessageBox, QSplitter, QFrame, QProgressBar,
    QCheckBox
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import backend

# Constants
NO_SELECTION_MSG = "Please select a template first!"
DELETE_NO_SELECTION_MSG = "Please select a template to delete!"
NO_SELECTION_TITLE = "No Selection"


class EmailThread(QThread):
    """Thread for sending emails to prevent GUI freezing."""
    progress = pyqtSignal(int)
    log_signal = pyqtSignal(str)
    finished = pyqtSignal(list)

    def __init__(self, smtp_config, recipients, subject, body, attachments):
        """Initialize the email thread."""
        super().__init__()
        self.smtp_config = smtp_config
        self.recipients = recipients
        self.subject = subject
        self.body = body
        self.attachments = attachments
        self.is_running = True

    def run(self):
        """Run the email sending process."""
        logs = []
        total = len(self.recipients)
        for i, recipient in enumerate(self.recipients):
            if not self.is_running:
                break
                
            success, error = backend.send_email(
                self.smtp_config, recipient, self.subject, self.body, self.attachments
            )
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            status = "Sent" if success else f"Failed: {error}"
            log_entry = {'timestamp': timestamp, 'recipient': recipient, 'status': status}
            logs.append(log_entry)

            self.progress.emit(int((i + 1) / total * 100))
            self.log_signal.emit(f"{timestamp} - {recipient} - {status}")

            # Throttle (but not after the last email)
            if i < total - 1 and self.is_running:
                time.sleep(random.uniform(2, 5))

            # Retry if failed
            if not success and self.is_running:
                time.sleep(random.uniform(1, 3))
                success, error = backend.send_email(
                    self.smtp_config, recipient, self.subject, self.body, self.attachments
                )
                timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                status = "Sent (Retry)" if success else f"Failed (Retry): {error}"
                log_entry = {'timestamp': timestamp, 'recipient': recipient, 'status': status}
                logs.append(log_entry)
                self.log_signal.emit(f"{timestamp} - {recipient} - {status}")

        self.finished.emit(logs)
        
    def stop(self):
        """Stop the email sending process."""
        self.is_running = False


class SmtpValidateThread(QThread):
    """Thread to validate SMTP config without blocking the GUI."""
    result = pyqtSignal(bool, str)

    def __init__(self, smtp_config):
        """Initialize the SMTP validation thread."""
        super().__init__()
        self.smtp_config = smtp_config

    def run(self):
        """Run the SMTP validation process."""
        valid, message = backend.validate_smtp_config(self.smtp_config)
        self.result.emit(valid, message)


class EmailApp(QMainWindow):
    """Main application window for the bulk email sender."""

    def __init__(self):
        """Initialize the application window."""
        super().__init__()
        self.setWindowTitle("Bulk Email Sender")
        self.setMinimumSize(1000, 700)

        # Fonts
        self.font_large = QFont("Segoe UI", 16, QFont.Weight.Bold)
        self.font_medium = QFont("Segoe UI", 14, QFont.Weight.DemiBold)

        # Theme
        self.dark_mode = False
        self.light_style = """
            QMainWindow, QWidget { background-color: #f0f0f0; color: #000; font-family: Segoe UI; font-size: 14px; }
            QTextEdit, QLineEdit, QListWidget { background-color: #fff; color: #000; border: 1px solid #ccc; border-radius: 4px; padding: 5px; }
            QLabel { background-color: transparent; color: #000; }
            QTabWidget::pane { border: 1px solid #ccc; background-color: #f0f0f0; }
            QTabBar::tab { background-color: #e0e0e0; padding: 8px 12px; border-top-left-radius: 4px; border-top-right-radius: 4px; }
            QTabBar::tab:selected { background-color: #f0f0f0; border-bottom: 2px solid #5a9bd5; }
            QPushButton { font-size: 14px; font-weight: bold; background-color: #5a9bd5; color: white; padding: 8px 15px; border-radius: 6px; }
            QPushButton:hover { background-color: #7bb1eb; }
            QPushButton#primary { background-color: #28a745; padding: 10px 20px; }
            QPushButton#primary:hover { background-color: #45c159; }
            QPushButton#toggle { font-size: 16px; background-color: #6c757d; padding: 10px 25px; border-radius: 10px; }
            QPushButton#toggle:hover { background-color: #868e96; }
            QPushButton#danger { background-color: #dc3545; }
            QPushButton#danger:hover { background-color: #e25563; }
            QProgressBar { border: 1px solid #ccc; border-radius: 4px; text-align: center; }
            QProgressBar::chunk { background-color: #28a745; }
            QListWidget::item:selected { background-color: #5a9bd5; color: white; }
            QCheckBox { font-size: 14px; }
        """
        self.dark_style = """
            QMainWindow, QWidget { background-color: #2b2b2b; color: #fff; font-family: Segoe UI; font-size: 14px; }
            QTextEdit, QLineEdit, QListWidget { background-color: #3b3b3b; color: #fff; border: 1px solid #555; border-radius: 4px; padding: 5px; }
            QLabel { background-color: transparent; color: #fff; }
            QTabWidget::pane { border: 1px solid #555; background-color: #2b2b2b; }
            QTabBar::tab { background-color: #353535; color: #fff; padding: 8px 12px; border-top-left-radius: 4px; border-top-right-radius: 4px; }
            QTabBar::tab:selected { background-color: #2b2b2b; border-bottom: 2px solid #5a9bd5; }
            QPushButton { font-size: 14px; font-weight: bold; background-color: #5a9bd5; color: white; padding: 8px 15px; border-radius: 6px; }
            QPushButton:hover { background-color: #7bb1eb; }
            QPushButton#primary { background-color: #28a745; padding: 10px 20px; }
            QPushButton#primary:hover { background-color: #45c159; }
            QPushButton#toggle { font-size: 16px; background-color: #6c757d; padding: 10px 25px; border-radius: 10px; }
            QPushButton#toggle:hover { background-color: #868e96; }
            QPushButton#danger { background-color: #dc3545; }
            QPushButton#danger:hover { background-color: #e25563; }
            QProgressBar { border: 1px solid #555; border-radius: 4px; text-align: center; color: white; }
            QProgressBar::chunk { background-color: #28a745; }
            QListWidget::item:selected { background-color: #5a9bd5; color: white; }
            QCheckBox { font-size: 14px; color: white; }
        """
        self.setStyleSheet(self.light_style)

        # Initialize instance variables
        self.attachments = []
        self.template_attachments = []
        self.email_thread = None
        self.current_logs = []
        self.validate_thread = None
        self.validate_thread2 = None
        self._pending_send = None
        
        # UI elements to be initialized in init_ui
        self.entries = []
        self.entry_server = None
        self.entry_port = None
        self.entry_email = None
        self.entry_password = None
        self.list_emails = None
        self.entry_subject = None
        self.text_body = None
        self.preview_body = None
        self.progress_bar = None
        self.text_logs = None
        self.list_templates = None
        self.entry_temp_subject = None
        self.text_temp_body = None
        self.preview_temp_body = None
        self.ssl_checkbox = None
        self.tls_checkbox = None

        self.init_ui()

    def init_ui(self):
        """Initialize the user interface."""
        try:
            self.tabs = QTabWidget()
            self.setCentralWidget(self.tabs)

            self.main_tab = QWidget()
            self.templates_tab = QWidget()
            self.tabs.addTab(self.main_tab, "Main")
            self.tabs.addTab(self.templates_tab, "Templates")

            toggle_container = QWidget()
            toggle_layout = QHBoxLayout()
            toggle_layout.setContentsMargins(0, 0, 10, 0)
            toggle_container.setLayout(toggle_layout)

            toggle_btn = QPushButton("Dark / Light Mode")
            toggle_btn.setObjectName("toggle")
            toggle_btn.clicked.connect(self.toggle_theme)
            toggle_layout.addWidget(toggle_btn)
            self.tabs.setCornerWidget(toggle_container, corner=Qt.Corner.TopRightCorner)

            self.init_main_tab()
            self.init_templates_tab()
        except Exception as e:
            QMessageBox.critical(self, "Initialization Error", f"Failed to initialize UI: {str(e)}")
            raise

    # ------------------- MAIN TAB -------------------
    def init_main_tab(self):
        """Initialize the main tab UI."""
        try:
            layout = QVBoxLayout()
            self.main_tab.setLayout(layout)

            # SMTP Configuration
            smtp_frame = QFrame()
            smtp_layout = QGridLayout()
            smtp_frame.setLayout(smtp_layout)
            layout.addWidget(smtp_frame)

            labels = ["SMTP Server:", "Port:", "Email:", "Password:"]
            self.entries = []
            for i, text in enumerate(labels):
                smtp_layout.addWidget(QLabel(text), i // 2, (i % 2) * 2)
                entry = QLineEdit()
                if "Password" in text:
                    entry.setEchoMode(QLineEdit.EchoMode.Password)
                smtp_layout.addWidget(entry, i // 2, (i % 2) * 2 + 1)
                self.entries.append(entry)

            # Fix unbalanced tuple unpacking
            if len(self.entries) >= 4:
                self.entry_server, self.entry_port, self.entry_email, self.entry_password = self.entries[:4]
            else:
                # Fallback initialization
                self.entry_server = QLineEdit()
                self.entry_port = QLineEdit()
                self.entry_email = QLineEdit()
                self.entry_password = QLineEdit()

            # Set some default values for Hostinger
            self.entry_server.setText("smtp.hostinger.com")
            self.entry_port.setText("587")  # Changed to 587 as it's more likely to work
            self.ssl_checkbox = QCheckBox("Use SSL")
            self.ssl_checkbox.setChecked(False)  # Changed to False for port 587
            self.tls_checkbox = QCheckBox("Use TLS")
            self.tls_checkbox.setChecked(True)  # Changed to True for port 587

            # Auto-adjust SSL/TLS when user changes port to reduce mistakes
            try:
                self.entry_port.textChanged.connect(self.on_port_changed)
            except Exception:
                pass
            
            smtp_layout.addWidget(self.ssl_checkbox, 2, 0, 1, 2)
            smtp_layout.addWidget(self.tls_checkbox, 3, 0, 1, 2)

            btn_validate_smtp = QPushButton("Validate SMTP Configuration")
            btn_validate_smtp.clicked.connect(self.validate_smtp)
            smtp_layout.addWidget(btn_validate_smtp, 4, 0, 1, 2)

            # Add connection troubleshooting tips
            tips_label = QLabel(
                "Troubleshooting tips:\n"
                "1. Check if your firewall allows outbound connections on SMTP ports\n"
                "2. Verify your SMTP server settings with your email provider\n"
                "3. Try different port/SSL combinations (587/TLS or 465/SSL)\n"
                "4. Ensure your credentials are correct"
            )
            tips_label.setWordWrap(True)
            smtp_layout.addWidget(tips_label, 5, 0, 1, 2)

            splitter = QSplitter(Qt.Orientation.Horizontal)
            layout.addWidget(splitter)

            # Recipients
            rec_frame = QFrame()
            rec_layout = QVBoxLayout()
            rec_frame.setLayout(rec_layout)
            splitter.addWidget(rec_frame)

            rec_layout.addWidget(QLabel("Recipients:"))
            self.list_emails = QListWidget()
            rec_layout.addWidget(self.list_emails)

            btn_load_csv = QPushButton("Load CSV/TXT")
            btn_load_csv.clicked.connect(self.load_csv)
            rec_layout.addWidget(btn_load_csv)

            btn_remove = QPushButton("Remove Selected")
            btn_remove.clicked.connect(self.remove_selected)
            rec_layout.addWidget(btn_remove)

            btn_clear_all = QPushButton("Clear All")
            btn_clear_all.clicked.connect(self.clear_all_emails)
            rec_layout.addWidget(btn_clear_all)

            rec_layout.addStretch()

            # Email Editor + Preview
            editor_frame = QFrame()
            editor_layout = QVBoxLayout()
            editor_frame.setLayout(editor_layout)
            splitter.addWidget(editor_frame)

            editor_layout.addWidget(QLabel("Subject:"))
            self.entry_subject = QLineEdit()
            editor_layout.addWidget(self.entry_subject)

            editor_layout.addWidget(QLabel("Body (Supports HTML/newsletter images):"))
            self.text_body = QTextEdit()
            editor_layout.addWidget(self.text_body)

            btn_insert_image = QPushButton("Insert Image")
            btn_insert_image.clicked.connect(self.insert_image_main)
            editor_layout.addWidget(btn_insert_image)

            btn_add_attachment = QPushButton("Add Attachments")
            btn_add_attachment.clicked.connect(self.add_attachment)
            editor_layout.addWidget(btn_add_attachment)

            # Preview Panel
            editor_layout.addWidget(QLabel("Preview:"))
            self.preview_body = QTextEdit()
            self.preview_body.setReadOnly(True)
            editor_layout.addWidget(self.preview_body)

            btn_preview = QPushButton("Render Preview")
            btn_preview.clicked.connect(self.render_preview_main)
            editor_layout.addWidget(btn_preview)

            # Progress bar
            self.progress_bar = QProgressBar()
            self.progress_bar.setVisible(False)
            editor_layout.addWidget(self.progress_bar)

            editor_layout.addWidget(QLabel("Logs:"))
            self.text_logs = QTextEdit()
            self.text_logs.setReadOnly(True)
            editor_layout.addWidget(self.text_logs)

            btn_send = QPushButton("Send Emails")
            btn_send.setObjectName("primary")
            btn_send.clicked.connect(self.send_email)
            editor_layout.addWidget(btn_send)

            btn_stop = QPushButton("Stop Sending")
            btn_stop.setObjectName("danger")
            btn_stop.clicked.connect(self.stop_sending)
            btn_stop.setVisible(False)
            editor_layout.addWidget(btn_stop)

            btn_export_logs = QPushButton("Export Logs")
            btn_export_logs.clicked.connect(self.export_logs)
            editor_layout.addWidget(btn_export_logs)

            splitter.setSizes([300, 700])
        except Exception as e:
            QMessageBox.critical(self, "UI Error", f"Failed to initialize main tab: {str(e)}")
            raise

    # ------------------- TEMPLATES TAB -------------------
    def init_templates_tab(self):
        """Initialize the templates tab UI."""
        try:
            layout = QHBoxLayout()
            self.templates_tab.setLayout(layout)

            # Left: Templates list
            left_widget = QWidget()
            left_layout = QVBoxLayout()
            left_widget.setLayout(left_layout)
            left_widget.setMinimumWidth(250)
            layout.addWidget(left_widget)

            left_layout.addWidget(QLabel("Saved Templates:"))
            self.list_templates = QListWidget()
            self.refresh_templates()
            left_layout.addWidget(self.list_templates)

            template_buttons_layout = QVBoxLayout()
            left_layout.addLayout(template_buttons_layout)

            btn_load_temp = QPushButton("Load Selected")
            btn_load_temp.clicked.connect(self.load_template)
            template_buttons_layout.addWidget(btn_load_temp)

            btn_use_temp = QPushButton("Use in Main Tab")
            btn_use_temp.clicked.connect(self.use_template_in_main)
            template_buttons_layout.addWidget(btn_use_temp)

            btn_save_temp = QPushButton("Save Current as Template")
            btn_save_temp.clicked.connect(self.save_template)
            template_buttons_layout.addWidget(btn_save_temp)

            btn_delete_temp = QPushButton("Delete Selected")
            btn_delete_temp.setObjectName("danger")
            btn_delete_temp.clicked.connect(self.delete_template)
            template_buttons_layout.addWidget(btn_delete_temp)
            template_buttons_layout.addStretch()

            # Right: Template Editor + Preview
            editor_frame = QFrame()
            editor_layout = QVBoxLayout()
            editor_frame.setLayout(editor_layout)
            layout.addWidget(editor_frame)

            editor_layout.addWidget(QLabel("Subject:"))
            self.entry_temp_subject = QLineEdit()
            editor_layout.addWidget(self.entry_temp_subject)

            editor_layout.addWidget(QLabel("Body (Supports HTML/newsletter images):"))
            self.text_temp_body = QTextEdit()
            editor_layout.addWidget(self.text_temp_body)

            btn_insert_img = QPushButton("Insert Image")
            btn_insert_img.clicked.connect(self.insert_image_template)
            editor_layout.addWidget(btn_insert_img)

            btn_add_template_attach = QPushButton("Add Attachments")
            btn_add_template_attach.clicked.connect(self.add_template_attachment)
            editor_layout.addWidget(btn_add_template_attach)

            # Preview Panel
            editor_layout.addWidget(QLabel("Preview:"))
            self.preview_temp_body = QTextEdit()
            self.preview_temp_body.setReadOnly(True)
            editor_layout.addWidget(self.preview_temp_body)

            btn_preview_temp = QPushButton("Render Preview")
            btn_preview_temp.clicked.connect(self.render_preview_template)
            editor_layout.addWidget(btn_preview_temp)
        except Exception as e:
            QMessageBox.critical(self, "UI Error", f"Failed to initialize templates tab: {str(e)}")
            raise

    # ------------------- FUNCTIONS -------------------
    def toggle_theme(self):
        """Toggle between light and dark themes."""
        self.dark_mode = not self.dark_mode
        self.setStyleSheet(self.dark_style if self.dark_mode else self.light_style)
        # Update previews background
        bg_color = "#3b3b3b" if self.dark_mode else "#fff"
        text_color = "#fff" if self.dark_mode else "#000"
        self.preview_body.setStyleSheet(f"background-color:{bg_color}; color:{text_color};")
        self.preview_temp_body.setStyleSheet(f"background-color:{bg_color}; color:{text_color};")

    def on_port_changed(self, value):
        """Automatically set sensible SSL/TLS defaults based on common SMTP ports.

        465 -> SSL only
        587 -> STARTTLS (TLS) only
        25  -> No SSL/TLS by default (can be upgraded manually)
        """
        port = value.strip()
        if port == "465":
            self.ssl_checkbox.setChecked(True)
            self.tls_checkbox.setChecked(False)
        elif port == "587":
            self.ssl_checkbox.setChecked(False)
            self.tls_checkbox.setChecked(True)
        elif port == "25":
            self.ssl_checkbox.setChecked(False)
            self.tls_checkbox.setChecked(False)

    def log(self, msg):
        """Add a message to the log."""
        self.text_logs.append(msg)

    # ---------- Updated: non-blocking validate ----------
    def validate_smtp(self):
        """Validate SMTP configuration (non-blocking)."""
        try:
            smtp_config = {
                "server": self.entry_server.text(),
                "port": int(self.entry_port.text()) if self.entry_port.text().isdigit() else 587,
                "email": self.entry_email.text(),
                "password": self.entry_password.text(),
                "ssl": self.ssl_checkbox.isChecked(),
                "tls": self.tls_checkbox.isChecked()
            }

            # run validation on background thread
            self.validate_thread = SmtpValidateThread(smtp_config)
            self.validate_thread.result.connect(self.on_validate_done)
            # disable cursor / button for UX
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            sender = self.sender()
            if isinstance(sender, QPushButton):
                sender.setEnabled(False)
            # restore states when finished
            def _cleanup():
                QApplication.restoreOverrideCursor()
                if isinstance(sender, QPushButton):
                    sender.setEnabled(True)
            self.validate_thread.finished.connect(_cleanup)
            self.validate_thread.start()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to validate SMTP: {str(e)}")

    def on_validate_done(self, valid, message):
        """Handle SMTP validation result."""
        if valid:
            QMessageBox.information(self, "SMTP Valid", message)
        else:
            # Provide more detailed troubleshooting for timeout errors
            if "timed out" in message or "connection" in message.lower():
                detailed_message = f"{message}\n\nTroubleshooting tips:\n"
                detailed_message += "1. Check if your firewall allows outbound connections on SMTP ports\n"
                detailed_message += "2. Verify your SMTP server settings with your email provider\n"
                detailed_message += "3. Try different port/SSL combinations (587/TLS or 465/SSL)\n"
                detailed_message += "4. Ensure your credentials are correct"
                QMessageBox.warning(self, "SMTP Connection Failed", detailed_message)
            else:
                QMessageBox.warning(self, "SMTP Invalid", message)

    def load_csv(self):
        """Load emails from a CSV or TXT file."""
        try:
            path, _ = QFileDialog.getOpenFileName(
                self, "Load CSV/TXT", "", "CSV Files (*.csv);;Text Files (*.txt)"
            )
            if path:
                count = backend.count_emails_in_file(path)
                if count == 0:
                    QMessageBox.warning(self, "No Valid Emails", 
                                       "The selected file contains no valid email addresses.")
                    return

                if QMessageBox.question(
                    self, "Confirm Load",
                    f"Found {count} valid emails. Load them?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                ) == QMessageBox.StandardButton.Yes:
                    emails = backend.load_emails(path)
                    self.list_emails.clear()
                    self.list_emails.addItems(emails)
                    self.log(f"Loaded {len(emails)} emails from {os.path.basename(path)}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load file: {str(e)}")

    def remove_selected(self):
        """Remove selected emails from the list."""
        for item in self.list_emails.selectedItems():
            self.list_emails.takeItem(self.list_emails.row(item))

    def clear_all_emails(self):
        """Clear all emails from the list."""
        if self.list_emails.count() > 0:
            if QMessageBox.question(
                self, "Confirm Clear",
                "Clear all recipients?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            ) == QMessageBox.StandardButton.Yes:
                self.list_emails.clear()

    def add_attachment(self):
        """Add attachments to the email."""
        files, _ = QFileDialog.getOpenFileNames(self, "Select Attachments/Images")
        if files:
            self.attachments.extend(files)
            self.log(f"Added {len(files)} attachments/images.")

    def insert_image_main(self):
        """Insert an image into the main email body."""
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Image", "", "Images (*.png *.jpg *.jpeg *.gif *.bmp)"
        )
        if file:
            html_tag = backend.inline_image(file)
            self.text_body.insertHtml(html_tag)

    def insert_image_template(self):
        """Insert an image into the template email body."""
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Image", "", "Images (*.png *.jpg *.jpeg *.gif *.bmp)"
        )
        if file:
            html_tag = backend.inline_image(file)
            self.text_temp_body.insertHtml(html_tag)

    def render_preview_main(self):
        """Render a preview of the main email."""
        html_content = self.text_body.toHtml()
        self.preview_body.setHtml(html_content)

    def render_preview_template(self):
        """Render a preview of the template email."""
        html_content = self.text_temp_body.toHtml()
        self.preview_temp_body.setHtml(html_content)

    # ---------- send flow: validate off-thread then start EmailThread ----------
    def send_email(self):
        """Send emails to all recipients."""
        try:
            recipients = [self.list_emails.item(i).text() for i in range(self.list_emails.count())]
            if not recipients:
                QMessageBox.warning(self, "No Recipients", "Load recipient emails first!")
                return

            subject = self.entry_subject.text()
            if not subject.strip():
                QMessageBox.warning(self, "No Subject", "Please enter an email subject!")
                return

            body = self.text_body.toHtml()
            if not body.strip():
                QMessageBox.warning(self, "No Body", "Please enter email content!")
                return

            # Prepare SMTP config
            smtp_config = {
                "server": self.entry_server.text(),
                "port": int(self.entry_port.text()) if self.entry_port.text().isdigit() else 587,
                "email": self.entry_email.text(),
                "password": self.entry_password.text(),
                "ssl": self.ssl_checkbox.isChecked(),
                "tls": self.tls_checkbox.isChecked()
            }

            # Validate first BUT non-blocking: use SmtpValidateThread then continue in callback
            self._pending_send = {
                "smtp_config": smtp_config,
                "recipients": recipients,
                "subject": subject,
                "body": body,
                "attachments": list(self.attachments)
            }

            # Disable UI controls while validating
            self.findChild(QPushButton, "primary").setEnabled(False)
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

            self.validate_thread2 = SmtpValidateThread(smtp_config)
            self.validate_thread2.result.connect(self._on_validate_before_send)
            # cleanup UI once validation attempt finishes
            self.validate_thread2.finished.connect(lambda: (
                QApplication.restoreOverrideCursor(),
            ))
            self.validate_thread2.start()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to send email: {str(e)}")
            self.findChild(QPushButton, "primary").setEnabled(True)

    def _on_validate_before_send(self, valid, message):
        """Handle validation result before sending emails."""
        if not valid:
            QMessageBox.warning(self, "SMTP Invalid", f"Cannot send emails: {message}")
            self.findChild(QPushButton, "primary").setEnabled(True)
            return

        # Proceed to start EmailThread
        pending = self._pending_send
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.findChild(QPushButton, "primary").setEnabled(False)
        self.findChild(QPushButton, "danger").setVisible(True)

        self.email_thread = EmailThread(
            pending["smtp_config"], 
            pending["recipients"], 
            pending["subject"], 
            pending["body"], 
            pending["attachments"]
        )
        self.email_thread.progress.connect(self.progress_bar.setValue)
        self.email_thread.log_signal.connect(self.log)
        self.email_thread.finished.connect(self.on_email_finished)
        self.email_thread.start()

        self.log("Started sending emails.")

    def stop_sending(self):
        """Stop the email sending process."""
        if self.email_thread and self.email_thread.isRunning():
            self.email_thread.stop()
            self.email_thread.wait()
            self.log("Email sending stopped by user")
            self.findChild(QPushButton, "primary").setEnabled(True)
            self.findChild(QPushButton, "danger").setVisible(False)
            self.progress_bar.setVisible(False)

    def on_email_finished(self, logs):
        """Handle completion of email sending."""
        self.current_logs.extend(logs)
        self.findChild(QPushButton, "primary").setEnabled(True)
        self.findChild(QPushButton, "danger").setVisible(False)
        self.progress_bar.setVisible(False)
        self.log("Email sending completed!")

    def export_logs(self):
        """Export logs to a CSV file."""
        if not self.current_logs:
            QMessageBox.warning(self, "No Logs", "There are no logs to export.")
            return

        path, _ = QFileDialog.getSaveFileName(self, "Export Logs", "", "CSV Files (*.csv)")
        if path:
            if backend.export_logs_to_csv(self.current_logs, path):
                QMessageBox.information(self, "Success", f"Logs exported to {path}")
            else:
                QMessageBox.critical(self, "Error", "Failed to export logs")

    # ------------------- TEMPLATE FUNCTIONS -------------------
    def refresh_templates(self):
        """Refresh the list of templates."""
        self.list_templates.clear()
        templates = backend.load_templates()
        for t in templates:
            self.list_templates.addItem(t["name"])

    def load_template(self):
        """Load a selected template."""
        selected = self.list_templates.selectedItems()
        if not selected:
            QMessageBox.warning(self, NO_SELECTION_TITLE, NO_SELECTION_MSG)
            return

        name = selected[0].text()
        template = backend.get_template_by_name(name)
        if template:
            self.entry_temp_subject.setText(template["subject"])
            self.text_temp_body.setHtml(template["body"])
            self.template_attachments = template.get("attachments", [])
            self.log(f"Loaded template '{name}'")
        else:
            QMessageBox.warning(self, "Error", f"Template '{name}' not found!")

    def use_template_in_main(self):
        """Load selected template into the main tab."""
        selected = self.list_templates.selectedItems()
        if not selected:
            QMessageBox.warning(self, NO_SELECTION_TITLE, NO_SELECTION_MSG)
            return

        name = selected[0].text()
        template = backend.get_template_by_name(name)
        if template:
            self.entry_subject.setText(template["subject"])
            self.text_body.setHtml(template["body"])
            self.attachments = template.get("attachments", [])
            self.tabs.setCurrentIndex(0)  # Switch to main tab
            self.log(f"Loaded template '{name}' into main editor")
        else:
            QMessageBox.warning(self, "Error", f"Template '{name}' not found!")

    def save_template(self):
        """Save the current email as a template."""
        name, ok = QFileDialog.getSaveFileName(self, "Save Template As", "", "JSON (*.json)")
        if not name or not ok:
            return

        if not name.endswith('.json'):
            name += '.json'

        success = backend.save_template(
            name,
            self.entry_temp_subject.text(),
            self.text_temp_body.toHtml(),
            self.template_attachments
        )

        if success:
            self.refresh_templates()
            self.log(f"Template saved as '{os.path.basename(name)}'")
        else:
            QMessageBox.critical(self, "Error", "Failed to save template!")

    def delete_template(self):
        """Delete a selected template."""
        selected = self.list_templates.selectedItems()
        if not selected:
            QMessageBox.warning(self, NO_SELECTION_TITLE, DELETE_NO_SELECTION_MSG)
            return

        name = selected[0].text()
        if QMessageBox.question(
            self, "Confirm Delete",
            f"Delete template '{name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        ) == QMessageBox.StandardButton.Yes:
            if backend.delete_template(name):
                self.refresh_templates()
                self.log(f"Deleted template '{name}'")
            else:
                QMessageBox.warning(self, "Error", f"Failed to delete template '{name}'")

    def add_template_attachment(self):
        """Add attachments to the template."""
        files, _ = QFileDialog.getOpenFileNames(self, "Select Attachments/Images")
        if files:
            self.template_attachments.extend(files)
            self.log(f"Added {len(files)} template attachments/images.")


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = EmailApp()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"Application error: {str(e)}")
        traceback.print_exc()
        input("Press Enter to exit...")
