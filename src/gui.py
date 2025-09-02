import time
import sys, os
import csv
import random
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QTextEdit, QListWidget, QFileDialog, QVBoxLayout, QHBoxLayout,
    QGridLayout, QTabWidget, QMessageBox, QSplitter, QFrame, QProgressBar
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import backend

# Constants
NO_SELECTION_MSG = "Please select a template first!"
DELETE_NO_SELECTION_MSG = "Please select a template to delete!"
NO_SELECTION_TITLE = "No Selection"

class EmailThread(QThread):
    """Thread for sending emails to prevent GUI freezing"""
    progress = pyqtSignal(int)
    log_signal = pyqtSignal(str)
    finished = pyqtSignal(list)
    
    def __init__(self, smtp_config, recipients, subject, body, attachments):
        super().__init__()
        self.smtp_config = smtp_config
        self.recipients = recipients
        self.subject = subject
        self.body = body
        self.attachments = attachments
    
    def run(self):
        logs = []
        total = len(self.recipients)
        for i, recipient in enumerate(self.recipients):
            success, error = backend.send_email(self.smtp_config, recipient, self.subject, self.body, self.attachments)
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            status = "Sent" if success else f"Failed: {error}"
            log_entry = {'timestamp': timestamp, 'recipient': recipient, 'status': status}
            logs.append(log_entry)
            
            self.progress.emit(int((i + 1) / total * 100))
            self.log_signal.emit(f"{timestamp} - {recipient} - {status}")
            
            # Throttle
            time.sleep(random.uniform(2, 5))
            
            # Retry if failed
            if not success:
                time.sleep(random.uniform(1, 3))
                success, error = backend.send_email(self.smtp_config, recipient, self.subject, self.body, self.attachments)
                timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                status = "Sent (Retry)" if success else f"Failed (Retry): {error}"
                log_entry = {'timestamp': timestamp, 'recipient': recipient, 'status': status}
                logs.append(log_entry)
                self.log_signal.emit(f"{timestamp} - {recipient} - {status}")
        
        self.finished.emit(logs)

class EmailApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Binder Bulk Email Automation")
        self.setMinimumSize(1350, 800)

        # Fonts
        self.font_large = QFont("Segoe UI", 16, QFont.Weight.Bold)
        self.font_medium = QFont("Segoe UI", 14, QFont.Weight.DemiBold)

        # Theme
        self.dark_mode = False
        self.light_style = """
            QMainWindow, QWidget { background-color: #f0f0f0; color: #000; font-family: Segoe UI; font-size: 16px; }
            QTextEdit, QLineEdit, QListWidget { background-color: #fff; color: #000; border: 1px solid #ccc; border-radius: 4px; padding: 5px; }
            QLabel { background-color: transparent; color: #000; }
            QTabWidget::pane { border: 1px solid #ccc; background-color: #f0f0f0; }
            QTabBar::tab { background-color: #e0e0e0; padding: 8px 12px; border-top-left-radius: 4px; border-top-right-radius: 4px; }
            QTabBar::tab:selected { background-color: #f0f0f0; border-bottom: 2px solid #5a9bd5; }
            QPushButton { font-size: 18px; font-weight: bold; background-color: #5a9bd5; color: white; padding: 8px 15px; border-radius: 6px; }
            QPushButton:hover { background-color: #7bb1eb; }
            QPushButton#primary { background-color: #28a745; padding: 10px 20px; }
            QPushButton#primary:hover { background-color: #45c159; }
            QPushButton#toggle { font-size: 20px; background-color: #6c757d; padding: 10px 25px; border-radius: 10px; }
            QPushButton#toggle:hover { background-color: #868e96; }
            QPushButton#danger { background-color: #dc3545; }
            QPushButton#danger:hover { background-color: #e25563; }
            QProgressBar { border: 1px solid #ccc; border-radius: 4px; text-align: center; }
            QProgressBar::chunk { background-color: #28a745; }
            QListWidget::item:selected { background-color: #5a9bd5; color: white; }
        """
        self.dark_style = """
            QMainWindow, QWidget { background-color: #2b2b2b; color: #fff; font-family: Segoe UI; font-size: 16px; }
            QTextEdit, QLineEdit, QListWidget { background-color: #3b3b3b; color: #fff; border: 1px solid #555; border-radius: 4px; padding: 5px; }
            QLabel { background-color: transparent; color: #fff; }
            QTabWidget::pane { border: 1px solid #555; background-color: #2b2b2b; }
            QTabBar::tab { background-color: #353535; color: #fff; padding: 8px 12px; border-top-left-radius: 4px; border-top-right-radius: 4px; }
            QTabBar::tab:selected { background-color: #2b2b2b; border-bottom: 2px solid #5a9bd5; }
            QPushButton { font-size: 18px; font-weight: bold; background-color: #5a9bd5; color: white; padding: 8px 15px; border-radius: 6px; }
            QPushButton:hover { background-color: #7bb1eb; }
            QPushButton#primary { background-color: #28a745; padding: 10px 20px; }
            QPushButton#primary:hover { background-color: #45c159; }
            QPushButton#toggle { font-size: 20px; background-color: #6c757d; padding: 10px 25px; border-radius: 10px; }
            QPushButton#toggle:hover { background-color: #868e96; }
            QPushButton#danger { background-color: #dc3545; }
            QPushButton#danger:hover { background-color: #e25563; }
            QProgressBar { border: 1px solid #555; border-radius: 4px; text-align: center; color: white; }
            QProgressBar::chunk { background-color: #28a745; }
            QListWidget::item:selected { background-color: #5a9bd5; color: white; }
        """
        self.setStyleSheet(self.light_style)

        self.attachments = []
        self.template_attachments = []
        self.email_thread = None
        self.current_logs = []

        self.initUI()

    def initUI(self):
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.main_tab = QWidget()
        self.templates_tab = QWidget()
        self.tabs.addTab(self.main_tab, "Main")
        self.tabs.addTab(self.templates_tab, "Templates")

        toggle_container = QWidget()
        toggle_layout = QHBoxLayout()
        toggle_layout.setContentsMargins(0,0,10,0)
        toggle_container.setLayout(toggle_layout)

        toggle_btn = QPushButton("Dark / Light Mode")
        toggle_btn.setObjectName("toggle")
        toggle_btn.clicked.connect(self.toggle_theme)
        toggle_layout.addWidget(toggle_btn)
        self.tabs.setCornerWidget(toggle_container, corner=Qt.Corner.TopRightCorner)

        self.init_main_tab()
        self.init_templates_tab()

    # ------------------- MAIN TAB -------------------
    def init_main_tab(self):
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
            smtp_layout.addWidget(QLabel(text), i//2, (i%2)*2)
            entry = QLineEdit()
            if "Password" in text:
                entry.setEchoMode(QLineEdit.EchoMode.Password)
            smtp_layout.addWidget(entry, i//2, (i%2)*2+1)
            self.entries.append(entry)

        self.entry_server, self.entry_port, self.entry_email, self.entry_password = self.entries

        # Set some default values for testing
        self.entry_server.setText("smtp.gmail.com")
        self.entry_port.setText("587")
        
        btn_validate_smtp = QPushButton("Validate SMTP Configuration")
        btn_validate_smtp.clicked.connect(self.validate_smtp)
        smtp_layout.addWidget(btn_validate_smtp, 2, 0, 1, 2)

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

        splitter.setSizes([350, 900])

    # ------------------- TEMPLATES TAB -------------------
    def init_templates_tab(self):
        layout = QHBoxLayout()
        self.templates_tab.setLayout(layout)

        # Left: Templates list
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        left_widget.setLayout(left_layout)
        left_widget.setMinimumWidth(300)
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

    # ------------------- FUNCTIONS -------------------
    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.setStyleSheet(self.dark_style if self.dark_mode else self.light_style)
        # Update previews background
        bg_color = "#3b3b3b" if self.dark_mode else "#fff"
        text_color = "#fff" if self.dark_mode else "#000"
        self.preview_body.setStyleSheet(f"background-color:{bg_color}; color:{text_color};")
        self.preview_temp_body.setStyleSheet(f"background-color:{bg_color}; color:{text_color};")

    def log(self, msg):
        self.text_logs.append(msg)

    def validate_smtp(self):
        """Validate SMTP configuration"""
        smtp_config = {
            "server": self.entry_server.text(),
            "port": int(self.entry_port.text()) if self.entry_port.text().isdigit() else 587,
            "email": self.entry_email.text(),
            "password": self.entry_password.text(),
            "tls": True
        }
        
        valid, message = backend.validate_smtp_config(smtp_config)
        if valid:
            QMessageBox.information(self, "SMTP Valid", message)
        else:
            QMessageBox.warning(self, "SMTP Invalid", message)

    def load_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load CSV/TXT", "", "CSV Files (*.csv);;Text Files (*.txt)")
        if path:
            try:
                count = backend.count_emails_in_file(path)
                if count == 0:
                    QMessageBox.warning(self, "No Valid Emails", "The selected file contains no valid email addresses.")
                    return
                
                if QMessageBox.question(self, "Confirm Load", 
                                       f"Found {count} valid emails. Load them?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
                    emails = backend.load_emails(path)
                    self.list_emails.clear()
                    self.list_emails.addItems(emails)
                    self.log(f"Loaded {len(emails)} emails from {os.path.basename(path)}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load file: {str(e)}")

    def remove_selected(self):
        for item in self.list_emails.selectedItems():
            self.list_emails.takeItem(self.list_emails.row(item))

    def clear_all_emails(self):
        if self.list_emails.count() > 0:
            if QMessageBox.question(self, "Confirm Clear", 
                                   "Clear all recipients?",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
                self.list_emails.clear()

    def add_attachment(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Attachments/Images")
        if files:
            self.attachments.extend(files)
            self.log(f"Added {len(files)} attachments/images.")

    def insert_image_main(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Image", "", "Images (*.png *.jpg *.jpeg *.gif *.bmp)")
        if file:
            html_tag = backend.inline_image(file)
            self.text_body.insertHtml(html_tag)

    def insert_image_template(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Image", "", "Images (*.png *.jpg *.jpeg *.gif *.bmp)")
        if file:
            html_tag = backend.inline_image(file)
            self.text_temp_body.insertHtml(html_tag)

    def render_preview_main(self):
        html_content = self.text_body.toHtml()
        self.preview_body.setHtml(html_content)

    def render_preview_template(self):
        html_content = self.text_temp_body.toHtml()
        self.preview_temp_body.setHtml(html_content)

    def send_email(self):
        recipients = [self.list_emails.item(i).text() for i in range(self.list_emails.count())]
        if not recipients:
            QMessageBox.warning(self, "No Recipients", "Load recipient emails first!")
            return
        
        # Validate SMTP configuration
        smtp_config = {
            "server": self.entry_server.text(),
            "port": int(self.entry_port.text()) if self.entry_port.text().isdigit() else 587,
            "email": self.entry_email.text(),
            "password": self.entry_password.text(),
            "tls": True
        }
        
        valid, message = backend.validate_smtp_config(smtp_config)
        if not valid:
            QMessageBox.warning(self, "SMTP Invalid", f"Cannot send emails: {message}")
            return
        
        subject = self.entry_subject.text()
        if not subject.strip():
            QMessageBox.warning(self, "No Subject", "Please enter an email subject!")
            return
        
        body = self.text_body.toHtml()
        if not body.strip():
            QMessageBox.warning(self, "No Body", "Please enter email content!")
            return
        
        # Start email sending thread
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.findChild(QPushButton, "primary").setEnabled(False)
        self.findChild(QPushButton, "danger").setVisible(True)
        
        self.email_thread = EmailThread(smtp_config, recipients, subject, body, self.attachments)
        self.email_thread.progress.connect(self.progress_bar.setValue)
        self.email_thread.log_signal.connect(self.log)
        self.email_thread.finished.connect(self.on_email_finished)
        self.email_thread.start()
        
        self.log("Started sending emails...")

    def stop_sending(self):
        if self.email_thread and self.email_thread.isRunning():
            self.email_thread.terminate()
            self.email_thread.wait()
            self.log("Email sending stopped by user")
            self.findChild(QPushButton, "primary").setEnabled(True)
            self.findChild(QPushButton, "danger").setVisible(False)
            self.progress_bar.setVisible(False)

    def on_email_finished(self, logs):
        self.current_logs.extend(logs)
        self.findChild(QPushButton, "primary").setEnabled(True)
        self.findChild(QPushButton, "danger").setVisible(False)
        self.progress_bar.setVisible(False)
        self.log("Email sending completed!")

    def export_logs(self):
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
        self.list_templates.clear()
        templates = backend.load_templates()
        for t in templates:
            self.list_templates.addItem(t["name"])

    def load_template(self):
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
        """Load selected template into the main tab"""
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
        selected = self.list_templates.selectedItems()
        if not selected:
            QMessageBox.warning(self, NO_SELECTION_TITLE, DELETE_NO_SELECTION_MSG)
            return
        
        name = selected[0].text()
        if QMessageBox.question(self, "Confirm Delete", 
                               f"Delete template '{name}'?",
                               QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            if backend.delete_template(name):
                self.refresh_templates()
                self.log(f"Deleted template '{name}'")
            else:
                QMessageBox.warning(self, "Error", f"Failed to delete template '{name}'")

    def add_template_attachment(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Attachments/Images")
        if files:
            self.template_attachments.extend(files)
            self.log(f"Added {len(files)} template attachments/images.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EmailApp()
    window.show()
    sys.exit(app.exec())