import os
import sqlite3
import imaplib 
import sys
import email
import re
import sqlite3
from email.message import EmailMessage
from email.header import decode_header
from PyQt5.QtWidgets import * 
from PyQt5.QtGui import * 
from PyQt5.QtCore import *

current_dir = os.getcwd()

DB_FILE = os.path.join(current_dir, "settings.sqlite3")

class Settings:

    def __init__(self, server, username, password, port, security, output_dir, keywords):
        self.server = server
        self.username = username
        self.password = password
        self.port = port
        self.security = security
        self.output_dir = output_dir
        self.keywords = keywords

    def save(self):
        # Check if the database exists.
        if not os.path.exists(DB_FILE):
            # Create the database.
            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                c.execute('''
                    CREATE TABLE settings (
                        server TEXT,
                        username TEXT,
                        password TEXT,
                        port INTEGER,
                        security TEXT,
                        output_dir TEXT,
                        keywords TEXT
                    )
                ''')

        # Connect to the database.
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("""
                INSERT INTO settings (server, username, password, port, security, output_dir, keywords)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (self.server, self.username, self.password, self.port, self.security, self.output_dir, ', '.join(self.keywords)))
            conn.commit()

    def update(self):
        # Check if the database exists.
        if not os.path.exists(DB_FILE):
            # Create the database.
            self.create_database()

        # Connect to the database.
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("""
                UPDATE settings
                SET server = ?,
                    username = ?,
                    password = ?,
                    port = ?,
                    security = ?,
                    output_dir = ?,
                    keywords = ?
            """, (self.server, self.username, self.password, self.port, self.security, self.output_dir, self.keywords))
            conn.commit() 

    @staticmethod
    def get():
        # Check if the database exists.
        if not os.path.exists(DB_FILE):
            return None

        # Connect to the database.
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            c.execute("SELECT * FROM settings")
            row = c.fetchone()

        # Close the connection to the database.
        conn.close()

        # If there are no settings in the database, return None.
        if row is None:
            return None

        # Create a Settings object from the row.
        settings = Settings(*row)

        # Return the Settings object.
        return settings
    
class EmailBot(QWidget):
    def __init__(self):
        super().__init__()

        # Initialize UI
        self.init_ui()

        # Load email credentials
        self.load_settings()

    def init_ui(self):
         # Create a grid layout for the main window.
        main_layout = QGridLayout()

        # Create a label and text field for the server input.
        server_label = QLabel("Server:")
        self.server_input = QLineEdit()
        main_layout.addWidget(server_label, 0, 0)
        main_layout.addWidget(self.server_input, 0, 1)

        # Create a label and text field for the username input.
        username_label = QLabel("Username:")
        self.username_input = QLineEdit()
        main_layout.addWidget(username_label, 1, 0)
        main_layout.addWidget(self.username_input, 1, 1)

        # Create a label and password field for the password input.
        password_label = QLabel("Password:")
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        main_layout.addWidget(password_label, 2, 0)
        main_layout.addWidget(self.password_input, 2, 1)

        # Create a label and text field for the encryption input.
        encryption_label = QLabel("Encryption:")
        encryption_options = ["SSL", "TLS"]
        self.encryption_input = QComboBox()
        self.encryption_input.addItems(encryption_options)
        main_layout.addWidget(encryption_label, 3, 0)
        main_layout.addWidget(self.encryption_input, 3, 1)

        # Create a label and text field for the port input.
        port_label = QLabel("Port:")
        self.port_input = QLineEdit()
        main_layout.addWidget(port_label, 4, 0)
        main_layout.addWidget(self.port_input, 4, 1)

        # Create a label and text field for the keywords input.
        keywords_label = QLabel("Keywords:")
        self.keywords_input = QLineEdit()
        main_layout.addWidget(keywords_label, 5, 0)
        main_layout.addWidget(self.keywords_input, 5, 1)

        # Create a label and text field for the output directory input.    
        output_dir_label = QLabel("Output Directory:")
        self.output_dir_input = QLineEdit()
        output_dir_button = QPushButton("Browse")
        output_dir_button.setToolTip("Select output directory")
        output_dir_layout = QHBoxLayout()
        output_dir_layout.addWidget(output_dir_label)
        output_dir_layout.addWidget(self.output_dir_input)

        # Add QLabel to display output directory path
        self.output_dir_path_label = QLabel("No output directory selected")
        output_dir_layout.addWidget(self.output_dir_path_label)

        output_dir_button.clicked.connect(self.browse_output_dir)
        output_dir_layout.addWidget(output_dir_button)
        
        # Add the output directory button to the main layout.
        main_layout.addWidget(output_dir_button, 6, 0)

        # Create a horizontal layout for the buttons.        
        buttons_layout = QHBoxLayout()
         
        # Create a label and add the logo image to it.
        logo_label = QLabel()
        logo_image = QPixmap(os.path.join(current_dir, "logo.png"))
        logo_label.setPixmap(logo_image)
        main_layout.addWidget(logo_label, 0, 2, 7, 1)

        self.setWindowIcon(QIcon(os.path.join(current_dir, "logo.png")))

        save_button = QPushButton("Save")
        save_button.setToolTip("Save settings")
        buttons_layout.addWidget(save_button)        
        save_button.clicked.connect(self.save_settings)
        
        load_button = QPushButton("Load")
        load_button.setToolTip("Load settings")
        buttons_layout.addWidget(load_button)
        load_button.clicked.connect(self.load_settings)        
        
        test_connection_button = QPushButton("Test Connection")
        test_connection_button.setToolTip("Test connection")
        buttons_layout.addWidget(test_connection_button)        
        test_connection_button.clicked.connect(self.test_connection)
        
        check_email_for_invoices_button = QPushButton("Check Email for Documents")
        check_email_for_invoices_button.setToolTip("Check email for Documents")
        buttons_layout.addWidget(check_email_for_invoices_button)
        check_email_for_invoices_button.clicked.connect(self.check_email_for_invoices)

        # Add the horizontal layout to the main layout.
        main_layout.addLayout(buttons_layout, 7, 0)

        # Set the main layout of the window.        
        self.setLayout(main_layout)

        # Set the size of the window to 600x600 pixels.
        self.setFixedSize(600, 600)

        # Get the screen resolution.
        screen_resolution = QApplication.desktop().screenGeometry()

        # Calculate the center of the screen.
        center_x = int((screen_resolution.width() - self.width()) / 2)
        center_y = int((screen_resolution.height() - self.height()) / 2)

        # Move the window to the center of the screen.
        self.move(center_x, center_y)

        self.setWindowTitle("Email Documents Downloader Settings")


    def browse_output_dir(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ShowDirsOnly
        output_dir = QFileDialog.getExistingDirectory(self, "Select Directory", options=options)
        if output_dir:
            self.output_dir_input.setText(output_dir)
            self.output_dir_path_label.setText(output_dir)

    def test_connection(self):
        # Load settings
        settings = Settings.get()
        if settings is None:
            QMessageBox.critical(self, "Error", "Please enter email credentials.")
            return

        # Connect to email server
        try:
            imap_server = imaplib.IMAP4_SSL(settings.server)
            imap_server.login(settings.username, settings.password)
            imap_server.select("inbox")
            QMessageBox.information(self, "Success", "Connection successful.")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            return

        # Logout of email server
        imap_server.logout()

    def load_settings(self):
       
        # Load the settings
        settings = Settings.get()

        # Populate the email credentials fields if they exist
        if settings is not None:
            self.server_input.setText(settings.server)
            self.username_input.setText(settings.username)
            self.password_input.setText(settings.password)
            self.encryption_input.setCurrentText("SSL" if settings.security else "TLS")
            self.port_input.setText(str(settings.port))

        # Populate the path field if it exists
        if settings is not None and settings.output_dir is not None:
            self.output_dir_input.setText(settings.output_dir)

        # Populate the keywords field if it exists
        if settings is not None and settings.keywords is not None:
            self.keywords_input.setText(settings.keywords)

        # Display a pop-up to confirm the loading
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("App loaded successfully.")
        msg.setWindowTitle("Documents Extrator")
        msg.exec_()

    def save_settings(self):
        # Get the email credentials
        server = self.server_input.text()
        username = self.username_input.text()
        password = self.password_input.text()
        security = self.encryption_input.currentText()
        port = int(self.port_input.text())
        output_dir = self.output_dir_input.text()
        keywords = self.keywords_input.text()

        # Validate the settings
        if not server or not username or not password:
            QMessageBox.critical(self, "Error", "Please enter a server, username, and password.")
            return

        if not output_dir:
            QMessageBox.critical(self, "Error", "Please enter a path.")
            return

        # Save the settings
        settings = Settings(server, username, password, port, security, output_dir, keywords)
        settings.update()

        QMessageBox.information(self, "Success", "Settings updated.")

    def check_email_for_invoices(self, keywords):
        # Load settings
        settings = Settings.get()
        if settings is None:
            QMessageBox.critical(self, "Error", "Please enter email credentials.")
            return

        # Connect to email server
        try:
            imap_server = imaplib.IMAP4_SSL(settings.server)
            imap_server.login(settings.username, settings.password)
            imap_server.select("inbox")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            return

        # Search for emails with attachments and keywords
        keywords = settings.keywords if settings.keywords else []
        search_query = " ".join(["(ALL)", "OR"] + [f'(SUBJECT "{k}")' for k in keywords] + [""])
        result, email_ids = imap_server.search(None, search_query)
 
        if result == "OK":
           # Process each email with attachments and keywords
           num_emails = len(email_ids[0].split())
           progress_dialog = QProgressDialog("Saving Documents...", "Cancel", 0, num_emails, self)
           progress_dialog.setWindowModality(Qt.WindowModal)
           progress_dialog.setWindowTitle("Checking Documents")
           progress_dialog.show()
           progress = 0

           for email_id in email_ids[0].split():
                if progress_dialog.wasCanceled():
                    break

                result, email_data = imap_server.fetch(email_id, "(RFC822)")
                email_message = email.message_from_bytes(email_data[0][1])
                attachments = self.get_attachments_from_email(email_message)
                if attachments:
                    # Extract sender email
                    sender = email_message["From"]
                    email_address = re.findall(r'<(.+?)>', sender)[0]

                    # Create folder for sender
                    output_dir = os.path.join(settings.output_dir, email_address)
                    if not os.path.exists(output_dir):
                        os.mkdir(output_dir)

                    for attachment in attachments:
                        # Check if file is a document
                        if attachment[0].split(".")[-1] in ["pdf", "xls", "xlsx","csv", "doc"]:
                            attachment_path = os.path.join(output_dir, attachment[0])
                            with open(attachment_path, "wb") as f:
                                f.write(attachment[1])

                    imap_server.store(email_id, "+FLAGS", "\\Seen")
                    progress += 1
                    progress_dialog.setValue(progress)
                    progress_dialog.setLabelText(f"Saving document from {email_address}...")

           progress_dialog.close()
           imap_server.logout()

    def get_attachments_from_email(self, email_message):
        attachments = []
        for part in email_message.walk():
            if part.get("Content-Disposition") is None:
                continue
            filename = decode_header(part.get_filename())[0][0]
            if isinstance(filename, bytes):
                filename = filename.decode()
            attachment = (filename, part.get_payload(decode=True))
            attachments.append(attachment)
        return attachments
    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    window = EmailBot()
    
    window.show()

    sys.exit(app.exec_())
