import os
import sqlite3
import datetime
import imaplib 
import sys
import email
import re
import sqlite3
from email.message import EmailMessage
from email.utils import parseaddr
from PyQt5.QtCore import Qt, QThread
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QComboBox,
    QGridLayout,
    QHBoxLayout,
    QProgressBar,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox
)

DB_FILE = "settings.sqlite3"

class Settings:
    """This class represents the settings for the email bot."""

    def __init__(self, server, username, password, port, security, output_dir, keywords):
        """Initialize the settings."""
        self.server = server
        self.username = username
        self.password = password
        self.port = port
        self.security = security
        self.output_dir = output_dir
        self.keywords = keywords

    def save(self):
        """Save the settings to the database."""
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

    @staticmethod
    def get():
        """Get the settings from the database."""
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
        output_dir_layout = QHBoxLayout()
        output_dir_layout.addWidget(output_dir_label)
        output_dir_layout.addWidget(self.output_dir_input)
        output_dir_button = QPushButton("Browse")
        output_dir_button.clicked.connect(self.browse_output_dir)
        output_dir_layout.addWidget(output_dir_button)    
        
        # Add the output directory button to the main layout.
        main_layout.addWidget(output_dir_button, 6, 0)

        # Create a horizontal layout for the buttons.        
        buttons_layout = QHBoxLayout()
        # Create the save, load, test connection, and check email for invoices buttons.

        save_button = QPushButton("Save")        
        load_button = QPushButton("Load")
        test_connection_button = QPushButton("Test Connection")        
        check_email_for_invoices_button = QPushButton("Check Email for Invoices")

        # Add the buttons to the horizontal layout.
        buttons_layout.addWidget(save_button)        
        buttons_layout.addWidget(load_button)
        buttons_layout.addWidget(test_connection_button)        
        buttons_layout.addWidget(check_email_for_invoices_button)
        
        # Add the horizontal layout to the main layout.
        main_layout.addLayout(buttons_layout, 7, 0)

        # Set the main layout of the window.        
        self.setLayout(main_layout)

        # Set the size of the window to 600x600 pixels.
        self.setFixedSize(600, 600)

        # Connect the buttons to their respective functions.        
        save_button.clicked.connect(self.save_settings)
        load_button.clicked.connect(self.load_settings)        
        test_connection_button.clicked.connect(self.test_connection)
        check_email_for_invoices_button.clicked.connect(self.check_email_for_invoices)

    def browse_output_dir(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ShowDirsOnly
        output_dir = QFileDialog.getExistingDirectory(self, "Select Directory", options=options)
        if output_dir:
            self.output_dir_input.setText(output_dir)

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
            self.keywords_input.setText(" ".join(settings.keywords))

        # Display a pop-up to confirm the loading
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("App loaded successfully.")
        msg.setWindowTitle("Confirmation")
        msg.exec_()

    def save_settings(self):
        # Get the email credentials
        server = self.server_input.text()
        username = self.username_input.text()
        password = self.password_input.text()
        security = self.encryption_input.currentText()
        port = int(self.port_input.text())

        # Get the path
        output_dir = self.output_dir_input.text()

        # Get the keywords
        keywords = self.keywords_input.text().split()

        # Validate the settings
        if not server or not username or not password:
            QMessageBox.critical(self, "Error", "Please enter a server, username, and password.")
            return

        if not output_dir:
            QMessageBox.critical(self, "Error", "Please enter a path.")
            return

        # Save the settings
        settings = Settings(server, username, password, port, security, output_dir, keywords)
        settings.save()

        QMessageBox.information(self, "Success", "Settings saved.")


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

        # Search for messages that contain the keywords
        status, messages = imap_server.search(None, "UNSEEN")         
        
        if status != "OK":         
          return False 

        # Iterate over the results and save the message IDs to a list
        invoice_count = 0
        for message_id in messages[0].split():
            # Fetch the bodies of the messages that were found
            status, data = imap_server.fetch(message_id, '(RFC822)')
            if status != 'OK':
                print('Error fetching message {}.'.format(message_id))
                continue

            message = email.message_from_bytes(data[0][1])

            # Extract the sender and the file extension
            sender = message.get('From')
            attachment = message.get_payload()[0]
            filename = attachment.get_filename()
            if not filename:
                continue

            extension = os.path.splitext(filename)[1][1:].lower()
            if extension not in ["pdf", "xls", "xlsx", "doc"]:
                continue

            # Create a new folder for each sender
            folder_name = os.path.join(settings.output_dir, sender)
            if not os.path.exists(folder_name):
                os.mkdir(folder_name)


        # Logout of email server
        imap_server.logout()

        # Show a message box with the number of invoices found
        if invoice_count == 0:
            QMessageBox.information(self, "Success", "Found {} invoices.".format(invoice_count))
        else:
            QMessageBox.information(self, "No invoices found", "No invoices were found with the given keywords.")

  
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    window = EmailBot()
    
    window.show()

    sys.exit(app.exec_())