from PyQt5.QtWidgets import  QSpacerItem, QAction, QFormLayout, QApplication, QMenuBar, QMainWindow, QFileDialog, QComboBox, QListWidget, QPushButton, QLabel, QTextEdit, QVBoxLayout, QHBoxLayout, QWidget, QMessageBox, QDialog
from PyQt5.QtCore import pyqtSignal, QObject, Qt
import sys


from .extract_functions import main


class OptionWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Select Program Type")

        layout = QVBoxLayout(self)

        # Add a spacer item to create space between the menu bar and the select dropdown
        spacer = QSpacerItem(20, 20)
        layout.addItem(spacer)

        # Add a form layout to neatly align the label and combo box
        form_layout = QFormLayout()
        layout.addLayout(form_layout)

        # Add a label for the program type combo box
        self.label = QLabel("Select Program Type:")
        self.program_type_combo = QComboBox()
        # self.program_type_combo.addItem("Extraction Data From Competition")
        self.program_type_combo.addItem("Extract Data from PDF")
 

        # Add the label and combo box to the form layout
        form_layout.addRow(self.label, self.program_type_combo)

        self.button_box = QHBoxLayout()
        self.button_box.setSpacing(10)

        self.select_button = QPushButton("Select")
        self.cancel_button = QPushButton("Cancel")
        self.button_box.addWidget(self.select_button)
        self.button_box.addWidget(self.cancel_button)
        layout.addLayout(self.button_box)

        self.select_button.clicked.connect(self.select_program_type)
        self.cancel_button.clicked.connect(self.reject)
        
        self.program_type = ""

        # Add instructions for the QComboBox
        self.program_type_combo.setWhatsThis("Select the type of program you want to run.")

        # Add instructions for the QPushButton
        self.select_button.setWhatsThis("Click this button to confirm your selection.")
        self.cancel_button.setWhatsThis("Click this button to cancel and close the window.")

        

    def select_program_type(self):
        self.program_type = self.program_type_combo.currentText()
        self.accept()

    
class DataExtractorWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Data Extractor")
        self.resize(800, 800)  # Set the initial dimensions of the window


        # Create widgets
        self.file_list_widget = QListWidget()
        self.output_file_label = QLabel()
        self.exit_button = QPushButton("Exit")
        self.clear_button = QPushButton("Clear")
        self.run_button = QPushButton("Run")

        self.output_text_edit = QTextEdit()  # Widget to display printed output

        # Configure layout
        central_widget = QWidget()

        layout = QVBoxLayout(central_widget)
        layout.addSpacing(50)

        select_pdf_label = QLabel("Select PDF Files:")
        layout.addWidget(select_pdf_label)
        layout.addWidget(self.file_list_widget)
        layout.addSpacing(10)
        self.browse_files_button = QPushButton("Select Files")
        layout.insertWidget(5, self.browse_files_button)

        layout.addSpacing(50)
        output_file_label = QLabel("Select Output Excel Files:")
        layout.addWidget(output_file_label)
        layout.addWidget(self.output_file_label)
        self.browse_folder_button = QPushButton("Browse")
        layout.insertWidget(8, self.browse_folder_button)


        layout.addSpacing(10)
        # Add the output text edit widget
        layout.addWidget(self.output_text_edit)

        # Create a QHBoxLayout for the buttons
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.exit_button)
        button_layout.addWidget(self.clear_button)
        button_layout.addWidget(self.run_button)

        # Add the button layout to the main layout
        layout.addLayout(button_layout)

        # Set central widget and connect signals
        self.setCentralWidget(central_widget)
        self.exit_button.clicked.connect(self.close)
        self.clear_button.clicked.connect(self.clear_fields)
        self.run_button.clicked.connect(self.run_process)
        

        # Initialize attributes
        self.files_list = []
        self.output_file = ""

        # Set the color of the "Run" button to light green
        self.run_button.setStyleSheet("background-color: lightgreen;")
        layout = self.centralWidget().layout()
        
    def clear_fields(self):
        self.file_list_widget.clear()
        self.output_file_label.clear()
        self.files_list = []
        self.output_file = ""

    def browse_files(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        file_dialog.setNameFilter("PDF Files (*.pdf);;All Files (*)")  # Add the PDF filter
        if file_dialog.exec_():
            file_names = file_dialog.selectedFiles()
            self.files_list = file_names
            self.file_list_widget.clear()
            self.file_list_widget.addItems(file_names)

    def browse_folder(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.AnyFile)
        file_dialog.setNameFilter("Excel Files (*.xlsx);;All Files (*)")
        if file_dialog.exec_():
            file_path = file_dialog.selectedFiles()[0]
            self.output_file = file_path
            self.output_file_label.setText(file_path)

    def run_process(self):
        if not self.files_list or not self.output_file:
            QMessageBox.warning(
                self,
                "Missing Inputs",
                "Please select PDF files, output excel file."
            )
            return

        loading_dialog = QDialog(self)
        loading_dialog.setWindowTitle("Data Transfer in Progress")
        loading_label = QLabel("Please wait while the data is being extracted...")
        layout = QVBoxLayout()
        layout.addWidget(loading_label)
        loading_dialog.setLayout(layout)
        loading_dialog.setFixedSize(300, 100)
        loading_dialog.setWindowModality(Qt.ApplicationModal)
        loading_dialog.show()


        # Redirect the output to the text edit widget
        sys.stdout = EmittingStream(textWritten=self.output_text_edit.append)
       
        try:
            # Perform your data transfer process here
            QApplication.processEvents()  # Allow the loading dialog to show up

            main(self.files_list, self.output_file)

            # After the process is completed, close the loading dialog
            loading_dialog.accept()

            # Show a completion message
            QMessageBox.information(
                self, "Process Completed", "All done!\nGo to the output folder to see the output file."
            )
            self.clear_fields()
        except PermissionError as pr:
            QMessageBox.critical(self, "Error", f"PermissionError: Please close this file \"{pr.filename}\" and try again.")
        except FileNotFoundError as ferr:
            QMessageBox.critical(self, "Error", f"FileNotFoundError: Problem with the following directory,\n{ferr.filename}\nCheck if this is the correct directory for your raw data and Excel file.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred: {str(e)}")
        finally:
            loading_dialog.accept()

    def setup_ui(self):
        
        
        self.browse_files_button.clicked.connect(self.browse_files)
        self.browse_folder_button.clicked.connect(self.browse_folder)

        
        
        

        # Create a menu bar
        self.menu_bar = QMenuBar(self)
        self.more_menu = self.menu_bar.addMenu("More Info")

        # Create a QAction for "How To Use" menu item
        self.more_action = QAction("How To Use", self)
        self.more_menu.addAction(self.more_action)
        self.more_action.triggered.connect(self.show_howtouse)  

    def show_howtouse(self):
        how_to_use_text = (
            "Welcome to the Data Transfer Application!\n\n"
            "To use this application, follow these steps:\n\n"
            "1. Click the 'Select PDF' button to choose the PDF file containing the data recived from projects reportersdd.\n\n"
            "2. Click the 'Select Genome Canada Excel Output' button to choose the Excel file that will contain all project details and the one you will send it to Genome Canada.\n\n"
            "3. Click the 'Select project Excel Output' button to choose the Excel file that will contain only same project details and you will keep it in your records.\n\n"
            "4. Once you've selected the necessary files, you can click the 'Run' button to start the data transfer process.\n\n"
            "5. If transfer success, \"Process Completed\" messages will displayed in the text area at the bottom.\n\n"
            "6. If you encounter any errors or issues, the application will display an appropriate error message.\n\n"
            "7. You can clear the selected files and output by clicking the 'Clear' button.\n\n"
            "8. When you're done using the application, click the 'Exit' button to close it.\n\n"
            "Enjoy using the Data Transfer Application!"
        )

        documentation_box = QMessageBox(self)
        documentation_box.setWindowTitle("How To Use")
        documentation_box.setIcon(QMessageBox.Information)
        documentation_box.setText(how_to_use_text)
        documentation_box.setStandardButtons(QMessageBox.Ok)
        documentation_box.exec_()


# Custom stream class to redirect printed output to the text edit widget
class EmittingStream(QObject):
    textWritten = pyqtSignal(str)

    def write(self, text):
        self.textWritten.emit(str(text))

    def flush(self):
        pass
