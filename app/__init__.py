from PyQt5.QtWidgets import QDialog,QApplication


from .program_ui import DataExtractorWindow, OptionWindow

app = QApplication([])

option_window = OptionWindow()
result = option_window.exec_()
program_type = option_window.program_type


if result == QDialog.Accepted:
    if program_type == "Extract Data from PDF":
    
        window = DataExtractorWindow()
        window.setup_ui()
        window.show()
        app.exec_()


