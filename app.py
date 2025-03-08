import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel
import os
import subprocess  # For opening files with default application
from create_schedule import create_original_schedule

class HomeScreen(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Create Schedule")

        self.layout = QVBoxLayout()

        # Button to upload Excel file
        self.upload_button = QPushButton("Upload Excel File")
        self.upload_button.clicked.connect(self.upload_file)

        # Label to show selected file
        self.file_label = QLabel("No file selected")

        # Button to generate schedule
        self.generate_button = QPushButton("Generate Schedule")
        self.generate_button.clicked.connect(self.generate_schedule)
        self.generate_button.setEnabled(False)  # Disable the button initially

        # Button to open the generated schedule
        self.open_button = QPushButton("Open Schedule")
        self.open_button.clicked.connect(self.open_schedule)
        self.open_button.setEnabled(False)  # Disable until schedule is generated

        self.layout.addWidget(self.upload_button)
        self.layout.addWidget(self.file_label)
        self.layout.addWidget(self.generate_button)
        self.layout.addWidget(self.open_button)

        self.setLayout(self.layout)

        self.uploaded_file = None
        self.generated_file = None

    def upload_file(self):
        # Open file dialog to select the Excel file
        file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")

        if file:
            self.uploaded_file = file
            self.file_label.setText(f"Selected file: {file}")
            self.generate_button.setEnabled(True)  # Enable generate button

    def generate_schedule(self):
        if not self.uploaded_file:
            self.file_label.setText("Error: No file uploaded.")
            return

        try:
            # Define input and output file paths
            input_file_path = self.uploaded_file
            output_file_path = r"C:\Users\morsh\Desktop\personal projects\soroka_solution\backend\generated_schedule.xlsx"

            # Call the scheduling logic
            create_original_schedule(input_file_path, output_file_path)

            # Save the generated schedule path
            self.generated_file = output_file_path

            # Enable the open button after the schedule is generated
            self.open_button.setEnabled(True)

            # Notify the user of the successful generation
            self.file_label.setText(f"Schedule generated: {self.generated_file}")

        except Exception as e:
            self.file_label.setText(f"Error: {str(e)}")

    def open_schedule(self):
        if not self.generated_file:
            self.file_label.setText("Error: No schedule generated to open.")
            return

        # Try to open the generated file with the default application
        try:
            if sys.platform == "win32":  # For Windows
                os.startfile(self.generated_file)
            elif sys.platform == "darwin":  # For macOS
                subprocess.call(["open", self.generated_file])
            else:  # For Linux
                subprocess.call(["xdg-open", self.generated_file])

            self.file_label.setText(f"Schedule opened: {self.generated_file}")

        except Exception as e:
            self.file_label.setText(f"Error opening schedule: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = HomeScreen()
    window.show()
    sys.exit(app.exec_())
