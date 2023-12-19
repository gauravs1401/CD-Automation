import sys
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel,
    QPushButton, QVBoxLayout, QWidget, QGridLayout
)
from PyQt5.QtGui import QPixmap, QFont
from PyQt5.QtCore import Qt


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.background_label = None
        self.setWindowTitle("Hewlett Packard Enterprise")
        self.setWindowState(Qt.WindowMaximized)

        # Set background image
        self.set_background_image("HPE_Wallpapers_2022_4K_3840x2160px_05.jpg")

        # Create a central widget and set the layout
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # Add header label
        header_label = QLabel("CONSOLIDATED DELIVERY AUTOMATION", self)
        header_label.setAlignment(Qt.AlignCenter)
        header_font = QFont("Arial", 40, QFont.Bold)  # Decreased font size for visibility
        header_label.setFont(header_font)
        header_label.setStyleSheet("color: white;")
        self.main_layout.addWidget(header_label)

        # Create a grid layout
        self.grid_layout = QGridLayout()
        self.grid_layout.setAlignment(Qt.AlignCenter)
        self.main_layout.addLayout(self.grid_layout)  # Add the grid layout to the main layout
        self.grid_layout.setVerticalSpacing(10)  # Adjust the vertical spacing
        self.grid_layout.setHorizontalSpacing(40)  # Adjust the horizontal spacing

        # Add buttons
        self.add_button("Execute", "CD-S4.py", 0)
        self.add_button("Execute - Alletra", "CD-S4 (Alletra).py", 1),


    def set_background_image(self, image_path):
        self.background_label = QLabel(self)
        pixmap = QPixmap(image_path)
        self.background_label.setPixmap(pixmap)
        self.background_label.setScaledContents(True)

    def add_button(self, button_text, python_file, column):
        button = QPushButton(button_text, self)
        button.setStyleSheet(
            "QPushButton {"
            "background-color: #6633BC;"
            "color: white;"
            "font-size: 24px;"  # Increase font size
            "font-family: Calibri Light;"
            "border-radius: 15px;"
            "padding: 20px 40px;"  # Increase padding
            "}"
            "QPushButton:hover {"
            "background-color: darkgray;"
            "}"
            "QPushButton:pressed {"
            "background-color: gray;"
            "}"
        )

        # Connect the button's clicked signal to the run_python_file slot
        button.clicked.connect(lambda: self.run_python_file(python_file))

        # Add the button to the grid layout
        self.grid_layout.addWidget(button, 0, column, alignment=Qt.AlignCenter)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.background_label is not None:
            # Resize the background label to fit the maximized window
            self.background_label.setGeometry(0, 0, self.width(), self.height())

    def run_python_file(self, python_file):
        command = ["python", python_file]
        subprocess.run(command)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
