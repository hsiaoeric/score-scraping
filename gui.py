import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QLineEdit

# Define the main window class
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        # Set window title
        self.setWindowTitle('自動查詢程式')
        layout = QVBoxLayout()

        # Set window dimensions
        self.setGeometry(100, 100, 300, 200)

        self.label = QLabel("請選擇excel", self)
        layout.addWidget(self.label)

        # self.line_edit = QLineEdit(self)
        # layout.addWidget(self.line_edit)


        
        # Create a button
        self.button = QPushButton('開始自動查榜', self)
        layout.addWidget(self.button)
        self.button.setGeometry(100, 80, 100, 30)

        # Connect button click to a method
        self.button.clicked.connect(self.on_button_click)

    def on_button_click(self):
        print("Button clicked!")

# Main function to execute the application
def main():
    # Create an application object
    app = QApplication(sys.argv)

    # Create an instance of the main window
    window = MainWindow()

    # Show the window
    window.show()

    # Start the application's event loop
    sys.exit(app.exec_())

# Run the main function
if __name__ == '__main__':
    main()