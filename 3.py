import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QVBoxLayout
from pylibdmtx.pylibdmtx import encode
class SimpleApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Simple PyQt5 App")
        self.setGeometry(100, 100, 300, 150)

        self.label = QLabel("Hello, PyQt5!", self)
        self.button = QPushButton("Click Me", self)
        self.button.clicked.connect(self.on_click)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.button)

        self.setLayout(layout)

    def on_click(self):
        self.label.setText("Button clicked!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SimpleApp()
    window.show()
    sys.exit(app.exec_())
