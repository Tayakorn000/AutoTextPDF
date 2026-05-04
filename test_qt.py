
import sys
from PySide6.QtWidgets import QApplication, QLabel
app = QApplication(sys.argv)
label = QLabel("Hello World")
label.show()
print("Label shown")
# Exit after 2 seconds to not block
from PySide6.QtCore import QTimer
QTimer.singleShot(2000, app.quit)
sys.exit(app.exec())
