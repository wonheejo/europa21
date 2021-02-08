import pywinauto
import time
app = pywinauto.application.Application()
print(app.Properties.print_control_identifiers())