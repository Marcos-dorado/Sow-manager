import sys
import os

# Asegurar que el directorio raíz esté en el path
# Funciona tanto como script como bundle de PyInstaller
if getattr(sys, "frozen", False):
    app_dir = os.path.dirname(sys.executable)
else:
    app_dir = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, app_dir)

import tkinter as tk
from src.ui import AppSOWs

if __name__ == "__main__":
    root = tk.Tk()
    app  = AppSOWs(root)
    root.mainloop()