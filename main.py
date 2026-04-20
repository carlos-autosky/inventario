
import sys
import os

# Asegura que el directorio del script esté en sys.path
# para que 'from app.xxx import ...' funcione sin importar
# desde dónde se llame a Python.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.ui import InventoryApp

if __name__ == "__main__":
    app = InventoryApp()
    app.mainloop()
