"""Punto de entrada web — AutoSky Inventario v3.70"""
import subprocess, sys, os, webbrowser, time

PORT = 8502  # Puerto cambiado — evita colisión con sesiones anteriores en 8501
os.chdir(os.path.dirname(os.path.abspath(__file__)))

cmd = [
    sys.executable, "-m", "streamlit", "run", "app_web.py",
    "--server.port", str(PORT),
    # Bind solo IPv4 (evita que Safari/iOS intente IPv6 primero y espere timeout
    # de Happy Eyeballs durante minutos antes de caer a IPv4)
    "--server.address", "0.0.0.0",
    "--server.headless", "true",
    # Desactivar XSRF/CORS: Safari + WebSocket de Streamlit se atascan con la
    # protección XSRF en LAN (la app no expone nada sensible a internet)
    "--server.enableXsrfProtection", "false",
    "--server.enableCORS", "false",
    "--browser.gatherUsageStats", "false",
    "--theme.base", "light",
    "--theme.primaryColor", "#0ea5e9",
    "--theme.backgroundColor", "#f0f9ff",
    "--theme.secondaryBackgroundColor", "#ffffff",
    "--theme.textColor", "#0f172a",
]

print(f"Iniciando AutoSky Inventario v3.70 en http://localhost:{PORT}")
proc = subprocess.Popen(cmd)
time.sleep(4)
webbrowser.open(f"http://localhost:{PORT}")
proc.wait()
