import sys
import os
import threading
import time
import socket
import subprocess
import urllib.request
import uvicorn
import webview
from app import app, load_projects_data

def kill_previous_instances():
    """Cierra cualquier proceso en segundo plano de Correspon para evitar conflictos o bloqueos."""
    current_pid = os.getpid()
    
    # 1. Liberar puerto 8000 en Windows si un proceso anterior lo tiene ocupado
    if sys.platform == "win32":
        try:
            output = subprocess.check_output('netstat -ano | findstr :8000', shell=True).decode()
            for line in output.strip().splitlines():
                parts = line.split()
                if len(parts) >= 5 and 'LISTENING' in parts:
                    pid = int(parts[-1])
                    if pid != current_pid and pid > 0:
                        subprocess.run(f'taskkill /F /PID {pid}', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass

        # 2. Terminar instancias duplicadas de Correspon.exe
        try:
            output = subprocess.check_output('tasklist /FI "IMAGENAME eq Correspon.exe" /FO CSV /NH', shell=True).decode()
            for line in output.strip().splitlines():
                parts = line.split(',')
                if len(parts) >= 2:
                    pid_str = parts[1].replace('"', '').strip()
                    if pid_str.isdigit():
                        pid = int(pid_str)
                        if pid != current_pid and pid > 0:
                            subprocess.run(f'taskkill /F /PID {pid}', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass

def find_free_port(default_port=8000):
    """Encuentra un puerto libre comenzando desde default_port."""
    for port in range(default_port, default_port + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(('127.0.0.1', port))
                return port
            except OSError:
                continue
    return default_port

def wait_for_server(url, max_retries=60):
    """Espera activamente a que el servidor FastAPI responda HTTP 200."""
    for _ in range(max_retries):
        try:
            with urllib.request.urlopen(url, timeout=0.5) as response:
                if response.status == 200:
                    return True
        except Exception:
            time.sleep(0.1)
    return False

def start_server(port):
    """Ejecuta el servidor FastAPI."""
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="error")

def on_window_closed():
    """Garantiza la terminación completa e inmediata de todos los procesos al cerrar la ventana."""
    os._exit(0)

if __name__ == "__main__":
    kill_previous_instances()
    load_projects_data()

    port = find_free_port(8000)
    server_url = f"http://127.0.0.1:{port}"

    # Iniciar servidor FastAPI en hilo secundario
    server_thread = threading.Thread(target=start_server, args=(port,), daemon=True)
    server_thread.start()

    # Esperar a que el servidor responda
    ready = wait_for_server(server_url)
    if not ready:
        print("Error: No se pudo conectar al servidor local.")
        sys.exit(1)

    # Abrir la ventana webview
    window = webview.create_window(
        title="Correspon - Automatizador de Documentos",
        url=server_url,
        width=1380,
        height=900,
        min_size=(900, 600),
        resizable=True,
        confirm_close=False
    )
    
    window.events.closed += on_window_closed
    webview.start()
    os._exit(0)
