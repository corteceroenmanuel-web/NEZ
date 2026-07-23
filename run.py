import webview
import threading
import os
from app import app

def iniciar_servidor():
    app.run(port=5000, debug=False)

if __name__ == '__main__':
    t = threading.Thread(target=iniciar_servidor)
    t.daemon = True
    t.start()

    # Ruta a tu icono .ico o .png
    ruta_icono = os.path.join(os.path.dirname(__file__), 'cochete.ico')

    # 1. Creamos la ventana SIN el argumento 'icon' aquí
    webview.create_window(
        'Sistema de Gestión y Finanzas 🚀',
        'http://127.0.0.1:5000',
        width=1280,
        height=800,
        resizable=True
    )
    
    # 2. Le pasamos el icono a webview.start()
    # (Si la imagen existe la usa, de lo contrario abre la ventana normalmente)
    if os.path.exists(ruta_icono):
        webview.start(icon=ruta_icono)
    else:
        webview.start()