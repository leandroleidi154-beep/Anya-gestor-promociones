import requests
import sys

# IMPORTANTE: Reemplaza esta URL por el enlace real de tu app de Streamlit
URL = "https://anya-gestor-promociones-n8sxjsw9o5vyvz4hih7qhh.streamlit.app/"

try:
    # Envía una petición a la app con un límite de tiempo de 30 segundos
    response = requests.get(URL, timeout=30)
    print(f"Estado recibido: {response.status_code}")
    
    if response.status_code == 200:
        print("La app respondió correctamente y sigue despierta.")
    else:
        print(f"La app devolvió el código de estado: {response.status_code}")

except Exception as e:
    # Si hay un fallo de conexión, se captura para evitar que GitHub mande un mail de error
    print(f"Error al intentar conectar con la app: {e}")
    sys.exit(0)
