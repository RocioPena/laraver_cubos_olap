import requests

url = "http://localhost:8080/medidas"  # IP del dispositivo 
try:
    response = requests.get(url, timeout=10) 
    if response.status_code == 200:
        data = response.json()
        print("✅ Datos recibidos:", data[:3])
    else:
        print(f"❌ Error de respuesta: {response.status_code}")
except Exception as e:
    print(f"❌ Error al conectar con el servidor: {e}")
