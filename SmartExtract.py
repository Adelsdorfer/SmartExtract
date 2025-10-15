import requests
import json

# Adresse deines Ollama-Servers
OLLAMA_SERVER = "http://md3fgqdc:11434"  # oder http://10.176.176.73:11434

# Modell
#MODEL = "qwen3:8b"
#MODEL = "gemma3:4b"
#MODEL = "deepseek-r1:8b"
#MODEL = "phi4-mini-reasoning"
MODEL = "granite4:tiny-h"
#MODEL = "deepseek-r1:8b"

prompt = "Schreibe einen Witz über künstliche Intelligenz auf Deutsch mit ca 137 wörter."
#prompt = "Berechne die Höhe eines Turms, von dem ein Stein fallen gelassen wird, wenn man den Aufprall nach 5 Sekunden hört."
def test_ollama_api():
    url = f"{OLLAMA_SERVER}/api/generate"
    payload = {
        "model": MODEL,
        "prompt": prompt,
        "stream": False  # Kein Stream, einfache Antwort
    }

    try:
        # 'json=payload' setzt automatisch den Header Content-Type: application/json
        response = requests.post(url, json=payload, timeout=560)
        response.raise_for_status()
        data = response.json()
        print("✅ Verbindung erfolgreich!")
        print("Antwort vom Modell:")
        print(data.get("response", "Keine Antwort erhalten."))
    except requests.exceptions.HTTPError as e:
        print("❌ HTTP-Fehler:", e.response.status_code)
        print("Antworttext:", e.response.text)
    except requests.exceptions.RequestException as e:
        print("❌ Verbindungsfehler:", e)

if __name__ == "__main__":
    test_ollama_api()
