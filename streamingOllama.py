from ollama import Client
import requests

def test_ollama_connection(host: str) -> bool:
    """Check if the Ollama server is reachable."""
    try:
        r = requests.get(f"{host}/api/version", timeout=5)
        if r.status_code == 200:
            print(f"✅ Ollama server reachable: {r.json()}")
            return True
        else:
            print(f"❌ Server responded with status {r.status_code}")
            return False
    except Exception as e:
        print(f"❌ Connection error: {e}")
        return False


def main():
    host = "http://md3fgqdc:11434"

    if not test_ollama_connection(host):
        return

    client = Client(host=host)
    print("🧠 Sending streaming request to model 'granite4:tiny-h'...\n")

    try:
        # Streaming response
        stream = client.chat(
            model="granite4:tiny-h",
            messages=[{"role": "user", "content": "Why is the sky blue?"}],
            stream=True,
        )

        for chunk in stream:
            if "message" in chunk and "content" in chunk["message"]:
                print(chunk["message"]["content"], end="", flush=True)

        print("\n\n✅ Streaming finished.")
    except Exception as e:
        print(f"❌ Ollama streaming request failed: {e}")


if __name__ == "__main__":
    main()
