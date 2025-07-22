import json
import threading
import time
from seleniumwire import webdriver

# === CONFIG ===
url_to_open = "URL"
output_file = "captured_requests.json"

# === GLOBALS ===
captured_data = []
seen_requests = set()
stop_capture = False

# === FUNCTION: Listen for "stop" command ===
def listen_for_stop():
    global stop_capture
    while True:
        user_input = input("Type 'stop' to end capturing and save the file:\n>>> ")
        if user_input.strip().lower() == "stop":
            stop_capture = True
            break

# === FUNCTION: Start capturing requests ===
def capture_requests():
    global stop_capture
    while not stop_capture:
        for request in driver.requests:
            if request.response and request.id not in seen_requests:
                seen_requests.add(request.id)
                try:
                    captured_data.append({
                        "url": request.url,
                        "method": request.method,
                        "headers": dict(request.headers),
                        "response_status": request.response.status_code,
                        "response_headers": dict(request.response.headers),
                        "body": request.body.decode(errors='ignore') if request.body else "",
                    })
                except Exception as e:
                    print(f"[!] Error capturing request: {e}")
        time.sleep(2)

# === SETUP SELENIUM DRIVER ===
print("🔄 Launching browser...")
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)

driver.get(url_to_open)
print(f"🌐 Opened {url_to_open}")
print("🟢 You can now interact with the page (e.g. type captcha, click buttons).")

# === START THREADS ===
stop_thread = threading.Thread(target=listen_for_stop)
stop_thread.start()

capture_requests()

# === SAVE OUTPUT ===
print("🛑 Stopping capture. Writing data to file...")

with open(output_file, "w", encoding="utf-8") as f:
    json.dump(captured_data, f, indent=2)

print(f"✅ Captured {len(captured_data)} requests saved to '{output_file}'")

# === CLEAN UP ===
driver.quit()
print("✅ Browser closed.")
