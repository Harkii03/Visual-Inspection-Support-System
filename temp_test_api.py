import urllib.request
import json
import threading
import time
import sys

# Start the flask app in a separate thread
def run_app():
    import app
    app.app.run(host="127.0.0.1", port=8765, debug=False, use_reloader=False)

threading.Thread(target=run_app, daemon=True).start()
time.sleep(2) # wait for server to start

def request(path, method="GET", data=None):
    url = f"http://127.0.0.1:8765{path}"
    req = urllib.request.Request(url, method=method)
    if data:
        req.add_header('Content-Type', 'application/json')
        data = json.dumps(data).encode('utf-8')
    try:
        with urllib.request.urlopen(req, data=data) as res:
            return json.loads(res.read().decode('utf-8'))
    except Exception as e:
        print(f"Error calling {path}: {e}")
        sys.exit(1)

# Test 1: Open the example data
res = request("/api/open", method="POST", data={"path": r"C:\Users\Masah\Desktop\目視確認アシスト\インプットデータ例\260721"})
print("Open:", res['ok'])

# Test 2: Get animals
res = request("/api/animals")
print("Animals:", [a['name'] for a in res['gridAnimals']])
first_animal = res['gridAnimals'][0]['name']

# Test 3: Get grid for the first animal
res = request(f"/api/grid/{urllib.parse.quote(first_animal)}")
print(f"Grid for {first_animal}: {len(res['rows'])} rows returned.")

# Test 4: Bulk confirm first 2 images
indices = [res['rows'][0]['index'], res['rows'][1]['index']]
res = request("/api/confirm", method="POST", data={"indices": indices})
print(f"Bulk confirm: {res['ok']}, confirmed count: {res['confirmed']}")

print("All tests passed!")
