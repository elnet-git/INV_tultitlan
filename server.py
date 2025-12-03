# server.py
from flask import Flask, jsonify, request
from flask_cors import CORS
from pathlib import Path
import json, os

# ===============================
# CONFIGURACIÓN
# ===============================
app = Flask(__name__)
CORS(app)  # Permitir CORS a todos los dominios

# CORS EXPLÍCITO (Render lo necesita)
@app.after_request
def agregar_cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
JSON_FILE = DATA_DIR / "inventario_render.json"

# ===============================
# FUNCIONES AUXILIARES
# ===============================
def cargar_inventario():
    """Carga el JSON de inventario desde disco"""
    if not JSON_FILE.exists():
        return []
    try:
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print("❌ Error cargando JSON:", e)
        return []

def guardar_inventario(data):
    """Guarda el JSON de inventario en disco"""
    try:
        with open(JSON_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        print("❌ Error guardando JSON:", e)
        return False

# ===============================
# RUTAS
# ===============================
@app.route("/inventario-json", methods=["GET"])
def inventario_json():
    """Devuelve el inventario en JSON"""
    data = cargar_inventario()
    return jsonify(data), 200

@app.route("/actualizar-inventario", methods=["POST"])
def actualizar_inventario():
    """Recibe un JSON para actualizar el inventario"""
    try:
        data = request.get_json()
        if not isinstance(data, list):
            return jsonify({"error": "Formato inválido, se espera lista de productos"}), 400
        guardar_inventario(data)
        return jsonify({"status": "ok", "message": "Inventario actualizado"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/limpiar", methods=["POST"])
def limpiar_inventario():
    """Vacía el inventario"""
    try:
        guardar_inventario([])
        return jsonify({"status": "ok", "message": "Inventario limpio"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ===============================
# INICIO DEL SERVIDOR
# ===============================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5051))
    app.run(host="0.0.0.0", port=port)
