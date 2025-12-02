from flask import Flask, request, jsonify
import json, os
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

INVENTARIO_FILE = "inventarios_recibidos.json"
JSON_AGENCIA = r"C:\Users\infob\Desktop\inventarios\agencias\tultitlan\inventario_render.json"

def cargar_inventario():
    if not os.path.exists(INVENTARIO_FILE):
        return []
    with open(INVENTARIO_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def guardar_inventario(data):
    with open(INVENTARIO_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

@app.route("/")
def home():
    return "Servidor PROXI JQ Motors TULTITLÁN activo."

@app.route("/inventario", methods=["POST"])
def recibir_inventario():
    data = request.get_json()
    agencia = data.get("agencia")
    inventario = data.get("inventario", [])
    inventario_final = [{"codigo": i.get("codigo",""), "descripcion": i.get("descripcion",""), "stock": i.get("stock",0), "agencia": agencia} for i in inventario]
    guardar_inventario(inventario_final)
    return jsonify({"status": "actualizado"})

@app.route("/inventario-json", methods=["GET"])
def obtener_inventario():
    return jsonify(cargar_inventario())

@app.route("/limpiar", methods=["POST"])
def limpiar_inventario():
    guardar_inventario([])
    return jsonify({"status": "inventario_limpiado"})

@app.route("/actualizar-matriz", methods=["POST"])
def actualizar_desde_matriz():
    if not os.path.exists(JSON_AGENCIA):
        return jsonify({"status": "no_encontrado"})
    with open(JSON_AGENCIA, "r", encoding="utf-8") as f:
        inventario_real = json.load(f)
    inventario_final = [{"codigo": i.get("codigo",""), "descripcion": i.get("descripcion",""), "stock": i.get("stock",0), "agencia": "tultitlan"} for i in inventario_real]
    guardar_inventario(inventario_final)
    return jsonify({"status": "actualizado"})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5004)
