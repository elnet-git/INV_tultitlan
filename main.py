from flask import Flask, request, jsonify
import json, os
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# =============================
# Configuración
# =============================
DATA_PATH = "data"
os.makedirs(DATA_PATH, exist_ok=True)

INVENTARIO_FILE = os.path.join(DATA_PATH, "inventarios_recibidos.json")
JSON_AGENCIA = os.path.join(DATA_PATH, "inventario_render.json")

# CAMBIA ESTE NOMBRE SEGÚN LA AGENCIA
NOMBRE_AGENCIA = "tultitlan"


# =============================
# Funciones
# =============================
def cargar_inventario():
    if not os.path.exists(INVENTARIO_FILE):
        return []
    try:
        with open(INVENTARIO_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return []


def guardar_inventario(data):
    with open(INVENTARIO_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


# =============================
# Rutas
# =============================
@app.route("/")
def home():
    return f"Servidor PROXI JQ Motors ({NOMBRE_AGENCIA}) activo."


@app.route("/inventario", methods=["POST"])
def recibir_inventario():
    data = request.get_json() or {}
    agencia = data.get("agencia", NOMBRE_AGENCIA)
    inventario = data.get("inventario", [])

    inventario_final = [
        {
            "codigo": i.get("codigo", ""),
            "descripcion": i.get("descripcion", ""),
            "stock": i.get("stock", 0),
            "agencia": agencia
        }
        for i in inventario
    ]

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

    try:
        with open(JSON_AGENCIA, "r", encoding="utf-8") as f:
            inventario_real = json.load(f)
    except:
        return jsonify({"status": "error_lectura"})

    inventario_final = [
        {
            "codigo": i.get("codigo", ""),
            "descripcion": i.get("descripcion", ""),
            "stock": i.get("stock", 0),
            "agencia": NOMBRE_AGENCIA
        }
        for i in inventario_real
    ]

    guardar_inventario(inventario_final)
    return jsonify({"status": "actualizado"})


# =============================
# Render requiere host=0.0.0.0
# =============================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
