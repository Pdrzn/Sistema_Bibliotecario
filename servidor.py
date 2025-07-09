# servidor.py
from flask import Flask, jsonify
import os

app = Flask(__name__)

IDS_AUTORIZADOS = [
    "144277ab828559cbb06db3022b27ad59",  # id do pedro
    
]

@app.route("/", methods=["GET"])
def autorizados():
    return jsonify({"autorizados": IDS_AUTORIZADOS})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Render define essa porta
    app.run(host="0.0.0.0", port=port)
