# servidor.py
from flask import Flask, jsonify

app = Flask(__name__)

IDS_AUTORIZADOS = [
    "144277ab828559cbb06db3022b27ad59",  # id do pedro
    
]

@app.route("/", methods=["GET"])
def autorizados():
    return jsonify({"autorizados": IDS_AUTORIZADOS})

if __name__ == "__main__":
    app.run()
