# SIATEMA DE SEGURANÇA


import socket
import os
import hashlib
import requests
from tkinter import messagebox

def gerar_id_unico():
    nome_pc = socket.gethostname()
    user = os.getenv("USERNAME") or os.getenv("USER")
    texto_base = f"{nome_pc}-{user}"
    return hashlib.md5(texto_base.encode()).hexdigest()

def verificar_id_remoto(id_gerado):
    try:
        url = "https://sistema-bibliotecario-u1nb.onrender.com"
        resposta = requests.get(url)
        if resposta.status_code == 200:
            dados = resposta.json()
            return id_gerado in dados.get("autorizados", [])
        else:
            return False
    except:
        return False

ID_ATUAL = gerar_id_unico()

if not verificar_id_remoto(ID_ATUAL):
    messagebox.showerror("Acesso Negado", f"Este computador não está autorizado.\n\nID: {ID_ATUAL}")
    exit()
