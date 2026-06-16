#!/usr/bin/env python3
"""
Esegui questo script una volta per generare l'hash della password admin.
Poi copia il risultato nella variabile d'ambiente ADMIN_PASSWORD_HASH su Render.

Uso:
    python genera_password.py
"""
from werkzeug.security import generate_password_hash
import getpass

pwd = getpass.getpass("Inserisci la password admin che vuoi usare: ")
pwd2 = getpass.getpass("Conferma password: ")
if pwd != pwd2:
    print("Le password non corrispondono.")
else:
    hashed = generate_password_hash(pwd)
    print("\n✅ Hash generato:\n")
    print(hashed)
    print("\nCopia questo valore nella variabile d'ambiente ADMIN_PASSWORD_HASH su Render.")
