import re
import os
import sys
import uuid
import qrcode
import yagmail
import pandas as pd
from datetime import datetime

# --- CONFIG ---
CARTELLA_QR = "qrcodes"
PSW = "buhs esbo bgaz cqje"
EMAIL = "companynoobie@gmail.com"

# --- ARGOMENTI LINEA DI COMANDO ---
if len(sys.argv) < 2:
    raise ValueError("Uso corretto: python script.py <file_excel>")

FILE_EXCEL = sys.argv[1]

# --- FUNZIONI UTILI ---
def email_valida(email: str) -> bool:
    return re.match(r"[^@]+@[^@]+\.[^@]+", email) is not None


# --- PREPARAZIONE CARTELLA ---
os.makedirs(CARTELLA_QR, exist_ok=True)

# --- LETTURA EXCEL ---
if not os.path.exists(FILE_EXCEL):
    raise FileNotFoundError(f"File Excel non trovato: {FILE_EXCEL}")

df = pd.read_excel(FILE_EXCEL).fillna("")
df.columns = df.columns.str.strip().str.upper()

# Controllo colonne richieste
COLONNE_RICHIESTE = {"NOME", "COGNOME", "MAIL", "GRUPPO"}
if not COLONNE_RICHIESTE.issubset(df.columns):
    raise ValueError(f"Il file Excel manca delle colonne richieste: {COLONNE_RICHIESTE}")

# --- CONNESSIONE SMTP ---
try:
    yag = yagmail.SMTP(EMAIL, PSW)
except Exception as e:
    raise ConnectionError(f"Errore nella connessione SMTP: {e}")


# --- PROCESSO PRINCIPALE ---
print("\n--- AVVIO INVIO QR ---\n")

for index, row in df.iterrows():

    try:
        nome = str(row["NOME"]).strip()
        cognome = str(row["COGNOME"]).strip()
        mail = str(row["MAIL"]).strip()
        gruppo = str(row["GRUPPO"]).strip()

        if not nome or not cognome:
            print(f"[SKIP] Riga {index}: nome o cognome vuoto.")
            continue

        if not email_valida(mail):
            print(f"[SKIP] Email non valida per {nome} {cognome}: {mail}")
            continue

        # --- Generazione ID univoco (NON salvato nell'Excel) ---
        id_unico = uuid.uuid4().hex  # sempre nuovo, sempre unico

        # --- Contenuto QR ---
        qr_text = (
            f"ID:{id_unico};"
            f"NOME:{nome};"
            f"COGNOME:{cognome};"
            f"GRUPPO:{gruppo}"
        )

        # --- Nome file univoco ---
        nome_file = f"{nome}_{cognome}_{id_unico}.png"
        percorso_file = os.path.join(CARTELLA_QR, nome_file)

        # Generazione QR (non controllo esistenza perché l’ID è sempre diverso)
        qr = qrcode.make(qr_text)
        qr.save(percorso_file)

        # --- Email ---
        subject = f"QR iscrizione open day - {nome}"
        body = f"""
        <p>Ciao <b>{nome}</b>,<br><br>
        la tua iscrizione è andata a buon fine.<br>
        Ecco il tuo <b>QR code personale</b> per il riconoscimento.<br><br>

        <b>Dati iscrizione:</b><br>
        • Nome: <b>{nome}</b><br>
        • Cognome: <b>{cognome}</b><br>
        • Gruppo: <b>{gruppo}</b><br><br>

        <i>Email generata automaticamente il {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</i>
        </p>
        """

        yag.send(
            to=mail,
            subject=subject,
            contents=[body, yagmail.inline(percorso_file)]
        )

        print(f"[OK] QR inviato a {mail}")

    except Exception as e:
        print(f"[ERRORE] Riga {index} ({nome} {cognome}): {e}")

print("\n--- COMPLETATO: tutti i QR sono stati processati ---\n")