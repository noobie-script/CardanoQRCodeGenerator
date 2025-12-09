import os
import qrcode
import yagmail
import pandas as pd

# --- CONFIG ---
CARTELLA_QR = "qrcodes"
PSW = "PASSWORD PER APP DELL'ACCOUNT GMAIL DA CREARE"
FILE_EXCEL = "NOMEFILE.xlsx"
EMAIL = "MAIL MITTENTE DEI QR CODE"

# preparazione cartela
os.makedirs(CARTELLA_QR, exist_ok=True)

# connesione SMTP
yag = yagmail.SMTP(EMAIL, PSW)

# lettura file 
df = pd.read_excel(FILE_EXCEL)

df.columns = df.columns.str.strip().str.upper()

for index, row in df.iterrows():
    cognome = str(row["COGNOME"]).strip()
    nome = str(row["NOME"]).strip()
    mail = str(row["MAIL"]).strip()
    gruppo = str(row["GRUPPO"]).strip()

    # contenuto qr
    qr_text = f"NOME:{nome};COGNOME:{cognome};GRUPPO:{gruppo}"

    # creazione qr
    qr = qrcode.make(qr_text)
    nome_file = f"{nome}_{cognome}.png"
    percorso_file = os.path.join(CARTELLA_QR, nome_file)
    qr.save(percorso_file)

    # preparazione email con qr inline
    subject = "QR iscrizione openday"
    body = f"""
    Ciao {nome}, <br><br>
    La tua iscrizione è andata a buon termine,
    ecco il tuo <b>QR code personale</b>: <br><br>

    Nome: <b>{nome}</b><br>
    Cognome: <b>{cognome}</b><br>
    Gruppo: <b>{gruppo}</b><br><br>
    """

    yag.send(
        to=mail,
        subject=subject,
        contents=[
            body,
            yagmail.inline(percorso_file)
        ]
    )

    print(f"qr inviato a {mail}")
print("tutti i qr code sono stati inviati")
