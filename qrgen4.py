import re
import os
import uuid
import time
import qrcode
import yagmail
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
from threading import Thread

# --- CONFIG ---
CARTELLA_QR = "qrcodes"
PSW = "buhs esbo bgaz cqje"
EMAIL = "companynoobie@gmail.com"
DELAY_OGNI_N_EMAIL = 20
DELAY_SECONDI = 60

# --- REGISTRO ID GLOBALE ---
id_generati = set()

# --- FUNZIONI UTILI ---
def email_valida(email: str) -> bool:
    return re.match(r"[^@]+@[^@]+\.[^@]+", email) is not None

def genera_id_univoco():
    """Genera un ID univoco che non è mai stato usato"""
    while True:
        nuovo_id = uuid.uuid4().hex
        if nuovo_id not in id_generati:
            id_generati.add(nuovo_id)
            return nuovo_id

def invia_singola_email(nome, cognome, mail, gruppo, callback_success, callback_error):
    """Invia una singola email con QR code"""
    try:
        os.makedirs(CARTELLA_QR, exist_ok=True)
        
        # Generazione ID univoco
        id_unico = genera_id_univoco()
        
        # Creazione QR
        qr_text = f"ID:{id_unico};NOME:{nome};COGNOME:{cognome};GRUPPO:{gruppo}"
        nome_file = f"{nome}_{cognome}_{id_unico}.png"
        percorso_file = os.path.join(CARTELLA_QR, nome_file)
        
        qr = qrcode.make(qr_text)
        qr.save(percorso_file)
        
        # Invio email
        yag = yagmail.SMTP(EMAIL, PSW)
        subject = f"QR iscrizione open day - {nome}"
        body = f"""
        <p>Ciao <b>{nome}</b>,<br><br>
        la tua iscrizione è andata a buon fine.<br>
        Ecco il tuo <b>QR code personale</b> per il riconoscimento.<br><br>
        <b>Dati iscrizione:</b><br>
        • Nome: <b>{nome}</b><br>
        • Cognome: <b>{cognome}</b><br>
        • Gruppo: <b>{gruppo}</b><br><br>
        <i>Email generata il {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</i>
        </p>
        """
        
        yag.send(
            to=mail,
            subject=subject,
            contents=[body, yagmail.inline(percorso_file)]
        )
        
        callback_success(f"Email inviata con successo a {mail}")
        
    except Exception as e:
        callback_error(f"Errore invio email: {str(e)}")

def invia_email_batch(file_excel, callback_progress, callback_complete):
    """Funzione che processa il file Excel e invia le email"""
    try:
        os.makedirs(CARTELLA_QR, exist_ok=True)
        
        if not os.path.exists(file_excel):
            callback_complete(False, f"File non trovato: {file_excel}")
            return
            
        df = pd.read_excel(file_excel).fillna("")
        df.columns = df.columns.str.strip().str.upper()
        
        COLONNE_RICHIESTE = {"NOME", "COGNOME", "MAIL", "GRUPPO"}
        if not COLONNE_RICHIESTE.issubset(df.columns):
            callback_complete(False, f"Colonne mancanti: {COLONNE_RICHIESTE}")
            return
        
        try:
            yag = yagmail.SMTP(EMAIL, PSW)
        except Exception as e:
            callback_complete(False, f"Errore connessione SMTP: {e}")
            return
        
        totale = len(df)
        inviati = 0
        errori = 0
        skippati = 0
        
        for index, row in df.iterrows():
            try:
                nome = str(row["NOME"]).strip()
                cognome = str(row["COGNOME"]).strip()
                mail = str(row["MAIL"]).strip()
                gruppo = str(row["GRUPPO"]).strip()
                
                if not nome or not cognome:
                    skippati += 1
                    callback_progress(index + 1, totale, f"Skip: nome/cognome vuoto (riga {index})", inviati, errori, skippati)
                    continue
                
                if not email_valida(mail):
                    skippati += 1
                    callback_progress(index + 1, totale, f"Skip: email non valida - {mail}", inviati, errori, skippati)
                    continue
                
                # Generazione ID univoco
                id_unico = genera_id_univoco()
                qr_text = f"ID:{id_unico};NOME:{nome};COGNOME:{cognome};GRUPPO:{gruppo}"
                
                nome_file = f"{nome}_{cognome}_{id_unico}.png"
                percorso_file = os.path.join(CARTELLA_QR, nome_file)
                
                qr = qrcode.make(qr_text)
                qr.save(percorso_file)
                
                subject = f"QR iscrizione open day - {nome}"
                body = f"""
                <p>Ciao <b>{nome}</b>,<br><br>
                la tua iscrizione è andata a buon fine.<br>
                Ecco il tuo <b>QR code personale</b> per il riconoscimento.<br><br>
                <b>Dati iscrizione:</b><br>
                • Nome: <b>{nome}</b><br>
                • Cognome: <b>{cognome}</b><br>
                • Gruppo: <b>{gruppo}</b><br><br>
                <i>Email generata il {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</i>
                </p>
                """
                
                yag.send(
                    to=mail,
                    subject=subject,
                    contents=[body, yagmail.inline(percorso_file)]
                )
                
                inviati += 1
                callback_progress(index + 1, totale, f"Inviato a {mail}", inviati, errori, skippati)
                
                if inviati % DELAY_OGNI_N_EMAIL == 0 and inviati < totale:
                    for sec in range(DELAY_SECONDI, 0, -1):
                        callback_progress(index + 1, totale, f"Pausa per sicurezza... {sec}s rimanenti", inviati, errori, skippati)
                        time.sleep(1)
                
            except Exception as e:
                errori += 1
                callback_progress(index + 1, totale, f"Errore riga {index}: {str(e)[:50]}", inviati, errori, skippati)
        
        callback_complete(True, f"Completato!\n\nInviati: {inviati}\nErrori: {errori}\nSkippati: {skippati}")
        
    except Exception as e:
        callback_complete(False, f"Errore generale: {e}")

# --- FINESTRA MODALE INVIO SINGOLO ---
class InvioSingoloDialog:
    def __init__(self, parent, callback_success, callback_error):
        self.top = tk.Toplevel(parent)
        self.top.title("Invio Email Singola")
        self.top.geometry("450x320")
        self.top.resizable(False, False)
        self.top.grab_set()
        
        self.callback_success = callback_success
        self.callback_error = callback_error
        
        # Centra la finestra
        self.top.transient(parent)
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.top.winfo_screenheight() // 2) - (320 // 2)
        self.top.geometry(f"+{x}+{y}")
        
        # Container
        main_frame = tk.Frame(self.top, bg="#f8f9fa", padx=25, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        # Titolo
        tk.Label(
            main_frame,
            text="Inserisci i dati per l'invio",
            font=("Arial", 13, "bold"),
            bg="#f8f9fa"
        ).grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Campi
        tk.Label(main_frame, text="Nome:", bg="#f8f9fa", font=("Arial", 10)).grid(row=1, column=0, sticky="w", pady=8)
        self.entry_nome = tk.Entry(main_frame, font=("Arial", 10), width=30)
        self.entry_nome.grid(row=1, column=1, pady=8, padx=(10, 0))
        
        tk.Label(main_frame, text="Cognome:", bg="#f8f9fa", font=("Arial", 10)).grid(row=2, column=0, sticky="w", pady=8)
        self.entry_cognome = tk.Entry(main_frame, font=("Arial", 10), width=30)
        self.entry_cognome.grid(row=2, column=1, pady=8, padx=(10, 0))
        
        tk.Label(main_frame, text="Email:", bg="#f8f9fa", font=("Arial", 10)).grid(row=3, column=0, sticky="w", pady=8)
        self.entry_email = tk.Entry(main_frame, font=("Arial", 10), width=30)
        self.entry_email.grid(row=3, column=1, pady=8, padx=(10, 0))
        
        tk.Label(main_frame, text="Gruppo:", bg="#f8f9fa", font=("Arial", 10)).grid(row=4, column=0, sticky="w", pady=8)
        self.entry_gruppo = tk.Entry(main_frame, font=("Arial", 10), width=30)
        self.entry_gruppo.grid(row=4, column=1, pady=8, padx=(10, 0))
        
        # Pulsanti
        button_frame = tk.Frame(main_frame, bg="#f8f9fa")
        button_frame.grid(row=5, column=0, columnspan=2, pady=(25, 0))
        
        tk.Button(
            button_frame,
            text="Annulla",
            command=self.top.destroy,
            bg="#6c757d",
            fg="white",
            font=("Arial", 10),
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2"
        ).pack(side="left", padx=5)
        
        tk.Button(
            button_frame,
            text="Invia Email",
            command=self.invia,
            bg="#007bff",
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2"
        ).pack(side="left", padx=5)
        
        # Focus sul primo campo
        self.entry_nome.focus()
    
    def invia(self):
        nome = self.entry_nome.get().strip()
        cognome = self.entry_cognome.get().strip()
        email = self.entry_email.get().strip()
        gruppo = self.entry_gruppo.get().strip()
        
        # Validazione
        if not nome or not cognome:
            messagebox.showwarning("Attenzione", "Nome e cognome sono obbligatori!")
            return
        
        if not email_valida(email):
            messagebox.showwarning("Attenzione", "Inserisci un'email valida!")
            return
        
        if not gruppo:
            messagebox.showwarning("Attenzione", "Il gruppo è obbligatorio!")
            return
        
        # Disabilita pulsanti durante invio
        self.top.withdraw()
        
        # Invio in thread separato
        Thread(
            target=invia_singola_email,
            args=(nome, cognome, email, gruppo, self.on_success, self.on_error),
            daemon=True
        ).start()
    
    def on_success(self, messaggio):
        self.top.destroy()
        self.callback_success(messaggio)
    
    def on_error(self, messaggio):
        self.top.deiconify()
        self.callback_error(messaggio)

# --- GUI PRINCIPALE ---
class QREmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR Code Email Sender")
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        self.root.configure(bg="#ffffff")
        
        self.file_path = tk.StringVar()
        self.is_running = False
        
        # Header
        header = tk.Frame(root, bg="#007bff", height=70)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text="QR Code Email Sender",
            font=("Arial", 18, "bold"),
            bg="#007bff",
            fg="white"
        ).pack(pady=20)
        
        # Container principale
        main_frame = tk.Frame(root, bg="#f8f9fa", padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        # Pulsanti azione principali
        action_frame = tk.Frame(main_frame, bg="#f8f9fa")
        action_frame.pack(fill="x", pady=(0, 20))
        
        tk.Button(
            action_frame,
            text="Invio totale (Excel)",
            command=self.scegli_file_batch,
            bg="#28a745",
            fg="white",
            font=("Arial", 11, "bold"),
            cursor="hand2",
            relief="flat",
            padx=25,
            pady=10,
            width=20
        ).pack(side="left", padx=(0, 10))
        
        tk.Button(
            action_frame,
            text="Invio Singolo",
            command=self.apri_invio_singolo,
            bg="#007bff",
            fg="white",
            font=("Arial", 11, "bold"),
            cursor="hand2",
            relief="flat",
            padx=25,
            pady=10,
            width=20
        ).pack(side="left")
        
        # File selezionato
        file_frame = tk.LabelFrame(main_frame, text="File Excel Selezionato", font=("Arial", 10, "bold"), bg="#f8f9fa", padx=15, pady=10)
        file_frame.pack(fill="x", pady=(0, 15))
        
        self.label_file = tk.Label(
            file_frame,
            textvariable=self.file_path,
            font=("Arial", 9),
            bg="#f8f9fa",
            fg="#6c757d",
            anchor="w"
        )
        self.label_file.pack(fill="x")
        
        # Progress
        progress_frame = tk.LabelFrame(main_frame, text="Progresso Invio totale", font=("Arial", 10, "bold"), bg="#f8f9fa", padx=15, pady=10)
        progress_frame.pack(fill="x", pady=(0, 15))
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress_bar.pack(fill="x", pady=(0, 10))
        
        stats_frame = tk.Frame(progress_frame, bg="#f8f9fa")
        stats_frame.pack(fill="x")
        
        self.label_inviati = tk.Label(stats_frame, text="Inviati: 0", font=("Arial", 9), bg="#f8f9fa", fg="#28a745")
        self.label_inviati.pack(side="left", padx=15)
        
        self.label_errori = tk.Label(stats_frame, text="Errori: 0", font=("Arial", 9), bg="#f8f9fa", fg="#dc3545")
        self.label_errori.pack(side="left", padx=15)
        
        self.label_skippati = tk.Label(stats_frame, text="Skippati: 0", font=("Arial", 9), bg="#f8f9fa", fg="#ffc107")
        self.label_skippati.pack(side="left", padx=15)
        
        # Log
        log_frame = tk.LabelFrame(main_frame, text="Log Attività", font=("Arial", 10, "bold"), bg="#f8f9fa", padx=15, pady=10)
        log_frame.pack(fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.log_text = tk.Text(
            log_frame,
            height=12,
            font=("Consolas", 9),
            yscrollcommand=scrollbar.set,
            bg="#ffffff",
            relief="solid",
            borderwidth=1
        )
        self.log_text.pack(fill="both", expand=True)
        scrollbar.config(command=self.log_text.yview)
        
        self.log("Applicazione pronta.")
    
    def scegli_file_batch(self):
        file = filedialog.askopenfilename(
            title="Seleziona file Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file:
            self.file_path.set(os.path.basename(file))
            self.file_excel_completo = file
            self.log(f"File selezionato: {os.path.basename(file)}")
            self.avvia_invio_batch()
    
    def apri_invio_singolo(self):
        InvioSingoloDialog(self.root, self.on_singolo_success, self.on_singolo_error)
    
    def on_singolo_success(self, messaggio):
        messagebox.showinfo("Successo", messaggio)
        self.log(messaggio)
    
    def on_singolo_error(self, messaggio):
        messagebox.showerror("Errore", messaggio)
        self.log(f"ERRORE: {messaggio}")
    
    def log(self, messaggio):
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert("end", f"[{timestamp}] {messaggio}\n")
        self.log_text.see("end")
        self.root.update()
    
    def aggiorna_progress(self, corrente, totale, messaggio, inviati, errori, skippati):
        percentuale = (corrente / totale) * 100
        self.progress_bar["value"] = percentuale
        self.label_inviati.config(text=f"Inviati: {inviati}")
        self.label_errori.config(text=f"Errori: {errori}")
        self.label_skippati.config(text=f"Skippati: {skippati}")
        self.log(messaggio)
    
    def invio_completato(self, successo, messaggio):
        self.is_running = False
        
        if successo:
            messagebox.showinfo("Completato", messaggio)
        else:
            messagebox.showerror("Errore", messaggio)
        
        self.log("=" * 60)
        self.log(messaggio)
    
    def avvia_invio_batch(self):
        if not hasattr(self, 'file_excel_completo'):
            messagebox.showwarning("Attenzione", "Nessun file selezionato!")
            return
        
        if self.is_running:
            messagebox.showwarning("Attenzione", "Invio già in corso!")
            return
        
        risposta = messagebox.askyesno(
            "Conferma",
            f"Avviare l'invio totale dal file:\n{self.file_path.get()}?\n\nPausa di {DELAY_SECONDI}s ogni {DELAY_OGNI_N_EMAIL} email."
        )
        
        if not risposta:
            return
        
        self.is_running = True
        self.progress_bar["value"] = 0
        self.log("Avvio invio...")
        
        thread = Thread(
            target=invia_email_batch,
            args=(self.file_excel_completo, self.aggiorna_progress, self.invio_completato),
            daemon=True
        )
        thread.start()

# --- MAIN ---
if __name__ == "__main__":
    root = tk.Tk()
    app = QREmailApp(root)
    root.mainloop()