import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import scrolledtext
import openpyxl

class MailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Enviador de correos")
        self.root.geometry('350x250')

        self.root.configure(bg='#ECECEC')
        style = ttk.Style()
        style.theme_create('custom', settings={
            'TLabel': {'configure': {'background': '#EC0000', 'foreground': '#FFFFFF', 'font': ('Helvetica', 11)}},
            'TFrame': {'configure': {'background': '#EC0000', 'borderwidth': 0, 'relief': 'flat'}},
            'TButton': {'configure': {'background':'#DEEDF2', 'font': ('Helvetica', 11), 'anchor': 'center'}}}
            )

        style.theme_use('custom')

        frame = ttk.Frame(root)
        frame.pack(expand=True, fill="both", padx=20, pady=20)

        ttk.Label(frame, text="Enviador de correos", font=("Helvetica", 16)).grid(row=0, column=0, columnspan=2, pady=10, sticky='en')

        ttk.Label(frame, text="Archivo excel:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        ttk.Button(frame, text="Seleccionar archivo", command=self.cargar_archivo("xlsx"), cursor="hand2").grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        
        ttk.Label(frame, text="Tu correo:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        self.correo_entry = ttk.Entry(frame, font=("Helvetica", 11))
        self.correo_entry.grid(row=3, column=1, padx=5, pady=5, sticky='ew')

        ttk.Label(frame, text="Tu contraseña:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        self.pass_entry = ttk.Entry(frame, font=("Helvetica", 11))
        self.pass_entry.grid(row=3, column=1, padx=5, pady=5, sticky='ew')

        ttk.Button(frame, text="Enviar", command=self.procesar,cursor="hand2", width=20).grid(row=4, column=1, columnspan=2, pady=10, sticky='n')

        self.excel_procesado = ""

    def procesar_organica(self, organica_path):
        wb_o = openpyxl.load_workbook(organica_path)
        ws_o = wb_o.active

        organica_c = {}
        for cell in ws_o[1]:
            if cell.value != None:
                organica_c[cell.value] = cell.column_letter

        self.organica_concats = set()
        self.concats_dict = {}

        for i in range(2,ws_o.max_row):
            ur = ws_o[f"{organica_c['UR']}{i}"]
            cargo = ws_o[f"{organica_c['Cargo']}{i}"]
            gls_cargo = ws_o[f"{organica_c['GlsCargo']}{i}"]
            gls_ur = ws_o[f"{organica_c['GlsUR']}{i}"]
            self.organica_concats.add(f"{ur.value}-{cargo.value}")
            self.concats_dict[f"{ur.value}-{cargo.value}"] = {'Cargo': cargo.value, 'GlsCargo': gls_cargo.value, 'UR': ur.value, 'GlsUR': gls_ur.value}
        
        self.organica_procesada = organica_path


    def procesar(self):
        if self.excel is not None:
            if self.organica_path != self.organica_procesada:
                self.procesar_organica(self.organica_path)
            if self.catalogo_path != self.catalogo_procesado:
                self.procesar_catalogo(self.catalogo_path)

            roles = self.organica_concats - self.catalogo_concats
            
            resultado_text = ""
            for rol in roles:
                resultado_text += f'Rol: {rol}\n'
                accesos = self.concats_dict[rol]
                for id, val in enumerate(accesos):
                    resultado_text += f"{id+1}. {val} {accesos[val]}\n"
                resultado_text += f"\n\n"
            self.mostrar_resultados(resultado_text)
        else:
            tk.messagebox.showwarning("Error", "Debe seleccionar el archivo.")

if __name__ == "__main__":
    root = tk.Tk()
    app = MailSenderApp(root)
    root.mainloop()

df = pd.read_excel('archivo.xlsx')

# Configuración del servidor SMTP de Hotmail (Outlook)
smtp_server = 'smtp.office365.com'
port = 587
sender_email = 'tucorreo@hotmail.com'
password = 'tupassword'

# Itera sobre las filas del DataFrame
for index, row in df.iterrows():
    destinatario = row['Destinatario']
    ccs = row['CCs']
    asunto = row['Asunto']
    mensaje = row['Mensaje']

    # Construye el mensaje
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = destinatario
    msg['Cc'] = ccs
    msg['Subject'] = asunto
    msg.attach(MIMEText(mensaje, 'plain'))

    # Inicia sesión en el servidor SMTP
    server = smtplib.SMTP(smtp_server, port)
    server.starttls()
    server.login(sender_email, password)

    # Envía el correo
    text = msg.as_string()
    server.sendmail(sender_email, destinatario.split(',') + ccs.split(','), text)

    # Cierra la conexión con el servidor
    server.quit()

print("Correos enviados correctamente")
