import os
import zipfile
import py7zr
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from docx2pdf import convert
import sys
import os

def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)



# ==========================================================
# CONFIGURAÇÕES GLOBAIS
# ==========================================================
ctk.set_appearance_mode("light")   # "dark" ou "light"
ctk.set_default_color_theme("blue")

APP_WIDTH = 820
APP_HEIGHT = 560


# ==========================================================
# FUNÇÕES
# ==========================================================
def pdf_para_word():
    arquivo = filedialog.askopenfilename(
        title="Selecionar PDF",
        filetypes=[("PDF", "*.pdf")]
    )
    if not arquivo:
        return

    destino = filedialog.askdirectory(title="Selecionar pasta de destino")
    if not destino:
        return

    nome = os.path.splitext(os.path.basename(arquivo))[0]
    saida = os.path.join(destino, f"{nome}.docx")

    try:
        status_label.configure(text="Convertendo PDF para Word...")
        app.update()
        cv = Converter(arquivo)
        cv.convert(saida)
        cv.close()
        status_label.configure(text="✔ Conversão concluída com sucesso")
        messagebox.showinfo("Sucesso", "PDF convertido para Word!")
    except Exception as e:
        messagebox.showerror("Erro", str(e))


def word_para_pdf():
    arquivo = filedialog.askopenfilename(
        title="Selecionar Word",
        filetypes=[("Word", "*.docx")]
    )
    if not arquivo:
        return

    destino = filedialog.askdirectory(title="Selecionar pasta de destino")
    if not destino:
        return

    try:
        status_label.configure(text="Convertendo Word para PDF...")
        app.update()
        convert(arquivo, destino)
        status_label.configure(text="✔ Conversão concluída com sucesso")
        messagebox.showinfo("Sucesso", "Word convertido para PDF!")
    except Exception as e:
        messagebox.showerror("Erro", str(e))


def extrair_arquivos():
    arquivos = filedialog.askopenfilenames(
        title="Selecionar arquivos",
        filetypes=[("Compactados", "*.zip *.7z")]
    )
    if not arquivos:
        return

    destino = filedialog.askdirectory(title="Selecionar pasta de destino")
    if not destino:
        return

    progress_bar.set(0)
    status_label.configure(text="Extraindo arquivos...")

    total = len(arquivos)
    for i, arquivo in enumerate(arquivos, start=1):
        nome = os.path.splitext(os.path.basename(arquivo))[0]
        pasta_destino = os.path.join(destino, nome)
        os.makedirs(pasta_destino, exist_ok=True)

        try:
            if arquivo.endswith(".zip"):
                with zipfile.ZipFile(arquivo, "r") as z:
                    z.extractall(pasta_destino)
            elif arquivo.endswith(".7z"):
                with py7zr.SevenZipFile(arquivo, "r") as z:
                    z.extractall(pasta_destino)
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            return

        progress_bar.set(i / total)
        app.update()

    status_label.configure(text="✔ Extração finalizada com sucesso")
    messagebox.showinfo("Finalizado", "Arquivos extraídos com sucesso!")


# ==========================================================
# JANELA PRINCIPAL
# ==========================================================
app = ctk.CTk()
app.title("Extrator do Campaner")
app.geometry(f"{APP_WIDTH}x{APP_HEIGHT}")
app.resizable(False, False)
app.iconbitmap(resource_path("icon.ico"))




# ==========================================================
# HEADER
# ==========================================================
header = ctk.CTkFrame(app, height=90, corner_radius=0)
header.pack(fill="x")

ctk.CTkLabel(
    header,
    text="Extractor Pro",
    font=("Segoe UI", 26, "bold")
).pack(pady=(18, 0))

ctk.CTkLabel(
    header,
    text="Conversão e descompactação em um único aplicativo",
    font=("Segoe UI", 13),
    text_color="gray"
).pack()


# ==========================================================
# CONTEÚDO
# ==========================================================
content = ctk.CTkFrame(app, corner_radius=16)
content.pack(padx=40, pady=25, fill="both", expand=True)


# ==========================================================
# CONVERSÃO
# ==========================================================
ctk.CTkLabel(
    content,
    text="Conversão de Documentos",
    font=("Segoe UI", 16, "bold")
).pack(pady=(20, 10))

row_convert = ctk.CTkFrame(content, fg_color="transparent")
row_convert.pack(pady=10)

ctk.CTkButton(
    row_convert,
    text="PDF → Word",
    width=200,
    height=48,
    font=("Segoe UI", 13, "bold"),
    command=pdf_para_word
).pack(side="left", padx=10)

ctk.CTkButton(
    row_convert,
    text="Word → PDF",
    width=200,
    height=48,
    font=("Segoe UI", 13, "bold"),
    command=word_para_pdf
).pack(side="left", padx=10)


# ==========================================================
# DIVISOR
# ==========================================================
ctk.CTkFrame(content, height=2, fg_color="#E5E7EB").pack(fill="x", pady=25)


# ==========================================================
# EXTRAÇÃO
# ==========================================================
ctk.CTkLabel(
    content,
    text="Descompactação de Arquivos",
    font=("Segoe UI", 16, "bold")
).pack(pady=(5, 10))

ctk.CTkButton(
    content,
    text="Selecionar arquivos (.zip / .7z)",
    width=420,
    height=52,
    font=("Segoe UI", 13, "bold"),
    command=extrair_arquivos
).pack(pady=10)


# ==========================================================
# PROGRESSO
# ==========================================================
progress_bar = ctk.CTkProgressBar(content, width=420)
progress_bar.set(0)
progress_bar.pack(pady=18)

status_label = ctk.CTkLabel(
    content,
    text="Aguardando ação do usuário",
    font=("Segoe UI", 12),
    text_color="gray"
)
status_label.pack()


# ==========================================================
# FOOTER
# ==========================================================
footer = ctk.CTkFrame(app, height=40, corner_radius=0)
footer.pack(fill="x", side="bottom")

ctk.CTkLabel(
    footer,
    text="Feito por Davi Campaner • 2025",
    font=("Segoe UI", 10),
    text_color="gray"
).pack(side="left", padx=20)

ctk.CTkLabel(
    footer,
    text="Python • CustomTkinter • Desktop App",
    font=("Segoe UI", 10),
    text_color="gray"
).pack(side="right", padx=20)


app.mainloop()
