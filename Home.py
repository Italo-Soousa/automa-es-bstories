import tkinter as tk
from tkinter import ttk
from tkinter import Toplevel

def abrir_ajuste_estoque():
    janela = Toplevel()
    janela.title("Ajuste de Estoque")
    janela.geometry("300x150")
    ttk.Label(janela, text="Página de Ajuste de Estoque\n(em construção)", font=("Segoe UI", 12), padding=20).pack()

def abrir_automacao_relatorio():
    janela = Toplevel()
    janela.title("Automação de Relatório")
    janela.geometry("300x150")
    ttk.Label(janela, text="Página de Automação de Relatório\n(em construção)", font=("Segoe UI", 12), padding=20).pack()

# Estilo da interface
root = tk.Tk()
root.title("Sistema - Menu Principal")
root.geometry("350x250")
root.configure(bg="#f0f0f0")

# Tema moderno usando ttk
style = ttk.Style()
style.theme_use("clam")

style.configure("TButton",
                font=("Segoe UI", 11),
                padding=10,
                relief="flat",
                background="#007acc",
                foreground="white")
style.map("TButton",
          background=[('active', '#005f99')])

style.configure("TLabel", background="#f0f0f0")

# Título
ttk.Label(root, text="Menu Principal", font=("Segoe UI", 16, "bold"), padding=(20, 10)).pack()

# Botões
ttk.Button(root, text="Ajuste de Estoque", command=abrir_ajuste_estoque).pack(pady=10, ipadx=10)
ttk.Button(root, text="Automação de Relatório", command=abrir_automacao_relatorio).pack(pady=10, ipadx=10)

root.mainloop()