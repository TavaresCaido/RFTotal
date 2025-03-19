import tkinter as tk
from tkinter import filedialog, messagebox

class Interface:
    def __init__(self, root, on_selecionar_diretorio, on_selecionar_arquivo, on_processar):
        self.root = root
        self.root.title("RFTotal - Automação de Relatórios de Obras")
        self.root.geometry("500x400")
        self.root.resizable(False, False)

        # Frame do topo
        frame_topo = tk.Frame(root, pady=20)
        frame_topo.pack()
        tk.Label(frame_topo, text="🔹 RFTotal - Automação de Relatórios", font=("Arial", 14, "bold")).pack()

        # Seção de seleção de diretório
        frame_diretorio = tk.Frame(root, pady=10)
        frame_diretorio.pack()
        btn_diretorio = tk.Button(frame_diretorio, text="📁 Selecionar Diretório de Imagens", command=on_selecionar_diretorio, width=35)
        btn_diretorio.pack()
        self.label_diretorio = tk.Label(frame_diretorio, text="Nenhum diretório selecionado", font=("Arial", 10), wraplength=400)
        self.label_diretorio.pack()

        # Seção de seleção de arquivo Word
        frame_arquivo = tk.Frame(root, pady=10)
        frame_arquivo.pack()
        btn_arquivo = tk.Button(frame_arquivo, text="📄 Selecionar Relatório Word", command=on_selecionar_arquivo, width=35)
        btn_arquivo.pack()
        self.label_arquivo = tk.Label(frame_arquivo, text="Nenhum arquivo selecionado", font=("Arial", 10), wraplength=400)
        self.label_arquivo.pack()

        # Botão para processar
        frame_processar = tk.Frame(root, pady=20)
        frame_processar.pack()
        btn_processar = tk.Button(frame_processar, text="🚀 Gerar Relatório", command=on_processar, bg="green", fg="white", font=("Arial", 12, "bold"))
        btn_processar.pack()

    def atualizar_label_diretorio(self, texto):
        self.label_diretorio.config(text=texto)

    def atualizar_label_arquivo(self, texto):
        self.label_arquivo.config(text=texto)

    def mostrar_erro(self, mensagem):
        messagebox.showerror("Erro", mensagem)

    def mostrar_sucesso(self, mensagem):
        messagebox.showinfo("Sucesso", mensagem)