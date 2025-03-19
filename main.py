import tkinter as tk
from interface import Interface
from logica import Logica

class Aplicacao:
    def __init__(self, root):
        self.logica = Logica()
        self.interface = Interface(
            root,
            on_selecionar_diretorio=self.selecionar_diretorio,
            on_selecionar_arquivo=self.selecionar_arquivo_word,
            on_processar=self.processar
        )

    def selecionar_diretorio(self):
        diretorio = self.logica.selecionar_diretorio()
        if diretorio:
            self.interface.atualizar_label_diretorio(f"📁 Diretório selecionado: {diretorio}")

    def selecionar_arquivo_word(self):
        arquivo = self.logica.selecionar_arquivo_word()
        if arquivo:
            self.interface.atualizar_label_arquivo(f"📄 Arquivo selecionado: {arquivo}")

    def processar(self):
        resultado = self.logica.inserir_imagens_no_word()
        if resultado.startswith("✅"):
            self.interface.mostrar_sucesso(resultado)
        else:
            self.interface.mostrar_erro(resultado)

if __name__ == "__main__":
    root = tk.Tk()
    app = Aplicacao(root)
    root.mainloop()