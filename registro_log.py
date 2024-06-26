import os
from datetime import datetime

class Logger:
    def __init__(self, caminho, log):
        self.caminho = caminho
        self.log = log

    def verificar_criar_caminho(self):
        if not os.path.exists(self.caminho):
            os.makedirs(self.caminho)

    def salvar_log(self, mensagem):
        mensagem = str(mensagem)
        traco = "------------------------------------------------"
        self.verificar_criar_caminho()
        data_atual = datetime.now()
        hora_formatada = data_atual.strftime("%H:%M:%S")
        nome_arquivo = f"log_{data_atual.strftime('%Y%b%d')}.txt"
        caminho_arquivo = os.path.join(self.caminho, nome_arquivo)
        with open(caminho_arquivo, 'a') as arquivo:
            if mensagem == "--" : 
                arquivo.write(f"{traco}\n")
            else:
                arquivo.write(f"{hora_formatada} - {mensagem}\n")
            


