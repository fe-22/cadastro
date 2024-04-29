from tkinter import *
from PIL import Image, ImageTk
from cad import carregar_imagem

class Util:
    def __init__(self):
        self.janela = Tk()
        self.janela.title("Visualizador de Imagem")

        # Carregar a imagem usando a função do módulo cad
        self.imagem = carregar_imagem()
        self.label_imagem = Label(self.janela, image=self.imagem)
        self.label_imagem.pack()

        # Adicionar campo para inserir endereço e CNPJ
        label_endereco = Label(self.janela, text="Endereço: Rua Ernesta Pelosine, 196 - Centro - SBC.")
        label_endereco.pack()
      

        label_cnpj = Label(self.janela, text="CNPJ:")
        label_cnpj.pack()
   

        # Exibir a data na janela
        self.label_data = Label(self.janela)
        self.label_data.pack()

        # Botão para atualizar a data
        botao_atualizar_data = Button(self.janela, text="Atualizar Data", command=self.atualizar_data)
        botao_atualizar_data.pack()

        self.janela.mainloop()

    def atualizar_data(self):
        # Obter data atual do formulário (aqui você precisa implementar a lógica para obter a data atual)
        data_atual = "Implemente a lógica para obter a data atual"
        # Atualizar rótulo com a data atual
        self.label_data.config(text=data_atual)

# Carregar a imagem fora da classe Util
imagem_original = Image.open("C:\\Users\\SAMSUNG\\OneDrive\\vs-python\\cadig.py\\Imagens\\imagem.gif")
imagem_tk = ImageTk.PhotoImage(imagem_original)

# Iniciar o aplicativo
app = Util()
