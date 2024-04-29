import matplotlib.pyplot as plt
from collections import Counter

class GeradorGrafico:
    def __init__(self, dados):
        self.dados = dados

    def criar_grafico_pizza_sexo(self):
        try:
            sexo_counts = Counter(self.dados['Sexo'])
            labels = list(sexo_counts.keys())
            sizes = list(sexo_counts.values())
            colors = ['blue', 'pink']

            fig, ax = plt.subplots()
            ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%')
            ax.set_title('Distribuição por Sexo')

            plt.show()  # Mostrar o gráfico diretamente ao invés de integrar ao Tkinter
        except Exception as e:
            print(f"Erro ao criar o gráfico de sexo: {e}")

    def criar_grafico_pizza_estado_civil(self):
        try:
            estado_civil_counts = Counter(self.dados['Estado Civil'])
            labels = list(estado_civil_counts.keys())
            sizes = list(estado_civil_counts.values())
            colors = ['green', 'orange', 'blue', 'red']  # Adicione mais cores se necessário

            fig, ax = plt.subplots()
            ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%')
            ax.set_title('Distribuição por Estado Civil')

            plt.show()  # Mostrar o gráfico diretamente ao invés de integrar ao Tkinter
        except Exception as e:
            print(f"Erro ao criar o gráfico de estado civil: {e}")
