import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os

# Função para gerar o gráfico
def gerar_grafico():
    # Verifica se o arquivo foi selecionado ou informado pelo usuário
    if arquivo.get():
        # Lê o arquivo Excel
        df = pd.read_excel(arquivo.get())
    else:
        # Mostra uma mensagem de erro caso o arquivo não seja selecionado ou informado
        messagebox.showerror("Erro", "Selecione ou informe um arquivo")
        return

    # Converte a coluna 'Consumidores Atendidos [y]' para o tipo numérico
    df['y'] = pd.to_numeric(df['y'], errors='coerce')

    # Remove linhas com valores ausentes
    df.dropna(subset=['y'], inplace=True)

    # Agrupa as vendas por região
    df_regiao = df.groupby('x')['y'].sum()

    # Cria o gráfico de barras com Seaborn
    sns.set_style('whitegrid')
    sns.set_palette('husl')
    plt.figure(figsize=(10, 6))
    if orientacao.get() == 'vertical':
        sns.barplot(x=df_regiao.index, y=df_regiao.values, orient='v')
    else:
        sns.barplot(x=df_regiao.values, y=df_regiao.index, orient='h')
    plt.title('Consumidores atendidos em 2020', fontsize=16, fontweight='bold')
    plt.xlabel('Consumidores Atendidos', fontsize=12)
    plt.ylabel('Ações', fontsize=12)
    plt.xticks(fontsize=11)
    plt.yticks(fontsize=11)

    # Adiciona os valores de venda acima das barras
    if orientacao.get() == 'vertical':
        for i, v in enumerate(df_regiao.values):
            plt.text(i, v+10, f'{v:,.0f}', ha='center', fontsize=11, fontweight='bold')
    else:
        for i, v in enumerate(df_regiao.values):
            plt.text(v+10, i, f'{v:,.0f}', va='center', fontsize=11, fontweight='bold')

    # Mostra o gráfico
    plt.tight_layout()
    plt.show()

# Função para selecionar o arquivo
def selecionar_arquivo():
    arquivo_path = filedialog.askopenfilename()
    arquivo.set(arquivo_path)

# Cria a janela principal
janela = tk.Tk()
janela.title("Gráfico de Vendas por Região")

# Variável para armazenar o nome do arquivo
arquivo = tk.StringVar()

# Variável para armazenar a orientação selecionada pelo usuário
orientacao = tk.StringVar(value='vertical')

# Cria o Frame para selecionar o arquivo
frame_arquivo = tk.Frame(janela)
frame_arquivo.pack(pady=10)

# Label e Entry para exibir o nome do arquivo selecionado
label_arquivo = tk.Label(frame_arquivo, text="Arquivo:")
label_arquivo.pack(side=tk.LEFT, padx=(0, 10))
entry_arquivo = tk.Entry(frame_arquivo, textvariable=arquivo)
entry_arquivo.pack(side=tk.LEFT)

# Botão para selecionar o arquivo
botao_arquivo = tk.Button(frame_arquivo, text="Selecionar Arquivo", command=selecionar_arquivo)
botao_arquivo.pack(side=tk.LEFT)

# Cria o Frame para selecionar a orientação do gráfico
frame_orientacao = tk.Frame(janela)
frame_orientacao.pack()

# Label para selecionar a orientação do gráfico
label_orientacao = tk.Label(frame_orientacao, text="Orientação do gráfico:")
label_orientacao.pack(side=tk.LEFT, padx=(0, 10))

# Radio buttons para selecionar a orientação do gráfico
radio_vertical = tk.Radiobutton(frame_orientacao, text="Vertical", variable=orientacao, value="vertical")
radio_vertical.pack(side=tk.LEFT)
radio_horizontal = tk.Radiobutton(frame_orientacao, text="Horizontal", variable=orientacao, value="horizontal")
radio_horizontal.pack(side=tk.LEFT)

# Botão para gerar o gráfico
botao_gerar = tk.Button(janela, text="Gerar Gráfico", command=gerar_grafico)
botao_gerar.pack(pady=10)

# Inicia a janela principal
janela.mainloop()