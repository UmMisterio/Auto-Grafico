import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os

# Function to generate the plot
def gerar_grafico():
    # Check if the file is selected or informed by the user
    if arquivo.get():
        # Read the Excel file
        df = pd.read_excel(arquivo.get())
    else:
        # Show an error message if the file is not selected or informed
        messagebox.showerror("Erro", "Selecione ou informe um arquivo")
        return
    
    # Convert the column 'y' to numeric type
    df['y'] = pd.to_numeric(df['y'], errors='coerce')
    
    # Remove rows with missing values
    df.dropna(subset=['y'], inplace=True)
    
    # Group sales by region
    df_regiao = df.groupby('x')['y'].sum()
    
    # Create the bar plot with Seaborn
    sns.set_style('whitegrid')
    sns.set_palette('husl')
    plt.figure(figsize=(10, 6))
    if orientacao.get() == 'vertical':
        sns.barplot(x=df_regiao.index, y=df_regiao.values, orient='v')
    else:
        sns.barplot(x=df_regiao.values, y=df_regiao.index, orient='h')
    plt.title(title.get(), fontsize=16, fontweight='bold')  # Use the custom title set by the user
    plt.xlabel(xlabel.get(), fontsize=12)  # Use the custom X-label set by the user
    plt.ylabel(ylabel.get(), fontsize=12)  # Use the custom Y-label set by the user
    plt.xticks(fontsize=11)
    plt.yticks(fontsize=11)

    # Add the sales values above the bars
    if orientacao.get() == 'vertical':
        for i, v in enumerate(df_regiao.values):
            plt.text(i, v+10, f'{v:,.0f}', ha='center', fontsize=11, fontweight='bold')
    else:
        for i, v in enumerate(df_regiao.values):
            plt.text(v+10, i, f'{v:,.0f}', va='center', fontsize=11, fontweight='bold')

    # Show the plot
    plt.tight_layout()
    plt.show()

# Function to select the file
def selecionar_arquivo():
    arquivo_path = filedialog.askopenfilename()
    arquivo.set(arquivo_path)

# Main window
janela = tk.Tk()
janela.title("Gráfico de Vendas por Região")

# Variable to store the file name
arquivo = tk.StringVar()

# Variable to store the orientation selected by the user
orientacao = tk.StringVar(value='vertical')

# Variables to store the custom title, X-label and Y-label set by the user
title = tk.StringVar(value="Consumidores atendidos em 2020")
xlabel = tk.StringVar(value="Consumidores Atendidos")
ylabel = tk.StringVar(value="Ações")

# Frame to select the file
frame_arquivo = tk.Frame(janela)
frame_arquivo.pack(pady=10)

# Label and Entry to display the selected file name
label_arquivo = tk.Label(frame_arquivo, text="Arquivo:")
label_arquivo.pack(side=tk.LEFT, padx=(0, 10))
entry_arquivo = tk.Entry(frame_arquivo, textvariable=arquivo)
entry_arquivo.pack(side=tk.LEFT)

# Button to select the file
botao_arquivo = tk.Button(frame_arquivo, text="Selecionar Arquivo", command=selecionar_arquivo)
botao_arquivo.pack(side=tk.LEFT)

# Frame to select the plot orientation
frame_orientacao = tk.Frame(janela)
frame_orientacao.pack()

# Label to select the plot orientation
label_orientacao = tk.Label(frame_orientacao, text="Orientação do gráfico:")
label_orientacao.pack(side=tk.LEFT, padx=(0, 10))

# Radio buttons to select the plot orientation
radio_vertical = tk.Radiobutton(frame_orientacao, text="Vertical", variable=orientacao, value="vertical")
radio_vertical.pack(side=tk.LEFT)
radio_horizontal = tk.Radiobutton(frame_orientacao, text="Horizontal", variable=orientacao, value="horizontal")
radio_horizontal.pack(side=tk.LEFT)

# Frame to customize the plot
frame_personalizar = tk.Frame(janela)
frame_personalizar.pack()

# Label and Entry to customize the title of the plot
label_title = tk.Label(frame_personalizar, text="Título:")
label_title.pack(side=tk.LEFT, padx=(0, 10))
entry_title = tk.Entry(frame_personalizar, textvariable=title)
entry_title.pack(side=tk.LEFT)

#Label e Entry para personalizar o eixo x do gráfico
label_xlabel = tk.Label(frame_personalizar, text="Eixo X:")
label_xlabel.pack(side=tk.LEFT, padx=(0, 10))
entry_xlabel = tk.Entry(frame_personalizar, textvariable=xlabel)
entry_xlabel.pack(side=tk.LEFT)

#Label e Entry para personalizar o eixo y do gráfico
label_ylabel = tk.Label(frame_personalizar, text="Eixo Y:")
label_ylabel.pack(side=tk.LEFT, padx=(0, 10))
entry_ylabel = tk.Entry(frame_personalizar, textvariable=ylabel)
entry_ylabel.pack(side=tk.LEFT)

#Função para gerar o gráfico
def gerar_grafico():
    
    # Lê o arquivo selecionado pelo usuário
    df = pd.read_csv(arquivo.get())

    # Cria o gráfico de barras
    if orientacao.get() == "vertical":
        ax = df.plot.bar(x='Região', y='Vendas')
    else:
        ax = df.plot.barh(x='Região', y='Vendas')

    # Personaliza o título e as legendas do eixo x e y
    ax.set_title(entry_title.get())
    ax.set_xlabel(entry_xlabel.get())
    ax.set_ylabel(entry_ylabel.get())

    # Exibe o gráfico
    plt.show()

#Botão para gerar o gráfico
botao_gerar = tk.Button(janela, text="Gerar Gráfico", command=gerar_grafico)
botao_gerar.pack(pady=10)

#Inicia a janela principal
janela.mainloop()