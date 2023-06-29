import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
import win32gui

# Impedido a janela do CMD abrir junto quando executado o app
win32gui.ShowWindow(win32gui.GetForegroundWindow(), 0)

# Lendo o arquivo Excel
df = pd.read_excel("PlanilhaCidadesCep.xlsx")

# Função para verificar o CEP
def verificar_cep():
    # Obtendo o CEP digitado pelo usuário
    cep = entrada_cep.get()

    # Verificando se o CEP possui 8 dígitos
    if len(cep) != 8:
        messagebox.showinfo("Erro", "Digite um CEP válido")
        entrada_cep.select_range(0, tk.END)
        return

    # Verificando se o CEP está dentro do range de algum registro na tabela
    if cep[0] in ['5', '6', '7']:
        resultado.set("Fora da área de coleta! NÃO TRABALHAMOS COM ESTA REGIÃO!!")
        messagebox.showinfo("Resultado", "Fora da área de coleta! \nNÃO TRABALHAMOS COM ESTA REGIÃO!!")
        entrada_cep.select_range(0, tk.END)
    elif cep[0] == '4':
        for i, row in df.iterrows():
            if row['Cep inicio'] <= int(cep) <= row['Cep fim']:
                cidade = row['Cidades']
                resultado.set(cidade + " possui coleta")
                messagebox.showinfo("Resultado", cidade + " - Possui coleta")
                entrada_cep.select_range(0, tk.END)
                break
        else:
            resultado.set("Fora da área de coleta! NÃO TRABALHAMOS COM ESTA REGIÃO!!")
            messagebox.showinfo("Resultado", "Fora da área de coleta! \nNÃO TRABALHAMOS COM ESTA REGIÃO!!")
            entrada_cep.select_range(0, tk.END)
    else:
        for i, row in df.iterrows():
            if row['Cep inicio'] <= int(cep) <= row['Cep fim']:
                cidade = row['Cidades']
                resultado.set(cidade + " possui coleta")
                messagebox.showinfo("Resultado", cidade + " - Possui coleta")
                entrada_cep.select_range(0, tk.END)
                break
        else:
            resultado.set("Fora da área de coleta!")
            messagebox.showinfo("Resultado", "Fora da área de coleta!")
            entrada_cep.select_range(0, tk.END)

def conferencia_em_massa():
    global df  # Adiciona essa linha para referenciar a variável global df

    # Abrir uma caixa de diálogo para selecionar o arquivo de entrada
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    
    if file_path:
        # Realizar a conferência em massa
        df_input = pd.read_excel(file_path)

        # Verificar se a coluna "CEP" existe no DataFrame
        if 'CEP' not in df_input.columns:
            messagebox.showinfo("Erro", "A coluna 'CEP' não foi encontrada no arquivo selecionado.")
            return

        resultados = []
        
        for _, row in df_input.iterrows():
            cep = str(row['CEP'])  # Converter o CEP para string

            cidade_encontrada = False

            for _, cidade_row in df.iterrows():
                if cidade_row['Cep inicio'] <= int(cep) <= cidade_row['Cep fim']:
                    cidade = cidade_row['Cidades']
                    resultados.append(cidade + " possui coleta")
                    cidade_encontrada = True
                    break

            if not cidade_encontrada:
                resultados.append("Fora da área de coleta!")

        # Salvar os resultados em um arquivo de saída
        output_file_path = filedialog.asksaveasfilename(defaultextension=".txt")
        if output_file_path:
            with open(output_file_path, 'w') as file:
                file.write('\n'.join(resultados))
            messagebox.showinfo("Conferência em Massa", "Conferência concluída! Resultados salvos com sucesso.")

# Criando a janela principal
janela = tk.Tk()
janela.title("Verificador de CEP")
janela.geometry("300x70")

# Criando a barra de menus
menu_bar = tk.Menu(janela)
janela.config(menu=menu_bar)

# Criando o menu "Arquivo"
menu_arquivo = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Arquivo", menu=menu_arquivo)

# Adicionando o item "Conferência em Massa" ao menu "Arquivo"
menu_arquivo.add_command(label="Conferência em Massa", command=conferencia_em_massa)

# Criando o rótulo e a entrada para o CEP
tk.Label(janela, text="Digite um CEP:").pack()
entrada_cep = tk.Entry(janela)
entrada_cep.pack()

# Criando o botão para verificar o CEP
botao_verificar = tk.Button(janela, text="Verificar", command=verificar_cep)
botao_verificar.pack()

# Vinculando a tecla Enter à função de verificação do CEP
entrada_cep.bind('<Return>', lambda event: verificar_cep())

# Criando o rótulo para exibir o resultado
resultado = tk.StringVar()

# Função para encerrar o programa
def fechar_janela():
    janela.quit()

# Definindo a função para encerrar o programa quando a janela for fechada
janela.protocol("WM_DELETE_WINDOW", fechar_janela)

# Iniciando o loop principal do Tkinter
janela.mainloop()
