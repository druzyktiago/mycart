import tkinter as tk
from tkinter import scrolledtext, filedialog
import os
import re
import pandas as pd
from tkinter import filedialog
from tkinter import messagebox
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def processarone():
    # Abrir a caixa de diálogo para selecionar o arquivo
    arquivo_path = filedialog.askopenfilename(filetypes=[("Arquivos de Texto", "*.txt")])
    if not arquivo_path:
        return

    # Processar o arquivo selecionado
    numeros, unidades = ler_arquivo_txt(arquivo_path)

    # Processar o arquivo selecionado
    with open(arquivo_path, encoding='utf-8') as arquivo:
        linhas = arquivo.read().splitlines()
        apartamentos = []
        matriculas = []
        blocos = []
        ultimo_apartamento = ''
        ultimo_bloco = ''
        apartamentos_duplicados = set()  # Para rastrear apartamentos duplicados

        for linha in linhas:
            palavras_cinco_digitos = re.findall(r'\b\d{5}\b', linha)
            match = re.search(r'(?i)\b[bB]loco\s(\d{2})\b', linha)
            palavras_tres_digitos = re.findall(r'(?<![a-zA-Z])([Nn]º\d{3})(?![a-zA-Z])', linha)

            match_apartamento = re.search(r'\b(\d{3})\b', linha)
            if match_apartamento and "apartamento" in linha.lower():
                ultimo_apartamento = match_apartamento.group(1)

            if match:
                ultimo_bloco = match.group(1)
            
            if palavras_cinco_digitos:
                # Verificar se o apartamento não está duplicado no mesmo bloco
                chave_duplicada = (ultimo_bloco, ultimo_apartamento)
                if chave_duplicada not in apartamentos_duplicados:
                    apartamentos_duplicados.add(chave_duplicada)
                    apartamentos.append(ultimo_apartamento)
                    matriculas.append(palavras_cinco_digitos[0])
                    blocos.append(ultimo_bloco)
            
            if palavras_tres_digitos:
                # Verificar se o apartamento não está duplicado no mesmo bloco
                chave_duplicada = (ultimo_bloco, ultimo_apartamento)
                if chave_duplicada not in apartamentos_duplicados:
                    apartamentos_duplicados.add(chave_duplicada)
                    apartamentos.append(ultimo_apartamento)
                    matriculas.append(palavras_tres_digitos[0])
                    blocos.append(ultimo_bloco)

    # Verificar se as listas têm o mesmo comprimento
    max_length = max(len(apartamentos), len(matriculas), len(blocos), len(unidades))
    apartamentos += [''] * (max_length - len(apartamentos))
    matriculas += [''] * (max_length - len(matriculas))
    blocos += [''] * (max_length - len(blocos))
    unidades += [''] * (max_length - len(unidades))

    # Criar um DataFrame com as colunas separadas
    df = pd.DataFrame({"APARTAMENTOS": apartamentos, "BLOCOS": blocos, "MATRICULAS": matriculas})
    
    # Criar uma coluna combinada com as informações
    df['INFORMACOES'] = df['BLOCOS'].astype(str) + ' - ' + df['APARTAMENTOS'].astype(str) + ' - ' + df['MATRICULAS'].astype(str)
    
    # Ordenar o DataFrame pela coluna combinada em ordem crescente
    df.sort_values("INFORMACOES", inplace=True, ignore_index=True)
    
    # Remover a coluna combinada se não for mais necessária
    df.drop(columns=['INFORMACOES'], inplace=True)
    
    # Remover todas as duplicatas
    df.drop_duplicates(inplace=True)

    # Atualizar o widget resultado_text com o resultado
    resultado_text.insert(tk.END, df.to_string(index=False))


def ler_arquivo_txt(nome_arquivo):
    with open(nome_arquivo, 'r', encoding='utf-8') as arquivo:
        linhas = arquivo.readlines()
        numeros = []
        unidades = []  

        padrao_residencia = re.compile(r"RESIDÊNCIA N\.\s*(\d+)", re.IGNORECASE)
        padrao_unidade = re.compile(r"UNIDADE N\.\s*(\d+)", re.IGNORECASE)

        for linha in linhas:
            matches_residencia = padrao_residencia.findall(linha)
            matches_unidade = padrao_unidade.findall(linha)

            for match in matches_residencia:
                numeros.append(match)
                unidades.append(match)
                
            for match in matches_unidade:
                numeros.append(match)
                unidades.append(match)

        return numeros, unidades

# Função para gerar a sequência com base nos dados inseridos
def processar(event=None):
    entrada = entry_numeros.get()

    # Verifica se a entrada contém dois números separados por um espaço
    numeros = entrada.split()
    if len(numeros) != 2:
        resultado.delete(1.0, tk.END)
        resultado.insert(tk.END, "Digite o número do primeiro apartamento um espaço e o último apartamento")
        return

    numero1 = int(numeros[0])
    numero2 = int(numeros[1])

    # Verifica se o campo de "Número de Blocos" foi preenchido
    blocos = entry_blocos.get()
    if not blocos.isdigit():
        resultado.delete(1.0, tk.END)
        resultado.insert(tk.END, "Digite um número válido para o campo Número de Blocos")
        return

    blocos = int(blocos)

    def inner(numero1, numero2):
        ultimo_digito_numero2 = numero2 % 10
        sequencia = []

        for num_sequencia in range(1, blocos + 1):
            numero1_temp = numero1  # Salva o valor original do primeiro número
            while numero1 <= numero2:
                if numero1 % 100 != 0:
                    sequencia.append((num_sequencia, numero1))

                ultimo_digito_numero1 = numero1 % 10

                if ultimo_digito_numero1 < ultimo_digito_numero2:
                    numero1 += 1
                else:
                    primeiro_digito_numero1 = numero1 // 100
                    primeiro_digito_numero1 += 1
                    numero1 = primeiro_digito_numero1 * 100

            numero1 = numero1_temp  # Restaura o valor original do primeiro número

        return sequencia

    sequencia = inner(numero1, numero2)
    resultado.delete(1.0, tk.END)
    for item in sequencia:
        # Adicione um zero à esquerda ao bloco se for menor que 10
        bloco_formatado = str(item[0]).zfill(2)
        resultado.insert(tk.END, f" {item[1]}  {bloco_formatado}\n")

    # Limpa o campo de entrada para o segundo número
    entry_numeros.delete(0, tk.END)


# Função para criar a tabela (a ser implementada no futuro)
def criar_tabela():
    pass

# Cria a janela principal
janela = tk.Tk()
janela.title("Gerador de Sequência")

# Criação de uma caixa de entrada para os números
label_numeros = tk.Label(janela, text="Apartamentos 1º e último:")
label_numeros.pack()

entry_numeros = tk.Entry(janela)
entry_numeros.pack()

# Criação de um campo de entrada para o número de blocos
label_blocos = tk.Label(janela, text="Último bloco:")
label_blocos.pack()

entry_blocos = tk.Entry(janela)
entry_blocos.pack()



def atualizar_terceiro_resultado():
    global resultado_final
    # Obtenha o conteúdo da segunda tabela (a tabela de referência)
    tabela_segundo_resultado = resultado_text.get("1.0", tk.END)

    # Divida o conteúdo da segunda tabela em linhas
    linhas_segunda_tabela = tabela_segundo_resultado.split('\n')

    # Crie uma lista para armazenar o conteúdo da primeira tabela (a tabela atual)
    linhas_primeira_tabela = resultado.get("1.0", tk.END).split('\n')

    # Crie uma lista para armazenar o resultado final
    resultado_final = []

    # Crie um conjunto para rastrear as linhas da segunda tabela que foram correspondidas
    linhas_correspondentes = set()

    # Percorra as linhas da primeira tabela
    for linha_primeira_tabela in linhas_primeira_tabela:
        # Divida a linha em partes
        partes_primeira_tabela = linha_primeira_tabela.split()

        # Verifique se a linha possui pelo menos 2 partes (bloco e apartamento)
        if len(partes_primeira_tabela) >= 2:
            bloco_primeira_tabela, apartamento_primeira_tabela = partes_primeira_tabela[:2]

            # Inicialize uma variável para controlar se encontrou uma correspondência
            encontrou_correspondencia = False

            # Percorra as linhas da segunda tabela
            for i, linha_segunda_tabela in enumerate(linhas_segunda_tabela):
                # Divida a linha da segunda tabela em partes
                partes_segunda_tabela = linha_segunda_tabela.split()

                # Verifique se a linha da segunda tabela possui pelo menos 2 partes (bloco e apartamento)
                if len(partes_segunda_tabela) >= 2:
                    bloco_segunda_tabela, apartamento_segunda_tabela = partes_segunda_tabela[:2]

                    # Verifique se o bloco e apartamento coincidem
                    if bloco_primeira_tabela == bloco_segunda_tabela and apartamento_primeira_tabela == apartamento_segunda_tabela:
                        # Se coincidirem, adicione a linha da segunda tabela ao resultado final
                        resultado_final.append(linha_segunda_tabela)
                        encontrou_correspondencia = True
                        linhas_correspondentes.add(i)
                        break  # Sai do loop interno

            # Se não encontrou correspondência, adicione uma linha formatada ao resultado final
            if not encontrou_correspondencia:
                # Determine a largura desejada para cada coluna
                largura_bloco = 7
                largura_apartamento = 10

                # Formate a linha da primeira tabela para que fique como as outras
                linha_formatada = f"         {bloco_primeira_tabela:{largura_bloco}} {apartamento_primeira_tabela:{largura_apartamento}}"
                resultado_final.append(linha_formatada)



def executar_tudo():
    # Executar as funções em ordem
    processar()
    processarone()
    criar_tabela()
    atualizar_terceiro_resultado()

    # Construa o resultado final na terceira tabela
    resultado_terceiro.delete("1.0", tk.END)
    resultado_terceiro.insert(tk.END, '\n'.join(resultado_final))

    # Salvar a terceira tabela em um arquivo Excel (.xlsx)
    salvar_terceira_tabela()

import openpyxl

# Função para salvar a terceira tabela com um cabeçalho personalizado
def salvar_terceira_tabela():
    global resultado_terceiro  # Certifique-se de que resultado_terceiro está definido em outro lugar

    # Crie uma janela Tkinter para obter o nome do cabeçalho do usuário
    nome_cabecalho = tk.simpledialog.askstring("Nome do Cabeçalho", "Digite o nome do cabeçalho:")

    # Verifique se o usuário inseriu um nome de cabeçalho
    if nome_cabecalho is None:
        return

    # Verifique se há dados na terceira tabela
    if resultado_terceiro.get("1.0", tk.END) == '\n':
        messagebox.showinfo("Aviso", "A terceira tabela está vazia. Nada a salvar.")
        return

    # Obtenha o conteúdo da terceira tabela
    conteudo_terceira_tabela = resultado_terceiro.get("1.0", tk.END)

    # Divida o conteúdo da terceira tabela em linhas
    linhas_terceira_tabela = conteudo_terceira_tabela.split('\n')

    # Crie uma lista de listas para os dados em células separadas
    dados_em_celulas = [linha.split() for linha in linhas_terceira_tabela]

    # Crie um DataFrame a partir das listas de dados
    df_terceira_tabela = pd.DataFrame(dados_em_celulas)

    # Adicione o nome do cabeçalho ao DataFrame
    df_terceira_tabela.columns = [nome_cabecalho] * len(df_terceira_tabela.columns)

    # Crie um arquivo Excel (.xlsx) e salve os dados nele
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        df_terceira_tabela.to_excel(file_path, index=False)

        # Abra o arquivo Excel recém-criado para mesclar as células do cabeçalho
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active

        # Mesclar as células do cabeçalho (de A1 até a última coluna)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_terceira_tabela.columns))

        # Salve as alterações no arquivo Excel
        wb.save(file_path)
        wb.close()



# Crie os campos de texto (não empacotados na janela principal)
resultado = scrolledtext.ScrolledText(janela, width=40, height=10, wrap=tk.WORD)
resultado_text = scrolledtext.ScrolledText(janela, width=40, height=10, wrap=tk.WORD)

# Crie a terceira tabela
resultado_terceiro = scrolledtext.ScrolledText(janela, width=40, height=10, wrap=tk.WORD)

# Crie o botão para executar tudo
executar_tudo_button = tk.Button(janela, text="Executar Tudo", command=executar_tudo)

# Empacote os widgets na janela principal
executar_tudo_button.pack()
resultado_terceiro.pack(fill=tk.BOTH, expand=True)

# Inicie a interface gráfica
janela.mainloop()