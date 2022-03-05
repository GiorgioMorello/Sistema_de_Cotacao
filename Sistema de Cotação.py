from tkinter import ttk
from tkcalendar import DateEntry
import tkinter as tk
from tkinter.filedialog import askopenfilename
import requests
import pandas as pd
import numpy as np
from datetime import datetime


janela = tk.Tk()
janela.title("Sistemas de Cotação")

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_json = requisicao.json()
lista_moeda = list(dicionario_json.keys())


def pegar_cotacao():
    moeda = combobox.get()
    data_cotacao = selecionar_data.get()
    dia = data_cotacao[:2]
    mes = data_cotacao[3:-5]
    ano = data_cotacao[6:]
    link = f"https://economia.awesomeapi.com.br/{moeda}-BRL/200?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    if len(cotacao) > 0:
        valor_moeda = float(cotacao[0]['bid'])
        resultado["text"] = f"Valor da Moeda: {moeda}: R${valor_moeda:.3f}"
    else:
        resultado["text"] = "Não foi possivel buscar a cotação"


def selecionar():
    caminho_arquivo = askopenfilename(title="Selecione um arquivo")
    var_caminho.set(caminho_arquivo)
    if caminho_arquivo:
        arquivo["text"] = f"Arquivo: {caminho_arquivo}"


def atualizar():
    try:
        #Ler o DataFrame de moedas
        df = pd.read_excel(var_caminho.get())
        moedas = df.iloc[:, 0]
        #Pegar a data de inicio e fim
        data_i = data_inicial.get()
        data_f = data_final.get()
        dia_i = data_i[:2]
        mes_i = data_i[3:-5]
        ano_i = data_i[6:]

        dia_f = data_f[:2]
        mes_f = data_f[3:-5]
        ano_f = data_f[6:]

        #Pra cada moeda
        for coin in moedas:
            #Pegar todas as cotações da moeda
            link = f"https://economia.awesomeapi.com.br/{coin}-BRL/200?start_date=" \
                   f"{ano_i}{mes_i}{dia_i}&" \
                   f"end_date={ano_f}{mes_f}{dia_f}"

            requisicao_moeda = requests.get(link)
            cotacoes = requisicao_moeda.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                valor_moeda = float(cotacao['bid'])
                data2 = datetime.fromtimestamp(timestamp)
                data2 = data2.strftime('%d/%m/%Y')
                if data2 not in df:
                    df[data2] = np.nan

                df.loc[df.iloc[:, 0] == coin, data2] = valor_moeda
        #Criar um novo arquivo com todas as cotações
        df.to_excel("Cotações das Moedas.xlsx")
        label_atualizar['text'] = "Arquivo Atualizado"

    except:
        label_atualizar['text'] = "Selecione apenas arquivos em Excel"


def proxima():
    janela.destroy()
    import Page2

label_cotacao = tk.Label(text="Cotação de uma Moeda Específica", borderwidth=3, relief='ridge')
label_cotacao.grid(row=0, column=0, padx=8, pady=8, sticky="NSWE", columnspan=3)


label_selecionar_moeda = tk.Label(text="Selecione a Moeda: ", anchor='e')
label_selecionar_moeda.grid(row=1, column=0, columnspan=2, padx=8, pady=8, sticky="NSWE")

combobox = ttk.Combobox(janela, values=lista_moeda)
combobox.grid(row=1, column=2, padx=8, pady=8, sticky="NSWE")


label_selecionar_data = tk.Label(text="Selecione o dia para pegar a cotação: ", anchor='e')
label_selecionar_data.grid(row=2, column=0, columnspan=2, padx=8, pady=8, sticky="NSWE")

selecionar_data = DateEntry(year=2022, locale='pt_br')
selecionar_data.grid(row=2, column=2, columnspan=2, padx=8, pady=8, sticky="NSWE")


resultado = tk.Label(text="")
resultado.grid(row=3, column=0, columnspan=2, padx=8, pady=8, sticky="NSWE")

botao_pegar_cotacao = tk.Button(text="Buscar Cotação", command=pegar_cotacao, borderwidth=3, relief='raised')
botao_pegar_cotacao.grid(row=3, column=2, columnspan=2, padx=8, pady=8, sticky="NSWE")



#SEGUNDA PARTE
label_multiplas = tk.Label(text="Cotação de Multiplas Moedas", borderwidth=3, relief='ridge')
label_multiplas.grid(row=4, column=0, padx=8, pady=8, sticky="NSWE", columnspan=3)


selecionar_arquivo = tk.Label(text="Selecione um arquivo Excel com as moedas na coluna A: ")
selecionar_arquivo.grid(row=5, column=0, padx=8, pady=8, sticky="NSWE", columnspan=2)

var_caminho = tk.StringVar()

botao_selecionar = tk.Button(text='Escolher Arquivo (Excel)', command=selecionar, borderwidth=3, relief='raised')
botao_selecionar.grid(row=5, column=2)


arquivo = tk.Label(text="Nenhum Arquivo Selecionado", anchor='e')
arquivo.grid(row=6, column=0, columnspan=3, padx=8, pady=8, sticky='news')

label_data_inicial = tk.Label(text='Data Inicial', anchor='e')
label_data_inicial.grid(row=7, column=0, padx=8, pady=8, sticky='news')

label_data_final = tk.Label(text="Data Final", anchor='e')
label_data_final.grid(row=8, column=0, padx=8, pady=8, sticky='news')

data_inicial = DateEntry(year=2022, locale='pt_br')
data_inicial.grid(row=7, column=1, padx=8, pady=8)
data_final = DateEntry(year=2022, locale='pt_br')
data_final.grid(row=8, column=1, padx=8, pady=8)


botao_atualizar = tk.Button(text="Atualizar Cotações", command=atualizar)
botao_atualizar.grid(row=9, column=0, padx=8, pady=8, sticky='news')

label_atualizar = tk.Label(text="")
label_atualizar.grid(row=9, column=1, columnspan=2, padx=8, pady=8, sticky='news')


botao_fechar = tk.Button(text='Fechar', command=janela.quit, borderwidth=3, relief='raised')
botao_fechar.grid(row=10, column=2, pady=8, padx=8, sticky='news')

#Próxima Página

label_proxima = tk.Button(text="Próxima Pagina", command=proxima, borderwidth=3, relief='raised')
label_proxima.grid(row=11, column=1, padx=8, pady=8, sticky='news')



janela.mainloop()
