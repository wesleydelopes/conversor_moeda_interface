import requests
import pandas as pd
from datetime import datetime
import time
from tkinter import *

def atualizar_cotacoes():
    try:
        # Requisição dos dados de cotação das moedas
        requisicao = requests.get("https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL,ETH-BRL,JPY-BRL,GBP-BRL,AUD-BRL,CHF-BRL")
        requisicao_dic = requisicao.json()

        cotacoes = {
            "Dólar": float(requisicao_dic["USDBRL"]["bid"]),
            "Euro": float(requisicao_dic["EURBRL"]["bid"]),
            "Iene": float(requisicao_dic["JPYBRL"]["bid"]),
            "Ethereum": float(requisicao_dic["ETHBRL"]["bid"].replace(",", ".")),
            "Bitcoin": float(requisicao_dic["BTCBRL"]["bid"].replace(",", ".")),
            "Libra Esterlina": float(requisicao_dic["GBPBRL"]["bid"]),
            "Dólar Australiano": float(requisicao_dic["AUDBRL"]["bid"]),
            "Franco Suíço": float(requisicao_dic["CHFBRL"]["bid"])
        }

        # Lendo a tabela de cotações atual
        tabela = pd.read_excel("Cotações.xlsx")

        # Atualizando as cotações na tabela
        tabela.loc[0, "Cotação"] = float(cotacoes["Dólar"])
        tabela.loc[1, "Cotação"] = float(cotacoes["Euro"])
        tabela.loc[2, "Cotação"] = float(cotacoes["Bitcoin"]) * 1000
        tabela.loc[3, "Cotação"] = float(cotacoes["Ethereum"]) * 1000
        tabela.loc[4, "Cotação"] = float(cotacoes["Iene"])
        tabela.loc[5, "Cotação"] = float(cotacoes["Libra Esterlina"])
        tabela.loc[6, "Cotação"] = float(cotacoes["Dólar Australiano"])
        tabela.loc[7, "Cotação"] = float(cotacoes["Franco Suíço"])

        # Atualizando a data de última atualização
        tabela.loc[0, "Data Última Atualização"] = datetime.now()

        # Salvando a tabela atualizada em um arquivo Excel
        tabela.to_excel("Cotações.xlsx", index=False)

        # Imprimindo a mensagem de atualização
        print(f"Cotação Atualizada. {datetime.now()}")
        print('-=' * 25)

        return cotacoes
    except Exception as e:
        print(f"Erro ao atualizar cotações: {e}")
        return None

def converter_moeda():
    moeda = entrada_moeda.get().lower()
    valor = float(entrada_valor.get())

    cotacoes = atualizar_cotacoes()

    if cotacoes:
        try:
            moeda_nome = None
            resultado = None

            if moeda.lower() == 'dólar' or moeda == 'dolar':
                moeda_nome = 'Dólar'
                resultado = valor / cotacoes["Dólar"]
            elif moeda.lower() == 'euro':
                moeda_nome = 'Euro'
                resultado = valor / cotacoes["Euro"]
            elif moeda.lower() == 'iene':
                moeda_nome = 'Iene'
                resultado = valor / cotacoes["Iene"]
            elif moeda.lower() == 'ethereum' or moeda.lower() == 'eth':
                moeda_nome = 'Eth'
                resultado = valor / cotacoes["Ethereum"]
            elif moeda.lower() == 'franco' or moeda.lower() == 'franco suíço':
                moeda_nome = 'Franco Suiço'
                resultado = valor / cotacoes["Franco Suíço"]
            elif moeda.lower() == 'libra' or moeda.lower() == 'libra esterlina':
                moeda_nome = 'Libra Esterlina'
                resultado = valor / cotacoes["Libra Esterlina"]
            elif moeda.lower() == 'dólar australiano' or moeda.lower() == 'dolar australiano':
                moeda_nome = 'Dólar Australiano'
                resultado = valor / cotacoes["Dólar Australiano"]
            elif moeda.lower() == 'btc' or moeda.lower() == 'bitcoin':
                moeda_nome = 'Bitcoin'
                resultado = valor / cotacoes["Bitcoin"]

            if moeda_nome is not None:
                resultado_label.config(text=f'Com R${valor:.2f}, você receberá {resultado:.2f} {moeda_nome} na cotação atual.')
            else:
                resultado_label.config(text="Moeda inválida.")
        except Exception as e:
            resultado_label.config(text="Ocorreu um erro durante a conversão.")

janela = Tk()
janela.title('Cotação atual das moedas')

texto_orientacao = Label(janela, text='Digite a moeda e o valor para conversão:')
texto_orientacao.grid(column=0, row=0, columnspan=2)

entrada_moeda_label = Label(janela, text='Moeda:')
entrada_moeda_label.grid(column=0, row=1)

entrada_moeda = Entry(janela)
entrada_moeda.grid(column=1, row=1)

entrada_valor_label = Label(janela, text='Valor:')
entrada_valor_label.grid(column=0, row=2)

entrada_valor = Entry(janela)
entrada_valor.grid(column=1, row=2)

converter_botao = Button(janela, text='Converter', command=converter_moeda)
converter_botao.grid(column=0, row=3, columnspan=2)

resultado_label = Label(janela, text='')
resultado_label.grid(column=0, row=4, columnspan=2)

janela.mainloop()

while True:
    cotacoes = atualizar_cotacoes()
    
    if cotacoes:
        # Pedindo a moeda e o valor para conversão
        moeda = input('Qual moeda você quer fazer a cotação? \n').lower()
        valor = float(input('Digite o valor para ver a cotação atual: \n'))
        
        converter_moeda(moeda, valor, cotacoes)

        print('-=' * 25)
    
    time.sleep(2)
