import time
import pandas as pd
import numpy as np
import threading


from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from tkinter import *
from tkinter import ttk
from ttkthemes import ThemedTk

root = Tk()


def selecionar_arquivo():
    arquivo_csv = filedialog.askopenfilename(title="Selecione o arquivo CSV")
    arquivo_csv_entry.delete(0, END)
    arquivo_csv_entry.insert(0, arquivo_csv)
    
Label(root, text="Arquivo CSV:").pack(pady=5)
arquivo_csv_entry = Entry(root, width=50)
arquivo_csv_entry.pack(pady=5)
Button(root, text="Selecionar arquivo", command=selecionar_arquivo).pack(pady=5)


def enviar_mensagens(tempo_min, tempo_max, arquivo_csv):
    # passo 1: entrar no WhatsApp
    driver = webdriver.Chrome()
    driver.get('https://web.whatsapp.com/')
    time.sleep(45)

    # passo 2: pegar lista de contatos e mensagens através de um arquivo excel
    tabela = pd.read_csv(arquivo_csv, encoding='ISO-8859-1', delimiter=';')

    # converte a coluna para o tipo string
    tabela['RE'] = tabela['RE'].astype(str)

    # substitui valores não numéricos por NaN
    tabela['RE'] = tabela['RE'].replace(['NaN', 'inf'], np.nan, regex=True)

    # converte a coluna para o tipo float e arredonda para o inteiro mais próximo
    tabela['RE'] = tabela['RE'].astype(float).round().astype('Int64', errors='ignore')

    # converte a coluna de volta para o tipo string
    tabela['RE'] = tabela['RE'].astype(str)

    contatos = tabela['RE'].tolist()
    mensagens = tabela['mensagemfinal'].tolist()

    # passo 3: enviar mensagens para cada contato
    for i in range(len(contatos)):
        contato = contatos[i]
        mensagem = mensagens[i] 

        # buscar contato
        buscar_contato = driver.find_element('xpath','//*[@id="side"]/div[1]/div/div/div[2]/div/div[1]/p')
        buscar_contato.click()

        #colocar contato no campo de busca
        buscar_contato.send_keys(contato)
        time.sleep(1)
        buscar_contato.send_keys(Keys.ENTER)
        time.sleep(2)

        # enviar mensagem
        mensagem_input = driver.find_element('xpath','//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]')
        mensagem_input.send_keys(mensagem)
        mensagem_input.send_keys(Keys.ENTER)

        #esperar aleatoriamente entre tempo_min e tempo_max segundos e de dez em dez for esperar entre tempo_max e tempo_max + 5 segundos
        if i % 10 == 0:
            time.sleep(np.random.randint(tempo_max, tempo_max + 5))
        else:
            time.sleep(np.random.randint(tempo_min, tempo_max))

from tkinter import *
from tkinter import ttk
import threading

def enviar_mensagens_thread():
    tempo_min = int(tempo_min_entry.get())
    tempo_max = int(tempo_max_entry.get())
    arquivo_csv = arquivo_csv_entry.get()
    t = threading.Thread(target=enviar_mensagens, args=(tempo_min, tempo_max, arquivo_csv))
    t.start()




root = ThemedTk(theme="equilux")
root.title("Envio de Relatórios por WhatsApp")

#inserir logo acima do campos centralizando
#parei aqui

#criar o root geometry de maneira flexivel
root.geometry("600x300+{}+{}".format(root.winfo_screenwidth()//2 - 250, root.winfo_screenheight()//2 - 150))




root.resizable(0, 0)
root.configure(bg="white")

#configurar dados_frame
dados_frame = ttk.LabelFrame(root, text="Dados", padding=10)
dados_frame.place(relx=0.5, rely=0.5, anchor=CENTER)

# Utilizar ttk.Entry ao invés de Entry padrão
arquivo_csv_entry = ttk.Entry(dados_frame, width=50)
arquivo_csv_entry.grid(row=0, column=1, padx=5, pady=5)

tempo_min_entry = ttk.Entry(dados_frame, width=10)
tempo_min_entry.grid(row=1, column=1, padx=5, pady=5)

tempo_max_entry = ttk.Entry(dados_frame, width=10)
tempo_max_entry.grid(row=2, column=1, padx=5, pady=5)

# Adicionar labels aos inputs
ttk.Label(dados_frame, text="Arquivo CSV:", font=("Arial", 12)).grid(row=0, column=0, padx=5, pady=5)
ttk.Label(dados_frame, text="Tempo mínimo de espera:", font=("Arial", 12)).grid(row=1, column=0, padx=5, pady=5)
ttk.Label(dados_frame, text="Tempo máximo de espera:", font=("Arial", 12)).grid(row=2, column=0, padx=5, pady=5)

enviar_button = ttk.Button(root, text="Enviar Mensagens", command=enviar_mensagens_thread)
enviar_button.place(relx=0.5, rely=0.7, anchor=CENTER)

root.mainloop()