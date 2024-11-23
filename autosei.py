
########################################################################
# busca todos os arquivos período explícito sem especificar a data
#########################################################################

import os
import csv
import re
import pandas as pd
import logging
import sys

from collections import OrderedDict
from datetime import date
from pathlib import Path

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException


MODULE_DIR = Path(__file__).parent
PASSWORD_FILE = Path.home().joinpath(".password.txt")


logging.basicConfig(level=logging.INFO)

def build_path(subfolder):
    current_folder = os.getcwd()
    folderpath = os.path.join(current_folder, subfolder)
    folderpath = os.path.abspath(folderpath)
    if not os.path.exists(folderpath): os.makedirs(folderpath)
    return folderpath

def busca_documentos(
    # TO DO: entrar com o endereço SEI 
    url, 
    # TO DO: selecionar o tipo
    #separa a lista de documentos
    doc_type, #tipo de documento
    start_date,
    end_date,
    output_dir="extraidos",
    charset="iso-8859-1",
    passwordfile=PASSWORD_FILE
):
    # nomes_arquivos = build_path('nomes_arquivos')
    _output_dir = build_path(output_dir)

    driver = webdriver.Chrome()
    driver.implicitly_wait(0.5)
    driver.get(url)
    

    # login page
    username_fld = driver.find_element("xpath", '//*[@id="txtUsuario"]')
    password_fld = driver.find_element("xpath", '//*[@id="pwdSenha"]')
    submit_button = driver.find_element("xpath", '//*[@id="sbmLogin"]')

    if not os.path.exists(passwordfile):
        message = "O arquivo de autenticação não existe. Colocar em .password/password.txt"
        logging.exception(message)
        raise Exception(message)

    with open(passwordfile) as f:
        _username, _password = f.read().strip().split(":", maxsplit=1)
        username_fld.send_keys(_username)
        password_fld.send_keys(_password)
        
    submit_button.click()
    del _password
    logging.info(f"Autenticando como {_username}")

    # home page (Chamar o item de menu de pesquisa)
    searching = driver.find_element("xpath", '//*[@id="main-menu"]/li[5]/a')
    driver.implicitly_wait(0.5)
    searching.click()

    # search page - Selecionando os widgets a serem preenchidos
    logging.info(f"Preenchendo o formulário de pesquisa")
    
    # Combo tipo de documento
    # TODO: Transformar esta opçao em argumento
    
    def documento_type():
        l_documentos = (driver.find_element("xpath", '//*[@id="divSeriePesquisa"]').text())
        selecao = input('s_doc')
        for documento in l_documentos: 
            print(documento)
            if documento.text == selecao:
                return doc_type 

        return doc_type
        
    
    driver.find_element("xpath", '//*[@id="selSeriePesquisa"]').send_keys(doc_type)
    driver.implicitly_wait(0.5)
    
    # Radio Data Documento - Período Explícito
    # TODO: Adicionar período de pesquisa como parâmetro
    driver.find_element("xpath", '//*[@id="optPeriodoExplicito"]').click()
    driver.implicitly_wait(10)
    
    
    
    # Wait
    #driver.implicitly_wait(0.5)
    
    # Manipulação de Período
    # if period:
    #     start_date, end_date = period
    
    #     if start_date:
    #         if not end_date:
    #             logging.warn("Não foi passado end_date: end_date será considerado a data de hoje")
    #             end_date = date.today()        
    #         else:
    #             if start_date >= end_date:
    #                 message = "A data de início e fim da pesquisa não pode ser menor ou igual"
    #                 logging.error(message)
    #                 raise Exception(message)            
    #     else:
    #         logging.warn("Não foi passado start_date: ignorando end_date, caso informado")
    #         end_date = None
        
    #     date_mask = r"%d/%m/%Y"        
        
    driver.find_element("xpath", '//*[@id="txtDataInicio"]').send_keys(start_date) #.strftime(date_mask)
    driver.implicitly_wait(5)
    driver.find_element("xpath", '//*[@id="txtDataFim"]').send_keys(end_date) #.strftime(date_mask)
    driver.implicitly_wait(50)
    

    # Buscar
    driver.find_element("xpath", '//*[@id="sbmPesquisar"]').click()

    # TODO: capturar na página de pesquisa a quantidade de documentos achados
    
    # getting files
    def get_files():       
        list_documents = []
        original_window = driver.current_window_handle 
        
        try:     
            page_docs_search = driver.find_element("xpath", '//*[@id="conteudo"]')
                        
            wait = WebDriverWait(driver, timeout=20)
            wait.until(lambda d : page_docs_search.is_displayed())
            
        except NoSuchElementException:
            logging.warn("Não tem resultados de pesquisa")
            sys.exit(-1)
           
        elements = page_docs_search.find_elements("xpath", 'table/tbody/tr[1]')
        
        for element in elements:
            pr_elemento = element.find_element("xpath", 'td[1]')
            
            # pega dados de cada documento na table de pesquisa
            processo = re.search('\d{5}\.\d{6}\/\d{4}\-\d{2}', pr_elemento.text).group()            
            documento = int(element.find_element("xpath", 'td[2]').text)
            url = pr_elemento.find_element("xpath", 'a[2]').get_attribute('href')
            
            # chama os links e salva os conteúdos em HTML
            driver.switch_to.new_window('tab')
            driver.get(url)
            
            out_html = build_path(os.path.join(_output_dir, "html"))
            html_extracted = os.path.join(out_html, f'documento_{documento}.html')
            
            # TODO: Procurar uma forma do selenium entregar o charset da página
            with open(html_extracted, 'w', encoding=charset) as file:
                file.write(driver.page_source)
            
            list_documents.append(
                {'processo': processo, 'documento': documento, 'url': url, "extraido": html_extracted, "charset": charset}
            )
            
            driver.close()
            driver.switch_to.window(original_window)
        
        return list_documents

    # pagination
    documentos = get_files()        
    
    while True:
        try:
            paginas_tag = driver.find_element(by=By.CLASS_NAME, value="paginas")
            
            proxima_pagina = paginas_tag.find_element("xpath", "span[last()]/a[@href]") 
            if not proxima_pagina.text.lower().strip().startswith("p"):
                logging.info(f"Paginação concluída.")
                break
            
            link = proxima_pagina.get_attribute('href')
            proxima_pagina.click()
            logging.info(f"Evento de click disparado: {link}") 
            
            documentos = documentos + get_files()
            
        except NoSuchElementException:
            logging.warn("A paginação acabou")
            break
        
        except StaleElementReferenceException:
            pass
            
        except Exception as e:
            logging.exception(f"Erro não esperado: {e}")
            sys.exit(-1)        

    driver.close()
    driver.quit()
    
    out_csv = os.path.join(_output_dir, "processos.csv")
    pd.DataFrame(documentos).to_csv(out_csv, index=False)
    
    return out_csv
    

############################################################# 
# para ler todos os arquivos em html e criar um DataFrame 
############################################################# 
def parse_csv_results(csvfile, output_csv = "resultado.csv"):
    with open(csvfile) as _csvfile:
        lista_df=[] 
        
        reader = csv.DictReader(_csvfile)
        
        for row in reader:
            extraido = row["extraido"]
            charset = row["charset"]
            
            with open(extraido, encoding=charset) as f: 
                soup = BeautifulSoup(f.read(), "html.parser") 
        
            tags = [tag for tag in soup.find("div", id="conteudo").children if len(tag.text.strip()) > 0 and not re.match(r"^\d+\.", tag.text.strip())] 
            # Cada HTML, um dicionário ordenado 
            dict_series = OrderedDict() 
        
            for index in range(len(tags)): 
                tag = tags[index] 
                
                if tag.name == "b": 
                    key = tag.text.strip().rstrip(":") 
                    value = tags[index + 1].text.strip() 
                    
                    dict_series[key] = value         
        
            # Empilhar todos os dicionários para criar o df e interpretar os dtypes 
            df = ( 
                pd.DataFrame([dict_series]) 
                    .apply(lambda x: pd.to_numeric(x.str.replace(",", "."), errors="ignore")) 
                    .apply(lambda x: x.replace("Sim", True).replace("Não", False))         
            ) 
            lista_df.append(df)
    final_df = pd.concat(lista_df, ignore_index=True)
    final_df.to_csv(output_csv, index = False)   

    print(f"Arquivo CSV salvo como: {output_csv}")
    return final_df


def carregar_janela_principal():    
    import tkinter as tk
    from tkinter import ttk   
    from pathlib import Path
    from tkcalendar import DateEntry
    # from ttkthemes import ThemedTk
    from threading import Thread
    
    class CustomThread(Thread):    
        def __init__(self, group=None, target=None, name=None,
                    args=(), kwargs={}, Verbose=None):
            Thread.__init__(self, group, target, name, args, kwargs)
            self._return = None

        def run(self):
            if self._target is not None:
                self._return = self._target(*self._args, **self._kwargs)
                
        def join(self, *args):
            Thread.join(self, *args)
            return self._return
    
    FONT = "Arial"

    def abrir_segunda_janela():
        janela2=tk.Toplevel()
        janela2.title("Janela de Login")
        janela2.config(bg="lightblue")
        janela2.resizable(False, False)
        janela2.iconbitmap(MODULE_DIR.joinpath("logo-sei.ico"))
        janela2.geometry("300x200")
        
        
        label_login = tk.Label(janela2, text = "Login", bg="lightblue", font = (FONT, 12, "bold"))
        label_login.grid(row = 0, column = 0, padx=1, pady=3)
        entry_login = ttk.Entry(janela2, width=15)
        entry_login.grid(row=0, column= 1, padx=2, pady=3)
        
        label_password = tk.Label(janela2, text = "Password", bg="lightblue", font = (FONT, 12, "bold"))
        label_password.grid(row = 1, column = 0, padx=1, pady=4)
        entry_password = ttk.Entry(janela2, width= 15, show = "*")
        entry_password.grid(row=1, column= 1, padx=2, pady=4)
        
        def save_arq():
            with open(PASSWORD_FILE, "w") as arq:
                arq.write(f'{entry_login.get()}:{entry_password.get()}')
                #fechar a janela de login
            janela2.destroy()
                               
        botao_voltar = ttk.Button(janela2, text = 'Enter ', style="big.TButton", command = save_arq)
        botao_voltar.grid(row = 2, column = 0, columnspan=2)
    
    janela = tk.Tk()
    janela.title("Tirando dados do SEI")  
    
    # janela.geometry("900x600")
    janela.config(bg="lightblue")
    janela.iconbitmap(MODULE_DIR.joinpath("logo-sei.ico"))

    img = tk.PhotoImage(file = MODULE_DIR.joinpath("logo-sei.png"), width=300, height=200)
    logo = tk.Label(janela, image=img, background="lightblue")

    logo.grid(row=0, column= 0, padx=1, pady=3)
    
    mensagem1 = tk.Label(janela,
                        text="Busca de dados dos documentos SEI",
                        fg="gray",
                        #  image= img,
                        bg ="lightblue",
                        font=(FONT,"20"),
                        width=30,
                        height=5
    )

    mensagem1.grid(row=0, column= 1, columnspan=20, padx=1, pady=3)

    #criar o Entry
    label_url = ttk.Label(janela,
                        text="URL do sistema SEI *",
                        background="lightblue",
                        font=(FONT)
    )

    label_url.grid(row=2,
                column= 0,
                padx=10,
                pady=5,
                sticky="nsew"
    )

    entry_url = ttk.Entry(janela,
                        width=15
    )

    entry_url.grid(row=2,
                column=1,
                columnspan=20,
                padx=10,
                pady=5,
                sticky="nsew"
    )

    label_doc_type = ttk.Label(janela,
                            text="Qual é o tipo de documento que quer os dados?*",
                            background="lightblue",
                            font=(FONT)
    )
    label_doc_type.grid(row=3,
                        column=0,
                        padx=10,
                        pady=5,
                        sticky="nsew"
    )

    entry_doc_type = ttk.Entry(janela,
                            width=20
    )

    entry_doc_type.grid(row=3,
                        column=1,
                        columnspan=20,
                        padx=10,
                        pady=5,
                        sticky="nsew"
    )
    
    label_obrigacao = ttk.Label(janela,
                        text="* São obrigatórios",
                        background="lightblue",
                        font= (FONT, 14, "bold"),
                        foreground="red"
    )
    
    label_obrigacao.grid(row=4,
                        # columnspan=2,
                        padx=10,
                        pady=5
    )
    label_datainicial = ttk.Label(janela,
                        text="Data de inicio da pesquisa:", 
                        background="lightblue",
                        font=(FONT)
    )
    label_datainicial.grid(row=5,
                        column=0,
                        padx=10,
                        pady=5,
                        sticky="nsew")

            
    entry_datainicial = DateEntry(janela,
                                date_pattern="dd/mm/yyyy"                                          
                                )

    entry_datainicial.grid(row=5,
                        column=1
                        # padx=8,
                        # pady=5
    )
       

    label_datafinal = ttk.Label(janela,
                        text="Data final da pesquisa:",
                        background="lightblue", 
                        font=(FONT)
    )
    label_datafinal.grid(row=6,
                        column=0,
                        padx=10,
                        pady=5,
                        sticky="nsew"
    )

    entry_datafinal = DateEntry(janela,
                                date_pattern="dd/mm/yyyy" 
    )
    # def select_datef():
    #     datef = entry_datafinal.get_date()
    #     date_label.config(text = datef)
    
    entry_datafinal.grid(row=6,
                        column=1
    )

  
    botao_login = ttk.Button(janela,
                            text = "Login",
                           #  font=(FONT, "12", "bold"),
                            style="big.TButton",
                            command=abrir_segunda_janela
    )
    botao_login.grid(row=7,
                    columnspan=2,
                    padx=15,
                    pady=5
    )
    
    # manipulador de evento click
    def schedule_check(t):
        """
        Schedule the execution of the `check_if_done()` function after
        one second.
        """
        janela.after(1000, check_if_done, t)

    def check_if_done(t):
        # If the thread has finished, re-enable the butt
        # on and show a message.
        if not t.is_alive():
            print(t.join())      
            label_entrada.config(text="\n Arquivos Baixados em \n \n /extraidos/html \n \n processos.csv",
                                 background="lightblue",
                                 font=(FONT, 14, "bold")
                                 )
            botao_entrada.config(state="normal")
            janela.after(100,janela.destroy)
        else:
            # Otherwise check again after one second.
            schedule_check(t)
        
    
    def botao_entrada_click():
        # Add label for waiting
        label_entrada.config(text="Buscando dados no SEI...",
                             font=(FONT, 14, "bold")
                             )
        # Disable the button while downloading the file.
        botao_entrada.config(state="disabled")        
        
        date_mask = r"%d/%m/%Y" 
        
        # Start the crawler in a new thread.
        url = entry_url.get()
        doc_type = entry_doc_type.get()
        data_inicial = entry_datainicial.get_date().strftime(date_mask)
        data_final = entry_datafinal.get_date().strftime(date_mask)
        t = CustomThread(target=busca_documentos, args=(url, doc_type, data_inicial, data_final))
        t.start()
        # Start checking periodically if the thread has finished.
        schedule_check(t)
        

    botao_entrada = ttk.Button(
        janela,
        text="Salvar os dados da Pesquisa ",
        style="big.TButton",
        command=botao_entrada_click
    )
    
    botao_entrada.grid(
        row=10,
        columnspan=2,
        padx=10,
        pady=5
    )

    # mostrar os dados para rodar o Script do SEI
    label_entrada = ttk.Label(
        janela,
        text="",
        font=(FONT, "20")
    )

    label_entrada.grid(
        row=15,
        column=0,
        padx=10,
        pady=5
    )

    # Mostrar a Janela 
    janela.mainloop()
    
    
def main():
    import argparse
    
    run_in_commandline = '--gui' not in sys.argv
    
    parser = argparse.ArgumentParser(
        description="Programa para capturar dados do SEI"
    )
    
    # argumentos obrigatórios para caso executar em linha de comando (gui = False)
    parser.add_argument("--url", required=run_in_commandline, help="URL do SEI a ser pesquisado")
    parser.add_argument("--doc", required=run_in_commandline, help="Tipo do documento SEI")
    parser.add_argument("--salvar", required=run_in_commandline, help="Diretório para salvar os resultados", type=Path)
    parser.add_argument("--passwordfile", required=run_in_commandline, help="Arquivo com a senha do SEI, em ASCII", default=PASSWORD_FILE)
    
    # argumentos opcionais para linha de comando
    parser.add_argument("-data_inicial", "--data-inicio", type=date, help="Data de início da pesquisa", dest="data_inicial")
    parser.add_argument("-data_final", "--data-fim", help="Data de fim da pesquisa", dest="data_final")
    parser.add_argument("--charset", help="codificação de caracteres", default="iso-8859-1")
    
    # inicializar em MainWindow 
    parser.add_argument(
        "--gui", 
        help="Carregar em modo janela. Caso esta opção for fornecida, todas as demais serão ignoradas", 
        action="store_true", 
        dest="gui",
        required=False
    )
    
    # Início do tratamento dos argumentos
    args = parser.parse_args()
    
    if args.gui:
        carregar_janela_principal()
        return 0
    
    else:    
        docs = busca_documentos(
            url=args.url,
            doc_type=args.doc,
            start_date = args.data_inicial,
            end_date = args.data_final,
            output_dir=args.salvar,
            charset=args.charset,
            passwordfile=args.passwordfile
            )
        
        results = parse_csv_results(docs)
        
        print(results)
        
if __name__ == '__main__':
    main()


