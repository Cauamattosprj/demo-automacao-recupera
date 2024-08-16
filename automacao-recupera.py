from selenium import webdriver
import selenium.common.exceptions
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pyautogui as pa
from time import sleep
from tkinter import *
import tkinter
import customtkinter
import logging
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from datetime import timedelta
from datetime import datetime


def formatarData(event, entry):
    try:
        texto = entry.get()
        texto_formatado = ''.join(filter(str.isdigit, texto))

        if len(texto_formatado) >= 8:
            dia = texto_formatado[:2]
            mes = texto_formatado[2:4]
            ano = texto_formatado[4:8]
            data_formatada = f"{dia}/{mes}/{ano}"
            entry.delete(0, END)
            entry.insert(0, data_formatada)
    except ValueError:
        pass

# essa função serve para não permitir datas inválidas
def validarData(data_inicial, data_final):
    try: 
        data_inicial_datetime = datetime.strptime(data_inicial, "%d/%m/%Y")
        data_final_datetime = datetime.strptime(data_final, "%d/%m/%Y")
        data_atual_datetime = datetime.today()


        if datetime.strptime(data_inicial, "%d/%m/%Y") and datetime.strptime(data_final, "%d/%m/%Y"):
            if data_inicial_datetime > data_atual_datetime and data_final_datetime > data_atual_datetime:
                return False
            else:
                return True
        else:
            return False
    except ValueError:
        return False

def startAutomacao():
    
    
    
    
    try: 
        # pega os valores inseridos na GUI e armazena nas variáveis abaixo.
        data_pgto_inicial = input.get()
        data_pgto_final = input_2.get()
        
        # array com o código de todos os credores que serão buscados
        codigos_credores = ["3","4","5","2","6","8","10","11","12"]

        # SE a data for válida, inicia o robô. SENÃO, pede pra reinserir a data
        if validarData(data_pgto_inicial, data_pgto_final):
            # pega os valores inseridos na GUI e armazena nas variáveis abaixo.
            data_pgto_inicial_str = input.get()
            data_pgto_final_str = input_2.get()

            # caixa de texto na GUI com as datas inseridas anteriormente pelo usuário
            data_pgto_inicial = datetime.strptime(data_pgto_inicial_str, "%d/%m/%Y")
            data_pgto_final = datetime.strptime(data_pgto_final_str, "%d/%m/%Y")

            textbox.insert(INSERT, data_pgto_inicial.strftime("%d/%m/%Y"))
            textbox.insert(INSERT, " - ")
            textbox.insert(INSERT, f'{data_pgto_final.strftime("%d/%m/%Y")}' + "\n")

            # essa linha coleta logs para informar em caso de erros.
            logging.basicConfig(level=logging.ERROR)

            # INICIO ROBÔ SELENIUM
            service = Service(ChromeDriverManager().install())
            options = Options()
            options.add_experimental_option("detach", True)
            options.add_experimental_option('prefs', {
            'download.prompt_for_download': True,  # Mostrar o prompt de download
            })
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            # acessar a página do recupera
            driver.get("https://aegea.recupera.com.br/recupera")
            pagina_inicial = driver.current_window_handle
            sleep(5)
            # trocar o foco do selenium webdriver para o popup do recupera
            for foco in driver.window_handles:
                if foco != pagina_inicial:
                    pagina_recupera = foco
            driver.switch_to.window(pagina_recupera)

            # realiza o login
            driver.find_element('xpath', '//*[@id="txtOperador"]').send_keys('USUARIO')
            driver.find_element('xpath', '//*[@id="txtSenha"]').send_keys('SENHA')
            driver.find_element('xpath', '//*[@id="cmdEmpresa"]').click()
            driver.find_element('xpath', '//*[@id="dtgEmpresa_TextBox1_0"]').click()
            pa.press('f11')
            driver.find_element('xpath', '//*[@id="cmdOk"]').click()

            sleep(4)
            
            # acessa, preenche os campos e realiza o download dos relatórios. O processo se repete para cada código de credor, até que todos os códigos tenham sido baixados.
            for codigo in codigos_credores:
                if codigo == "3":
                    codigos_unidades = ["2"]
                if codigo == "4":
                    codigos_unidades = ["37", "36", "38", "39", "28", "27", "12", "15", "8", "18", "19", "20"]
                if codigo == "5":
                    codigos_unidades = ["32", "35", "42", "30", "9", "10", "11", "13", "14", "7", "29", "16", "17"]
                if codigo == "2":
                    codigos_unidades = ["3"]            
                if codigo == "6":
                    codigos_unidades = ["31", "44"]
                if codigo == "8":
                    codigos_unidades = ["51"]
                if codigo == "10":
                    codigos_unidades = ["41"]
                if codigo == "11":
                    codigos_unidades = ["49"]
                if codigo == "12":
                    codigos_unidades = ["45", "46", "48", "50"]



                for codigo_unidade in codigos_unidades:
                    data_pgto_inicial = datetime.strptime(data_pgto_inicial_str, "%d/%m/%Y")
                    while data_pgto_inicial <= data_pgto_final: 
                        dia_pgto_inicial_relatorio = data_pgto_inicial
                        dia_pgto_final_relatorio = data_pgto_inicial + timedelta(days=2)

                        driver.find_element('xpath', '//*[@id="tbCor_3_1"]/tbody/tr/td[1]').click()
                        driver.find_element('xpath', '//*[@id="dtgLista"]/tbody/tr[5]/td[1]/a').click()
                        driver.find_element('xpath', '//*[@id="cmdArquivoTexto"]').click()

                        # pyautogui assume para preencher campos que não foram possíveis preencher com selenium
                        pa.moveTo(751,225, 1.5)
                        pa.click(751,225, clicks=5, interval=1)
                        print('clicou')
                        # preenche datas
                        pa.write(dia_pgto_inicial_relatorio.strftime("%d/%m/%Y"))
                        pa.press('tab', presses=2)
                        pa.write(dia_pgto_final_relatorio.strftime("%d/%m/%Y"))

                        # - Preencher COD_CREDOR de acordo com o array de códigos
                        pa.press('tab', presses=2)
                        pa.write(codigo)
                        pa.press('tab', presses=2)
                        pa.write(codigo_unidade)
                        pa.press('tab', presses=2)
                        pa.write('cobrart')
                        
                        # Atualiza a data_pgto_inicial para a próxima iteração
                        data_pgto_inicial = dia_pgto_final_relatorio + timedelta(days=1)     


                        # deve ser editado por se tratar de um layout diferente
                        i = 1
                        while i <= 12:
                            pa.hotkey('shift', 'tab')
                            i += 1
                        pa.press('enter')

                        # função para verificar se há processos em andamento antes de realizar o download
                        driver.find_element('xpath', '//*[@id="cmdConsultarEx"]').click()
                        driver.find_element('xpath', '//*[@id="rdbAnteriores"]').click()
                        # verifica se o relatório está em execução ou já está finalizado
                        status_execucao = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="dtgVisualizar"]/tbody/tr[2]/td[5]'))).get_attribute("innerHTML")
                        while status_execucao == "Executando":
                            try:
                                print("IDENTIFICOU QUE HÁ PROCESSOS EM EXECUÇÃO E ENTROU NO LOOP")
                                sleep(5)
                                driver.find_element('xpath', '//*[@id="cmdAtualizar"]').click()
                                print("CLICOU PARA ATUALIZAR OS PROCESSOS EM EXECUÇÃO")
                                status_execucao = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="dtgVisualizar"]/tbody/tr[2]/td[5]'))).get_attribute("innerHTML")
                            except selenium.common.exceptions.TimeoutException:
                                break
                        # verifica se o relatório deu status de retorno 0 (êxito) ou 99 (falha)
                        status_retorno = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="dtgVisualizar"]/tbody/tr[2]/td[6]'))).get_attribute("innerHTML")
                        print(f"STATUS RETORNO: {status_retorno}")
                        if status_retorno == "0":
                            driver.find_element('xpath', '//*[@id="tbCor_3_1"]/tbody/tr/td[1]').click()
                            try:
                                driver.find_element('xpath', '//*[@id="mnuvid1"]').click()
                            except:
                                print("NÃO HOUVE NECESSIDADE DE CLICAR NO VIDRO")

                            # é realizado o download do arquivo
                            driver.find_element('xpath', '//*[@id="cmdArquivo"]').click()
                            driver.find_element('xpath', '//*[@id="dtgBuscar"]/tbody/tr[2]/td[1]/a').click()
                            driver.find_element('xpath', '//*[@id="cmdDownload"]').click()
                            sleep(7)

                            # formata o nome do arquivo para download (Cred2_Unidade32_01.05-03.05) 
                            pa.write(f"Cred{codigo}_Unidade{codigo_unidade}_{dia_pgto_inicial_relatorio.day}.{dia_pgto_inicial_relatorio.month}-{dia_pgto_final_relatorio.day}.{dia_pgto_final_relatorio.month}", interval=0.1)
                            pa.press('enter')

                            data_pgto_inicial = dia_pgto_final_relatorio + timedelta(days=1)
                        else:
                            data_pgto_inicial = dia_pgto_final_relatorio + timedelta(days=1)
                            continue
            # escreve na caixa de texto informando que todos os relatórios foram baixados para a data solicitada
            textbox.insert(INSERT, "(CONCLUIDO)\n")

        else:
            # pede para colocar data válida
            textbox.insert(INSERT, f"Data inválida: '{data_pgto_inicial} - {data_pgto_final}'.\nTente novamente\n")
    except:
        print(logging.exception("ERROR: "))

# GUI CUSTOMTKINTER
# iniciando o elemento raiz ctk
root = customtkinter.CTk()
# window config
root.title("Cobrart - Robô Relatórios Recupera")
root.iconbitmap("")
# essa formatação permite que a GUI sempre surja no meio da tela
height = 500
width = 400
y = (root.winfo_screenheight()//2-(width//2))
x = (root.winfo_screenwidth()//2-(height//2))
root.geometry('{}x{}+{}+{}'.format(width, height, x, y)) 
# adiciona o tema escuro
customtkinter.set_appearance_mode("light")
# cor dos componentes
verde_cobrart = "#0c4d02"

# ctk label
label = customtkinter.CTkLabel(root, text="Data inicial de pagamento: ")
label.pack()
# ctk input
data_pgto_inicial = tkinter.StringVar()
input = customtkinter.CTkEntry(root, width=300, height=40, textvariable=data_pgto_inicial)
input.pack(pady=(0, 20))
input.bind('<KeyRelease>', lambda event: formatarData(event, input))
# ctk label
label_2 = customtkinter.CTkLabel(root, text="Data final de pagamento: ")
label_2.pack()
# ctk input
data_pgto_final = tkinter.StringVar()
input_2 = customtkinter.CTkEntry(root, width=300, height=40, textvariable=data_pgto_final)
input_2.pack()
input_2.bind('<KeyRelease>', lambda event: formatarData(event, input_2))
# ctk button
button = customtkinter.CTkButton(root, text='Confirmar', command=startAutomacao, fg_color = verde_cobrart)
button.pack(padx=10, pady=10)
# caixa de texto com log do sistema
em_proceesso = customtkinter.CTkLabel(root, text="Histórico: ")
em_proceesso.pack(padx=10, pady=10)
textbox = customtkinter.CTkTextbox(root, width=250)
textbox.pack()
# loop para que a GUI não feche
root.mainloop()
