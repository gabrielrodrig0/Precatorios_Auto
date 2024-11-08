from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
import pandas as pd
import time
import pyautogui
from webdriver_manager.chrome import ChromeDriverManager
import os
from base64 import b64decode
import shutil

# Caminho para a pasta 'pdfs'
pdfs_folder = os.path.join(os.getcwd(), "pdfs")

temp_folder = os.path.join(os.getcwd(), "temp")

# Verifica se a pasta 'pdfs' existe. Se existir, remove e recria.
if os.path.exists(pdfs_folder):
    shutil.rmtree(pdfs_folder)  # Remove a pasta e todo o seu conteúdo
    print("Pasta 'pdfs' antiga removida com sucesso.")
    
if os.path.exists(temp_folder):
    shutil.rmtree(temp_folder)
    print("Pasta 'pdfs' antiga removida com sucesso.")

# Cria a pastas
os.makedirs(pdfs_folder)
print("Nova pasta 'pdfs' criada com sucesso.")
os.makedirs(temp_folder)
print("Nova pasta 'temp' criada com sucesso.")

# Leitura de arquivo XLS
nomeArq = input('Digite o nome do arquivo "xls" a ser processado: ')
print(nomeArq)
df = pd.read_excel(nomeArq)

# Lista para armazenar os nomes dos registros processados
registros_boleto = []

# Iniciando o ChromeDriver
service = ChromeService(ChromeDriverManager().install())
browser = webdriver.Chrome(service=service)

browser.get('https://www63.bb.com.br/portalbb/djo/id/IdDeposito,802,4647,4648,0,1.bbx?pk_vid=a8c030459b6e546b1624899559a5e218&pk_vid=a8c030459b6e546b1624899559a5e218&pk_vid=a8c030459b6e546b1624900138a5e218&#8203;&pk_vid=c232b04747645307166878141098691e&pk_vid=c232b04747645307166878141098691e&pk_vid=c232b04747645307166878141098691e')

if 'HTTP_PROXY' in os.environ:
    del os.environ['HTTP_PROXY']
if 'HTTPS_PROXY' in os.environ:
    del os.environ['HTTPS_PROXY']

time.sleep(1.5)

# DIV PRECATORIOS JUDICIAIS	PRECATORIO	2021.06875-1	 R$ 12.345,00 	 RÉU 	Dummy	FÍSICA	MUNICIPIO DO RIO DE JANEIRO	JURÍDICA	15.744.077/0001-43

counter = 1

# REPETE ENQUANTO HOUVER REGISTRO P/ PROCESSAR
for registro, Orgao_de_Justica in enumerate(df["Orgao_de_ Justica"]): 
    # TIPO DE JUSTIÇA (FIXO - ESTADUAL)
    search_input_radio_button = browser.find_element(By.XPATH, '//*[@id="formularioIdDeposito:justica:0"]')
    search_input_radio_button.click()
    time.sleep(0.5)

    # PRÉ CADASTRAMENTO (FIXO - PRIMEIRO DEPÓSITO)
    search_input_escolha = browser.find_element(By.XPATH, '//*[@id="formularioIdDeposito:optTipoCadastramento"]')
    search_input_escolha.send_keys('PRIMEIRO DEPÓSITO')
    time.sleep(0.5)

    # CONTINUAR
    search_input_continuar = browser.find_element(By.XPATH, '//*[@id="formularioIdDeposito:btnContinuar"]')
    search_input_continuar.click()
    time.sleep(0.5)

    # UNIDADE DA FEDERAÇÃO (FIXO - RIO DE JANEIRO) 
    search_input_UF = browser.find_element(By.XPATH, '//*[@id="formulario:comboUf"]')
    search_input_UF.send_keys('RJ - RIO DE JANEIRO')
    time.sleep(0.5)

    # TRIBUNAL (FIXO - TRIBUNAL DE JUSTIÇA)
    search_input_tribunal = browser.find_element(By.XPATH, '//*[@id="formulario:cmbTribunal"]')
    search_input_tribunal.send_keys('TRIBUNAL DE JUSTICA')
    time.sleep(0.2)

    # COMARCA (FIXO - RIO DE JANEIRO)
    search_input_comarca = browser.find_element(By.XPATH, '//*[@id="formulario:cmbComarca"]')
    search_input_comarca.send_keys('RIO DE JANEIRO')
    time.sleep(0.2)

    # CONTINUAR
    search_input_continuar = browser.find_element(By.XPATH, '//*[@id="formulario:botaoContinuar"]')
    search_input_continuar.click()
    time.sleep(0.5)

    # ÓRGÃO DE JUSTIÇA (FIXO - DIV PRECATÓRIOS JUDICIAIS)
    search_input_orgao_de_justica = browser.find_element(By.XPATH, '//*[@id="formulario:cmbOrgaos"]')
    search_input_orgao_de_justica.send_keys('DIV PRECATORIOS JUDICIAIS')
    time.sleep(0.5)

    # NATUREZA DA AÇÃO (FIXO - PRECATÓRIO)
    search_input_natureza_da_acao = browser.find_element(By.XPATH, '//*[@id="formulario:cmbAcoes"]')
    search_input_natureza_da_acao.send_keys('PRECATORIO')
    time.sleep(0.5)

    # NÚMERO DO PROCESSO JUDICIAL (VARIÁVEL)
    search_input_numero_processo = browser.find_element(By.XPATH, '//*[@id="formulario:numeroProcesso"]')
    search_input_numero_processo.send_keys(df.loc[registro, "Numero_do_Processo_Judicial"])
    time.sleep(0.5)

    # VALOR DO DEPÓSITO JUDICIAL (VARIÁVEL)
    search_input_valor_do_deposito_judicial = browser.find_element(By.XPATH, '//*[@id="formulario:valorDeposito"]')
    valor = df.loc[registro, "Valor_do_deposito_ judicial"]
    ss = f'{valor:.2f}'
    search_input_valor_do_deposito_judicial.send_keys(ss)
    time.sleep(0.15)

    # DEPOSITANTE (FIXO - Réu)
    search_input_depositante = browser.find_element(By.XPATH, '//*[@id="formulario:cmbTipoDepositante"]')
    search_input_depositante.send_keys('Réu')
    time.sleep(0.1)

    # NOME AUTOR (VARIÁVEL)
    search_input_nome_autor = browser.find_element(By.XPATH, '//*[@id="formulario:nomeAutor"]')
    search_input_nome_autor.send_keys(df.loc[registro, "Autor"])
    time.sleep(0.2)

    # TIPO DE PESSOA AUTOR (VARIÁVEL)
    search_input_tipo_de_pessoa_autor = browser.find_element(By.XPATH, '//*[@id="formulario:cmbTipoPessoa"]')
    time.sleep(0.2) 

    pessoaAutor = df.loc[registro, "TIPO_DE_PESSOA_AUTOR"]
    if pessoaAutor == 'Física':
        search_input_tipo_de_pessoa_autor.send_keys('Física')
    else:
        search_input_tipo_de_pessoa_autor.send_keys('Jurídica')
    time.sleep(0.5)

    # NOME RÉU (VARIÁVEL)
    search_input_nome_reu = browser.find_element(By.XPATH, '//*[@id="formulario:nomeReu"]')
    search_input_nome_reu.send_keys(df.loc[registro, "REU"])
    time.sleep(0.15)

    # TIPO DE PESSOA RÉU (FIXO - Jurídica)
    search_input_tipo_de_pessoa_reu = browser.find_element(By.XPATH, '//*[@id="formulario:cmbTipoPessoaReu"]')
    search_input_tipo_de_pessoa_reu.send_keys('Jurídica')
    time.sleep(0.2) 

    # CPF RÉU (VARIÁVEL)
    search_input_cnpj_reu = browser.find_element(By.XPATH, '//*[@id="formulario:cnpjReu"]')
    search_input_cnpj_reu.send_keys(df.loc[registro, "CNPJ"])
    time.sleep(0.2)

    # GERAR ID
    search_input_gerar_id = browser.find_element(By.XPATH, '//*[@id="formulario:btnGerarID"]')
    search_input_gerar_id.click()
    time.sleep(0.5)

    # GERAR BOLETO
    search_input_gerar_boleto = browser.find_element(By.XPATH, '//*[@id="botaoImprimirBoleto"]')
    search_input_gerar_boleto.click()
    time.sleep(0.5)

    # SALVAR A URL DA PÁGINA CORRENTE
    guarda_url = browser.current_url

    # CONFIRMAR
    search_input_confirmar = browser.find_element(By.XPATH, '//*[@id="formulario:confirmar"]')
    search_input_confirmar.click()

    time.sleep(5)
    
    all_windows = browser.window_handles
    
    # Captcha Escape
    if counter == 1:
        # Trocar para a segunda janela aberta
        browser.switch_to.window(all_windows[1])
        
        # Realizando o download do arquivo
        pyautogui.hotkey('ctrl', 's')
        time.sleep(2)
        
        autor = df.loc[registro, "Autor"]
        numero_do_processo_judicial = df.loc[registro, "Numero_do_Processo_Judicial"]
        pyautogui.typewrite(f"{temp_folder}\\{numero_do_processo_judicial} - {autor}.pdf")
        
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(8)
        shutil.rmtree(temp_folder)
               
        time.sleep(1)
        counter += 1
        # Fechar a segunda janela
        browser.close()
        
        # Voltar para a primeira janela
        browser.switch_to.window(all_windows[0])
        
        # Navegar para a página principal novamente
        browser.get('https://www63.bb.com.br/portalbb/djo/id/IdDeposito,802,4647,4648,0,1.bbx?pk_vid=a8c030459b6e546b1624899559a5e218&pk_vid=a8c030459b6e546b1624899559a5e218&pk_vid=a8c030459b6e546b1624900138a5e218&#8203;&pk_vid=c232b04747645307166878141098691e&pk_vid=c232b04747645307166878141098691e&pk_vid=c232b04747645307166878141098691e')
        time.sleep(1)
        
    else:
        # Para outras janelas, realizar o download diretamente
        if len(all_windows) > 1:
            # Trocar para a nova janela
            browser.switch_to.window(all_windows[1])
            time.sleep(10)
            # Fechar a nova janela
            autor = df.loc[registro, "Autor"]
            numero_do_processo_judicial = df.loc[registro, "Numero_do_Processo_Judicial"]
            path = os.path.join(pdfs_folder, f"{numero_do_processo_judicial} - {autor}.pdf")
            
            boleto = browser.find_element(By.XPATH, '/html/body/div/div[3]/div/div/div[1]/apw-ng-app/app-template/bb-layout/div/div/div/div/div/bb-layout-column/app-breadcrump/div/div[2]/div/app-impressao-boleto/div/div/div[2]/div/div[1]/section/bb-card[2]/bb-card-body/div/div[1]/div[1]/bb-text-chip/div/div[2]/span[2]')
            boleto = boleto.text
            
            registros_boleto.append({
                "Orgao_de_Justica": df.loc[registro, "Orgao_de_ Justica"],
                "Natureza_da_Acao": df.loc[registro, "Natureza_da_Acao"],
                "Numero_do_Processo_Judicial": df.loc[registro, "Numero_do_Processo_Judicial"],
                "Valor_do_deposito_judicial": df.loc[registro, "Valor_do_deposito_ judicial"],
                "Depositante": df.loc[registro, "Depositante"],
                "Autor": df.loc[registro, "Autor"],
                "TIPO_DE_PESSOA_AUTOR": df.loc[registro, "TIPO_DE_PESSOA_AUTOR"],
                "REU": df.loc[registro, "REU"],
                "TIPO_DE_PESSOA_REU": df.loc[registro, "TIPO_DE_PESSOA_REU"],
                "CNPJ": df.loc[registro, "CNPJ"],
                "Boleto": boleto
            })
            
            print(registros_boleto)
        
            elemento = browser.find_element(By.XPATH, '/html/body/div/div[3]/div/div/div[1]/apw-ng-app/app-template/bb-layout/div/div/div/div/div/bb-layout-column/app-breadcrump/div/div[2]/div/app-impressao-boleto/div/div/div[2]/div/div[2]/section/bb-pdf-viewer/bb-card/bb-card-header/bb-card-header-action/a[2]')
            base64_data = elemento.get_attribute("href")
            base64_data = base64_data.replace("data:application/pdf;base64,", "")
            with open(path, "wb") as file:
                file.write(b64decode(base64_data))
            
            browser.close()
            # Voltar para a aba principal
            browser.switch_to.window(all_windows[0])
            print("Voltamos para a janela principal!")
            
        browser.get('https://www63.bb.com.br/portalbb/djo/id/IdDeposito,802,4647,4648,0,1.bbx?pk_vid=a8c030459b6e546b1624899559a5e218&pk_vid=a8c030459b6e546b1624899559a5e218&pk_vid=a8c030459b6e546b1624900138a5e218&#8203;&pk_vid=c232b04747645307166878141098691e&pk_vid=c232b04747645307166878141098691e&pk_vid=c232b04747645307166878141098691e')
        counter += 1
        time.sleep(1)

# Fechar o navegador ao final do processo

browser.quit()

df_boleto = pd.DataFrame(registros_boleto)

# Gerar o arquivo Excel com os registros
output_file = 'registros_boleto.xlsx'
df_boleto.to_excel(output_file, index=False)

print(f'Arquivo {output_file} gerado com sucesso!')