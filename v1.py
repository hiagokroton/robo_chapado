import pandas as pd
import datetime
import time
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys






######################################### FUNCOES
# funcao para verificar se a pagina carregou
def verificar_pagina(elem_x):
    result = None
    while result is None:
        try:
            driver.find_element_by_xpath(elem_x)
            result = 1
        except:
            pass









###################################### VARIAVEIS GERAIS
# PREENCHER COM OS DADOS DE QUEM FOR ABRIR CHAMADO
matricula_solicitante = '00000000'
nome_solicitante = 'Thais'
email_solicitante = 'thais@kroton.com.br'
telefone_solicitante = '35464587'
texto_padrao = 'ajshdkajshdsjkahdakjshdajksdhakjshdkj aksdhaksjdh {} ajksdhaskjdha akjsdha kjhsd jka'







##################################### CARREGAR ATE A TELA DE CHAMADO
# carregar chrome
driver = webdriver.Chrome(executable_path=r"chromedriver.exe")

# esperar 
driver.implicitly_wait(1)


# maximizar
driver.maximize_window()

# carregar tela inicial
pagina = "http://suporte.kroton.com.br/qualitor/logoutSolicitante.php"
driver.get(pagina)

# clicar em relogar e esperar carregar
elem_relogar = '/html/body/table/tbody/tr/td/font'
element = driver.find_element_by_xpath(elem_relogar).click()

elem_cks_corporativo = '//*[@id="divContainer"]/fieldset/table[2]/tbody/tr[2]/td/table[1]/tbody/tr[1]/td[1]/table/tbody/tr/td[2]/font/font/span'
verificar_pagina(elem_cks_corporativo)

# clicar em CKS - Corporativo e esperar carregar
element = driver.find_element_by_xpath(elem_cks_corporativo).click()

elem_cks_administracao = '//*[@id="btnActionService2429"]'
verificar_pagina(elem_cks_administracao)

# clicar em CKS - Administração Pessoal e esperar carregar
element = driver.find_element_by_xpath(elem_cks_administracao).click()

elem_admissao = '//*[@id="btnActionService3422"]'
verificar_pagina(elem_admissao)

# clicar em Admissao e esperar carregar
element = driver.find_element_by_xpath(elem_admissao).click()

elem_inclusao = '//*[@id="dsalternativa" and @value="Inclusão de Admissão Kroton<SEPARA_INFO>118170<SEPARA_INFO>142988<SEPARA_INFO>"]'
verificar_pagina(elem_inclusao)





























########################################### VARIAVEIS PLANILHA
nome = 'teste'
cpf = '405.758.291-20'


# clicar em Inclusão de Admissão korotn
element = driver.find_element_by_xpath(elem_inclusao).click()

# clicar em Avançar e esperar carregar
elem_avancar = '//*[@id="btnScriptNext"]'
element = driver.find_element_by_xpath(elem_avancar).click()

verificar_pagina('//*[@id="vlinformation6"]/option[74]')
verificar_pagina('//*[@id="vlinformation963"]/option[3]')

########################################### ABRIR CHAMADO
# selecionar unidade
elem_mantenedora = '//*[@id="vlinformation6"]/option[74]'
element = driver.find_element_by_xpath(elem_mantenedora).click() # selecionar opcao

# colocar matricula
elem_matricula = '//*[@id="vlinformation341"]'
element = driver.find_element_by_xpath(elem_matricula).send_keys(matricula_solicitante)

# colocar nome
elem_nome = '//*[@id="vlinformation247"]'
element = driver.find_element_by_xpath(elem_nome).send_keys(nome)

# colocar cpf
elem_cpf = '//*[@id="vlinformation246"]'
element = driver.find_element_by_xpath(elem_cpf).send_keys(cpf)

# colocar nome solicitante
elem_nome_solicitante = '//*[@id="vlinformation33"]'
element = driver.find_element_by_xpath(elem_nome_solicitante).send_keys(nome_solicitante)

# colocar email solicitante
elem_email_solicitante = '//*[@id="vlinformation34"]'
element = driver.find_element_by_xpath(elem_email_solicitante).send_keys(email_solicitante)

# colocar telefone solicitante
elem_telefone_solicitante = '//*[@id="vlinformation35"]'
element = driver.find_element_by_xpath(elem_telefone_solicitante).send_keys(telefone_solicitante)

# selecionar tipo de admissao
elem_autonomo = '//*[@id="vlinformation963"]/option[3]'
element = driver.find_element_by_xpath(elem_autonomo).click()

# preencher data de admissao
d = datetime.date.today() # data de hoje
mes = str(d.month)

if len(mes) == 1:
    mes = '0' + mes

data = '01' + '/' + mes + '/2018'

elem_data_admissao = '//*[@id="vlinformation353"]'
element = driver.find_element_by_xpath(elem_data_admissao).send_keys(data)

# clicar em avancar e esperar carregar
element = driver.find_element_by_xpath(elem_avancar).click()

elem_texto_solicitacao = '//*[@id="dsalternativa"]'
verificar_pagina(elem_texto_solicitacao)

# enviar texto padrao
result = None
while result is None:
    try:
        element = driver.find_element_by_xpath(elem_texto_solicitacao).send_keys(texto_padrao.format(nome))
        result = 1
    except:
        pass

# clicar em avancar e esperar carregar
element = driver.find_element_by_xpath(elem_avancar).click()

elem_anexo = '//*[@id="divParentgridAnexos"]'
verificar_pagina(elem_anexo)





########################### fazer loop para anexo
anexos = os.listdir('./'+ nome)

# elementos do anexo
elem_upload = '//*[@id="btnAnexoChamado"]'


def clicar_ok():
    elem_ok = '//*[@id="btnOK"]'
    result = None
    while result is None:
        try:
            element = driver.find_element_by_xpath(elem_ok).click()
            result = 1
        except:
            pass

def clicar_avancar():
    result = None
    while result is None:
        try:
            element = driver.find_element_by_xpath(elem_avancar).click()
            result = 1
        except:
            pass


for i in range(len(anexos)):
    # ler arquivo
    arquivo = anexos[i]

    # tentar carregar arquivo
    time.sleep(2)
    result = None
    while result is None:
        try:
            driver.find_element_by_xpath('//*[@id="fileAnexoChamado[]"]').send_keys(os.getcwd()+'/teste/' + arquivo)
            result = 1
        except:
            pass

    time.sleep(1)
    # clicar em upload
    result = None
    while result is None:
        try:
            element = driver.find_element_by_xpath(elem_upload).click()
            result = 1
        except:
            pass

    time.sleep(1)
    # clicar em OK
    clicar_ok()

    time.sleep(1)
    # clicar em avançar
    clicar_avancar()


if len(anexos) == 10:
    pass
else:
    for j in range(10-len(anexos)):
        # tentar carregar arquivo
        time.sleep(2)
        result = None
        while result is None:
            try:
                driver.find_element_by_xpath('//*[@id="fileAnexoChamado[]"]').send_keys(os.getcwd()+'/Capturar.PNG')
                result = 1
            except:
                pass

        time.sleep(1)
        # clicar em upload
        result = None
        while result is None:
            try:
                element = driver.find_element_by_xpath(elem_upload).click()
                result = 1
            except:
                pass

        time.sleep(1)
        # clicar em OK
        clicar_ok()

        time.sleep(1)
        # clicar em avançar
        clicar_avancar()


# salvar numero do chamado
elem_chamado = '//*[@id="divAlert"]/div/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]'
chamado_num = driver.find_element_by_xpath(elem_chamado).text

# tratar
chamado_num_tratado = chamado_num.replace('Serviço número ', '')
chamado_num_tratado = chamado_num_tratado.replace(' registrado.\n\nDeseja exibi-lo?', '')
chamado_num_tratado = int(chamado_num_tratado)

# clicar para nao exibir chamado
elem_nao = '//*[@id="btnNO"]'
result = None
while result is None:
        try:
            element = driver.find_element_by_xpath(elem_nao).click()
            result = 1
        except:
            pass

