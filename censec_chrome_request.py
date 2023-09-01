from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchWindowException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
from openpyxl import Workbook
from webdriver_manager.chrome import ChromeDriverManager
import os

class censec:
    def __init__(self) -> None:
        pass

    def action_click(self, ac_driver, ac_element):
        actions = ActionChains(ac_driver)
        actions.move_to_element(ac_element)
        actions.click(ac_element)
        actions.perform()
        print("Usou actionschains \n")

    def carregou_pagina(self, driver_navegador, by_content, by=By.ID, timeout = 60, to_sleep = 0, loop = True):
        fica_no_loop = True
        while (fica_no_loop):
            fica_no_loop = loop
            try:
                #element_present = EC.element_to_be_clickable(by_content)
                wait = WebDriverWait(driver_navegador, timeout, poll_frequency=1)
                element = wait.until(EC.presence_of_element_located((by, by_content)))
                fica_no_loop = False
            except TimeoutException:
                print("A página ainda não carregou")
                if not fica_no_loop: 
                    return False
            except NoSuchWindowException:
                sleep(1)
                print("Elemento esperado ainda não foi localizado")
                #counter_no_such_window += 1
                #if counter_no_such_window >= timeout:
                #    return False
            if (to_sleep > 0):
                sleep(to_sleep)
            if not fica_no_loop:
                sleep(1)
                return True

    def scrap(self, lista, nome_lista = ''):
        alvos = lista
        url = 'https://censec.org.br/private/s/00000000-0000-0000-0000-000000000000/cep'
        

        lista_campos = ['Nome pesquisado', 'CPF / CNPJ / OAB', 'Identidade', 'Cartorio', 'Município', 'UF', 'CNS', 'Livro', 'Complemento Livro', 'Folha', 'Complemento Folha', 
                'Valor', 'Data do Ato', 'Tipo do Ato', 'Natureza do Ato',  'Fonte']
        
        userprofile = os.environ['USERPROFILE']
        dirname = userprofile + r'\Downloads\Notarius'
        if not os.path.exists(dirname):
            os.makedirs(dirname)

        # Prepara o Excel
        wb = Workbook() # Cria a pasta de trabalho
        ws = wb.active # Pega a planilha ativa
        ws.append(lista_campos) # Junta cabeçalho
        
        chromeOptions = webdriver.ChromeOptions()
        chromeOptions.add_extension('web_pki.zip')

        # driver = webdriver.Chrome(ChromeDriverManager(path = dirname).install(), options=chromeOptions) Esta linha está desabilitada devido ao problema de atualização.
        # O código novo que contornou o problema segue abaixo:
        
        ###########################################################################################################
        #import urllib.request
        from selenium.webdriver.chrome.service import Service
        service = Service(path = dirname)
        
        #try:
        #    service = Service(ChromeDriverManager(path = dirname).install())
        #except ValueError:
        #    latest_chromedriver_version_url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
        #    latest_chromedriver_version = urllib.request.urlopen(latest_chromedriver_version_url).read().decode('utf-8')
        #    service = Service(ChromeDriverManager(path = dirname, version = latest_chromedriver_version).install())

        driver = webdriver.Chrome(service=service, options=chromeOptions) 
        ###########################################################################################################
        # fim do código novo
        
        # Acessa o site
        driver.get(url)
        driver.implicitly_wait(10)  # in seconds
        
        # Clica no campo de CPF e preenche o CPF conforme a lista
        xpath_campo_cpf = '/html/body/app-root/spa-root/div/app-layout/mat-sidenav-container/mat-sidenav-content/app-public-query/spa-box/div/form/div[2]/mat-form-field[1]/div/div[1]/div/input'
        self.carregou_pagina(self, driver, xpath_campo_cpf, by=By.XPATH)
        campo_cpf = driver.find_element(By.XPATH, xpath_campo_cpf)
        
        # Preenche o CPF:
        try: 
            campo_cpf.clear()
        except: 
            self.carregou_pagina(self, driver, xpath_campo_cpf, by=By.XPATH)
            campo_cpf = driver.find_element(By.XPATH, self, xpath_campo_cpf)
            campo_cpf.clear()
                
        sleep(0.5)
        campo_cpf.send_keys(alvos) 
        sleep(1)

        # Clica no botão Buscar para pesquisar o CPF
        botao_buscar = '/html/body/app-root/spa-root/div/app-layout/mat-sidenav-container/mat-sidenav-content/app-public-query/spa-box/div/form/div[5]/spa-submit-button/button'
        self.carregou_pagina(self, driver, botao_buscar, by=By.XPATH)
        try:
            driver.find_element(By.XPATH, botao_buscar).click()
        except:
            sleep(4)
            self.carregou_pagina(self, driver, botao_buscar, by=By.XPATH)
            botao_buscar2 = driver.find_element(By.XPATH, '/html/body/app-root/spa-root/div/app-layout/mat-sidenav-container/mat-sidenav-content/app-public-query/spa-box/div/form/div[5]/spa-submit-button/button')
            self.action_click(self, driver, botao_buscar2)
        sleep(3)
        
        all_cookies=driver.get_cookies();
        
        return all_cookies
    

    


