from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchWindowException
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import os
from webdriver_manager.chrome import ChromeDriverManager

class signo:
    def __init__(self) -> None:
        pass

    def carregou_pagina(self, driver_navegador, by_content, by=By.ID, timeout = 90, to_sleep = 0, loop = True):
        fica_no_loop = True
        #counter_no_such_window = 0
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
            if (to_sleep > 0):
                sleep(to_sleep)
            if not fica_no_loop:
                sleep(1)
                return True

    def automatiza_senha(self):
        pass
    
    def busca_authorization(self, reqs): # Obtem o token da sessão autenticada
        for req in reqs:
            if 'https://backend.signo.org.br/api/ato/consultas-internas/listar-tipos-atos?content=&size=' in req.url:
                for x in req.headers._headers:
                    if x[0] == 'authorization':
                        return x[1]

    def scrap(self, lista, nome_lista = ''):
        #os.environ['WDM_LOCAL'] = '1'
        
        userprofile = os.environ['USERPROFILE']
        dirname = userprofile + r'\Downloads\Notarius'
        if not os.path.exists(dirname):
            os.makedirs(dirname)
        
        alvos = lista
        url = 'https://signo.org.br/#/login'
        
        lista_campos = ['Nome pesquisado', 'CPF / CNPJ / OAB', 'Identidade', 'Cartorio', 'Município', 'UF', 'CNS', 'Livro', 'Complemento Livro', 'Folha', 'Complemento Folha', 
                'Valor', 'Data do Ato', 'Tipo do Ato', 'Natureza do Ato',  'Fonte']
        
        chromeOptions = webdriver.ChromeOptions()
        chromeOptions.add_extension('web_pki.zip')
        driver = webdriver.Chrome(ChromeDriverManager(path = dirname).install(), options=chromeOptions)

        # Acessa o site
        driver.get(url)
        driver.implicitly_wait(10)  # in seconds
        self.carregou_pagina(driver, '/html/body/app-root/app-interno/app-sidebar/nav/div[2]/div[1]/i', by=By.XPATH)
        sleep(5)
        
        # Clica em Relatórios na seta pra baixo
        link_relatorio = '/html/body/app-root/app-interno/app-sidebar/nav/div[2]/ul/li[3]/div/i[2]'
        self.carregou_pagina(driver, link_relatorio, by=By.XPATH)
        driver.find_element(By.XPATH, link_relatorio).click()
        
        # Clica em Consultar atos
        cons_atos = '/html/body/app-root/app-interno/app-sidebar/nav/div[2]/ul/li[3]/app-submenu/ul/li/div/span'
        self.carregou_pagina(driver, cons_atos, by=By.XPATH)
        driver.find_element(By.XPATH, cons_atos).click()
        
        # Clica em Central
        link_central = '/html/body/app-root/app-interno/div/div/app-consultas-internas/app-fieldset[1]/div/div[2]/div[1]/div[1]/app-dropdown/div[1]/i'
        self.carregou_pagina(driver, link_central, by=By.XPATH)
        driver.find_element(By.XPATH, link_central).click()
        
        # Escolhe CEP
        botao_cep = '/html/body/app-root/app-interno/div/div/app-consultas-internas/app-fieldset[1]/div/div[2]/div[1]/div[1]/app-dropdown/div[2]/div[2]/div[1]/i'
        self.carregou_pagina(driver, botao_cep, by=By.XPATH)
        driver.find_element(By.XPATH, botao_cep).click()
        seta_botao_cep = '/html/body/app-root/app-interno/div/div/app-consultas-internas/app-fieldset[1]/div/div[2]/div[1]/div[1]/app-dropdown/div[1]/i[2]'
        self.carregou_pagina(driver, seta_botao_cep, by=By.XPATH)
        driver.find_element(By.XPATH, seta_botao_cep).click()
        
        # Clica em CPF e digita o primeiro CPF
        campo_cpf = '/html/body/app-root/app-interno/div/div/app-consultas-internas/app-fieldset[1]/div/div[2]/div[2]/div[1]/app-input/input'
        self.carregou_pagina(driver, campo_cpf, by=By.XPATH)
        driver.find_element(By.XPATH, campo_cpf).send_keys(alvos[0])
        
        # Clica em Pesquisar
        botao_pesquisar = '/html/body/app-root/app-interno/div/div/app-consultas-internas/app-fieldset[1]/div/div[2]/div[5]/div/app-button[1]/button'
        self.carregou_pagina(driver, botao_pesquisar, by=By.XPATH)
        driver.find_element(By.XPATH, botao_pesquisar).click()
        
        # Carrega Nenhum ato encontrado ou sleep 5 (preferencialmente esse último)
        sleep(5)
        
        
        requisicao = driver.requests
        auth = self.busca_authorization(requisicao)
        
        all_cookies=driver.get_cookies();
        driver.close()
        return auth
    


