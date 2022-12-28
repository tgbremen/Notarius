import PySimpleGUI as sg
from time import sleep
from PySimpleGUI.PySimpleGUI import Output
import CensecRequest4 as CR
import SignoRequest3 as SR

instrucoes = """Instruções: \n
1. Os CPFs e CNPJs não são validados pelo programa. Se houver erro, não encontrará nada ou dará erro no programa. \n
2. Os CPFs e CNPJs podem ser escritos com ou sem pontos, hífens e barras. \n
3. Os zeros iniciais não podem ser omitidos ou o CPF / CNPJ não será encontrado. \n
4. Após clicar no botão Iniciar Extração, só é possível iniciar nova busca fechando o programa ou quando terminar a pesquisa. \n
5. Não use ou feche a janela do chrome durante a execução do programa, apenas a minimize. Não abra outra aba nessa janela. \n
6. O arquivo de saída será censec_nomedalista_datahora.xlsx ou signo_nomedalista_datahora.xlsx \n
7. O arquivo de saída será salvo na pasta c:\\users\\nomedousuário\\Downloads\\Notarius. \n
8. É possível executar o programa duas ou mais vezes e em cada programa, executar uma parte da lista, otimizando o tempo de extração.
"""

class TelaPython:
    def __init__(self) -> None:
        
        sg.theme('Reddit')
        
        layout_splash = [
                        [sg.Image(r'vilavelha.png', expand_y = True, expand_x = True)]
            
        ]
        
        layout_text_splash = [

                        [sg.Text('Carregando Notarius')]
        ]
        
        
        sg.splash = sg.Window(title="Notarius", layout = layout_splash, finalize=True, no_titlebar=True)
        sg.splash_text = sg.Window(title="Notarius", layout = layout_text_splash, finalize=True, no_titlebar=True, alpha_channel = 0.3)
        sg.splash_text.read(timeout = 2000)
        sg.splash.read(timeout=2000)
        sg.splash.Size
        sg.splash.close()
        sg.splash_text.close()
        

        layout_tela =   [
                    [sg.Text("Fonte dos Dados:"), sg.Radio('CENSEC', 'fonte', key= 'Fonte_censec', default = True), sg.Radio('SIGNO', 'fonte', key = 'Fonte_signo', default = False), sg.Radio('Ambos', 'fonte', key= 'Ambos')],
                    [sg.HSeparator()],
                    [sg.Text("Nome da lista:")], [sg.Input(size=(25,1), key = 'nomedalista')],
                    [sg.HSeparator()],
                    [sg.Text("Lista de CPFs ou CNPJs:")],
                    [sg.Multiline(size=(25,18), key = 'cpf_cnpj', expand_y = True, autoscroll = True), sg.Text(text = instrucoes)], 
                    [sg.Button('Iniciar Extração', key =  "Iniciar"), sg.Button('Pausar Extração', key = "Pause", visible = False), sg.Button('Parar Extração', key = 'Stop', visible = False, disabled = True)], 
                    [sg.Output(size=(150,12), key= "Output")]
                    ]


        self.janela = sg.Window(title="Notarius", layout = layout_tela, finalize=True)
        
    
    def Iniciar(self):
        Output.expand_y = True
        #self.janela.read()
        print ('Scraper Notarius')
        print ("Desenvolvido no Núcleo de Inovação, Prospecção e Análise de Dados (CGU-ES/NAE/NIPAD)")
        print ("Atualizações disponíveis em http://github.com/tgbremen/Notarius")
        print ("-------------------------------------------------------------------------")

        while(True):
            event, self.values = self.janela.read()
            if event == sg.WIN_CLOSED:
                break
            
            if event == "Iniciar":

                # Armazena valores da tela
                lista_CPF_CNPJ = self.values['cpf_cnpj']
                scraper_censec = self.values['Fonte_censec']
                scraper_signo = self.values['Fonte_signo']
                scraper_ambos = self.values['Ambos']
                nome_lista = self.values['nomedalista']

                # Remove os itens da lista que representam linhas em branco
                lista_CPF_CNPJ = list(filter(None,(lista_CPF_CNPJ.split('\n'))))

                # Atualiza os botão
                self.janela['Iniciar'].update("Iniciar Execução", disabled = True)
  
                # Executa o scraper

                if scraper_censec or scraper_ambos:
                    print("Iniciando CENSEC")
                    scraper = CR.censec_request()
                    scraper.scrap(lista_CPF_CNPJ, nome_lista)

                if scraper_signo or scraper_ambos:
                    print("Iniciando SIGNO")
                    scraper2 = SR.signo_request()
                    scraper2.scrap(lista_CPF_CNPJ, nome_lista)
                
                self.janela['Iniciar'].update("Iniciar Execução", disabled = False)
        self.janela.close()       

tela = TelaPython()
tela.Iniciar()




