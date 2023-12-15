import PySimpleGUI as sg
from time import sleep
import CensecRequest4 as CR
import SignoRequest3 as SR

instrucoes = """Instruções: \n
1. Os CPFs e CNPJs não são validados pelo Notarius. Se houver erro, não encontrará nada. \n
2. Os CPFs e CNPJs podem ser escritos com ou sem pontos, hífens e barras. \n
3. Os zeros iniciais não podem ser omitidos ou o CPF / CNPJ não será encontrado. \n
4. Não use ou feche o chrome durante a execução do Notarius, apenas minimize. Não abra outra aba. \n
5. Os arquivos de saída serão censec_nomedalista_datahora.xlsx ou signo_nomedalista_datahora.xlsx \n
6. Os arquivos de saída serão salvos na pasta c:\\users\\nomedousuário\\Downloads\\Notarius. \n
"""

texto_conversor = """O conversor transforma uma planilha Excel gerada pelo Notarius em outro formato para uso no Macros ou I2. \n 
Os formatos são JSON (compatível com o Macros2) e ANX (compatível com o IBM I2). \n
O arquivo será gerado na pasta c:\\users\\nomedousuário\\Downloads\\Notarius. \n
O conversor permite que se edite as planilhas do Notarius para exportar apenas as relações de interesse. \n
"""

texto_editor = """O editor cria relações novas para serem inseridas em sistemas que os visualizem. \n
O arquivo de saída será JSON (compatível com o Macros 2) ou ANX (compatível com IBM I2).
O arquivo será gerado na pasta c:\\users\\nomedousuário\\Downloads\\Notarius. 
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
        
        layout_conversor = [
                            [sg.Text("Arquivo Excel: ")], 
                            [sg.Input(key='-ARQ_EXCEL-'), sg.FileBrowse("Buscar arquivo")], 
                            [sg.Checkbox("JSON (compatível com Macros2)", default = True, key = 'JSON_conversor'), sg.Checkbox("ANX (compatível com IBM I2)", key = 'ANX_conversor'),
                             ], 
                            [sg.Button("Converter Arquivo", key = "-CONVERSOR-")], 
                            [sg.Text(texto_conversor)],
                            #[sg.Output(size=(105,10), expand_y = True, key= "Output")]
                            ]
        
        cabecalho = ['Nome', 'ID', 'Tipo', 'Sexo']
        cabecalho2 = ['Origem', 'Destino', 'Vínculo']
        linhas = []
        #linhas = [['1', '', '', ''], ['2', '', '', '']]
        
        
        layout_editor_ligacoes = [[sg.Text("Tipo: "), sg.OptionMenu(default_value ='PF',values=('PF', 'PJ'),key='-TIPOPESSOA-'), sg.Text("Nome: "), sg.Input(s=15, key ='-NOME-' ), 
                                   sg.Text("CPF / CNPJ: "), sg.Input(s=15, key ='-CPFCNPJ-'), sg.Text("Sexo: "), 
                                   sg.OptionMenu(default_value ='1 - Masculino', values =('1 - Masculino', '2 - Feminino', '0 - Desconhecido'), key='-SEXO-') ],
                                  [sg.Button('Adicionar Item', key =  "-ADICIONAR1-"), sg.Button('Remover item selecionado', key =  "-REMOVER1-"), sg.Button('Limpar', key =  "-LIMPAR1-")],
                                  [sg.Table(values=linhas, headings=cabecalho, auto_size_columns=True, display_row_numbers=False, justification='center', key='-TABELA1-', 
                                            selected_row_colors='red on yellow', enable_events=True, expand_x=True, expand_y=True, enable_click_events=True, num_rows=5)],
                                  [sg.Text("Origem (CPF/CNPJ): "), sg.Input(s=15, key ='-ORIGEM-'), sg.Text("Destino (CPF/CNPJ): "), sg.Input(s=15, key ='-DESTINO-'), 
                                   sg.Text("Descrição do Vínculo: "), sg.Input(s=15, key ='-VINCULO-')],
                                  [sg.Button('Adicionar Item', key =  "-ADICIONAR2-"), sg.Button('Remover item selecionado', key =  "-REMOVER2-"), sg.VSeparator(), 
                                   sg.Button('Copiar ID para Origem', key =  "-COPIAR_ORIGEM-"), sg.Button('Copiar ID para Destino', key =  "-COPIAR_DESTINO-"), 
                                   sg.Button('Limpar', key =  "-LIMPAR2-")],
                                  [sg.Table(values=linhas, headings=cabecalho2, auto_size_columns=True, display_row_numbers=False, justification='center', key='-TABELA2-', 
                                            selected_row_colors='red on yellow', enable_events=True, expand_x=True, expand_y=True, enable_click_events=True, num_rows=5)],
                                  [sg.Checkbox("JSON (compatível com Macros2)", default = True, key = 'JSON_editor'), sg.Checkbox("ANX (IBM I2)", key = 'ANX_editor'), sg.VSeparator(),
                                   sg.Text("Nome do arquivo: "), sg.Input(s=15, key ='-NOME_ARQUIVO-'), sg.Button("Gerar Arquivo", key = "-MAKE_FILE-")]                                
                                  ]

        layout_cartorios =   [
                    [sg.Text("Fonte dos Dados:"), sg.Radio('CENSEC', 'fonte', key= 'Fonte_censec', default = True), sg.Radio('SIGNO', 'fonte', key = 'Fonte_signo', default = False), sg.Radio('Ambos', 'fonte', key= 'Ambos')],
                    [sg.Text("Arquivos de saída:"), sg.Checkbox('XLSX (Planilha Excel)', default = True, disabled = True, key = 'Excel'), sg.Checkbox('JSON (MACROS2)', default = True, key = 'Macros2'), sg.Checkbox('ANX (IBM I2)', default = False, key = 'I2')], 
                    [sg.Text("Nome da lista:"), sg.Input(size=(25,1), key = 'nomedalista')],
                    [sg.HSeparator()],
                    [sg.Text("Lista de CPFs ou CNPJs:")],
                    [sg.Multiline(size=(20,8), key = 'cpf_cnpj', expand_y = True, autoscroll = True), sg.Text(text = instrucoes)], 
                    [sg.Button('Iniciar Extração', key =  "Iniciar"), sg.Button('Pausar Extração', key = "Pause", visible = False), sg.Button('Parar Extração', key = 'Stop', visible = False, disabled = True)], 
                    #[sg.Output(size=(105,10), expand_y = True, key= "Output")]
                    ]

        layout_tab = [[sg.TabGroup([[sg.Tab('Cartórios', layout_cartorios, key='-mykey-'), sg.Tab('Conversor', layout_conversor), sg.Tab('Editor', layout_editor_ligacoes)
                ]], key='-group2-', title_color='black', selected_title_color='blue')], [sg.Output(size=(110,8), expand_y = True, key= "Output")]]

        self.janela = sg.Window(title="Notarius", layout = layout_tab, finalize=True)
        
    
    def Iniciar(self):
        print ('Scraper Notarius')
        print ("Desenvolvido no Núcleo de Inovação, Prospecção e Análise de Dados (CGU-ES/NIPAD)")
        print ("Atualizações disponíveis em http://github.com/tgbremen/Notarius")
        print ("-------------------------------------------------------------------------")
        
        #linhas = [['1', '', '', ''], ['2', '', '', '']]
        linhas = []
        #tabela, tabela2 = [['1', '', '', ''], ['2', '', '', '']], [['1', '', '', ''], ['2', '', '', '']]
        tabela = []
        tabela2 = []
        
        while(True):
            event, self.values = self.janela.read()
            if event == sg.WIN_CLOSED:
                break
            
            if event == "Iniciar": #Botão da tela de cartórios
                # Armazena valores da tela de cartórios
                lista_CPF_CNPJ = self.values['cpf_cnpj']
                scraper_censec = self.values['Fonte_censec']
                scraper_signo = self.values['Fonte_signo']
                scraper_ambos = self.values['Ambos']
                nome_lista = self.values['nomedalista']
                macros2 = self.values['Macros2']
                I2 = self.values['I2']

                # Remove os itens da lista que representam linhas em branco
                lista_CPF_CNPJ = list(filter(None,(lista_CPF_CNPJ.split('\n'))))

                # Atualiza os botão
                self.janela['Iniciar'].update("Iniciar Execução", disabled = True)
  
                # Executa o scraper
                if scraper_censec or scraper_ambos:
                    print("Iniciando CENSEC")
                    scraper = CR.censec_request()
                    scraper.scrap(lista_CPF_CNPJ, nome_lista, gerar_json = macros2, gerar_anx = I2)

                if scraper_signo or scraper_ambos:
                    print("Iniciando SIGNO")
                    scraper2 = SR.signo_request()
                    scraper2.scrap(lista_CPF_CNPJ, nome_lista, gerar_json = macros2, gerar_anx = I2)
                
                self.janela['Iniciar'].update("Iniciar Execução", disabled = False)
            
            # Evento da tela do Conversor, para converter o XLSX em JSON ou ANX. 
            if event == "-CONVERSOR-":
                macros2 = self.values['JSON_conversor']
                I2 = self.values['ANX_conversor']
                arquivo = self.values['-ARQ_EXCEL-']
                if macros2: ### REMOVER
                    print("Gerando o arquivo JSON")
                if I2: ### REMOVER
                    print("Gerando o arquivo I2")
                if (macros2 or I2):
                    conversor = CR.censec_request()
                    conversor.grava_json_anx(arquivo, macros2, I2)
            
            
            # Eventos para a tela do Editor de ligações
            if event == '-ADICIONAR1-':
                cpfcnpj_no = self.values['-CPFCNPJ-']
                nome_no = self.values['-NOME-']
                sexo = self.values['-SEXO-']
                
                #Ajusta o valor do sexo para o código correspondente
                match sexo : 
                    case '1 - Masculino':
                        sexo = 1
                    case '2 - Feminino':
                        sexo = 2
                    case '0 - Desconhecido':
                        sexo = 0

                tipo_pessoa = self.values['-TIPOPESSOA-']
                linha = [nome_no, cpfcnpj_no, tipo_pessoa, sexo]
                tabela.append(linha)
                self.janela['-TABELA1-'].update(values = tabela)

            if event == '-REMOVER1-':
                num_linhas = self.janela['-TABELA1-'].SelectedRows
                num_linhas.sort(reverse=True)
                for linha in num_linhas:
                    tabela.pop(linha)
                self.janela['-TABELA1-'].update(values = tabela)

            if event == '-ADICIONAR2-':
                cpfcnpj_origem = self.values['-ORIGEM-']
                cpfcnpj_destino = self.values['-DESTINO-']
                vinculo = self.values['-VINCULO-']
                linha2 = [cpfcnpj_origem, cpfcnpj_destino, vinculo]
                tabela2.append(linha2)
                self.janela['-TABELA2-'].update(values = tabela2)
                
            if event == '-REMOVER2-':
                num_linhas2 = self.janela['-TABELA2-'].SelectedRows
                num_linhas2.sort(reverse=True)
                for linha2 in num_linhas2:
                    tabela2.pop(linha2)
                self.janela['-TABELA2-'].update(values = tabela2)
                
            if event == '-MAKE_FILE-':
                macros2 = self.values['JSON_editor']
                I2 = self.values['ANX_editor']
                arq = self.values['-NOME_ARQUIVO-']
                if macros2:
                    print("Gerando o arquivo JSON")
                if I2:
                    print("Gerando o arquivo I2")
                if (macros2 or I2):
                    import editor_arq as EA
                    EA.cria_arquivos(tabela, tabela2, macros2, I2, arq)
                if (not macros2 and not I2):
                    print("É necessário selecionar ao menos 1 tipo de arquivo (JSON ou ANX) para continuar.")

            if event == '-LIMPAR1-':
                self.janela['-NOME-'].update('')
                self.janela['-CPFCNPJ-'].update('')
                self.janela['-TIPOPESSOA-'].update('PF')
                self.janela['-SEXO-'].update('1 - Masculino')
                
            if event == '-LIMPAR2-':
                self.janela['-ORIGEM-'].update('')
                self.janela['-DESTINO-'].update('')
                self.janela['-VINCULO-'].update('')
                
            if event == '-COPIAR_ORIGEM-':
                #Obtém a primeira linha selecionada da tabela1, coluna 2 (ID - CPF/CNPJ)
                self.janela['-ORIGEM-'].update(tabela[self.janela['-TABELA1-'].SelectedRows[0]][1])
                
            if event == '-COPIAR_DESTINO-':
                self.janela['-DESTINO-'].update(tabela[self.janela['-TABELA1-'].SelectedRows[0]][1])
                
            
                 
                
            
                

        self.janela.close()       

tela = TelaPython()
tela.Iniciar()




