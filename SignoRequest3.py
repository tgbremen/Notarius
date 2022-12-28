import requests as r
import pandas as pd
import datetime
import cookies_headers_signo as chs
from time import sleep
import json
import signo_chrome_request as scr
import os

class signo_request():    
    def __init__(self) -> None:
        pass
    
    def toca_beep(self, frequency=2500, duration=1000):   
        import winsound
        winsound.Beep(frequency, duration)

    def persiste(self, filename, dataf, aba, outro_df=False, dataf2=None, aba2=None):
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        dataf.to_excel(writer, sheet_name=aba, index=False, header=True)
        workbook  = writer.book

        worksheet = writer.sheets[aba]
        worksheet.set_column(1,30, 18)
        worksheet.set_column(0, 0, 50)
        worksheet.set_column(10, 11, 50)
        worksheet.set_column(18, 18, 50)
        worksheet.set_column(2, 4, 30)
        worksheet.set_column(12, 12, 23)
        if outro_df: 
            dataf2.to_excel(writer, sheet_name=aba2, index=False, header=True)
            worksheet = writer.sheets[aba2]
            worksheet.set_column(1,30, 18)
            worksheet.set_column(0, 0, 50)
            worksheet.set_column(18, 18, 50)
            worksheet.set_column(2, 4, 30)
            worksheet.set_column(12, 13, 23)
            worksheet.set_column(10, 11, 50)
        

        writer.close()
        pass
    
    def salva_arquivo(self, pd_final, filename):
        # Cria campos que não existem no Signo e que existem no Censec
        pd_final['outraNatureza'] = ''
        pd_final['numeroDocumentoParte'] = ''
        pd_final['fonte'] = 'SIGNO/CEP - Central de Escritura e Procuração'
        pd_final['uf'] = 'SP'
        
        # Padroniza a informação dentro de livro e folha conforme modelo do Censec. No signo, toda informação estava aglutinada em livrocomplemento e movemos pra livro (idem pra folha)
        pd_final['livro'] = pd_final['livroComplemento']
        pd_final['folha'] = pd_final['folhaComplemento']
        pd_final['livroComplemento'] = ''
        pd_final['folhaComplemento'] = ''
        
        # Seleciona as colunas e as reordena
        df_final = pd_final[['nome', 'cpfPesquisado', 'tipoDocumentoPesquisado', 'nomeTipoDocumentoPesquisado', 'documentoPesquisado', 'tipoAto',  'naturezaAto',  
                             'outraNatureza', 'valorOperacao',  'dataAto', 'fonte', 'nomeParte',  'numeroDocumentoParte',  
                             'tipoDocumentoParte', 'qualificacaoParte', 'idParte', 'orgaoEmissor', 'cpfConjuge', 'nomeCartorio', 
                             'municipio', 'uf', 'numeroCns', 'livro', 'livroComplemento', 'folha', 'folhaComplemento', 'id', 'idAto']] 
        df_final = df_final.drop_duplicates()
        
        # Renomeia as colunas (padrão signo - censec)
        df_final.columns = ['nomePesquisado', 'cpfCnpjPesquisado', 'documentoPesquisado', 'tipoDocumentoPesquisado', 'nomeTipoDocumentoPesquisado', 'tipoAto', 'naturezaEscritura', 'outraNatureza', 
                    'valorAto',  'dataAto', 'fonte', 'nomeParte', 'numeroDocumentoParte', 'tipoDocumentoParte', 'qualidadeParte', 'identidadeParte',
                    'orgaoEmissorParte', 'conjugeParte' , 'cartorioNome', 'cartorioMunicípio', 'cartorioUf', 'cartorioCns', 
                    'livro', 'livroComplemento', 'folha', 'folhaComplemento', 'IdPesquisado', 'IdAto']
        
        # Coloca data com separador de / ao invés de -
        df_final['dataAto'].replace('-', '/', regex = True, inplace = True)
        
        # Cria uma cópia simplificada dos dados, sem as partes
        df_final2 = df_final[['nomePesquisado', 'cpfCnpjPesquisado', 'documentoPesquisado', 'tipoDocumentoPesquisado', 'nomeTipoDocumentoPesquisado', 'tipoAto', 'naturezaEscritura', 'outraNatureza', 
                    'valorAto',  'dataAto', 'fonte', 'cartorioNome', 'cartorioMunicípio', 'cartorioUf', 'cartorioCns', 
                    'livro', 'livroComplemento', 'folha', 'folhaComplemento', 'IdPesquisado', 'IdAto']]
        df_final2 = df_final2.drop_duplicates()
        
        # Grava o arquivo. 
        self.persiste(filename, df_final, 'Principal', True, df_final2, 'Simplificado')
        print ('Arquivo salvo:', filename)

        
    #Fazer a classe e a passagem de parâmetros
    def start_scrap(self, alvos, nome_lista):
        # Cria a pasta C:\users\usuário\Dowloads\Notarius se ela não existir. Armazena esse caminho. Cria a estrutura do nome do arquivo. 
        userprofile = os.environ['USERPROFILE']
        now = datetime.datetime.now()
        dirname = userprofile + r'\Downloads\Notarius'
        if not os.path.exists(dirname):
            os.makedirs(dirname)
        filename = dirname + "\\" + "signo_" + nome_lista + '_' + str(now.year) + str(now.month) + str(now.day) + str(now.hour) + str(now.minute) + str(now.second) + ".xlsx"
        
        # Inicia o scrap
        params = chs.params
        headers = chs.headers
        params['size'] = 2000
        driver = scr.signo()
        token = driver.scrap([alvos[0]], '')
        headers['authorization'] = token
        lista2 = []
        s = r.Session()
        num_alvos = len(alvos)
        url = r'https://backend.signo.org.br/api/ato/consultas-internas/listar-atos?content=&size=10000&centrais=8' #URL de pesquisa do cpf / cnpj para o signo. 
        url_detalhe_ato = r'https://backend.signo.org.br/api/ato/ato-cep/ato-detailed?idAto={0}'

        for num, alvo in enumerate(alvos):
            params['cpf'] = alvo
            # Executa a consulta para o cpf / nome
            print("Obtendo dados do {0}º CPF/CNPJ de um total de {1}. CPF/CNPJ: {2}".format(num+1, num_alvos, alvo))
            consulta_parte = s.get(url, params=params, headers=headers)
            while consulta_parte.status_code != 200:
                #self.toca_beep()
                sleep(10)
                token = driver.scrap([alvo], '')
                headers['authorization'] = token
                try:
                    consulta_parte = s.get(url, params=params, headers=headers)
                except:
                    pass
                
            json_consulta_parte = json.loads(consulta_parte.text)['content']

            
            # Consulta os detalhes do ato
            for ato in json_consulta_parte:
                id_ato = ato['idAto']
                
                # substitui documento, cpf e tipoDocumento por chaves que não vão ser alteradas pela consulta seguinte.
                doc_pesquisado = ato['documento']
                tipo_documento_pesquisado = ato['tipoDocumento']
                nome_tipo_documento_pesquisado = ato['nomeTipoDocumento']
                cpf_pesquisado = ato['cpf']
                ato['documentoPesquisado'] = doc_pesquisado
                ato['tipoDocumentoPesquisado'] = tipo_documento_pesquisado
                ato['nomeTipoDocumentoPesquisado'] = nome_tipo_documento_pesquisado
                ato['cpfPesquisado'] = cpf_pesquisado
                
                # Consulta detalhes do ato (obtém as partes e outros dados)
                detalhe_ato = s.get(url_detalhe_ato.format(id_ato), params=params, headers=headers)
                while detalhe_ato.status_code != 200:
                    #self.toca_beep()
                    sleep(10)
                    token = driver.scrap([alvo], '')
                    headers['authorization'] = token
                    try:
                        detalhe_ato = s.get(url_detalhe_ato, params=params, headers=headers)
                    except:
                        pass
                    
                json_detalhe_ato = json.loads(detalhe_ato.text)
                json_detalhe_ato.update({'cpfPesquisado' : alvo})
                partes = json_detalhe_ato['partes']
                del json_detalhe_ato['partes']
                for parte in partes:
                    parte['tipoDocumentoParte'] = parte['tipoDocumento']
                    del parte['tipoDocumento']
                    lista = {}
                    lista.update(ato)
                    lista.update(json_detalhe_ato)
                    lista.update(parte)
                    lista2.append(lista)
                    l2 = pd.DataFrame(lista2)
            
            if ((num+1) % 100 == 0) or (num == (num_alvos -1)):
                if len(lista2) == 0:
                    colunas = ['nomePesquisado', 'cpfCnpjPesquisado', 'documentoPesquisado', 'tipoDocumentoPesquisado', 'nomeTipoDocumentoPesquisado', 
                    'tipoAto', 'naturezaEscritura', 'outraNatureza', 'valorAto',  'dataAto', 'fonte', 'nomeParte', 'numeroDocumentoParte', 'tipoDocumentoParte', 'qualidadeParte', 
                    'identidadeParte', 'orgaoEmissorParte', 'conjugeParte', 'cartorioNome', 'cartorioMunicípio', 'cartorioUf', 'cartorioCns',
                    'livro', 'livroComplemento', 'folha', 'folhaComplemento', 'IdPesquisado', 'IdAto']
                    df = pd.DataFrame(columns = colunas)
                    
                    colunas2 = ['nomePesquisado', 'cpfCnpjPesquisado', 'documentoPesquisado', 'tipoDocumentoPesquisado', 'nomeTipoDocumentoPesquisado', 'tipoAto', 'naturezaEscritura', 'outraNatureza', 
                    'valorAto',  'dataAto', 'fonte', 'cartorioNome', 'cartorioMunicípio', 'cartorioUf', 'cartorioCns', 
                    'livro', 'livroComplemento', 'folha', 'folhaComplemento', 'IdPesquisado', 'IdAto']
                    df2 = pd.DataFrame(columns = colunas2)
                    self.persiste(filename, df, 'Principal', True, df2, 'Simplificado')
                    
                
                else:
                    colunas = list(lista.keys())
                    df = pd.DataFrame(columns = colunas)
                    df = pd.concat([df, l2])
                    self.salva_arquivo(df, filename)
                
    
    def scrap(self, lista_alvos, nome_lista = ''):
        start = datetime.datetime.now()
        print (start) 
        self.start_scrap(lista_alvos, nome_lista)                 
        stop = datetime.datetime.now()
        print(stop)
        print ('Tempo de execução: ', stop - start)