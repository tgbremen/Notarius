import requests as r
import json
import pandas as pd
import datetime
import cookies_headers_censec as chc
import censec_chrome_request as censec_selenium
import os

class censec_request:
    def __init__(self) -> None:
        pass
        
    def define_params(documento):
        params = {
        'cpfCnpj': documento,
        }
        return params

    def pega_cookies(alvo_inicial):
        
        selenium = censec_selenium
        cookies2 = selenium.censec.scrap(censec_selenium.censec, alvo_inicial)
        cookies_dict = {}
        for cookie in cookies2:
            cookies_dict.update({cookie['name'] : cookie['value']})
            
        return cookies_dict
    
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
    
    def df_simplificado(self, df):
        df_simplif = df[['nomePesquisado', 'cpfCnpjPesquisado', 'documentoPesquisado', 'tipoDocumentoPesquisado', 'nomeTipoDocumentoPesquisado', 'tipoAto', 'naturezaEscritura', 'outraNatureza', 'valorAto',  'dataAto', 'fonte', 'cartorioNome', 'cartorioMunicípio', 'cartorioUf', 'cartorioCns', 'livro', 'livroComplemento', 'folha', 'folhaComplemento', 'IdPesquisado', 'IdAto']]
        df_simplif = df_simplif.drop_duplicates()
        return df_simplif

    def start_scrap(self, alvos, arquivo=''):
        alvo1 = [alvos[0]]  
        cookies = chc.cookies
        headers = chc.headers
        cookies = censec_request.pega_cookies(alvo1)
        headers['authorization'] = 'Bearer ' + cookies['access_token']
        df_partes = pd.DataFrame()
        df_atos = pd.DataFrame()
        global start2
        start2 = datetime.datetime.now() 
        df_final = pd.DataFrame()

        for num, doc in enumerate(alvos):
            params = censec_request.define_params(doc)    
            s = r.Session()
            url = 'https://censec.org.br/api/atos/cep/busca'
            
            # Faz a requisição consultando o CPF ou CNPJ
            pesquisa_cpf_cnpj = s.get(url, params=params, cookies=cookies, headers=headers)
            
            # Se a autenticação quebrou, vai dar status code 401 e esse while loga de novo no censec. 
            while pesquisa_cpf_cnpj.status_code == 401:
                cookies = censec_request.pega_cookies(doc)
                headers = chc.headers
                headers['authorization'] = 'Bearer ' + cookies['access_token']
                pesquisa_cpf_cnpj = s.get(url, params=params, cookies=cookies, headers=headers)
            
            # Converte pesquisa_cpf_cnpj em JSON
            json_pesquisa_cpf_cnpj = json.loads(pesquisa_cpf_cnpj.text)   #JSONDecodeError

            num_atos = len(json_pesquisa_cpf_cnpj['atos'])
            print("O {0}º CPF possui {1} atos".format(num +1, num_atos))
            
            # Salva a cada 100, se o df_final possuir conteúdo. 
            if ((((num+1) % 10) == 0) and (len(df_final) > 0)):
            # Cria cópia do DataFrame sem as partes e removendo duplicatas
                df_final2 = self.df_simplificado(df_final)        
                # Grava o arquivo. 
                self.persiste(arquivo, df_final, 'Principal', True, df_final2, 'Simplificado')
                print ('Arquivo INTERMEDIARIO salvo:', arquivo)
            
            
            # Loop para pegar detalhes do ato
            for num_ato, ato in enumerate(range(num_atos)):
                # Pega o ID e valor do nome do pesquiado de cada ato
                val_nome = json_pesquisa_cpf_cnpj['atos'][ato]['nome']
                id = json_pesquisa_cpf_cnpj['atos'][ato]['id']
                
                # Ajusta o JSON para evitar ambiguidades
                json_copia_pesquisado = json_pesquisa_cpf_cnpj['atos'][ato]
                json_copia_pesquisado['fonte'] = 'CENSEC/CEP - Central de Escritura e Procuração'
                json_copia_pesquisado['nomePesquisado'] = json_copia_pesquisado['nome']
                json_copia_pesquisado['IdPesquisado'] = json_copia_pesquisado['parteId']
                del json_copia_pesquisado['nome']
                del json_copia_pesquisado['parteId']
                del json_copia_pesquisado['cargaRef']
                del json_copia_pesquisado['cartorioId']
                
                # Pega detalhes do ID do ato            
                url_detalhes = 'https://censec.org.br/api/atos/cep/{0}?full=false'.format(id)
                dados_detalhes = s.get(url_detalhes, cookies=cookies, headers=headers)
                while dados_detalhes.status_code == 401:
                    cookies = censec_request.pega_cookies(doc)
                    headers = chc.headers
                    headers['authorization'] = 'Bearer ' + cookies['access_token']
                    dados_detalhes = s.get(url_detalhes, params=params, cookies=cookies, headers=headers)
                
                
                json_dados_detalhes = json.loads(dados_detalhes.text)
                
                # Acrescenta dados ao primeiro Json
                json_copia_pesquisado['valorAto'] = json_dados_detalhes['valor']
                json_copia_pesquisado['naturezaEscritura'] = json_dados_detalhes['naturezaEscritura']
                json_copia_pesquisado['outraNatureza'] = json_dados_detalhes['outraNatureza']
                json_copia_pesquisado['idAto'] = id
                del json_copia_pesquisado['naturezaAto']
                del json_copia_pesquisado['id']
                del json_copia_pesquisado['cargaId']
                del json_copia_pesquisado['carga']
                del json_copia_pesquisado['tipoDocumento']
                            
                # Pega os IDs das partes
                url_dados_partes = 'https://censec.org.br/api/busca-atos/cep/{0}/partes?q=&offset=0&limit=1000&order=desc&orderBy=nome'.format(id)
                dados_partes = s.get(url_dados_partes, cookies=cookies, headers=headers)
                while dados_partes == 401:
                    cookies = censec_request.pega_cookies(doc)
                    headers = chc.headers
                    headers['authorization'] = 'Bearer ' + cookies['access_token']
                    dados_partes = s.get(url_dados_partes, params=params, cookies=cookies, headers=headers)
                
                
                json_dados_partes = json.loads(dados_partes.text)
                num_partes = json_dados_partes['totalCount']

                for num_parte, parte in enumerate(range(num_partes)):
                    json_unificado = json_copia_pesquisado
                    
                    #Salva os dados básicos da parte em variáveis
                    val_nome_parte = json_dados_partes['items'][parte]['nome']
                    val_id_parte = json_dados_partes['items'][parte]['id']
                    
                    #Pega os detalhes da parte, junta com os dados básicos e formata o json 
                    url_detalhes_parte = 'https://censec.org.br/api/busca-atos/cep/{0}/partes/{1}'.format(id, val_id_parte)
                    detalhes_parte = s.get(url_detalhes_parte, cookies=cookies, headers=headers)
                    while detalhes_parte.status_code == 401:
                        cookies = censec_request.pega_cookies(doc)
                        headers = chc.headers
                        headers['authorization'] = 'Bearer ' + cookies['access_token']
                        detalhes_parte = s.get(url_detalhes_parte, params=params, cookies=cookies, headers=headers)

                    # Converte para Json e ajusta nome das chaves                        
                    json_detalhes_parte = json.loads(detalhes_parte.text)
                    del json_detalhes_parte['id'] #apaga a chave ID para evitar ambiguidade
                    json_detalhes_parte['conjuge'] = json_dados_partes['items'][parte]['conjuge']
                    json_detalhes_parte['nomeParte'] = json_detalhes_parte['nome']
                    del json_detalhes_parte['nome']
                    
                    # Converte para DF e faz outro ajuste de chaves
                    json_unificado.update(json_detalhes_parte)
                    df_unificado= pd.DataFrame([json_unificado])
                    df_unificado['tipoDocumentoPesquisado'] = ''
                    df_unificado['nomeTipoDocumentoPesquisado'] = ''
                    
                    	

                    # Renomeia colunas
                    df_unificado.columns = ['cpfCnpjPesquisado', 'documentoPesquisado', 'cartorioNome', 'cartorioMunicípio', 'cartorioUf', 'cartorioCns', 
                                            'tipoAto', 'livro', 'folha', 'livroComplemento', 'folhaComplemento', 'dataAto', 'fonte', 'nomePesquisado', 
                                            'IdPesquisado', 'valorAto', 'naturezaEscritura', 'outraNatureza', 'IdAto', 'qualidadeParte', 
                                            'identidadeParte', 'orgaoEmissorParte', 'numeroDocumentoParte', 'tipoDocumentoParte', 'conjugeParte', 'nomeParte', 
                                            'tipoDocumentoPesquisado', 'nomeTipoDocumentoPesquisado']

                    #Reordena colunas
                    df_unificado = df_unificado[['nomePesquisado', 'cpfCnpjPesquisado', 'documentoPesquisado', 'tipoDocumentoPesquisado', 'nomeTipoDocumentoPesquisado', 
                    'tipoAto', 'naturezaEscritura', 'outraNatureza', 'valorAto',  'dataAto', 'fonte', 'nomeParte', 'numeroDocumentoParte', 'tipoDocumentoParte', 'qualidadeParte', 
                    'identidadeParte', 'orgaoEmissorParte', 'conjugeParte', 'cartorioNome', 'cartorioMunicípio', 'cartorioUf', 'cartorioCns',
                    'livro', 'livroComplemento', 'folha', 'folhaComplemento', 'IdPesquisado', 'IdAto']]

                    
                    #Salva a linha no dataframe final
                    df_final = pd.concat([df_final, df_unificado])
                    #df_final.append(df_unificado, ignore_index = True)
                    del json_unificado
                    del df_unificado
                    
                    print("Finalizado o {0}º CPF/CNPJ: {1} - {2}, {3}º ato de {4}, {5}ª parte de {6}, nome: {7}".format(num+1, doc, val_nome, num_ato+1, num_atos, num_parte+1, num_partes, val_nome_parte))
        
        # Cria cópia do DataFrame sem as partes e removendo duplicatas
        if len(df_final) == 0:   
            colunas = ['nomePesquisado', 'cpfCnpjPesquisado', 'documentoPesquisado', 'tipoDocumentoPesquisado', 'nomeTipoDocumentoPesquisado', 
            'tipoAto', 'naturezaEscritura', 'outraNatureza', 'valorAto',  'dataAto', 'fonte', 'nomeParte', 'numeroDocumentoParte', 'tipoDocumentoParte', 'qualidadeParte', 
            'identidadeParte', 'orgaoEmissorParte', 'conjugeParte', 'cartorioNome', 'cartorioMunicípio', 'cartorioUf', 'cartorioCns',
            'livro', 'livroComplemento', 'folha', 'folhaComplemento', 'IdPesquisado', 'IdAto']
            df_final = pd.DataFrame(df_final, columns = colunas) 
        
        df_final2 = self.df_simplificado(df_final)      
        
        # Grava o arquivo. 
        self.persiste(arquivo, df_final, 'Principal', True, df_final2, 'Simplificado')
        print ('Arquivo final salvo:', arquivo)
        
            

        
    def scrap(self, lista_alvos, nome_lista = ''):
        scraper = censec_request()
        userprofile = os.environ['USERPROFILE']
        dirname = userprofile + r'\Downloads\Notarius'
        now = datetime.datetime.now()
        if not os.path.exists(dirname):
                os.makedirs(dirname)
        filename = dirname + "\\" + "censec_" + nome_lista + '_' + str(now.year) + str(now.month) + str(now.day) + str(now.hour) + str(now.minute) + str(now.second) + ".xlsx"
        print (filename)
        start = datetime.datetime.now()
        global start2
        print (start)
        scraper.start_scrap(lista_alvos, filename)                
        stop = datetime.datetime.now()
        print(stop)
        print ('Tempo de execução: ', stop - start)
        print ('Tempo de execução descontado o selenium: ', stop - start2)
        print ('Arquivo final:', filename)
