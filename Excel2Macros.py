#Converte um arquivo XLSX proveniente do Notarius para JSON no padrão do MACROS2. 

import openpyxl
import json_macros as jm

def ler_planilha(nome_arquivo):
    #Inicializa as variáveis
    no = []
    ligacao = []
    lista_nomes = []
    ato_atual = ''
    
    # Carregar a planilha
    workbook = openpyxl.load_workbook(nome_arquivo)
    sheet = workbook['Principal']

    # Iterar a planilha linha a linha, a partir da linha 2.
    for num, row in enumerate(sheet.iter_rows()):
        if num == 0:
            continue
        #Lê o pesquisado e grava o nó
        nome_pesquisado = row[0].value if row[0].value else ''
        cpf_cnpj_pesquisado = row[1].value if row[1].value else ''
        if nome_pesquisado not in lista_nomes:
            lista_nomes.append(nome_pesquisado)
            jm.cria_no_pessoa(no, nome_pesquisado, cpf_cnpj_pesquisado)
        
        #Lê o ato (escritura ou procuração) e grava
        Id_Ato = row[27].value if row[27].value else '' #IdAto
        if ato_atual != Id_Ato:
            ato_atual = Id_Ato
            id_Ato_reduzido = str(Id_Ato)[-5:]
            col_J = row[9].value[:10] if row[9].value[:10] else '' #dataAto, os 10 primeiros caracteres
            id_data_Ato = col_J + '_' + id_Ato_reduzido
            col_F = row[5].value if row[5].value else '' #tipoAto
            col_G = row[6].value if row[6].value else '' #NaturezaEscritura
            if col_F == 'Escritura':
                Ato = col_F + ' - ' + col_G
                jm.cria_no_ESC(no, Ato, id_data_Ato)
            if ((col_F == 'Procuracao') or (col_F == 'Procuração')):
                Ato = col_F
                jm.cria_no_PRC(no, Ato, id_data_Ato)                
                
        #Lê o nome da parte e grava
        nome_parte = row[11].value if row[11].value else ''
        cpf_cnpj_parte = row[12].value if row[12].value else ''
        if nome_parte not in lista_nomes:
            lista_nomes.append(nome_parte)
            jm.cria_no_pessoa(no, nome_parte, cpf_cnpj_parte)
        
        #Cria a relação entre parte e ato
        qualidadeParte = row[14].value if row[14].value else ''
        if qualidadeParte == "Outorgado":
            ligacao = jm.cria_ligacao(ligacao, id_data_Ato, cpf_cnpj_parte, qualidadeParte)
        else: 
            ligacao = jm.cria_ligacao(ligacao, cpf_cnpj_parte, id_data_Ato, qualidadeParte)
    
    return no, ligacao

