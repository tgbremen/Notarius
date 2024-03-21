def cria_arquivo_json(tab1, tab2, nome_arq):
    no = []
    import json_macros as jm
    for linha in tab1:
        if linha[2] == 'PF':
            no = jm.cria_no_PF(no, linha[0], linha[1], linha[3])
        elif linha[2] == 'PJ':
            no = jm.cria_no_PJ(no, linha[0], linha[1], linha[3])
    
    ligacao = []
    for linha2 in tab2:
        ligacao = jm.cria_ligacao(ligacao, linha2[0], linha2[1], linha2[2])
    
    jm.grava_json(no, ligacao, nome_arq)

def cria_arquivo_anx(tab1, tab2, arquivo):
    no = []
    import anx_macros as anx
    
    #Cria parte inicial, meio e fim do XML    
    no = anx.inicia_xml()
    meio = anx.meio_xml()
    fim = anx.fim_xml()
    
    #Acrescenta cada nó da tabela 1    
    for linha in tab1:
        if linha[2] == 'PF':
            no = anx.cria_no_PF(no, linha[0], linha[1], linha[3])
        elif linha[2] == 'PJ':
            no = anx.cria_no_PJ(no, linha[0], linha[1])
    
    #Acrescenta cada ligação da tabela 2    
    ligacao = ''
    for linha2 in tab2:
        ligacao = anx.cria_ligacao(ligacao, linha2[0], linha2[1], linha2[2])
    
    texto_xml = no + meio + ligacao + fim
    anx.grava_anx(arquivo, texto_xml)
    

def cria_arquivos(tab1, tab2, json, anx, nome_arq):
    import os
    userprofile = os.environ['USERPROFILE']
    dirname = userprofile + r'\Downloads\Notarius'
    nome_arquivo = dirname + '\\' + nome_arq
    if json:
        arquivo = nome_arquivo + '.json'
        cria_arquivo_json(tab1, tab2, arquivo)
    if anx:
        arquivo = nome_arquivo + '.anx'
        cria_arquivo_anx(tab1, tab2, arquivo)