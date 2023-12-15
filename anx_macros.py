import unidecode
import re

def texto_puro(texto):
    # Remove acentuação e substitui caracteres acentuados pelos equivalentes sem acento
    texto_sem_acentos = unidecode.unidecode(texto)

    # Remove caracteres especiais (exceto letras e números)
    texto_sem_especiais = re.sub(r'[^\w\s]', '', texto_sem_acentos)

    return texto_sem_especiais

def inicia_xml():
    info_inicial = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Chart GridVisibleOnAllViews="true">
    <EntityTypeCollection>
        <EntityType PreferredRepresentation="RepresentAsIcon" IconFile="Person" Name="Person" Colour="0"/>
        <EntityType PreferredRepresentation="RepresentAsIcon" IconFile="Woman" Name="Woman" Colour="0"/>
        <EntityType PreferredRepresentation="RepresentAsIcon" IconFile="Office" Name="Office" Colour="0"/>
        <EntityType PreferredRepresentation="RepresentAsIcon" IconFile="Workshop" Name="Workshop" Colour="0"/>
        <EntityType PreferredRepresentation="RepresentAsIcon" IconFile="Office" Name="Office" Colour="0"/>
        <EntityType PreferredRepresentation="RepresentAsIcon" IconFile="Phone" Name="Phone" Colour="0"/>
        <EntityType PreferredRepresentation="RepresentAsIcon" IconFile="Place" Name="Place" Colour="0"/>
        <EntityType PreferredRepresentation="RepresentAsIcon" IconFile="Account" Name="Account" Colour="0"/>
        <EntityType PreferredRepresentation="RepresentAsIcon" IconFile="Cabinet" Name="Cabinet" Colour="0"/>
        <EntityType PreferredRepresentation="RepresentAsIcon" IconFile="Office" Name="Office" Colour="0"/>
    </EntityTypeCollection>
    <LinkTypeCollection>
        <LinkType Colour="0" Name="Link"/>
    </LinkTypeCollection>
    <ChartItemCollection>'''
    return(info_inicial)

def meio_xml():
    info_meio = """
    </ChartItemCollection>
    <ChartItemCollection>"""
    return info_meio

def fim_xml():
    info_fim = """
    </ChartItemCollection>
</Chart>"""
    return info_fim

def cria_no_pessoa(no, nome, cpf_cnpj):
    
    if len(cpf_cnpj) == 11:
        no = cria_no_PF(no, nome, cpf_cnpj)
    elif len(cpf_cnpj) == 14:
        no = cria_no_PJ(no, nome, cpf_cnpj)
    return no 


def cria_no_PF(no, nome, cpf, sexo=0):
    nome = texto_puro(nome)
    cpf = texto_puro(cpf)
    label = cpf + ' - ' + nome
    if sexo == 0 or sexo == 1:
        tipo = 'Person'
    elif sexo == 2:
        tipo = 'Woman'
    no_item = f"""
            <ChartItem Label="{label}">
                <End>
                    <Entity EntityId="{cpf}" LabelIsIdentity="false" Identity="{cpf}">
                        <Icon TextY="16" TextX="0">
                            <IconStyle Type="{tipo}"/>
                        </Icon>
                    </Entity>
                </End>
            </ChartItem>"""
        
    no = no + no_item
    return no

def cria_no_PJ(no, nome, cnpj):
    cnpj = texto_puro(cnpj)
    nome = texto_puro(nome)
    label = cnpj + ' - ' + nome
    no_item = f"""
            <ChartItem Label="{label}">
                <End>
                    <Entity EntityId="{cnpj}" LabelIsIdentity="false" Identity="{cnpj}">
                        <Icon TextY="16" TextX="0">
                            <IconStyle Type="Office"/>
                        </Icon>
                    </Entity>
                </End>
            </ChartItem>"""
        
    no = no + no_item
    return no

def cria_no_ESC(no, nome, id):
    nome = texto_puro(nome)
    id = texto_puro(id)
    no_item = f"""
            <ChartItem Label="{nome}">
                <End>
                    <Entity EntityId="{id}" LabelIsIdentity="false" Identity="{id}">
                        <Icon TextY="16" TextX="0">
                            <IconStyle Type="Cabinet"/>
                        </Icon>
                    </Entity>
                </End>
            </ChartItem>"""
    
    no = no + no_item
    return no

def cria_no_PRC(no, nome, id):
    nome = texto_puro(nome)
    id = texto_puro(id)
    no_item = f"""
            <ChartItem Label="{nome}">
                <End>
                    <Entity EntityId="{id}" LabelIsIdentity="false" Identity="{id}">
                        <Icon TextY="16" TextX="0">
                            <IconStyle Type="Cabinet"/>
                        </Icon>
                    </Entity>
                </End>
            </ChartItem>"""
    
    no = no + no_item
    return no


def cria_ligacao(ligacao, origem, destino, descricao):
    origem = texto_puro(origem)
    destino = texto_puro(destino)
    descricao = texto_puro(descricao)
    ligacao_item = f"""
        <ChartItem Label="{descricao}">
            <Link End2Id="{destino}" End1Id="{origem}">
                <LinkStyle Type="Link" ArrowStyle="ArrowOnHead"/>
            </Link>
        </ChartItem>"""
    
    ligacao = ligacao + ligacao_item
    return ligacao


def grava_anx(nome_arquivo, texto):
    with open(nome_arquivo, "w") as f:
        f.write(texto)
    print("Arquivo para IBM I2 salvo: ", nome_arquivo)




