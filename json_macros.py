import json
import unidecode
import re

def texto_puro(texto):
  # Remove acentuação e substitui caracteres acentuados pelos equivalentes sem acento
  texto_sem_acentos = unidecode.unidecode(texto)

  # Remove caracteres especiais (exceto letras e números)
  texto_sem_especiais = re.sub(r'[^\w\s]', '', texto_sem_acentos)

  return texto_sem_especiais

def cria_no_pessoa(no, nome, cpf_cnpj):
  if len(cpf_cnpj) == 11:
    cria_no_PF(no, nome, cpf_cnpj)
  elif len(cpf_cnpj) == 14:
    cria_no_PJ(no, nome, cpf_cnpj)

def cria_no_PF(no, nome, cpf, sexo=0):
  nome = texto_puro(nome)
  cpf = texto_puro(cpf)    
  no_item = {
      "id": cpf,
      "tipo": "PF",
      "sexo": sexo,
      "label": nome,
      "camada": 1,
      "situacao": 0,
      "m1": 0,
      "m2": 0,
      "m3": 0,
      "m4": 0,
      "m5": 0,
      "m6": 0,
      "m7": 0,
      "m8": 0,
      "m9": 0,
      "m10": 0,
      "m11": 0,
      "posicao": {
          "x": -51.059146533761975,
          "y": 49.78295662546248
      }
      }
  
  no.append(no_item)
  return no
  
def cria_no_PJ(no, nome, cnpj):
  nome = texto_puro(nome)
  cnpj = texto_puro(cnpj) 
  no_item = {
      "id": cnpj,
      "tipo": "PJ",
      "sexo": 1,
      "label": nome,
      "camada": 1,
      "situacao": 0,
      "m1": 0,
      "m2": 0,
      "m3": 0,
      "m4": 0,
      "m5": 0,
      "m6": 0,
      "m7": 0,
      "m8": 0,
      "m9": 0,
      "m10": 0,
      "m11": 0,
      "posicao": {
          "x": -51.059146533761975,
          "y": 49.78295662546248
      }
      }
  
  no.append(no_item)
  return no
  
def cria_no_ESC(no, nome, id):
  nome = texto_puro(nome)
  id = texto_puro(id) 
  no_item = {
      "id": id,
      "tipo": "ESC",
      "sexo": 0,
      "label": nome,
      "camada": 1,
      "situacao": 0,
      "m1": 0,
      "m2": 0,
      "m3": 0,
      "m4": 0,
      "m5": 0,
      "m6": 0,
      "m7": 0,
      "m8": 0,
      "m9": 0,
      "m10": 0,
      "m11": 0,
      "posicao": {
          "x": -51.059146533761975,
          "y": 49.78295662546248
      }
      }
  
  no.append(no_item)
  return no
  
def cria_no_PRC(no, nome, id):
  nome = texto_puro(nome)
  id = texto_puro(id)
  no_item = {
      "id": id,
      "tipo": "PRC",
      "sexo": 0,
      "label": nome,
      "camada": 1,
      "situacao": 0,
      "m1": 0,
      "m2": 0,
      "m3": 0,
      "m4": 0,
      "m5": 0,
      "m6": 0,
      "m7": 0,
      "m8": 0,
      "m9": 0,
      "m10": 0,
      "m11": 0,
      "posicao": {
          "x": -51.059146533761975,
          "y": 49.78295662546248
      }
      }
  
  no.append(no_item)
  return no

def cria_ligacao(ligacao, origem, destino, descricao):
  origem = texto_puro(origem)
  destino = texto_puro(destino)
  descricao = texto_puro(descricao)
    
  ligacao_item = {
    "origem": origem,
    "destino": destino,
    "cor": "Black",
    "camada": 1,
    "tipoDescricao": {
      "0": descricao
    }
  }
  
  ligacao.append(ligacao_item)
  return ligacao

def grava_json(no, ligacao, nome_arq):

  x = json.dumps(no, ensure_ascii=False)
  x2 = '{"no":'+ x +  '}'
  x3 = json.loads(x2)
  x3.update({"ligacao" : ligacao})
  x4 = json.dumps(x3)
  with open(nome_arq, "w") as f: 
    f.write(x4) 
  print ("Arquivo JSON (compatível com MACROS2) salvo: ", nome_arq)

