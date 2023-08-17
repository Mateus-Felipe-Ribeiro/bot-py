import xmltodict
import os
import json
import pandas as pd

# Função para ler arquivos da pasta
def pegar_infos(arquivo, valores):
    with open(f'NF/{arquivo}', "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        try:
            if "NFe" in dic_arquivo:
                infos_nf = dic_arquivo["NFe"]["infNFe"]
            else:
                infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
            numero_nota = infos_nf["@Id"]
            empresa_emissora = infos_nf["emit"]["xNome"]
            nome_cliente = infos_nf["dest"]["xNome"]
            endereco = infos_nf["dest"]["enderDest"]
            if "vol" in infos_nf["transp"]:
                peso_bruto = infos_nf["transp"]["vol"]["pesoB"]
            else:
                peso_bruto = 0

            valores.append([numero_nota, empresa_emissora, nome_cliente, endereco, peso_bruto])
        except Exception as e:
            print(e)
            # print usado para verificar qual a estrutura do xml
            print(json.dumps(dic_arquivo, indent=4))

# Pega a pasta onde localiza-se os arquivos
arquivos = os.listdir("NF")

colunas = ["numero_nota", "empresa_emissora", "nome_cliente", "endereco", "peso_bruto"]
valores = []

# Leitura de cada arquivo da lista
for arquivo in arquivos:
    pegar_infos(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel("NotasFiscais.xlsx", index=False)