import os
import json
import xmltodict
import openpyxl

opens = openpyxl.Workbook()
opens.create_sheet('notasfiscais')
sheet_notas = opens['notasfiscais']
sheet_notas['A1'].value = 'Numero da nota'
sheet_notas['B1'].value = 'Empresa Emissora'
sheet_notas['C1'].value = "Nome do cliente"
sheet_notas['D1'].value = "Peso"

def pegar_dados(arquivos):
    print(f"esse são os arquivos {arquivos}")
    with open(f"nfs/{arquivos}", "rb") as arquivo_xml:
        try:
            dic_arquivo = xmltodict.parse(arquivo_xml)
            if 'NFe' in dic_arquivo:
                infos_nf = dic_arquivo['NFe']['infNFe']
            else:
                infos_nf = dic_arquivo['nfeProc']['NFe']['infNFe']
            numero_nota = infos_nf['@Id']
            empresa_emissora = infos_nf['emit']['xNome']
            nome_cliente = infos_nf['dest']['xNome']
            if "vol" in infos_nf['transp']:
                pesos = infos_nf['transp']['vol']['pesoB']
            else:
                pesos = "Não informado"
            sheet_notas.append([numero_nota, empresa_emissora,nome_cliente,pesos])
            opens.save('notas.xlsx')
            
        except Exception as e:
            print(e)
            #print(json.dumps(dic_arquivo, indent=4))

lista_xml = os.listdir('nfs')
for arquivo in lista_xml:
    pegar_dados(arquivo)