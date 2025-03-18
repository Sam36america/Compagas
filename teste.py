import os
import pandas as pd
import shutil
import xml.etree.ElementTree as ET

DIST = 'Cegás'
NAMESPACE = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

def extrair_informacoes_xml(caminho_do_xml):
    informacoes = {}
    try:
        tree = ET.parse(caminho_do_xml)
        root = tree.getroot()
        
        # Extrair informações básicas do XML
        informacoes['cnpj'] = root.find('.//nfe:emit/nfe:CNPJ', NAMESPACE).text
        informacoes['valor_total'] = root.find('.//nfe:total/nfe:ICMSTot/nfe:vNF', NAMESPACE).text
        informacoes['volume_total'] = root.find('.//nfe:det/nfe:prod/nfe:qCom', NAMESPACE).text
        informacoes['data_emissao'] = root.find('.//nfe:ide/nfe:dhEmi', NAMESPACE).text.split('T')[0]
        
        inf_cpl_element = root.find('.//nfe:infAdic/nfe:infCpl', NAMESPACE)
        inf_cpl = inf_cpl_element.text if inf_cpl_element is not None else ''
        informacoes['data_inicio'] = inf_cpl.split(' ')[2] if inf_cpl else ''
        informacoes['data_fim'] = inf_cpl.split(' ')[4] if inf_cpl else ''
        
        informacoes['numero_fatura'] = root.find('.//nfe:ide/nfe:nNF', NAMESPACE).text
        informacoes['valor_icms'] = root.find('.//nfe:total/nfe:ICMSTot/nfe:vICMS', NAMESPACE).text

        # Buscar PCS em várias localizações possíveis
        pcs = None
        
        # 1. Tentar encontrar na seção de combustíveis
        comb = root.find('.//nfe:det/nfe:prod/nfe:comb', NAMESPACE)
        if comb is not None:
            for elem in comb.iter():
                if 'PCS' in elem.tag or (elem.text and 'PCS' in elem.text):
                    pcs = elem.text
                    break
        
        # 2. Procurar nas informações adicionais do produto
        if not pcs:
            det = root.find('.//nfe:det', NAMESPACE)
            if det is not None:
                for elem in det.iter():
                    if 'PCS' in elem.tag or (elem.text and 'PCS' in elem.text):
                        pcs = elem.text
                        break
        
        informacoes['correcao_pcs'] = pcs if pcs else ''

        return informacoes
    except Exception as e:
        print(f"Erro ao extrair informações do XML: {caminho_do_xml}, erro: {e}")
        return {}

def adicionar_na_planilha(informacoes, caminho_planilha, nome_arquivo):
    try:
        df = pd.read_excel(caminho_planilha)
    except FileNotFoundError:
        print(f"O arquivo '{caminho_planilha}' não foi encontrado. Criando um novo.")
        df = pd.DataFrame(columns=['CNPJ', 'Valor Total', 'Volume Total', 'Data Emissão', 'Data Início', 'Data Fim', 'Número Fatura', 'Valor ICMS', 'Correção PCS', 'Distribuidora', 'Nome do Arquivo'])

    cnpj = informacoes['cnpj']
    valor_total = pd.to_numeric(informacoes['valor_total'].replace('.', '').replace(',', '.'))
    volume_total = pd.to_numeric(informacoes['volume_total'].replace('.', '').replace(',', '.'))
    data_inicio = informacoes['data_inicio']
    data_fim = informacoes['data_fim']
    valor_icms = pd.to_numeric(informacoes['valor_icms'].replace('.', '').replace(',', '.'))
    correcao_pcs = informacoes['correcao_pcs']

    nova_linha = pd.DataFrame([{
        'CNPJ': cnpj,
        'Valor Total': valor_total,
        'Volume Total': volume_total,
        'Data Emissão': informacoes.get('data_emissao', ''),
        'Data Início': data_inicio,
        'Data Fim': data_fim,
        'Número Fatura': informacoes.get('numero_fatura', ''),
        'Valor ICMS': valor_icms,
        'Correção PCS': correcao_pcs,
        'Distribuidora': DIST,
        'Nome do Arquivo': nome_arquivo
    }])
    
    df = pd.concat([df, nova_linha], ignore_index=True)
    df.to_excel(caminho_planilha, index=False)
    return True

def mover_arquivo(origem, destino):
    shutil.move(origem, destino)
    print(f"Arquivo movido para {destino}")

def processar_xml(file, caminho_planilha, diretorio_destino):
    informacoes = extrair_informacoes_xml(file)
    
    # Verifica se todos os campos obrigatórios foram extraídos
    campos_obrigatorios = ['cnpj', 'valor_total', 'volume_total', 'data_emissao', 
                           'data_inicio', 'data_fim', 'numero_fatura', 'valor_icms']
    campos_faltantes = [campo for campo in campos_obrigatorios if not informacoes.get(campo)]
    
    if campos_faltantes:
        print(f"Campos faltantes no arquivo {file}: {', '.join(campos_faltantes)}")
        return

    nome_arquivo = os.path.basename(file)
    inserido = adicionar_na_planilha(informacoes, caminho_planilha, nome_arquivo)
    print(informacoes)

    if inserido:
        destino = os.path.join(diretorio_destino, nome_arquivo)
        mover_arquivo(file, destino)
    else:
        print('Arquivo não foi inserido na planilha devido a dados faltantes ou duplicados. Não será movido.')

# Exemplo de uso
file_path = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Compagás\Faturas'
diretorio_destino = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Compagás\Lidos'
caminho_planilha = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\00 Faturas Lidas\COMPAGÁS.xlsx'

for arquivo in os.listdir(file_path):
    if arquivo.lower().endswith('.xml'):
        arquivo_full = os.path.join(file_path, arquivo)
        processar_xml(arquivo_full, caminho_planilha, diretorio_destino)