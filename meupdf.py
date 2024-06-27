import os
import fitz  # PyMuPDF
import pandas as pd
# 99
# Função para extrair informações de um PDF usando PyMuPDF
def extrair_informacoes_pdf(caminho_pdf):
    documento = fitz.open(caminho_pdf)
    conteudo = ''
    for pagina in documento:
        conteudo += pagina.get_text()

    # Dividir o conteúdo em linhas
    linhas = conteudo.split('\n')

    # Dicionário para armazenar dados temporários
    dados = {
        "NF-e": None,
        "DATA DE EMISSÃO": None,
        "MUNICÍPIO": None,
        "CEP": None,
        "UF": None,
        "VALOR TOTAL DOS PRODUTOS": None,
        "CÓD. PROD.": None,
        "DESCRIÇÃO": None,
        "QUANT": None,
        "UNIT.": None,
        "VALOR TOTAL DA NOTA": None,
    }

    encontrou_primeira = False

    for i, linha in enumerate(linhas):
        if "NF-e" in linha and not encontrou_primeira:
            dados["NF-e"] = linhas[i+1].strip()[3:]
            encontrou_primeira = True

        elif "DATA DE EMISSÃO" in linha:
            dados["DATA DE EMISSÃO"] = linhas[i+1].strip()

        elif "MUNICÍPIO" in linha and dados["MUNICÍPIO"] is None:
            if i + 1 < len(linhas):
                municipio_possivel = linhas[i+1].strip()
                # Adicionar verificação para evitar atribuição errada
                if len(municipio_possivel) <= 20:  # ajuste conforme necessário
                    dados["MUNICÍPIO"] = municipio_possivel
        
        elif "CEP" in linha:
            dados["CEP"] = linhas[i+1].strip()
        
        elif "UF" in linha and dados["UF"] is None:
            dados["UF"] = linhas[i+1].strip()
        
        elif "VALOR TOTAL DOS PRODUTOS" in linha:
            dados["VALOR TOTAL DOS PRODUTOS"] = linhas[i+1].strip()
        
        elif "CÓD. PROD." in linha and dados["CÓD. PROD."] is None:
            dados["CÓD. PROD."] = linhas[i+21].strip()
        
        elif "DESCRIÇÃO" in linha:
            descricao = dados["DESCRIÇÃO"] = linhas[i+21].strip()
            j = i + 1
            descricao = ''
            j = i + 21
            if len(linhas) < 141:
                descricao = linhas[i+21].strip()   # Apenas uma linha de descrição
            else:
                while j < len(linhas):
                    descricao += linhas[j].strip()+ ', '
                    j += 13
                    print(j)
                    if j > len(linhas):
                        break  # Encerra a captura se encontrar uma nova chave
                     # Adiciona a linha e um espaço
            dados["DESCRIÇÃO"] = descricao.strip()  # Remove espaços extras do início e fim
        
        elif "QUANT" in linha:
            dados["QUANT"] = linhas[i+20].strip()
        
        elif "UNIT." in linha and dados["UNIT."] is None:
            dados["UNIT."] = linhas[i+20].strip()
        
        elif "VALOR TOTAL DA NOTA" in linha:
            dados["VALOR TOTAL DA NOTA"] = linhas[i+1].strip()

    # Adiciona os dados na lista de informações
    return dados

# Lista para armazenar os dados de todas as notas
dados_notas = []

# Diretório onde os PDFs estão armazenados
diretorio_pdfs = 'pdf'

# Contador de arquivos lidos
arquivos_lidos = 0

# Iterar sobre todos os arquivos no diretório
for nome_arquivo in os.listdir(diretorio_pdfs):
    if nome_arquivo.endswith('.pdf'):
        caminho_pdf = os.path.join(diretorio_pdfs, nome_arquivo)
        informacoes_pdf = extrair_informacoes_pdf(caminho_pdf)
        dados_notas.append(informacoes_pdf)
        arquivos_lidos += 1

# Criar um DataFrame do pandas com os dados das notas
df = pd.DataFrame(dados_notas)

# Definir a ordem das colunas
colunas = [
    "NF-e", "DATA DE EMISSÃO", "MUNICÍPIO", "CEP", "UF",
    "VALOR TOTAL DOS PRODUTOS", "CÓD. PROD.",
    "DESCRIÇÃO", "QUANT", "UNIT.", "VALOR TOTAL DA NOTA"
]

# Ajustar DataFrame para garantir a ordem das colunas
df = df[colunas]

# Salvar o DataFrame em um arquivo Excel
caminho_excel = 'notas_fiscais.xlsx'
df.to_excel(caminho_excel, index=False)

# Exibir uma mensagem com o número de arquivos lidos
caminho_absoluto = os.path.abspath(caminho_excel)
print(f"{arquivos_lidos} arquivo(s) lido(s) e salvo(s) em {caminho_absoluto}.")
