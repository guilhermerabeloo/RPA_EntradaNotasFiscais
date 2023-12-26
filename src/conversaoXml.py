import gzip
import base64
import json

def converteGzipParaXml(caminhoArquivoGzip, nomeArquivoXml):
    # lendo arquivo gzip
    with gzip.open(caminhoArquivoGzip, 'rt') as arquivo:
        conteudoBase64 = arquivo.read()

    # extraindo conteudo de base64 para string
    conteudoBytes = base64.b64decode(conteudoBase64)
    conteudoDescomprimido = gzip.decompress(conteudoBytes)
    conteudoXml = conteudoDescomprimido.decode('utf-8')

    # criando arquivo XML
    with open("../config/config.json", "r", encoding="utf-8") as file:
        sensitive_data = json.load(file)
        pastaXml = sensitive_data["enderecoXml"]

    caminhoArquivoXml = f'{pastaXml}{nomeArquivoXml}.xml'

    with open(caminhoArquivoXml, 'w', encoding='utf-8') as arquivo:
        arquivo.write(conteudoXml)

    return caminhoArquivoXml
