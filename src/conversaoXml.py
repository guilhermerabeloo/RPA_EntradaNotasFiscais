import gzip
import base64
import json

def converteGzipParaXml(conteudoGzip, nomeArquivoXml):
    # extraindo conteudo de base64 para string
    conteudoBytes = base64.b64decode(conteudoGzip)
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
