import gzip
import base64

def converteGzipParaXml(caminhoArquivoGzip):
    # lendo arquivo gzip
    with gzip.open(caminhoArquivoGzip, 'rt') as arquivo:
        conteudoBase64 = arquivo.read()

    # extraindo conteudo de base64 para string
    conteudoBytes = base64.b64decode(conteudoBase64)
    conteudoDescomprimido = gzip.decompress(conteudoBytes)
    conteudoXml = conteudoDescomprimido.decode('utf-8')

    # criando arquivo XML
    with open('arquivo2.xml', 'w', encoding='utf-8') as arquivo:
        arquivo.write(conteudoXml)


nome_arquivo_gzip = 'C:\\Users\\guilherme.rabelo\\Documents\\RPA_docs\\EntradaDaNota\\newGzip.txt.gz'
converteGzipParaXml(nome_arquivo_gzip)