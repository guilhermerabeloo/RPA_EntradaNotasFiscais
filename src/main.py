from login import loginDealernet
from sql import sqlPool
from mudarEmpresa import selecionarEmpresa
from entradaDaNota import preenchimentoCapaNota
from conversaoXml import converteGzipParaXml
import warnings
import json
import subprocess
import time

warnings.filterwarnings("ignore", category=UserWarning)
with open("../config/config.json", "r", encoding="utf-8") as file:
    sensitive_data = json.load(file)
    dealernetLogin = sensitive_data["acessoDealernet"]
    senha = dealernetLogin['senha']

    dealernetModulo = sensitive_data['modulosDealernet']['Estoque']
    executavel = dealernetModulo['executavel']

loginDealernet(executavel, senha)
selecionarEmpresa('FME BOM JESUS - 0018/07')

nome_arquivo_gzip = 'C:\\Users\\guilherme.rabelo\\Documents\\RPA_docs\\EntradaDaNota\\newGzip.txt.gz'
nome_arquivo_xml = converteGzipParaXml(nome_arquivo_gzip, 'arquivo')

preenchimentoCapaNota(nome_arquivo_xml)
