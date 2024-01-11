from login import loginDealernet
from sql import sqlPool
from mudarEmpresa import selecionarEmpresa
from rodape import preencheRodape
from tabulacaoImpostos import tabulaItens
from entradaDaNota import preenchimentoCapaNota
from conversaoXml import converteGzipParaXml
import warnings
import json
import subprocess
import datetime

print(datetime.datetime.now())

warnings.filterwarnings("ignore", category=UserWarning)
with open("../config/config.json", "r", encoding="utf-8") as file:
    sensitive_data = json.load(file)
    dealernetLogin = sensitive_data["acessoDealernet"]
    senha = dealernetLogin['senha']

    dealernetModulo = sensitive_data['modulosDealernet']['Estoque']
    executavel = dealernetModulo['executavel']

# loginDealernet(executavel, senha)
empresas = sqlPool("SELECT", """
                    SELECT 
                        emp_cd,
                        REPLACE(emp_ds, 'LUIS', 'LUÍS') AS emp_ds,
                        emp_banco
                    FROM [BD_MTZ_FOR]..ger_emp
                    WHERE 
                        emp_cd IN ('01')
                    ORDER BY emp_ds
                """)

for empresa in empresas:
    codEmpresa = empresa[0]
    empresa = empresa[1]
    notas = sqlPool('SELECT', f"SELECT * FROM DIA_Automate.nfemaster.DWIN_entradaNFeProdutoXML WHERE codigo_empresa = '{codEmpresa}'")

    if len(notas):
        # selecionarEmpresa(empresa)
        for nota in notas:
            dados = {
                'idNota': nota[0],
                'numeroNf': nota[1],
                'cnpj': nota[3],
                'gzip': nota[4],
                'codFornecedor': nota[5],
                'natureza': nota[6],
                'tipoDocumento': nota[7],
                'departamento': nota[8],
                'almoxarifado': nota[9]
            }
            
        nome_arquivo_xml = converteGzipParaXml(dados['gzip'], dados['numeroNf'])
        preenchimentoCapaNota(nome_arquivo_xml, dados['idNota'], dados['natureza'], dados['tipoDocumento'], dados['departamento'], dados['almoxarifado'])
        preencheRodape(dados['idNota'])
        tabulaItens(dados['idNota'])

print(datetime.datetime.now())

subprocess.run(["powershell", "-Command", "Stop-process -Name ead"], shell=True)

