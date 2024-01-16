from login import loginDealernet
from sql import sqlPool
from rodape import preencheRodape
from tabulacaoImpostos import tabulaItens
from entradaDaNota import preenchimentoCapaNota
from conversaoXml import converteGzipParaXml
import warnings
import json
import subprocess
import datetime
import logging
import time


warnings.filterwarnings("ignore", category=UserWarning)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('../log/logsDeErro.log'), 
    ]
)

print(datetime.datetime.now())
with open("../config/config.json", "r", encoding="utf-8") as file:
    sensitive_data = json.load(file)
    dealernetLogin = sensitive_data["acessoDealernet"]
    senha = dealernetLogin['senha']

    dealernetModulo = sensitive_data['modulosDealernet']['Estoque']
    executavel = dealernetModulo['executavel']

retornoNota = sqlPool('SELECT', f"""
                SELECT TOP 1
                    NF.* 
                    , REPLACE(E.emp_ds, 'LUIS', 'LUÍS') AS Empresa
                FROM [nfemaster].[DWIN_entradaNFeProdutoXML] AS NF
                inner join [BD_MTZ_FOR]..ger_emp AS E ON E.emp_cd = NF.codigo_empresa
                WHERE
                      --integrado = 'E'
                    integrado is null OR integrado NOT IN ('S', 'E')
                ORDER BY NF.data_insert                
                """)

if len(retornoNota):
    nota = retornoNota[0]
    codEmpresa = nota[2]
    empresa = nota[13]

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
        
    try: 
        loginDealernet(executavel, empresa, senha)
        nome_arquivo_xml = converteGzipParaXml(dados['gzip'], dados['numeroNf'])
        preenchimentoCapaNota(nome_arquivo_xml, dados['idNota'], dados['natureza'], dados['tipoDocumento'], dados['departamento'], dados['almoxarifado'])
        preencheRodape(dados['idNota'])
        tabulaItens(dados['idNota'])

        sqlPool("INSERT", f"""
                EXEC nfemaster.DWIN_insere_log_entradaNFe '{dados['idNota']}', 'S', '{codEmpresa}', '{dados['codFornecedor']}', '{dados['numeroNf']}', '1'
        """)
        
    except Exception as err:
        sqlPool("INSERT", f"""
                EXEC nfemaster.DWIN_insere_log_entradaNFe '{dados['idNota']}', 'E', '{codEmpresa}', '{dados['codFornecedor']}', '{dados['numeroNf']}', '0'
        """)

        logging.info(f'NOTA {dados['numeroNf']} - ERRO: {err}')
        subprocess.run(["powershell", "-Command", "Stop-process -Name ead"], shell=True)
        time.sleep(7)
        

else:
    logging.info(f'Nada para integrar')
    print('Nada para integrar')
    subprocess.run(["powershell", "-Command", "Stop-process -Name ead"], shell=True)

print(datetime.datetime.now())

subprocess.run(["powershell", "-Command", "Stop-process -Name ead"], shell=True)

