from login import loginDealernet
from sql import sqlPool
from rodape import preencheRodape
from tabulacaoImpostos import tabulaItens
from entradaDaNota import preenchimentoCapaNota
from conversaoXml import converteGzipParaXml
import warnings
import json
import subprocess
import logging
import time


warnings.filterwarnings("ignore", category=UserWarning)

execucao_handler = logging.FileHandler('../log/logsDeExecucao.log')
execucao_handler.setLevel(logging.INFO)
execucao_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
erro_handler = logging.FileHandler('../log/logsDeErro.log')
erro_handler.setLevel(logging.ERROR)
erro_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))

logging.basicConfig(
    level=logging.DEBUG,
    handlers=[execucao_handler, erro_handler]
)
logger = logging.getLogger(__name__)

with open("../config/config.json", "r", encoding="utf-8") as file:
    sensitive_data = json.load(file)
    dealernetLogin = sensitive_data["acessoDealernet"]
    senha = dealernetLogin['senha']

    dealernetModulo = sensitive_data['modulosDealernet']['Estoque']
    executavel = dealernetModulo['executavel']

logger.info("=-=-=-=-=-=-=-=-INICIO DA EXECUCAO=-=-=-=-=-=-=-=-")
retornoNota = sqlPool('SELECT', f"""
                SELECT TOP 1
                    NF.* 
                    , REPLACE(E.emp_ds, 'LUIS', 'LUÍS') AS Empresa
                FROM [nfemaster].[DWIN_entradaNFeProdutoXML] AS NF
                inner join [BD_MTZ_FOR]..ger_emp AS E ON E.emp_cd = NF.codigo_empresa
                WHERE
                    integrado = 'P'
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
        logger.info(f'ID {dados["idNota"]} - Realizando login no Dealernet')
        loginDealernet(executavel, empresa, senha)

        logger.info(f'ID {dados["idNota"]} - Realizando conversao de XML')
        nome_arquivo_xml = converteGzipParaXml(dados['gzip'], dados['numeroNf'])

        logger.info(f'ID {dados["idNota"]} - Realizando preenchimento da capa da nota')
        preenchimentoCapaNota(nome_arquivo_xml, dados['idNota'], dados['natureza'], dados['tipoDocumento'], dados['departamento'], dados['almoxarifado'])
        
        logger.info(f'ID {dados["idNota"]} - Realizando preenchimento do rodape')
        preencheRodape(dados['idNota'])

        logger.info(f'ID {dados["idNota"]} - Realizando tabulacao dos itens')
        tabulaItens(dados['idNota'])

        sqlPool("INSERT", f"""
                EXEC nfemaster.DWIN_insere_log_entradaNFe '{dados['idNota']}', 'S', '{codEmpresa}', '{dados['codFornecedor']}', '{dados['numeroNf']}', '1'
        """)

        logger.info("=-=-=-=-=-=-=-=-FIM DA EXECUCAO=-=-=-=-=-=-=-=-\n")
        
    except Exception as err:
        sqlPool("INSERT", f"""
                EXEC nfemaster.DWIN_insere_log_entradaNFe '{dados['idNota']}', 'E', '{codEmpresa}', '{dados['codFornecedor']}', '{dados['numeroNf']}', '0'
        """)

        logging.error(f'NOTA {dados['numeroNf']} - ERRO: {err}')
        subprocess.run(["powershell", "-Command", "Stop-process -Name ead"], shell=True)
        time.sleep(7)
        
        logger.info("=-=-=-=-=-=-=-=-FIM DA EXECUCAO=-=-=-=-=-=-=-=-\n")

else:
    logging.info(f'Nada para integrar')
    logger.info("=-=-=-=-=-=-=-=-FIM DA EXECUCAO=-=-=-=-=-=-=-=-\n")
    subprocess.run(["powershell", "-Command", "Stop-process -Name ead"], shell=True)

subprocess.run(["powershell", "-Command", "Stop-process -Name ead"], shell=True)

