import requests

def criaInstancia(empresa, departamento, idNota):
    try:
        url = 'http://integrador.grupofornecedora.com.br:8975/api/v2/DWIN_lancamentoNFZeev'

        body = {
            "cod_setor": f"{departamento}",
            "cod_emp": f"{empresa}",
            "sistema_id": 1,
            "id_nota": idNota
        }
        
        response = requests.post(url, json=body)
        if response.status_code != 200:
            raise Exception(f"Erro na requisição: {response.text}")
    except Exception as err:
        print(err)

