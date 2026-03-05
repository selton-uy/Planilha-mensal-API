import requests
import time
import random

session = requests.Session()

headers_ejud = {
    "Content-Type": "application/json; charset=UTF-8",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": "https://www3.tjrj.jus.br",
    "Referer": "https://www3.tjrj.jus.br/ejud/ConsultaProcesso.aspx"
}
headers_processo = {
    "Content-Type": "application/json",
    "Accept": "application/json, text/plain, */*",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Referer": "https://www3.tjrj.jus.br/consultaprocessual/",
    "Origin": "https://www3.tjrj.jus.br"
}


def consultar_processo_por_numero(numero_cnj):
    url_base = "https://www3.tjrj.jus.br/consultaprocessual/api/processos/por-numeracao-unica"
    payload_base = {"tipoProcesso": "1", "codigoProcesso": numero_cnj}

    try:
        # Pausa aleatória entre 1.5 e 3 segundos antes de cada requisição
        time.sleep(random.uniform(1.5, 3.0))

        res_base = session.post(url_base, json=payload_base, headers=headers_ejud)
        dados_lista = res_base.json()

        if isinstance(dados_lista, dict) and "Message" in dados_lista:
            return None

        n_antigo = None
        if isinstance(dados_lista, list):
            for item in dados_lista:
                if 'numProcesso' in item and '.' in item['numProcesso']:
                    n_antigo = item['numProcesso']
                    break

        if not n_antigo:
            return None

        url_ws = "https://www3.tjrj.jus.br/ejud/WS/ConsultaEjud.asmx/DadosProcesso_1"
        payload_ws = {"nAntigo": n_antigo, "pCPF": "", "pLogin": ""}

        res_final = session.post(url_ws, json=payload_ws, headers=headers_ejud)
        return res_final.json().get('d', res_final.json())

    except Exception:
        return None


def consultar_processo_por_nome_e_data(nome, ano_inicio, ano_fim):
    url_base = "https://www3.tjrj.jus.br/consultaprocessual/api/processos/por-nome-parte"

    payload = {
        "anoFinal": int(ano_fim),
        "anoInicial": int(ano_inicio),
        "isPortal": "S",
        "isPortalDeServicos": 1,
        "nome": nome.upper(),
        "origem": "2",
        "pIsProcAtivos": "N",
        "radio": "1",
        "secao": "RJ",
        "tipoConsulta": "publica",
        "tipoSegundaInstancia": "1",
        "totalProcessoPesquisa": 300,
        "validarSecao": True
    }

    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Referer": "https://www3.tjrj.jus.br/consultaprocessual/",
        "Origin": "https://www3.tjrj.jus.br"
    }

    try:
        response = session.post(url_base, json=payload, headers=headers)

        if response.status_code == 200:
            dados = response.json()
            return dados
        else:
            print(f"Erro {response.status_code}: {response.text}")
            return None
    except Exception as e:
        print(f"Erro na requisição: {e}")
        return None

def extrair_dados(dados):
    if not dados or not isinstance(dados, dict):
        return {'RECORRENTE': '', 'TURMA': '', 'RELATOR': '', 'SUMULA': ''}

    try:
        info = dados.get('d', dados)

        nome_recorrente = (
                info.get('Autor') or
                info.get('APELANTE') or
                info.get('AGTE') or
                ''
        )

        decisao = info.get('DadosJulgamento')
        texto_decisao = decisao[-1].get('Txt', '') if (decisao and isinstance(decisao, list)) else ''

        return {
            'RECORRENTE': nome_recorrente,
            'TURMA': info.get('OrgaoJulgador', ''),
            'RELATOR': info.get('Relator', ''),
            'SUMULA': texto_decisao
        }
    except Exception:
        return {'RECORRENTE': '', 'TURMA': '', 'RELATOR': '', 'SUMULA': ''}