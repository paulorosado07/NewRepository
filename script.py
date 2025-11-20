import sys
import os
import re
import logging
import unicodedata
from openpyxl import load_workbook
import pandas as pd


REPLACEMENTS = {
    r"\bAV\.(?=[ ,])": "AVENIDA",
    r"\bLIC\.(?=[ ,])": "LICENCIADO",
    r"\bSTA\.(?=[ ,])": "SANTA",
    r"\bNO\.(?=[ ,])": "NUMERO",
    r"\bURB\.(?=[ ,])": "URBANIZACION",
    r"\bGRAL\.(?=[ ,])": "GENERAL",
    r"\bAV\b": "AVENIDA",
    r"\bC(?=(\s|,|$))": "CALLE",
    r"\bCALZ\b": "CALZADA",
    r"\bCARR\b": "CARRETERA",
    r"\bCAM\b": "CAMINO",
    r"\bCDA\b": "CERRADA",
    r"\bGOB\b": "GOBERNADOR",
    r"\bCDAD\.?\b": "CIUDAD",
    r"\bNU#EZ\b": "NUNEZ",
    r"\bJ\.\s*ENRIQUE\s*PESTALOZZI\b": "JUAN ENRIQUE PESTALOZZI",
    r"\bJ\s*ENRIQUE\s*PESTALOZZI\b": "JUAN ENRIQUE PESTALOZZI",
    r"\bSTA\b": "SANTA",
    r"\bSECC\b": "SECCION",
    r"\b1RA\b": "PRIMERA",
    r"\bLIC\b": "LICENCIADO",
}

TIPOS = [
    "AVENIDA",
    "CALLE",
    "CALZADA",
    "CAMINO",
    "CERRADA",
    "CARRETERA",
    "PASEO",
    "BOULEVARD",
    "BLVD",
]

STATE_MAP = {
  "CDMX": "CDMX.",
  "SIN": "SIN.",
  "MEX": "MÉX.",
  "NL":"N.L.",
  "GTO":"GTO.",
  "PUE": "PUE.",
  "SON": "SON.",
  "BC":"B.C",
  "AGS": "AGS.",
  "YUC": "YUC.",
  "VER": "VER.",
  "TAMPS": "TAMPS.",
  "HGO": "HGO.",
  "QR": "Q.R.",
  "COAH": "COAH.",
  "MICH": "MICH.",
  "MEXICO": "MÉX.",
  "OAX": "OAX.",
  
}


def configurar_logger(path_log="log_enderecos.txt"):
    logger = logging.getLogger("enderecos")
    logger.setLevel(logging.INFO)

    file_handler = logging.FileHandler(path_log, encoding="utf-8")
    file_handler.setLevel(logging.INFO)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    if not logger.handlers:
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

    return logger


def remover_acentos(texto):
    if not isinstance(texto, str):
        return texto
    return "".join(
        caracter
        for caracter in unicodedata.normalize("NFD", texto)
        if unicodedata.category(caracter) != "Mn"
    )


def normalizar_endereco(endereco_bruto):
    
    if not isinstance(endereco_bruto, str):
        return ""
    
    endereco_normalizado = endereco_bruto.strip()
    endereco_normalizado = re.sub(r"\s+", " ", endereco_normalizado)

    endereco_normalizado = endereco_normalizado.upper()
    endereco_normalizado = re.sub(r"\s+", " ", endereco_normalizado)

    for padrao_regex, substituicao in REPLACEMENTS.items():
        endereco_normalizado = re.sub(padrao_regex, substituicao, endereco_normalizado)

    endereco_normalizado = re.sub(r"\s+", " ", endereco_normalizado).strip()

    texto_anterior = None
    while texto_anterior != endereco_normalizado:
        texto_anterior = endereco_normalizado
        endereco_normalizado = re.sub(r",\s*,", ",", endereco_normalizado)

    return endereco_normalizado


def extrair_geo(endereco_normalizado):
    
    if not endereco_normalizado:
        return endereco_normalizado, None, None
    
    match_parenteses = re.search(r"\(([^()]*)\)\s*$", endereco_normalizado)

    if not match_parenteses:
        return endereco_normalizado.strip(), None, None
    
    conteudo_parenteses = match_parenteses.group(1)
    numeros_encontrados = re.findall(r"-?\d+,\d+", conteudo_parenteses)

    latitude = None
    longitude = None
    
    if len(numeros_encontrados) >= 2:
        latitude = (numeros_encontrados[0].replace(",", "."))
        longitude = (numeros_encontrados[1].replace(",", "."))

    endereco_sem_geo = endereco_normalizado[: match_parenteses.start()].strip().rstrip(",")

    return endereco_sem_geo, latitude, longitude


def limpar_pontos(sigla):
    return re.sub(r"[^\w]", "", sigla)


def mapear_estado(token_estado):
    if token_estado is None:
        return None

    token_estado_normalizado = token_estado.strip().upper()
    token_estado_normalizado = remover_acentos(token_estado_normalizado)
    sigla_sem_pontos = limpar_pontos(token_estado_normalizado)

    if not sigla_sem_pontos:
        return None

    return STATE_MAP.get(sigla_sem_pontos, sigla_sem_pontos)


def parse_canonical_address(endereco_canonico):
    
    resultado_parse = {
        "TIPO": None,
        "LOGRADOURO": None,
        "NUMERO": None,
        "COMPLEMENTO": None,
        "BAIRRO": None,
        "CIDADE": None,
        "ESTADO": None,
        "PAIS": "MÉXICO",
    }

    if not endereco_canonico:
        return resultado_parse

    partes_endereco = [parte.strip() for parte in endereco_canonico.split(",") if parte.strip()]
    primeira_parte = partes_endereco[0] if partes_endereco else endereco_canonico.strip()

    tokens_primeira_parte = primeira_parte.split()
    tipo_logradouro = None
    indice_inicio_logradouro = 0

    if tokens_primeira_parte and tokens_primeira_parte[0] in TIPOS:
        tipo_logradouro = tokens_primeira_parte[0]
        indice_inicio_logradouro = 1

    tokens_logradouro = tokens_primeira_parte[indice_inicio_logradouro:]
    indice_numero = None

    for indice_token, token in enumerate(tokens_logradouro):
        if re.fullmatch(r"\d+(-\d+)?|S/N", token):
            indice_numero = indice_token

    tokens_nome_logradouro = []
    numero_endereco = None
    tokens_complemento = []

    if indice_numero is None:
        tokens_nome_logradouro = tokens_logradouro
    else:
        tokens_nome_logradouro = tokens_logradouro[:indice_numero]
        numero_endereco = tokens_logradouro[indice_numero]
        tokens_complemento = tokens_logradouro[indice_numero + 1 :]

    resultado_parse["TIPO"] = tipo_logradouro
    resultado_parse["LOGRADOURO"] = " ".join(tokens_nome_logradouro) if tokens_nome_logradouro else None
    resultado_parse["NUMERO"] = numero_endereco
    resultado_parse["COMPLEMENTO"] = " ".join(tokens_complemento) if tokens_complemento else None

    # BAIRRO
    if len(partes_endereco) >= 2:
        bairro_parte = partes_endereco[1]
        if isinstance(bairro_parte, str):
            resultado_parse["BAIRRO"] = re.sub(r"\d{5}", "", bairro_parte).strip()
        else:
            resultado_parse["BAIRRO"] = bairro_parte

    # CIDADE / ESTADO / PAIS
    if len(partes_endereco) >= 3:
        partes_cidade_estado = partes_endereco[2:]
        indice_pais = None

        for indice in reversed(range(len(partes_cidade_estado))):
            if "MEXICO" in partes_cidade_estado[indice]:
                indice_pais = indice
                break

        if indice_pais is not None:
            partes_cidade_estado = partes_cidade_estado[:indice_pais]

        if partes_cidade_estado:
            token_estado = partes_cidade_estado[-1]
            resultado_parse["ESTADO"] = mapear_estado(token_estado)

            partes_cidade = partes_cidade_estado[:-1]
            if partes_cidade:
                cidade_bruta = " ".join(partes_cidade)
                match_cep_cidade = re.match(r"\d+\s+(.*)", cidade_bruta)
                if match_cep_cidade:
                    cidade_bruta = match_cep_cidade.group(1)
                cidade_index = cidade_bruta.find("CIUDAD")
                cidade_modificada = ( cidade_bruta[cidade_index:] if cidade_index != -1 else cidade_bruta ).replace("CIUDAD.", "CIUDAD ").strip()
                resultado_parse["CIDADE"] = cidade_modificada

    return resultado_parse


def process_dataframe(df_enderecos, logger=None):
    """
    Fluxo:
    - Normalizar endereços
    - Extrair geo (LAT/LON) sem deixar coordenadas no endereço
    - Deduplicar usando regra: endereco_completo.find(endereco_incompleto) != -1
    - Preencher ENDEREÇO CANÔNICO, LATITUDE, LONGITUDE
    - Fazer parsing em TIPO, LOGRADOURO, etc.
    """
    lista_normalizado = []
    lista_endereco_sem_geo = []
    lista_latitude = []
    lista_longitude = []

    for _, endereco_original in df_enderecos["ENDEREÇO ORIGINAL"].items():
        endereco_normalizado = normalizar_endereco(endereco_original)
        endereco_sem_geo, latitude, longitude = extrair_geo(endereco_normalizado)

        lista_normalizado.append(endereco_normalizado)
        lista_endereco_sem_geo.append(endereco_sem_geo)
        lista_latitude.append(latitude)
        lista_longitude.append(longitude)

    df_tratado = df_enderecos.copy()
    df_tratado["_NORMALIZADO"] = lista_normalizado
    df_tratado["_ADDR"] = lista_endereco_sem_geo
    df_tratado["_LAT_TMP"] = lista_latitude
    df_tratado["_LON_TMP"] = lista_longitude

    quantidade_linhas = len(df_tratado)

    lista_endereco_canonico = [None] * quantidade_linhas
    lista_latitude_final = [None] * quantidade_linhas
    lista_longitude_final = [None] * quantidade_linhas

    # Nova deduplicação baseada em "find"
    for indice_linha in range(quantidade_linhas):
        endereco_base = df_tratado["_ADDR"].iat[indice_linha]
        latitude_base = df_tratado["_LAT_TMP"].iat[indice_linha]
        longitude_base = df_tratado["_LON_TMP"].iat[indice_linha]

        if not isinstance(endereco_base, str) or not endereco_base.strip():
            lista_endereco_canonico[indice_linha] = endereco_base
            lista_latitude_final[indice_linha] = latitude_base
            lista_longitude_final[indice_linha] = longitude_base
            continue

        endereco_mais_completo = endereco_base
        latitude_mais_completa = latitude_base
        longitude_mais_completa = longitude_base

        for indice_outro in range(quantidade_linhas):
            if indice_outro == indice_linha:
                continue

            endereco_comparacao = df_tratado["_ADDR"].iat[indice_outro]
            latitude_comparacao = df_tratado["_LAT_TMP"].iat[indice_outro]
            longitude_comparacao = df_tratado["_LON_TMP"].iat[indice_outro]

            if not isinstance(endereco_comparacao, str) or not endereco_comparacao.strip():
                continue

            # Se "endereco_comparacao" contém "endereco_base",
            # então "endereco_comparacao" é mais completo para "endereco_base"
            if endereco_comparacao.find(endereco_base) != -1:
                if len(endereco_comparacao) > len(endereco_mais_completo or ""):
                    # Escolhe endereço mais completo
                    if (
                        logger
                        and endereco_mais_completo is not None
                        and len(endereco_comparacao) == len(endereco_mais_completo)
                    ):
                        # Mesmo tamanho mas textos diferentes => ambíguo
                        logger.warning(
                            f"Deduplicacao ambigua entre '{endereco_mais_completo}' "
                            f"e '{endereco_comparacao}' para linha {indice_linha}"
                        )

                    endereco_mais_completo = endereco_comparacao
                    latitude_mais_completa = latitude_comparacao
                    longitude_mais_completa = longitude_comparacao

                elif (
                    len(endereco_comparacao) == len(endereco_mais_completo or "")
                    and endereco_comparacao != endereco_mais_completo
                ):
                    # Mesmo tamanho, textos diferentes e ambos contém o endereco_base:
                    # considerar ambíguo
                    if logger:
                        logger.warning(
                            f"Deduplicacao ambigua para endereco '{endereco_base}' "
                            f"entre '{endereco_mais_completo}' e '{endereco_comparacao}' "
                            f"(linhas relacionadas: {indice_linha})"
                        )

        lista_endereco_canonico[indice_linha] = endereco_mais_completo
        lista_latitude_final[indice_linha] = latitude_mais_completa
        lista_longitude_final[indice_linha] = longitude_mais_completa

    df_tratado["ENDEREÇO CANÔNICO"] = lista_endereco_canonico
    df_tratado["LATITUDE"] = lista_latitude_final
    df_tratado["LONGITUDE"] = lista_longitude_final

    lista_tipos = []
    lista_logradouros = []
    lista_numeros = []
    lista_complementos = []
    lista_bairros = []
    lista_cidades = []
    lista_estados = []
    lista_paises = []

    for endereco_canonico in df_tratado["ENDEREÇO CANÔNICO"]:
        resultado_parse = parse_canonical_address(endereco_canonico)

        lista_tipos.append(resultado_parse["TIPO"])
        lista_logradouros.append(resultado_parse["LOGRADOURO"])
        lista_numeros.append(resultado_parse["NUMERO"])
        lista_complementos.append(resultado_parse["COMPLEMENTO"])
        lista_bairros.append(resultado_parse["BAIRRO"])
        lista_cidades.append(resultado_parse["CIDADE"])
        lista_estados.append(resultado_parse["ESTADO"])
        lista_paises.append(resultado_parse["PAIS"])

    df_tratado["TIPO"] = lista_tipos
    df_tratado["LOGRADOURO"] = lista_logradouros
    df_tratado["NÚMERO "] = lista_numeros
    df_tratado["COMPLEMENTO"] = lista_complementos
    df_tratado["BAIRRO"] = lista_bairros
    df_tratado["CIDADE"] = lista_cidades
    df_tratado["ESTADO "] = lista_estados
    df_tratado["PAIS"] = lista_paises

    df_tratado.drop(
        columns=["_NORMALIZADO", "_ADDR", "_LAT_TMP", "_LON_TMP"],
        inplace=True,
    )

    return df_tratado


def carregar_planilha(caminho_arquivo, logger):
    if not os.path.exists(caminho_arquivo):
        logger.error(f"Arquivo '{caminho_arquivo}' nao encontrado.")
        raise FileNotFoundError(caminho_arquivo)
    try:
        dataframe = pd.read_excel(caminho_arquivo, sheet_name="POC")
    except ValueError:
        logger.error("Planilha 'POC' nao encontrada no arquivo.")
        raise
    return dataframe


def salvar_planilha(df_tratado, caminho_saida, logger):
    try:
        df_tratado.to_excel(caminho_saida, sheet_name="POC", index=False)

        workbook = load_workbook(caminho_saida)
        worksheet = workbook.active

        # Ajusta a largura das colunas
        worksheet.column_dimensions["A"].width = 123
        worksheet.column_dimensions["B"].width = 115
        worksheet.column_dimensions["C"].width = 10
        worksheet.column_dimensions["D"].width = 45
        worksheet.column_dimensions["E"].width = 10
        worksheet.column_dimensions["F"].width = 13
        worksheet.column_dimensions["G"].width = 38
        worksheet.column_dimensions["H"].width = 45
        worksheet.column_dimensions["I"].width = 10
        worksheet.column_dimensions["J"].width = 9
        worksheet.column_dimensions["K"].width = 12
        worksheet.column_dimensions["L"].width = 12

        workbook.save(caminho_saida)

        logger.info(f"Arquivo de saida gerado em '{caminho_saida}'.")
    except Exception as erro:
        logger.error(f"Erro ao salvar arquivo de saida: {erro}")
        raise


def run_tests():
    # Teste básico de normalização e replaces
    endereco_teste_1 = "Av 5 de Mayo 17"
    endereco_normalizado_1 = normalizar_endereco(endereco_teste_1)
    assert endereco_normalizado_1 == "AVENIDA 5 DE MAYO 17"

    endereco_teste_2 = "C Ignacio Ramirez 97"
    endereco_normalizado_2 = normalizar_endereco(endereco_teste_2)
    assert endereco_normalizado_2 == "CALLE IGNACIO RAMIREZ 97"

    # Teste de remoção de vírgulas duplicadas
    endereco_com_virgulas = "CALLE A,, B, , ,C"
    endereco_com_virgulas_normalizado = normalizar_endereco(endereco_com_virgulas)
    assert endereco_com_virgulas_normalizado == "CALLE A,B,C"

    # Teste de extração de geo
    endereco_com_geo = normalizar_endereco(
        "Calle Tuxpan 29, Roma Sur, 6760 Ciudad de México, CDMX, Mexico (19,4051498, -99,167649)"
    )
    endereco_sem_geo, latitude, longitude = extrair_geo(endereco_com_geo)
    assert "CALLE TUXPAN 29" in endereco_sem_geo
    assert latitude is not None and longitude is not None

    # Teste parsing simples
    resultado_parse = parse_canonical_address("AVENIDA 5 DE MAYO 17")
    assert resultado_parse["TIPO"] == "AVENIDA"
    assert resultado_parse["LOGRADOURO"] == "5 DE MAYO"
    assert resultado_parse["NUMERO"] == "17"

    print("Todos os testes passaram.")


def main():
    logger = configurar_logger()

    entrada = "teste_Paulo.xlsx"
    saida = "teste_Paulo_tratado.xlsx"

    logger.info("Iniciando processamento de enderecos.")

    df = carregar_planilha(entrada, logger)
    logger.info(f"Planilha carregada com {len(df)} linhas.")

    df_tratado = process_dataframe(df, logger=logger)

    salvar_planilha(df_tratado, saida, logger)

    logger.info("Processamento concluido.")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        run_tests()
    else:
        main()
