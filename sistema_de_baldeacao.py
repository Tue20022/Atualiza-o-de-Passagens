import json
import os
import re
import time
import unicodedata
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.parse import quote, urlencode
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import pandas as pd

ARQUIVO_ENTRADA = r"C:\Users\Matheus\Downloads\trajetos_atualizados_deonibus_parte_5_de_5.xlsx"
ARQUIVO_SAIDA = r"C:\Users\Matheus\Downloads\trajetos_com_baldeacao.xlsx"

ABA_ENTRADA = "viagens_none"
ABA_ROTAS = "resultados"

HUBS = [
    "RIO DE JANEIRO-RJ",
    "MACAE-RJ",
    "CAMPOS DOS GOYTACAZES-RJ",
    "VITORIA-ES",
    "BELO HORIZONTE-MG",
    "JUIZ DE FORA-MG",
    "CAMPINAS-SP",
    "SAO PAULO-SP",
    "SALVADOR-BA",
    "ARACAJU-SE",
]

ALIASES_ORIGEM = ["Endereço", "Endereco", "origem"]
ALIASES_DESTINO = ["Destino", "destino"]
ALIASES_PRECO = ["PRECO_ATUAL", "preco", "Preço", "Preco"]

SEPARADOR_ROTA = " -> "
OSRM_BASE_URL = os.getenv("OSRM_BASE_URL", "https://router.project-osrm.org").rstrip("/")
NOMINATIM_BASE_URL = os.getenv("NOMINATIM_BASE_URL", "https://nominatim.openstreetmap.org").rstrip("/")
HTTP_TIMEOUT = 60
PESO_HORA_REAIS = float(os.getenv("PESO_HORA_REAIS", "20"))
OSRM_PROFILE = os.getenv("OSRM_PROFILE", "driving").strip().lower() or "driving"
NOMINATIM_DELAY_SEGUNDOS = float(os.getenv("NOMINATIM_DELAY_SEGUNDOS", "0.2"))
USER_AGENT = "sistema-de-baldeacao/1.0 (contato-local)"
ARQUIVO_CACHE_GEOCODING = Path(__file__).with_name("cache_geocoding_osrm.json")
ARQUIVO_CACHE_ROTAS = Path(__file__).with_name("cache_rotas_osrm.json")


def normalizar_texto(texto):
    texto = str(texto).strip()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r"\s+", " ", texto)
    return texto.upper()


def resolver_nome_coluna(colunas, nome_esperado):
    alvo = normalizar_texto(nome_esperado)
    for coluna in colunas:
        if normalizar_texto(coluna) == alvo:
            return coluna
    raise KeyError(
        f"Coluna esperada '{nome_esperado}' nao encontrada. "
        f"Colunas disponiveis: {list(colunas)}"
    )


def resolver_nome_coluna_por_aliases(colunas, aliases):
    for alias in aliases:
        try:
            return resolver_nome_coluna(colunas, alias)
        except KeyError:
            continue
    raise KeyError(
        f"Nenhuma das colunas esperadas foi encontrada. "
        f"Aliases testados: {aliases}. Colunas disponiveis: {list(colunas)}"
    )


def possui_colunas_esperadas(colunas, nomes_esperados):
    colunas_normalizadas = {normalizar_texto(coluna) for coluna in colunas}
    return all(normalizar_texto(nome) in colunas_normalizadas for nome in nomes_esperados)


def possui_alguma_coluna(colunas, aliases):
    colunas_normalizadas = {normalizar_texto(coluna) for coluna in colunas}
    return any(normalizar_texto(alias) in colunas_normalizadas for alias in aliases)


def carregar_planilha_entrada(caminho_arquivo, aba_preferencial, colunas_esperadas):
    excel = pd.ExcelFile(caminho_arquivo)

    if aba_preferencial in excel.sheet_names:
        df_aba = pd.read_excel(caminho_arquivo, sheet_name=aba_preferencial)
        if possui_colunas_esperadas(df_aba.columns, colunas_esperadas):
            return df_aba, aba_preferencial

    for nome_aba in excel.sheet_names:
        df_aba = pd.read_excel(caminho_arquivo, sheet_name=nome_aba)
        if possui_colunas_esperadas(df_aba.columns, colunas_esperadas):
            return df_aba, nome_aba

    raise KeyError(
        "Nenhuma aba do arquivo possui todas as colunas esperadas: "
        f"{colunas_esperadas}. Abas encontradas: {excel.sheet_names}"
    )


def listar_arquivos_partes(caminho_arquivo):
    caminho = Path(caminho_arquivo)
    nome_match = re.match(r"^(.*)_parte_\d+_de_\d+(\.xlsx)$", caminho.name, flags=re.IGNORECASE)
    if not nome_match:
        return [str(caminho)]

    prefixo = nome_match.group(1)
    sufixo = nome_match.group(2)
    padrao = f"{prefixo}_parte_*_de_*{sufixo}"
    arquivos = sorted(caminho.parent.glob(padrao))
    return [str(arquivo) for arquivo in arquivos] or [str(caminho)]


def carregar_planilhas_consolidadas(caminhos_arquivos, aba_preferencial, colunas_esperadas):
    dataframes = []
    abas_utilizadas = []

    for caminho_arquivo in caminhos_arquivos:
        df_aba, aba_utilizada = carregar_planilha_entrada(
            caminho_arquivo,
            aba_preferencial,
            colunas_esperadas,
        )
        df_aba = df_aba.copy()
        df_aba["ARQUIVO_FONTE"] = Path(caminho_arquivo).name
        dataframes.append(df_aba)
        abas_utilizadas.append(f"{Path(caminho_arquivo).name}:{aba_utilizada}")

    df_consolidado = pd.concat(dataframes, ignore_index=True) if dataframes else pd.DataFrame()
    return df_consolidado, abas_utilizadas


def carregar_cache_json(caminho_arquivo):
    if not caminho_arquivo.exists():
        return {}
    try:
        with caminho_arquivo.open("r", encoding="utf-8") as arquivo:
            return json.load(arquivo)
    except Exception:
        return {}


def salvar_cache_json(caminho_arquivo, dados):
    with caminho_arquivo.open("w", encoding="utf-8") as arquivo:
        json.dump(dados, arquivo, ensure_ascii=False, indent=2)


def normalizar_local(valor):
    if pd.isna(valor):
        return None

    texto = normalizar_texto(valor)
    if not texto or texto == "NAN":
        return None
    return texto


def normalizar_preco(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, str):
        texto = valor.strip()
        if not texto:
            return None

        texto = texto.replace("R$", "").replace(" ", "")
        if "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        else:
            partes = texto.split(".")
            if len(partes) > 2:
                texto = "".join(partes)

        try:
            valor = Decimal(texto)
        except InvalidOperation:
            return None
    else:
        try:
            valor = Decimal(str(valor))
        except InvalidOperation:
            return None

    if valor <= 0:
        return None
    return float(valor)


def montar_endereco_google(local):
    cidade, separador, uf = local.rpartition("-")
    if separador and uf:
        return f"{cidade.strip()}, {uf.strip()}, Brasil"
    return f"{local}, Brasil"


def fazer_requisicao_json(url):
    request = Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "application/json",
        },
    )

    try:
        with urlopen(request, timeout=HTTP_TIMEOUT) as response:
            return json.load(response)
    except HTTPError as exc:
        corpo = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Erro HTTP ao consultar {url}: {exc.code} {corpo}") from exc
    except URLError as exc:
        raise RuntimeError(f"Falha de conexao ao consultar {url}: {exc}") from exc


def parse_duration_seconds(duration_texto):
    if not duration_texto:
        return None
    if duration_texto.endswith("s"):
        duration_texto = duration_texto[:-1]
    try:
        return float(duration_texto)
    except ValueError:
        return None


def formatar_brl(valor):
    if valor is None:
        return ""
    texto = f"{valor:,.2f}"
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def formatar_duracao(segundos):
    if segundos is None:
        return ""
    total_minutos = int(round(segundos / 60))
    horas, minutos = divmod(total_minutos, 60)
    return f"{horas}h{minutos:02d}"


def formatar_distancia(metros):
    if metros is None:
        return ""
    return f"{metros / 1000:.1f} km"


def montar_trechos(caminho):
    return [f"{origem}{SEPARADOR_ROTA}{destino}" for origem, destino in zip(caminho, caminho[1:])]


def montar_rota_resumida(caminho, duracao_segundos, distancia_metros):
    partes = []
    if duracao_segundos is not None:
        partes.append(formatar_duracao(duracao_segundos))
    if distancia_metros is not None:
        partes.append(formatar_distancia(distancia_metros))
    if partes:
        return f"{SEPARADOR_ROTA.join(caminho)} ({', '.join(partes)})"
    return SEPARADOR_ROTA.join(caminho)


def obter_coordenadas_osm(local, cache_geocoding):
    if local in cache_geocoding:
        return cache_geocoding[local]

    params = urlencode(
        {
            "q": montar_endereco_google(local),
            "format": "jsonv2",
            "limit": 1,
            "countrycodes": "br",
        }
    )
    url = f"{NOMINATIM_BASE_URL}/search?{params}"
    resposta = fazer_requisicao_json(url)
    time.sleep(NOMINATIM_DELAY_SEGUNDOS)

    if not resposta:
        cache_geocoding[local] = None
        return None

    primeiro = resposta[0]
    coordenadas = {
        "lat": float(primeiro["lat"]),
        "lon": float(primeiro["lon"]),
    }
    cache_geocoding[local] = coordenadas
    return coordenadas


def obter_metricas_osrm(origem, destino, cache_metricas, cache_geocoding):
    chave = f"{origem}|{destino}|{OSRM_PROFILE}"
    if chave in cache_metricas:
        return cache_metricas[chave]

    origem_coord = obter_coordenadas_osm(origem, cache_geocoding)
    destino_coord = obter_coordenadas_osm(destino, cache_geocoding)
    if origem_coord is None or destino_coord is None:
        cache_metricas[chave] = None
        return None

    coordenadas = (
        f"{origem_coord['lon']},{origem_coord['lat']};"
        f"{destino_coord['lon']},{destino_coord['lat']}"
    )
    params = urlencode({"overview": "false", "steps": "false"})
    url = f"{OSRM_BASE_URL}/route/v1/{quote(OSRM_PROFILE)}/{coordenadas}?{params}"
    resposta = fazer_requisicao_json(url)

    rotas_resposta = resposta.get("routes") or []
    if not rotas_resposta:
        cache_metricas[chave] = None
        return None

    primeira_rota = rotas_resposta[0]
    metricas = {
        "distancia_metros": primeira_rota.get("distance"),
        "duracao_segundos": primeira_rota.get("duration"),
    }
    cache_metricas[chave] = metricas
    return metricas


def obter_metricas_caminho(caminho, cache_metricas, cache_geocoding):
    distancia_total = 0.0
    duracao_total = 0.0

    for origem, destino in zip(caminho, caminho[1:]):
        metricas = obter_metricas_osrm(origem, destino, cache_metricas, cache_geocoding)
        if metricas is None:
            return None
        if metricas["distancia_metros"] is None or metricas["duracao_segundos"] is None:
            return None
        distancia_total += metricas["distancia_metros"]
        duracao_total += metricas["duracao_segundos"]

    return {
        "distancia_metros": round(distancia_total, 2),
        "duracao_segundos": round(duracao_total, 2),
    }


def montar_resultado(tipo, caminho, cache_metricas, cache_geocoding):
    metricas = obter_metricas_caminho(caminho, cache_metricas, cache_geocoding)
    duracao_segundos = metricas["duracao_segundos"] if metricas else None
    distancia_metros = metricas["distancia_metros"] if metricas else None
    score_tempo = round(duracao_segundos / 3600, 2) if duracao_segundos is not None else None
    trechos = montar_trechos(caminho)

    if len(caminho) <= 2:
        baldeacao = "DIRETO"
    else:
        baldeacao = SEPARADOR_ROTA.join(caminho[1:-1])

    return {
        "TIPO_ROTA": tipo,
        "BALDEACAO": baldeacao,
        "ROTA_COMPLETA": SEPARADOR_ROTA.join(caminho),
        "CUSTO_TOTAL_ROTA": None,
        "TEMPO_TOTAL_HORAS": round(duracao_segundos / 3600, 2) if duracao_segundos is not None else None,
        "DISTANCIA_TOTAL_KM": round(distancia_metros / 1000, 2) if distancia_metros is not None else None,
        "SCORE_PRECO_TEMPO": score_tempo,
        "ROTA_RESUMIDA": montar_rota_resumida(caminho, duracao_segundos, distancia_metros),
        "TRECHOS_PARA_COTACAO": " | ".join(trechos),
        "TRECHO_1": trechos[0] if len(trechos) >= 1 else None,
        "TRECHO_2": trechos[1] if len(trechos) >= 2 else None,
        "TRECHO_3": trechos[2] if len(trechos) >= 3 else None,
        "_NUM_BALDEACOES": max(len(caminho) - 2, 0),
        "_TEMPO_SEGUNDOS": duracao_segundos,
    }


def sem_rota(origem=None, destino=None):
    caminho = [item for item in [origem, destino] if item]
    return {
        "TIPO_ROTA": "SEM ROTA",
        "BALDEACAO": "SEM ROTA",
        "ROTA_COMPLETA": SEPARADOR_ROTA.join(caminho) if caminho else "",
        "CUSTO_TOTAL_ROTA": None,
        "TEMPO_TOTAL_HORAS": None,
        "DISTANCIA_TOTAL_KM": None,
        "SCORE_PRECO_TEMPO": None,
        "ROTA_RESUMIDA": "",
        "TRECHOS_PARA_COTACAO": "",
        "TRECHO_1": None,
        "TRECHO_2": None,
        "TRECHO_3": None,
        "_NUM_BALDEACOES": None,
        "_TEMPO_SEGUNDOS": None,
    }


def encontrar_melhor_rota(origem, destino, conexoes, hubs_validos, cache_metricas, cache_geocoding):
    if not origem or not destino:
        return sem_rota(origem, destino)

    candidatos = []

    if (origem, destino) in conexoes:
        candidatos.append(
            montar_resultado(
                tipo="DIRETO",
                caminho=[origem, destino],
                cache_metricas=cache_metricas,
                cache_geocoding=cache_geocoding,
            )
        )

    for hub in hubs_validos:
        if hub in (origem, destino):
            continue

        if (origem, hub) not in conexoes or (hub, destino) not in conexoes:
            continue

        candidatos.append(
            montar_resultado(
                tipo="1 BALDEACAO",
                caminho=[origem, hub, destino],
                cache_metricas=cache_metricas,
                cache_geocoding=cache_geocoding,
            )
        )

    for hub1 in hubs_validos:
        if hub1 in (origem, destino):
            continue

        if (origem, hub1) not in conexoes:
            continue

        for hub2 in hubs_validos:
            if hub2 in (origem, destino) or hub2 == hub1:
                continue

            if (hub1, hub2) not in conexoes or (hub2, destino) not in conexoes:
                continue

            candidatos.append(
                montar_resultado(
                    tipo="2 BALDEACOES",
                    caminho=[origem, hub1, hub2, destino],
                    cache_metricas=cache_metricas,
                    cache_geocoding=cache_geocoding,
                )
            )

    if not candidatos:
        return sem_rota(origem, destino)

    melhor = min(
        candidatos,
        key=lambda item: (
            0 if item["_TEMPO_SEGUNDOS"] is not None else 1,
            item["_TEMPO_SEGUNDOS"] if item["_TEMPO_SEGUNDOS"] is not None else float("inf"),
            item["_NUM_BALDEACOES"] if item["_NUM_BALDEACOES"] is not None else float("inf"),
        ),
    )
    melhor.pop("_NUM_BALDEACOES", None)
    melhor.pop("_TEMPO_SEGUNDOS", None)
    return melhor


def diferenca(row):
    preco_atual = row.get("_PRECO_NORMALIZADO")
    custo_rota = normalizar_preco(row.get("CUSTO_TOTAL_ROTA"))
    if preco_atual is None or custo_rota is None:
        return None
    return round(custo_rota - preco_atual, 2)


df, aba_utilizada = carregar_planilha_entrada(
    ARQUIVO_ENTRADA,
    ABA_ENTRADA,
    ["origem", "destino"],
)
df.columns = df.columns.astype(str).str.strip()

arquivos_partes = listar_arquivos_partes(ARQUIVO_ENTRADA)
df_rotas, abas_rotas_utilizadas = carregar_planilhas_consolidadas(
    arquivos_partes,
    ABA_ROTAS,
    ["origem", "destino", "preco"],
)
df_rotas.columns = df_rotas.columns.astype(str).str.strip()

col_origem_entrada = resolver_nome_coluna_por_aliases(df.columns, ALIASES_ORIGEM)
col_destino_entrada = resolver_nome_coluna_por_aliases(df.columns, ALIASES_DESTINO)
col_preco_entrada = (
    resolver_nome_coluna_por_aliases(df.columns, ALIASES_PRECO)
    if possui_alguma_coluna(df.columns, ALIASES_PRECO)
    else None
)

col_origem_rotas = resolver_nome_coluna_por_aliases(df_rotas.columns, ALIASES_ORIGEM)
col_destino_rotas = resolver_nome_coluna_por_aliases(df_rotas.columns, ALIASES_DESTINO)
col_preco_rotas = resolver_nome_coluna_por_aliases(df_rotas.columns, ALIASES_PRECO)

df["_ORIGEM_NORMALIZADA"] = df[col_origem_entrada].apply(normalizar_local)
df["_DESTINO_NORMALIZADO"] = df[col_destino_entrada].apply(normalizar_local)
if col_preco_entrada:
    df["_PRECO_NORMALIZADO"] = df[col_preco_entrada].apply(normalizar_preco)
else:
    df["_PRECO_NORMALIZADO"] = pd.Series([None] * len(df), index=df.index)

df_rotas["_ORIGEM_NORMALIZADA"] = df_rotas[col_origem_rotas].apply(normalizar_local)
df_rotas["_DESTINO_NORMALIZADO"] = df_rotas[col_destino_rotas].apply(normalizar_local)
df_rotas["_PRECO_NORMALIZADO"] = df_rotas[col_preco_rotas].apply(normalizar_preco)

df_rotas_validas = df_rotas.dropna(
    subset=["_ORIGEM_NORMALIZADA", "_DESTINO_NORMALIZADO"]
)
conexoes = {
    (row["_ORIGEM_NORMALIZADA"], row["_DESTINO_NORMALIZADO"])
    for _, row in df_rotas_validas.iterrows()
}

cidades_base = set(df_rotas["_ORIGEM_NORMALIZADA"].dropna())
cidades_base.update(df_rotas["_DESTINO_NORMALIZADO"].dropna())
hubs_validos = [normalizar_texto(hub) for hub in HUBS if normalizar_texto(hub) in cidades_base]

cache_metricas = carregar_cache_json(ARQUIVO_CACHE_ROTAS)
cache_geocoding = carregar_cache_json(ARQUIVO_CACHE_GEOCODING)
resultados_lista = []

for indice, row in df.iterrows():
    resultado = encontrar_melhor_rota(
        row["_ORIGEM_NORMALIZADA"],
        row["_DESTINO_NORMALIZADO"],
        conexoes,
        hubs_validos,
        cache_metricas,
        cache_geocoding,
    )
    resultados_lista.append(resultado)

    if (indice + 1) % 5 == 0 or (indice + 1) == len(df):
        print(
            f"Processadas {indice + 1}/{len(df)} linhas | "
            f"cache geocoding: {len(cache_geocoding)} | cache rotas: {len(cache_metricas)}"
        )
        salvar_cache_json(ARQUIVO_CACHE_GEOCODING, cache_geocoding)
        salvar_cache_json(ARQUIVO_CACHE_ROTAS, cache_metricas)

resultados = pd.Series(resultados_lista, index=df.index)

df["TIPO_ROTA"] = resultados.apply(lambda x: x["TIPO_ROTA"])
df["BALDEACAO"] = resultados.apply(lambda x: x["BALDEACAO"])
df["ROTA_COMPLETA"] = resultados.apply(lambda x: x["ROTA_COMPLETA"])
df["CUSTO_TOTAL_ROTA"] = resultados.apply(lambda x: x["CUSTO_TOTAL_ROTA"])
df["TEMPO_TOTAL_HORAS"] = resultados.apply(lambda x: x["TEMPO_TOTAL_HORAS"])
df["DISTANCIA_TOTAL_KM"] = resultados.apply(lambda x: x["DISTANCIA_TOTAL_KM"])
df["SCORE_PRECO_TEMPO"] = resultados.apply(lambda x: x["SCORE_PRECO_TEMPO"])
df["ROTA_RESUMIDA"] = resultados.apply(lambda x: x["ROTA_RESUMIDA"])
df["TRECHOS_PARA_COTACAO"] = resultados.apply(lambda x: x["TRECHOS_PARA_COTACAO"])
df["TRECHO_1"] = resultados.apply(lambda x: x["TRECHO_1"])
df["TRECHO_2"] = resultados.apply(lambda x: x["TRECHO_2"])
df["TRECHO_3"] = resultados.apply(lambda x: x["TRECHO_3"])
df["DIFERENCA_VS_PRECO_ATUAL"] = df.apply(diferenca, axis=1)

salvar_cache_json(ARQUIVO_CACHE_GEOCODING, cache_geocoding)
salvar_cache_json(ARQUIVO_CACHE_ROTAS, cache_metricas)

df = df.drop(columns=["_ORIGEM_NORMALIZADA", "_DESTINO_NORMALIZADO", "_PRECO_NORMALIZADO"])
df.to_excel(ARQUIVO_SAIDA, index=False)

print("FINALIZADO")
print(f"Arquivo salvo em: {ARQUIVO_SAIDA}")
print(f"Aba lida: {aba_utilizada}")
print(f"Arquivos de rotas lidos: {len(arquivos_partes)}")
print(f"Abas de rotas: {abas_rotas_utilizadas}")
print(f"Hubs considerados: {hubs_validos}")
print(f"Total de conexoes conhecidas: {len(conexoes)}")
print(f"OSRM base URL: {OSRM_BASE_URL}")
print(f"Nominatim base URL: {NOMINATIM_BASE_URL}")
print(f"Perfil OSRM: {OSRM_PROFILE}")
print(f"Peso por hora aplicado no score: R$ {PESO_HORA_REAIS:.2f}")
print(f"Cache geocoding: {ARQUIVO_CACHE_GEOCODING}")
print(f"Cache rotas: {ARQUIVO_CACHE_ROTAS}")
