import json
from collections import Counter
from math import ceil
import re
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path
from urllib.request import Request, urlopen

import pandas as pd
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

ARQUIVO_ENTRADA = r"C:\Users\Matheus\Downloads\ENDEREÇOS COLABORADORES _ teste- Copiar (2).xlsx"
ARQUIVO_SAIDA = r"C:\Users\Matheus\Downloads\trajetos_atualizados_deonibus.xlsx"
TOTAL_PARTES = 5

COL_ORIGEM = "Endereco"
COL_DESTINO = "Destino"
URL_BASIC_STOPS = "https://deonibus.com/api/partner/deonibus/basic-stops"

df = pd.read_excel(ARQUIVO_ENTRADA)
df.columns = df.columns.astype(str).str.strip()


def normalizar_texto(texto: str):
    texto = str(texto).strip()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = texto.lower()
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def resolver_nome_coluna(colunas, nome_esperado: str):
    alvo = normalizar_texto(nome_esperado)
    for coluna in colunas:
        if normalizar_texto(coluna) == alvo:
            return coluna
    raise KeyError(f"Coluna esperada '{nome_esperado}' nao encontrada. Colunas disponiveis: {list(colunas)}")


def split_cidade_estado(valor: str):
    texto = normalizar_texto(valor).replace(" - ", "-")
    partes = [parte.strip() for parte in texto.rsplit("-", 1)]
    if len(partes) == 2 and len(partes[1]) == 2:
        return partes[0], partes[1]
    return texto, ""


def normalizar_display_name(valor: str):
    texto = normalizar_texto(valor)
    texto = texto.replace("(todos)", "todos")
    texto = re.sub(r"[(),]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def carregar_stops():
    req = Request(
        URL_BASIC_STOPS,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json, text/plain, */*",
            "Referer": "https://deonibus.com/",
        },
    )
    with urlopen(req, timeout=60) as response:
        stops = json.load(response)

    for stop in stops:
        stop["display_norm"] = normalizar_display_name(stop["displayName"])
        stop["uri_norm"] = normalizar_texto(stop["uri"]).replace(" ", "-")
    return stops


def escolher_stop(valor: str, stops):
    cidade, uf = split_cidade_estado(valor)
    cidade_tokens = [token for token in re.split(r"[^a-z0-9]+", cidade) if token]

    candidatos = []
    for stop in stops:
        display = stop["display_norm"]
        score = 0

        if uf and f" {uf}" in f" {display}":
            score += 3
        if all(token in display for token in cidade_tokens):
            score += 5
        if display == f"{cidade} {uf} todos".strip():
            score += 10
        if display == f"{cidade} {uf}".strip():
            score += 8
        if "todos" in display:
            score += 2
        score -= min(int(stop.get("priority", 99999999)), 99999999) / 100000000

        if score >= 8:
            candidatos.append((score, stop))

    if not candidatos:
        return None

    candidatos.sort(key=lambda item: item[0], reverse=True)
    return candidatos[0][1]


def montar_url_resultado(origem_uri: str, destino_uri: str, data_viagem: str):
    return f"https://deonibus.com/passagens-de-onibus/{origem_uri}-para-{destino_uri}?departureDate={data_viagem}"


def extrair_menor_preco(page):
    loc = page.locator('[itemprop="price"]')
    valores = []
    for i in range(loc.count()):
        try:
            texto = loc.nth(i).inner_text(timeout=500).strip()
            valor = float(texto.replace(".", "").replace(",", "."))
            if 20 <= valor <= 5000:
                valores.append(valor)
        except Exception:
            pass
    return min(valores) if valores else None


def extrair_preco_mais_frequente(page):
    loc = page.locator('[itemprop="price"]')
    valores = []
    for i in range(loc.count()):
        try:
            texto = loc.nth(i).inner_text(timeout=500).strip()
            valor = float(texto.replace(".", "").replace(",", "."))
            if 20 <= valor <= 5000:
                valores.append(valor)
        except Exception:
            pass

    if not valores:
        return None

    contagem = Counter(valores)
    preco, _ = sorted(contagem.items(), key=lambda item: (-item[1], item[0]))[0]
    return preco


def registrar_resultados(resultados, selecoes, sem_preco, origem, destino, preco, status, selecao):
    base = {
        "origem": origem,
        "destino": destino,
        "preco": preco,
        "status": status,
    }

    selecoes.append(
        {
            **base,
            "origem_input": selecao.get("origem_input"),
            "origem_selecionada": selecao.get("origem_selecionada"),
            "origem_uri": selecao.get("origem_uri"),
            "destino_input": selecao.get("destino_input"),
            "destino_selecionado": selecao.get("destino_selecionado"),
            "destino_uri": selecao.get("destino_uri"),
            "data_viagem": selecao.get("data_viagem"),
            "url_resultado": selecao.get("url_resultado"),
        }
    )

    if preco is None:
        sem_preco.append(base)
        return

    resultados.append(
        {
            "origem": origem,
            "destino": destino,
            "preco": preco,
        }
    )


def exportar_excel(caminho_saida, resultados, selecoes, sem_preco):
    abas = {
        "resultados": pd.DataFrame(resultados, columns=["origem", "destino", "preco"]),
        "selecoes": pd.DataFrame(
            selecoes,
            columns=[
                "origem",
                "destino",
                "preco",
                "status",
                "origem_input",
                "origem_selecionada",
                "origem_uri",
                "destino_input",
                "destino_selecionado",
                "destino_uri",
                "data_viagem",
                "url_resultado",
            ],
        ),
        "viagens_none": pd.DataFrame(sem_preco, columns=["origem", "destino", "preco", "status"]),
    }

    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        for nome_aba, dataframe in abas.items():
            dataframe.to_excel(writer, sheet_name=nome_aba, index=False)


def gerar_caminho_parte(caminho_base: str, parte_atual: int, total_partes: int):
    base = Path(caminho_base)
    return str(base.with_name(f"{base.stem}_parte_{parte_atual}_de_{total_partes}{base.suffix}"))


def buscar_preco_deonibus(page, origem: str, destino: str, data_viagem: str, indice_linha: int, stops):
    selecao = {
        "origem_input": origem,
        "origem_selecionada": None,
        "origem_uri": None,
        "destino_input": destino,
        "destino_selecionado": None,
        "destino_uri": None,
        "data_viagem": data_viagem,
        "url_resultado": None,
    }

    origem_stop = escolher_stop(origem, stops)
    destino_stop = escolher_stop(destino, stops)

    if not origem_stop or not destino_stop:
        return None, "stop_nao_encontrado", selecao

    selecao["origem_selecionada"] = origem_stop["displayName"]
    selecao["origem_uri"] = origem_stop["uri"]
    selecao["destino_selecionado"] = destino_stop["displayName"]
    selecao["destino_uri"] = destino_stop["uri"]
    selecao["url_resultado"] = montar_url_resultado(origem_stop["uri"], destino_stop["uri"], data_viagem)

    try:
        page.goto(selecao["url_resultado"], wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(5000)

        corpo = page.locator("body").inner_text(timeout=10000)
        if "Ops! Não encontramos nenhuma viagem" in corpo or "não encontramos nenhuma viagem" in corpo.lower():
            return None, "nao_encontrado", selecao

        preco = extrair_preco_mais_frequente(page)
        if preco is not None:
            return preco, "ok", selecao

        try:
            page.screenshot(path=f"C:/Users/Matheus/Downloads/deonibus_sem_preco_{indice_linha}.png")
        except Exception:
            pass
        return None, "nao_encontrado", selecao

    except PlaywrightTimeoutError:
        try:
            page.screenshot(path=f"C:/Users/Matheus/Downloads/deonibus_timeout_{indice_linha}.png")
        except Exception:
            pass
        return None, "timeout", selecao

    except Exception as e:
        print(f"Erro DeOnibus {origem} -> {destino}: {e}")
        try:
            page.screenshot(path=f"C:/Users/Matheus/Downloads/deonibus_erro_{indice_linha}.png")
        except Exception:
            pass
        return None, "erro", selecao


COL_ORIGEM = resolver_nome_coluna(df.columns, COL_ORIGEM)
COL_DESTINO = resolver_nome_coluna(df.columns, COL_DESTINO)
trajetos_unicos = df[[COL_ORIGEM, COL_DESTINO]].drop_duplicates().reset_index(drop=True)
stops = carregar_stops()
tam_parte = ceil(len(trajetos_unicos) / TOTAL_PARTES)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True, args=["--disable-blink-features=AutomationControlled"])
    context = browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/145.0.0.0 Safari/537.36"
        ),
        locale="pt-BR",
        timezone_id="America/Sao_Paulo",
    )
    page = context.new_page()
    data_viagem = (datetime.now() + timedelta(days=7)).strftime("%d/%m/%Y")

    for indice_parte in range(TOTAL_PARTES):
        inicio = indice_parte * tam_parte
        fim = min(inicio + tam_parte, len(trajetos_unicos))
        if inicio >= fim:
            break

        resultados_exportacao = []
        selecoes_exportacao = []
        viagens_none_exportacao = []
        trajetos_parte = trajetos_unicos.iloc[inicio:fim].reset_index(drop=True)
        caminho_parte = gerar_caminho_parte(ARQUIVO_SAIDA, indice_parte + 1, TOTAL_PARTES)

        print(f"PROCESSANDO PARTE {indice_parte + 1}/{TOTAL_PARTES} - trajetos {inicio + 1} a {fim}")

        for i, row in trajetos_parte.iterrows():
            indice_global = inicio + i
            origem = str(row[COL_ORIGEM]).strip()
            destino = str(row[COL_DESTINO]).strip()

            preco, status, selecao = buscar_preco_deonibus(page, origem, destino, data_viagem, indice_global, stops)
            print(f"[parte {indice_parte + 1} | {i + 1}/{len(trajetos_parte)} | global {indice_global + 1}/{len(trajetos_unicos)}] {origem} -> {destino} = {preco} ({status})")

            registrar_resultados(
                resultados_exportacao,
                selecoes_exportacao,
                viagens_none_exportacao,
                origem,
                destino,
                preco,
                status,
                selecao,
            )

            if (i + 1) % 25 == 0:
                exportar_excel(caminho_parte, resultados_exportacao, selecoes_exportacao, viagens_none_exportacao)
                print(f"progresso salvo em {caminho_parte}")

        exportar_excel(caminho_parte, resultados_exportacao, selecoes_exportacao, viagens_none_exportacao)
        print(f"PARTE {indice_parte + 1} FINALIZADA: {caminho_parte}")

    browser.close()

print(f"FINALIZADO EM {TOTAL_PARTES} PARTES")
