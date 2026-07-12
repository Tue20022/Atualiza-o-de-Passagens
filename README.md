# Atualização de Passagens

Conjunto de scripts em Python para automatizar a cotação de passagens rodoviárias e o cálculo de rotas com baldeação entre os endereços de uma planilha de colaboradores.

O fluxo é dividido em duas etapas independentes, encadeadas pelos arquivos Excel que uma gera para a outra:

1. **`atualizar_passagens.py`** — busca o preço da passagem direta (origem → destino) no site [deonibus.com](https://deonibus.com), usando automação de navegador (Playwright).
2. **`sistema_de_baldeacao.py`** — para os trajetos sem passagem direta, monta rotas alternativas passando por até dois hubs (terminais intermediários), calculando distância e tempo via OSRM/Nominatim.

Há também um script solto de exploração/prototipagem (`from playwright.py`) usado para investigar o autocomplete de origem do site ClickBus — não faz parte do pipeline principal.

## Como funciona

### 1. `atualizar_passagens.py`

- Lê a planilha de entrada (`ARQUIVO_ENTRADA`), identificando automaticamente as colunas de origem (`Endereco`) e destino (`Destino`), tolerando acentos/maiúsculas diferentes.
- Deduplica os trajetos únicos (origem, destino).
- Baixa a lista de paradas (`stops`) da API pública do deonibus.com e usa um algoritmo de pontuação (`escolher_stop`) para casar cada cidade/UF da planilha com a parada correta do site.
- Abre um navegador Chromium headless via Playwright, monta a URL de resultados para a data de viagem (hoje + 7 dias) e extrai o preço mais frequente entre as ofertas exibidas (`extrair_preco_mais_frequente`).
- Processa os trajetos em `TOTAL_PARTES` (padrão: 5) blocos, salvando o progresso a cada 25 linhas e exportando um Excel por parte com três abas:
  - `resultados`: origem, destino e preço encontrado.
  - `selecoes`: detalhe de qual parada foi selecionada para cada origem/destino, URL usada e status da busca.
  - `viagens_none`: trajetos sem preço encontrado (sem rota direta, timeout ou erro), que alimentam a etapa de baldeação.
- Em caso de falha ou preço não encontrado, tira um screenshot da página para depuração.

### 2. `sistema_de_baldeacao.py`

- Lê a aba `viagens_none` (trajetos sem passagem direta) gerada pela etapa anterior, junto com a aba `resultados` de **todas** as partes exportadas (consolidando os arquivos `*_parte_N_de_M.xlsx` automaticamente).
- Usa os trajetos com preço conhecido como grafo de conexões diretas entre cidades.
- Define uma lista fixa de `HUBS` (Rio de Janeiro, Macaé, Campos dos Goytacazes, Vitória, Belo Horizonte, Juiz de Fora, Campinas, São Paulo, Salvador, Aracaju) e busca, para cada trajeto sem rota direta, a melhor combinação com **1 ou 2 baldeações** passando por esses hubs.
- Para cada rota candidata, calcula distância e duração reais via **OSRM** (roteamento) e **Nominatim** (geocodificação), com cache local em `cache_geocoding_osrm.json` e `cache_rotas_osrm.json` para evitar chamadas repetidas.
- Escolhe a rota de menor tempo total (com menor número de baldeações como critério de desempate).
- Gera colunas extras na planilha final: tipo de rota (`DIRETO` / `1 BALDEACAO` / `2 BALDEACOES` / `SEM ROTA`), hub(s) de baldeação, rota completa, tempo e distância totais, trechos individuais para cotação manual, e a diferença entre o preço atual e o custo estimado da rota.
- Salva o resultado consolidado em `ARQUIVO_SAIDA` (`trajetos_com_baldeacao.xlsx`).

### 3. `from playwright.py`

Script avulso de testes, usado para inspecionar visualmente (`headless=False`) o comportamento do campo de autocomplete de origem no site ClickBus. Não é chamado pelos outros scripts.

## Requisitos

- Python 3.10+
- Dependências: `pandas`, `openpyxl`, `playwright`

```bash
pip install pandas openpyxl playwright
playwright install chromium
```

## Uso

Os caminhos de entrada/saída estão fixos no topo de cada script (`ARQUIVO_ENTRADA` / `ARQUIVO_SAIDA`) e devem ser ajustados conforme o ambiente antes da execução.

```bash
# 1. Cotar passagens diretas no deonibus.com
python atualizar_passagens.py

# 2. Calcular rotas com baldeação para os trajetos sem passagem direta
python sistema_de_baldeacao.py
```

Variáveis de ambiente opcionais para `sistema_de_baldeacao.py`:

| Variável | Padrão | Descrição |
| --- | --- | --- |
| `OSRM_BASE_URL` | `https://router.project-osrm.org` | Servidor OSRM usado para calcular rotas |
| `NOMINATIM_BASE_URL` | `https://nominatim.openstreetmap.org` | Servidor Nominatim usado para geocodificação |
| `OSRM_PROFILE` | `driving` | Perfil de roteamento do OSRM |
| `PESO_HORA_REAIS` | `20` | Peso (R$/hora) usado no score preço x tempo |
| `NOMINATIM_DELAY_SEGUNDOS` | `0.2` | Intervalo entre chamadas ao Nominatim, para respeitar limites de uso |

## Arquivos de cache

`cache_geocoding_osrm.json` e `cache_rotas_osrm.json` armazenam, respectivamente, coordenadas geocodificadas e métricas de rota já calculadas, evitando refazer requisições às APIs públicas em execuções futuras.
