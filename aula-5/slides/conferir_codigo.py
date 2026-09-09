# -*- coding: utf-8 -*-
"""
Confere os slides da Aula 05 contra o codigo real de aula-5/exemplo-rotas.

Os slides sao o guia de construcao do aluno: ele tem que conseguir digitar o
projeto inteiro so olhando para o telao. Este script verifica as duas metades
dessa promessa:

  FIDELIDADE  todo codigo que aparece num painel existe mesmo no projeto
              (pega erro de digitacao, nome de chave trocado, aspas erradas);

  COBERTURA   todo codigo do projeto aparece em algum painel
              (pega o arquivo que ninguem mostrou e o aluno nao tem como
              escrever).

A comparacao normaliza: tira comentarios e todo espaco em branco. Assim uma
linha longa que o painel quebrou em duas continua batendo com o arquivo.

Uso:
    python conferir_codigo.py        (sai 1 se achar problema)
"""

import io
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import gerar_slides as g  # noqa: E402

SRC = Path(__file__).resolve().parents[1] / "exemplo-rotas" / "src"
falhas = []


def ler(nome):
    return io.open(SRC / nome, encoding="utf-8").read()


def normalizar(texto):
    """Tira comentarios e espacos: sobra so o codigo que importa."""
    texto = re.sub(r"//[^\n]*", "", texto)
    texto = re.sub(r"\{\s*/\*.*?\*/\s*\}", "", texto, flags=re.S)   # {/* JSX */}
    texto = re.sub(r"/\*.*?\*/", "", texto, flags=re.S)             # CSS e JS
    return re.sub(r"\s+", "", texto)


def linhas_de_codigo(bloco):
    """As linhas de um painel que sao codigo: sem vazias, '...' e comentarios."""
    fora = []
    for linha in bloco:
        t = linha.strip()
        if not t or t in ("...", "…") or t.startswith(("//", "/*", "*")):
            continue
        if t.endswith("*/"):
            continue
        fora.append(linha)
    return fora


# ---------------------------------------------------------------------------
# Que painel retrata que arquivo.
#
# A aula e incremental: um slide do comeco mostra o arquivo como ele estava
# NAQUELE momento, e nao como ficou no fim. Quando um painel mostra uma linha
# que a propria aula vai substituir depois, ela entra em `so_neste_estagio` —
# com o slide onde isso acontece. Fora dessa lista, divergencia e erro.
# ---------------------------------------------------------------------------

PAINEIS = [
    ("main.jsx", "CODE_MAIN_TOPO", g.CODE_MAIN_TOPO, []),
    ("main.jsx", "CODE_MAIN_RENDER", g.CODE_MAIN_RENDER, []),
    ("App.jsx", "CODE_APP", g.CODE_APP, [
        # slides 08 e 09: a Home ainda esta solta; o guardiao chega no slide 22
        '<Route path="/home" element={<Home />} />',
    ]),
    ("App.jsx", "CODE_APP_GUARDA", g.CODE_APP_GUARDA, []),
    ("App.jsx", "CODE_APP_404", g.CODE_APP_404, []),
    ("Login.jsx", "CODE_LOGIN_TOPO", g.CODE_LOGIN_TOPO, []),
    ("Login.jsx", "CODE_NAVIGATE", g.CODE_NAVIGATE, []),
    ("Login.jsx", "CODE_LOGIN_SESSAO", g.CODE_LOGIN_SESSAO, []),
    ("Home.jsx", "CODE_HOME", g.CODE_HOME, [
        # slides 13 e 14: o Sair ainda e um Link; vira botao no slide 23
        'import { Link } from "react-router-dom";',
        '<Link to="/">Sair</Link>',
    ]),
    ("Home.jsx", "CODE_HOME_SAIR", g.CODE_HOME_SAIR, []),
    ("Home.css", "CSS_HOME_A", g.CSS_HOME_A, []),
    ("Home.css", "CSS_HOME_B", g.CSS_HOME_B, []),
    ("Home.css", "CSS_SAIR", g.CSS_SAIR, []),
    ("RotaPrivada.jsx", "CODE_ROTA_PRIVADA", g.CODE_ROTA_PRIVADA, []),
    ("NaoEncontrada.jsx", "CODE_NAO_ENCONTRADA", g.CODE_NAO_ENCONTRADA, []),
]

# Linhas que o aluno ve fora de um painel (na caixa de nota do slide).
NAS_NOTAS = {
    "Home.jsx": ["export default Home;"],                      # nota do slide 13
    "NaoEncontrada.jsx": ["export default NaoEncontrada;"],    # nota do slide 27
}

# O Login.jsx inteiro veio da aula 4. Desta aula, so precisa aparecer no telao
# o que tem a ver com rota e sessao.
SO_O_QUE_E_DESTA_AULA = {
    "Login.jsx": ("navigate", "localStorage", "react-router-dom"),
}

ARQUIVOS = ["main.jsx", "App.jsx", "Login.jsx", "Home.jsx", "Home.css",
            "RotaPrivada.jsx", "NaoEncontrada.jsx"]


# ------------------------------------------------------------- FIDELIDADE ---

print("FIDELIDADE — o que esta no slide existe no projeto")
for arquivo, nome, bloco, estagio in PAINEIS:
    alvo = normalizar(ler(arquivo))
    perdoadas = {normalizar(l) for l in estagio}
    for linha in estagio:
        if normalizar(linha) in alvo:
            falhas.append(f"{nome}: {linha!r} esta marcada como codigo de um "
                          f"estagio anterior, mas continua em {arquivo} — "
                          f"tire da lista so_neste_estagio")
    fora = [l.strip() for l in linhas_de_codigo(bloco)
            if normalizar(l) not in alvo and normalizar(l) not in perdoadas]
    if fora:
        for l in fora:
            falhas.append(f"{nome} mostra o que nao existe em {arquivo}: {l}")
    print(f"  {nome:22s} -> {arquivo:20s} "
          f"{len(linhas_de_codigo(bloco)):2d} linhas de codigo"
          f"{'   <-- PROBLEMA' if fora else ''}")


# -------------------------------------------------------------- COBERTURA ---

print("\nCOBERTURA — o que esta no projeto aparece em algum slide")
for arquivo in ARQUIVOS:
    mostrado = normalizar("\n".join(
        "\n".join(bloco) for arq, _, bloco, _ in PAINEIS if arq == arquivo))
    mostrado += normalizar("\n".join(NAS_NOTAS.get(arquivo, [])))
    filtro = SO_O_QUE_E_DESTA_AULA.get(arquivo)

    ausentes, total = [], 0
    for linha in linhas_de_codigo(ler(arquivo).split("\n")):
        if filtro and not any(p in linha for p in filtro):
            continue
        total += 1
        if normalizar(linha) not in mostrado:
            ausentes.append(linha.strip())
    for l in ausentes:
        falhas.append(f"{arquivo}: nenhum slide mostra esta linha: {l}")
    escopo = "  (so o que e desta aula)" if filtro else ""
    print(f"  {arquivo:20s} {total - len(ausentes):2d}/{total:2d} linhas nos "
          f"slides{escopo}{'   <-- FALTA MOSTRAR' if ausentes else ''}")


# ------------------------------------------------------------ CABE NO SLIDE --

print("\nCABE NO PAINEL — largo: 14 linhas x 49 colunas; duplo: 19 x 50")
for nome, bloco in sorted((n, b) for n, b in vars(g).items()
                          if n.startswith(("CODE_", "CSS_"))
                          and isinstance(b, list)):
    duplo = nome in ("CSS_HOME_A", "CSS_HOME_B")
    max_l, max_c = (19, 50) if duplo else (14, 49)
    larg = max(len(l) for l in bloco)
    ruim = len(bloco) > max_l or larg > max_c
    print(f"  {nome:22s} {len(bloco):2d} linhas, coluna maxima {larg:2d}"
          f"{'   <-- ESTOURA' if ruim else ''}")
    if ruim:
        falhas.append(f"{nome}: {len(bloco)} linhas x {larg} colunas nao cabem")

# a numeracao do slide 15 tem que ser corrida: A termina onde B comeca
juntas = g.CSS_HOME_A + g.CSS_HOME_B
reais = ler("Home.css").split("\n")
for i, (slide, real) in enumerate(zip(juntas, reais), start=1):
    if slide != real:
        falhas.append(f"Home.css linha {i} do slide 15 diverge do arquivo: "
                      f"slide={slide!r} arquivo={real!r}")

print()
if falhas:
    print(f"{len(falhas)} PROBLEMA(S):")
    for f in falhas:
        print("  -", f)
    sys.exit(1)
print("OK: os slides sao fieis ao projeto e mostram o projeto inteiro.")
