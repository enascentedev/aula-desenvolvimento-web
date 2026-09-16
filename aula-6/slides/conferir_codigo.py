# -*- coding: utf-8 -*-
"""
Confere os slides da Aula 06 contra o projeto real de aula-6/api-login.

Os slides sao o guia de construcao do aluno: ele tem que conseguir montar a
API inteira so olhando para o telao. Este script verifica essa promessa:

  FIDELIDADE  todo codigo que aparece num painel existe mesmo no projeto
              (pega erro de digitacao, aspas trocadas, nome de campo errado);

  COBERTURA   todo codigo do main.py e do pyproject.toml aparece em algum
              painel (pega a linha que ninguem mostrou e o aluno nao tem como
              escrever);

  VERSAO      a versao do FastAPI digitada nos comandos dos slides e a mesma
              fixada no pyproject.toml;

  CABE        nenhuma linha passa da largura do painel — em Python uma linha
              quebrada pelo PowerPoint parece indentacao, e indentacao e codigo.

A comparacao normaliza: tira comentarios e todo espaco em branco.

Uso:
    python conferir_codigo.py        (sai 1 se achar problema)
"""

import io
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import gerar_slides as g  # noqa: E402

PROJETO = Path(__file__).resolve().parents[1] / "api-login"
falhas = []


def ler(nome):
    return io.open(PROJETO / nome, encoding="utf-8").read()


def normalizar(texto):
    """Tira comentarios # e espacos: sobra so o codigo que importa."""
    texto = re.sub(r"#[^\n]*", "", texto)
    return re.sub(r"\s+", "", texto)


def linhas_de_codigo(bloco):
    """As linhas de um painel que sao codigo: sem vazias, '...' e comentarios."""
    return [l for l in bloco
            if l.strip() and l.strip() not in ("...", "…")
            and not l.strip().startswith("#")]


# ---------------------------------------------------------------------------
# Que painel retrata que arquivo — e em que painel ele e desenhado, para
# saber quantas colunas cabem. Largura em polegadas, fonte em pontos.
# ---------------------------------------------------------------------------

LARGO = (g.PAINEL_LARGO, 13, 14)
DUPLO_A = (5.3, 10.5, 19)
DUPLO_B = (g.W - 2 * g.M - 0.3 - 5.3, 10.5, 19)
COMPARAR_B = (g.W - 2 * g.M - 0.34 - 4.9, 13, 8)

PAINEIS = [
    ("main.py", "CODE_TOPO", g.CODE_TOPO, LARGO),
    ("main.py", "CODE_MODELO", g.CODE_MODELO, LARGO),
    ("main.py", "CODE_ROTA", g.CODE_ROTA, LARGO),
    ("main.py", "CODE_IF", g.CODE_IF, LARGO),
    ("main.py", "CODE_RAISE", g.CODE_RAISE, LARGO),
    ("main.py", "CODE_MAIN_A", g.CODE_MAIN_A, DUPLO_A),
    ("main.py", "CODE_MAIN_B", g.CODE_MAIN_B, DUPLO_B),
    ("main.py", "PY_DICIONARIO", g.PY_DICIONARIO, COMPARAR_B),
    ("pyproject.toml", "TOML_PYPROJECT", g.TOML_PYPROJECT, LARGO),
]

TERMINAIS = [(n, b) for n, b in vars(g).items()
             if n.startswith("TERM_") and isinstance(b, list)]


# ------------------------------------------------------------- FIDELIDADE ---

print("FIDELIDADE — o que esta no slide existe no projeto")
for arquivo, nome, bloco, _ in PAINEIS:
    alvo = normalizar(ler(arquivo))
    fora = [l.strip() for l in linhas_de_codigo(bloco)
            if normalizar(l) not in alvo]
    for l in fora:
        falhas.append(f"{nome} mostra o que nao existe em {arquivo}: {l}")
    print(f"  {nome:16s} -> {arquivo:15s} "
          f"{len(linhas_de_codigo(bloco)):2d} linhas de codigo"
          f"{'   <-- PROBLEMA' if fora else ''}")


# -------------------------------------------------------------- COBERTURA ---

print("\nCOBERTURA — o que esta no projeto aparece em algum slide")
for arquivo in ("main.py", "pyproject.toml"):
    mostrado = normalizar("\n".join(
        "\n".join(bloco) for arq, _, bloco, _ in PAINEIS if arq == arquivo))
    codigo = linhas_de_codigo(ler(arquivo).split("\n"))
    ausentes = [l.strip() for l in codigo if normalizar(l) not in mostrado]
    for l in ausentes:
        falhas.append(f"{arquivo}: nenhum slide mostra esta linha: {l}")
    print(f"  {arquivo:15s} {len(codigo) - len(ausentes):2d}/{len(codigo):2d} "
          f"linhas nos slides{'   <-- FALTA MOSTRAR' if ausentes else ''}")

# o slide 30 e o arquivo inteiro: A + B tem que ser o main.py, linha a linha,
# com a indentacao exata (em Python ela e codigo)
juntas = g.CODE_MAIN_A + g.CODE_MAIN_B
reais = ler("main.py").rstrip("\n").split("\n")
if len(juntas) != len(reais):
    falhas.append(f"slide 30 tem {len(juntas)} linhas; o main.py tem {len(reais)}")
for i, (slide, real) in enumerate(zip(juntas, reais), start=1):
    if slide != real:
        falhas.append(f"main.py linha {i} do slide 30 diverge do arquivo: "
                      f"slide={slide!r} arquivo={real!r}")


# quem digita tudo o que ve nao pode quebrar o arquivo: a continuidade e
# "# ..." (comentario); um "..." solto com recuo e IndentationError
for _, nome, bloco, _ in PAINEIS:
    for l in bloco:
        if l.strip() in ("...", "…"):
            falhas.append(f"{nome}: use '# ...' em vez de {l!r} — digitado, "
                          f"vira IndentationError")


# ----------------------------------------------------------------- VERSAO ---

print("\nVERSAO — o FastAPI dos comandos e o do pyproject.toml")
fixada = re.search(r'"fastapi\[standard\]==([\d.]+)"', ler("pyproject.toml"))
if not fixada:
    falhas.append("pyproject.toml nao fixa fastapi[standard]==versao")
else:
    fonte = io.open(Path(__file__).with_name("gerar_slides.py"),
                    encoding="utf-8").read()
    citadas = set(re.findall(r"fastapi\[standard\]==([\d.]+)", fonte))
    for v in sorted(citadas - {fixada.group(1)}):
        falhas.append(f"os slides citam fastapi {v}, mas o projeto fixa "
                      f"{fixada.group(1)}")
    print(f"  pyproject.toml: {fixada.group(1)}   slides: "
          f"{', '.join(sorted(citadas)) or 'nenhuma'}")


# ------------------------------------------------------------ CABE NO SLIDE --

def colunas(largura, fonte):
    """Quantos caracteres de Cascadia Mono cabem num painel dessa largura."""
    return int((largura - 1.10) / (fonte * 0.6 / 72)) - 1


print("\nCABE NO PAINEL")
medidas = [(nome, bloco, *painel) for _, nome, bloco, painel in PAINEIS]
medidas += [(nome, bloco, g.W - 2 * g.M, g.TERM_SIZE.pt, 8)
            for nome, bloco in TERMINAIS]
for nome, bloco, largura, fonte, max_l in medidas:
    max_c = colunas(largura, fonte)
    larg = max(len(l) for l in bloco)
    ruim = len(bloco) > max_l or larg > max_c
    print(f"  {nome:18s} {len(bloco):2d}/{max_l:2d} linhas, "
          f"coluna maxima {larg:3d}/{max_c:3d}"
          f"{'   <-- ESTOURA' if ruim else ''}")
    if ruim:
        falhas.append(f"{nome}: {len(bloco)} linhas x {larg} colunas nao cabem "
                      f"(limite {max_l} x {max_c})")

print()
if falhas:
    print(f"{len(falhas)} PROBLEMA(S):")
    for f in falhas:
        print("  -", f)
    sys.exit(1)
print("OK: os slides sao fieis ao projeto e mostram o projeto inteiro.")
