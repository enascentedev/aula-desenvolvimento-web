# -*- coding: utf-8 -*-
"""
Gera a versao COMPLETA (guia passo a passo) dos slides da Aula 04.

Diferenca para gerar_slides.py: aquele e o deck da exposicao em aula, enxuto.
Este e o material de consulta que o aluno leva para casa — mostra a estrutura
de arquivos, o que precisa existir em cada arquivo e monta o Login.jsx em
etapas, uma peca de codigo por vez, com o codigo completo no final.

Reaproveita toda a identidade visual e os helpers de gerar_slides.py.

Uso:
    python gerar_slides_final.py

Saida:
    Aula04-Primeira-Pagina-React-final.pptx
    (o PDF sai desse .pptx pelo PowerPoint)
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import gerar_slides as gs
from gerar_slides import (  # noqa: F401
    ACCENT, ACCENT_2, ACCENT_DEEP, BG, BORDER, C_COMMENT, C_FUNC, C_KEYWORD,
    C_NUM, C_SELECTOR, C_STRING, C_TAG, C_VAR, ERR, ESCOLA, H, HILITE, M, MONO,
    OK, PROFESSOR, SANS, SURFACE, SURFACE_2, TEXT, TEXT_DIM, TEXT_FAINT, W,
    code_panel, eyebrow_pill, header, new_deck, oval, para, rect, rich_para,
    rule, slide_capa, slide_cartoes, slide_codigo, slide_conceito,
    slide_passos, textbox, body_block, add_slide, chrome,
)
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


TOTAL = 41
gs.TOTAL_SLIDES = TOTAL          # a regua de progresso do rodape usa isso

AREA_TOP = gs.AREA_TOP
AREA_H = gs.AREA_H


# ------------------------------------------------- LAYOUTS ADICIONAIS ------

TREE_LH = Pt(17)
TREE_LH_IN = 17 / 72.0

CORES_ARVORE = {
    "pasta": C_TAG,
    "arquivo": TEXT,
    "novo": OK,
    "editar": C_FUNC,
    "dim": TEXT_FAINT,
}


def painel_arvore(slide, x, y, w, h, titulo, itens, size=Pt(12)):
    """Painel de janela com uma arvore de arquivos dentro.

    itens = [(texto, tipo, nota), ...] — tipo em CORES_ARVORE.
    O texto ja vem com os tracos da arvore; eles saem apagados e so o nome
    do arquivo recebe cor, para o olho achar o arquivo antes do desenho.
    """
    bar_h = 0.42
    pad_x = 0.34
    pad_y = 0.24

    rect(slide, x, y, w, h, fill=SURFACE, line=BORDER,
         radius=0.035, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(slide, x, y, w, bar_h, fill=SURFACE_2, radius=0.09,
         shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    rect(slide, x, y + bar_h - 0.02, w, 0.02, fill=BORDER)
    for i, dot in enumerate((ERR, gs.RGBColor(0xE0, 0xB0, 0x60), OK)):
        oval(slide, x + 0.22 + i * 0.22, y + 0.155, 0.11, 0.11, fill=dot)
    tf = textbox(slide, x + 0.95, y + 0.09, w - 1.2, 0.26)
    para(tf, titulo, Pt(11), TEXT_DIM, font=MONO, first=True)

    tf = textbox(slide, x + pad_x, y + bar_h + pad_y, w - 2 * pad_x,
                 h - bar_h - 2 * pad_y)
    # as notas comecam todas na mesma coluna, senao viram escadinha
    col_nota = max((len(it[0]) for it in itens if len(it) > 2 and it[2]),
                   default=0) + 3
    for i, item in enumerate(itens):
        texto, tipo = item[0], item[1]
        nota = item[2] if len(item) > 2 else ""
        # separa os tracos da arvore do nome
        corte = 0
        while corte < len(texto) and texto[corte] in "│├└─ ":
            corte += 1
        pecas = []
        if corte:
            pecas.append((texto[:corte], TEXT_FAINT))
        pecas.append((texto[corte:], CORES_ARVORE[tipo],
                      tipo in ("novo", "editar")))
        if nota:
            recuo = max(3, col_nota - len(texto))
            pecas.append((" " * recuo + nota, TEXT_FAINT))
        rich_para(tf, pecas, size, font=MONO, line_spacing=TREE_LH,
                  first=(i == 0))
    return tf


def slide_arvore(prs, n, eyebrow, termo, titulo_painel, itens, blocos,
                 legenda=None):
    """Arvore de arquivos a esquerda, explicacao a direita."""
    slide = add_slide(prs)

    eyebrow_pill(slide, eyebrow)
    tf = textbox(slide, M, 1.02, 6.55, 0.66)
    para(tf, termo, Pt(31), TEXT, font=gs.DISPLAY, bold=True, first=True)
    rule(slide, M, 1.72)

    alt = min(AREA_H, 0.42 + 0.50 + len(itens) * TREE_LH_IN)
    topo = AREA_TOP + (AREA_H - alt) / 2
    painel_arvore(slide, M, topo, 6.55, alt, titulo_painel, itens)

    x = 7.55
    w = W - M - x
    rect(slide, x, AREA_TOP, 0.045, AREA_H, fill=BORDER)
    body_block(slide, x + 0.34, AREA_TOP + 0.08, w - 0.34,
               AREA_H - (1.30 if legenda else 0.20), blocos, size=Pt(14.5))

    if legenda:
        ly = AREA_TOP + AREA_H - 1.12
        rect(slide, x + 0.34, ly, w - 0.34, 1.12, fill=SURFACE, line=BORDER,
             radius=0.05, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        tfl = textbox(slide, x + 0.58, ly + 0.16, w - 0.82, 0.86)
        for i, (marca, cor, texto) in enumerate(legenda):
            rich_para(tfl, [(marca + "  ", cor, True, MONO), (texto, TEXT_DIM)],
                      Pt(12), line_spacing=1.3, first=(i == 0))

    chrome(slide, n)
    return slide


def slide_duplo(prs, n, eyebrow, termo, nome_a, col_a, nome_b, col_b,
                lang="js", size=Pt(11.5), line_h=Pt(15.5),
                first_a=1, first_b=None, rodape=None):
    """Listagem longa em duas colunas, com numeracao continua entre elas."""
    slide = add_slide(prs)

    eyebrow_pill(slide, eyebrow)
    tf = textbox(slide, M, 1.02, 8.6, 0.62)
    para(tf, termo, Pt(31), TEXT, font=MONO, bold=True, first=True)

    if first_b is None:
        first_b = first_a + len(col_a)

    gap = 0.3
    cw = (W - 2 * M - gap) / 2
    alt_max = 5.05 if not rodape else 4.55
    # a altura acompanha a coluna mais longa e o bloco fica centralizado —
    # nada de painel com meio palmo de vazio embaixo
    alt = min(alt_max, 0.42 + 0.52 + max(len(col_a), len(col_b)) * line_h.pt / 72)
    bloco = alt + (0.68 if rodape else 0.0)
    top = 1.62 + max(0.0, (4.94 - bloco) / 2)
    code_panel(slide, M, top, cw, alt, nome_a, col_a, lang=lang,
               size=size, line_h=line_h, first_line=first_a)
    code_panel(slide, M + cw + gap, top, cw, alt, nome_b, col_b, lang=lang,
               size=size, line_h=line_h, first_line=first_b)

    if rodape:
        ry = top + alt + 0.16
        rect(slide, M, ry, W - 2 * M, 0.52, fill=SURFACE, line=BORDER,
             radius=0.12, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        tf = textbox(slide, M + 0.28, ry, W - 2 * M - 0.56, 0.52,
                     anchor=MSO_ANCHOR.MIDDLE)
        rich_para(tf, rodape, Pt(12.5), first=True)

    chrome(slide, n)
    return slide


# ------------------------------------------------------ CODIGO COMPLETO ----

# Login.jsx da aula (sem o botao do trabalho) — 55 linhas.
# Todas as listagens parciais dos slides saem do mesmo codigo, para nao
# existir versao divergente do arquivo dentro do material.
JSX_FULL = [
    'import { useState } from "react";',
    'import "./Login.css";',
    "",
    "function Login() {",
    '  const [email, setEmail] = useState("");',
    '  const [senha, setSenha] = useState("");',
    '  const [erro, setErro] = useState("");',
    "",
    "  function handleSubmit(e) {",
    "    e.preventDefault();",
    "",
    "    if (!email || !senha) {",
    '      setErro("Preencha todos os campos.");',
    "      return;",
    "    }",
    "",
    '    if (email === "user@email.com" &&',
    '        senha === "123456") {',
    '      alert("Login realizado com sucesso!");',
    '      setErro("");',
    "    } else {",
    '      setErro("E-mail ou senha incorretos.");',
    "    }",
    "  }",
    "",
    "  return (",
    '    <div className="container">',
    "      <form onSubmit={handleSubmit}>",
    "        <h2>Login</h2>",
    "",
    "        <input",
    '          type="email"',
    '          placeholder="Email"',
    "          value={email}",
    "          onChange={(e) =>",
    "            setEmail(e.target.value)}",
    "        />",
    "",
    "        <input",
    '          type="password"',
    '          placeholder="Senha"',
    "          value={senha}",
    "          onChange={(e) =>",
    "            setSenha(e.target.value)}",
    "        />",
    "",
    '        {erro && <p className="erro">{erro}</p>}',
    "",
    '        <button type="submit">Entrar</button>',
    "      </form>",
    "    </div>",
    "  );",
    "}",
    "",
    "export default Login;",
]

CSS_FULL = [
    ".container {",
    "  display: flex;",
    "  justify-content: center;",
    "  align-items: center;",
    "  height: 100vh;",
    "  background: #f0f2f5;",
    "}",
    "",
    "form {",
    "  background: white;",
    "  padding: 2rem;",
    "  border-radius: 8px;",
    "  box-shadow: 0 2px 10px rgba(0,0,0,0.1);",
    "  display: flex;",
    "  flex-direction: column;",
    "  gap: 1rem;",
    "  width: 300px;",
    "}",
    "",
    "input {",
    "  padding: 0.75rem;",
    "  border: 1px solid #ccc;",
    "  border-radius: 6px;",
    "  font-size: 1rem;",
    "}",
    "",
    "button {",
    "  padding: 0.75rem;",
    "  background: #4f46e5;",
    "  color: white;",
    "  border: none;",
    "  border-radius: 6px;",
    "  font-size: 1rem;",
    "  cursor: pointer;",
    "}",
    "",
    "button:hover {",
    "  background: #4338ca;",
    "}",
    "",
    ".erro {",
    "  color: red;",
    "  font-size: 0.85rem;",
    "  margin: 0;",
    "}",
]

CSS_TRABALHO = [
    "/* botão mostrar / ocultar senha */",
    "",
    ".campo-senha {",
    "  position: relative;",
    "  display: flex;",
    "}",
    "",
    ".campo-senha input {",
    "  width: 100%;",
    "  padding-right: 5rem;",
    "}",
    "",
    ".ver-senha {",
    "  position: absolute;",
    "  right: 0.4rem;",
    "  top: 50%;",
    "  transform: translateY(-50%);",
    "  background: transparent;",
    "  color: #4f46e5;",
    "  border: none;",
    "  font-size: 0.8rem;",
    "  padding: 0.3rem 0.5rem;",
    "  cursor: pointer;",
    "  width: auto;",
    "}",
    "",
    ".ver-senha:hover {",
    "  background: #eef0fe;",
    "  border-radius: 4px;",
    "}",
]


# ------------------------------------------------ LISTAGENS DAS ETAPAS -----

COD_INDEX_HTML = [
    "<!doctype html>",
    '<html lang="pt-BR">',
    "  <head>",
    '    <meta charset="UTF-8" />',
    "    <title>Login — Aula 04</title>",
    "  </head>",
    "  <body>",
    '    <div id="root"></div>',
    '    <script type="module"',
    '            src="/src/main.jsx"></script>',
    "  </body>",
    "</html>",
]

COD_MAIN = [
    'import React from "react";',
    'import ReactDOM from "react-dom/client";',
    'import App from "./App";',
    'import "./index.css";',
    "",
    "ReactDOM.createRoot(",
    '  document.getElementById("root")',
    ").render(",
    "  <React.StrictMode>",
    "    <App />",
    "  </React.StrictMode>",
    ");",
]

COD_INDEX_CSS = [
    "* {",
    "  box-sizing: border-box;",
    "}",
    "",
    "body {",
    "  margin: 0;",
    "  font-family: system-ui, sans-serif;",
    "}",
]

E1_ESQUELETO = [
    "function Login() {",
    "  return (",
    "    <h2>Login</h2>",
    "  );",
    "}",
    "",
    "export default Login;",
]

E2_IMPORT = [
    'import { useState } from "react";',
    'import "./Login.css";',
    "",
    "function Login() {",
    "  ...",
    "}",
    "",
    "export default Login;",
]

E3_ESTADO = [
    "function Login() {",
    '  const [email, setEmail] = useState("");',
    '  const [senha, setSenha] = useState("");',
    '  const [erro, setErro] = useState("");',
    "",
    "  ...",
    "}",
]

E4_HANDLE = [
    "  ...",
    "",
    "  function handleSubmit(e) {",
    "    e.preventDefault();",
    "",
    "    ...",
    "  }",
    "",
    "  ...",
]

E5_VAZIO = [
    "  function handleSubmit(e) {",
    "    e.preventDefault();",
    "",
    "    if (!email || !senha) {",
    '      setErro("Preencha todos os campos.");',
    "      return;",
    "    }",
    "",
    "    ...",
    "  }",
]

E6_IFELSE = [
    "    ...",
    "",
    '    if (email === "user@email.com" &&',
    '        senha === "123456") {',
    '      alert("Login realizado com sucesso!");',
    '      setErro("");',
    "    } else {",
    '      setErro("E-mail ou senha incorretos.");',
    "    }",
    "  }",
]

E7_RETURN = [
    "  ...",
    "",
    "  return (",
    '    <div className="container">',
    "",
    "      ...",
    "",
    "    </div>",
    "  );",
    "}",
]

E8_FORM = [
    "  ...",
    "",
    '    <div className="container">',
    "      <form onSubmit={handleSubmit}>",
    "        <h2>Login</h2>",
    "",
    "        ...",
    "",
    "      </form>",
    "    </div>",
]

E9_EMAIL = [
    "        ...",
    "",
    "        <input",
    '          type="email"',
    '          placeholder="Email"',
    "          value={email}",
    "          onChange={(e) =>",
    "            setEmail(e.target.value)}",
    "        />",
    "",
    "        ...",
]

E10_SENHA = [
    "        ...",
    "",
    "        <input",
    '          type="password"',
    '          placeholder="Senha"',
    "          value={senha}",
    "          onChange={(e) =>",
    "            setSenha(e.target.value)}",
    "        />",
    "",
    "        ...",
]

E11_ERRO = [
    "        ...",
    "",
    '        {erro && <p className="erro">{erro}</p>}',
    "",
    "        ...",
]

E12_BOTAO = [
    "        ...",
    "",
    '        <button type="submit">Entrar</button>',
    "      </form>",
    "    </div>",
    "  );",
    "}",
    "",
    "export default Login;",
]

COD_APP = [
    'import Login from "./Login";',
    "",
    "function App() {",
    "  return <Login />;",
    "}",
    "",
    "export default App;",
]

T1_VERSENHA = [
    "  ...",
    '  const [erro, setErro] = useState("");',
    "  // guarda se a senha está visível",
    "  const [verSenha, setVerSenha] =",
    "    useState(false);",
    "",
    "  ...",
]

T2_TYPE = [
    "        ...",
    "",
    "        <input",
    '          type={verSenha ? "text" : "password"}',
    '          placeholder="Senha"',
    "          value={senha}",
    "          onChange={(e) =>",
    "            setSenha(e.target.value)}",
    "        />",
    "",
    "        ...",
]

T3_BOTAO = [
    "        ...",
    "",
    "        <button",
    '          type="button"',
    '          className="ver-senha"',
    "          onClick={() => setVerSenha(!verSenha)}",
    "        >",
    '          {verSenha ? "Ocultar" : "Mostrar"}',
    "        </button>",
    "",
    "        ...",
]

T4_COMPLETO = [
    '<div className="campo-senha">',
    "  <input",
    '    type={verSenha ? "text" : "password"}',
    '    placeholder="Senha"',
    "    value={senha}",
    "    onChange={(e) => setSenha(e.target.value)}",
    "  />",
    "",
    '  <button type="button" className="ver-senha"',
    "    onClick={() => setVerSenha(!verSenha)}>",
    '    {verSenha ? "Ocultar" : "Mostrar"}',
    "  </button>",
    "</div>",
]


# --------------------------------------------------------------- ARVORES ---

ARV_VITE = [
    ("meu-app/", "pasta"),
    ("├── node_modules/", "dim", "instalado pelo npm"),
    ("├── public/", "pasta"),
    ("├── src/", "pasta"),
    ("│   ├── assets/", "dim"),
    ("│   ├── App.css", "dim"),
    ("│   ├── App.jsx", "arquivo"),
    ("│   ├── index.css", "arquivo"),
    ("│   └── main.jsx", "arquivo"),
    ("├── .gitignore", "dim"),
    ("├── index.html", "arquivo"),
    ("├── package.json", "arquivo"),
    ("├── package-lock.json", "dim"),
    ("└── vite.config.js", "dim"),
]

ARV_AULA = [
    ("meu-app/", "pasta"),
    ("├── index.html", "arquivo", "já existe"),
    ("├── package.json", "arquivo", "já existe"),
    ("└── src/", "pasta"),
    ("    ├── main.jsx", "arquivo", "já existe"),
    ("    ├── index.css", "editar", "limpar"),
    ("    ├── App.jsx", "editar", "editar"),
    ("    ├── Login.jsx", "novo", "CRIAR"),
    ("    └── Login.css", "novo", "CRIAR"),
]


# --------------------------------------------------------------- BUILD -----

def build(destino):
    prs = new_deck()
    n = 1

    def passo():
        nonlocal n
        n += 1
        return n

    # 01 — capa
    slide_capa(
        prs,
        eyebrow="Programação Web 2  ·  Aula 04  ·  Guia completo",
        titulo_1="Primeira página",
        titulo_2="passo a passo",
        subtitle="Do projeto ao trabalho, arquivo por arquivo",
        author=f"{PROFESSOR}  ·  {ESCOLA}",
    )

    # 02 — roteiro
    slide_passos(
        prs, passo(), "Roteiro", "O caminho da aula", "",
        [
            "Criar o projeto com o Vite e entender a estrutura de arquivos.",
            [("Escrever o ", TEXT), ("Login.jsx", C_FUNC, True, MONO),
             (" em 12 etapas — uma peça de código por vez.", TEXT)],
            [("Estilizar no ", TEXT), ("Login.css", C_FUNC, True, MONO),
             (" e montar o componente no ", TEXT),
             ("App.jsx", C_FUNC, True, MONO), (".", TEXT)],
            "Fazer o trabalho: o botão mostrar / ocultar senha.",
        ],
    )

    # 03 — antes de comecar
    slide_cartoes(
        prs, passo(), "Preparação", "Antes de começar",
        [
            {
                "titulo": "Node.js",
                "texto": "É o Node que roda o Vite e o npm. Rode os dois "
                         "comandos abaixo no terminal: se as versões "
                         "aparecerem, está tudo certo.",
                "comando": ["node -v", "npm -v"],
                "nota": "Sem resposta? Instale em nodejs.org (versão LTS).",
            },
            {
                "titulo": "VS Code",
                "texto": "É onde vamos escrever o código. Instale também a "
                         "extensão \"ES7+ React snippets\" — ela ajuda, mas "
                         "não é obrigatória.",
                "comando": ["code ."],
                "nota": "O comando abre a pasta atual dentro do VS Code.",
            },
            {
                "titulo": "Uma pasta só sua",
                "texto": "Escolha onde o projeto vai morar e entre nessa pasta "
                         "pelo terminal antes de criar o projeto. Evite pastas "
                         "sincronizadas na nuvem.",
                "comando": ["cd Documentos/Projetos"],
                "nota": "Caminho sem acento e sem espaço evita dor de cabeça.",
            },
        ],
    )

    # 04 — criando o projeto
    slide_cartoes(
        prs, passo(), "Passo 1  ·  Montando o projeto", "Criando o projeto",
        [
            {
                "titulo": "Criar",
                "texto": "O Vite monta a pasta do projeto já com o React "
                         "configurado. Rode o comando e responda às perguntas "
                         "que ele fizer no terminal.",
                "comando": ["npm create vite@latest",
                            "  meu-app -- --template react"],
                "nota": "Troque \"meu-app\" pelo nome que você quiser.",
            },
            {
                "titulo": "Instalar",
                "texto": "Entre na pasta que acabou de nascer e instale as "
                         "dependências. O npm install cria a node_modules — "
                         "demora um pouco na primeira vez.",
                "comando": ["cd meu-app", "npm install"],
                "nota": "A node_modules não se edita e não se envia na entrega.",
            },
            {
                "titulo": "Rodar",
                "texto": "Sobe o servidor de desenvolvimento. Deixe rodando: "
                         "cada vez que você salvar um arquivo, o navegador "
                         "atualiza sozinho.",
                "comando": ["npm run dev"],
                "nota": "Abra o endereço http://localhost:5173 no navegador.",
            },
        ],
    )

    # 05 — arvore do Vite
    slide_arvore(
        prs, passo(), "Passo 1  ·  Montando o projeto",
        "O que o Vite criou", "meu-app", ARV_VITE,
        ["Essa é a pasta recém-criada. Parece muita coisa, mas hoje só quatro "
         "arquivos importam.",
         [("A ", TEXT_DIM), ("node_modules", TEXT, True, MONO),
          (" guarda as bibliotecas instaladas: nunca se edita nada lá dentro.",
           TEXT_DIM)],
         [("Todo o código que a gente escreve vive dentro da pasta ", TEXT_DIM),
          ("src", C_TAG, True, MONO), (" — é lá que os nossos arquivos vão "
          "nascer.", TEXT_DIM)],
         "Os arquivos apagados na lista ao lado são de configuração: deixe "
         "como estão."],
    )

    # 06 — arvore da aula
    slide_arvore(
        prs, passo(), "Passo 1  ·  Montando o projeto",
        "A estrutura da nossa aula", "meu-app", ARV_AULA,
        ["Ao final da aula a tua pasta src precisa ter exatamente esses cinco "
         "arquivos.",
         [("Dois deles você vai ", TEXT_DIM), ("criar do zero", TEXT, True),
          (": o componente e o estilo dele. Dois já existem e a gente ajusta. "
           "Um fica intocado.", TEXT_DIM)],
         "Pode apagar o App.css e a pasta assets: nesta aula eles não são "
         "usados."],
        legenda=[("CRIAR ", OK, "arquivo novo, feito por você"),
                 ("editar", C_FUNC, "arquivo que já existe e vamos mexer"),
                 ("cinza ", TEXT_FAINT, "não precisa tocar hoje")],
    )

    # 07 — o caminho ate a tela
    slide_passos(
        prs, passo(), "Como funciona", "O caminho até a tela", "",
        [
            [("index.html", C_FUNC, True, MONO),
             (" tem uma div vazia com id=\"root\". É a página que o navegador "
              "abre.", TEXT)],
            [("main.jsx", C_FUNC, True, MONO),
             (" pega essa div e manda o React desenhar a aplicação dentro "
              "dela.", TEXT)],
            [("App.jsx", C_FUNC, True, MONO),
             (" é o componente raiz: ele decide o que aparece na tela.", TEXT)],
            [("Login.jsx", C_FUNC, True, MONO),
             (" é o nosso componente — a tela que vamos construir hoje.",
              TEXT)],
        ],
    )

    # 08 — index.html
    slide_codigo(
        prs, passo(), "Arquivos que já existem", "index.html",
        "index.html", COD_INDEX_HTML, (7, 7),
        ["Esse arquivo já veio pronto e a gente quase não mexe nele.",
         [("A linha em destaque é a única que importa: a ", TEXT_DIM),
          ('<div id="root">', C_TAG, True, MONO),
          (" é o buraco vazio onde o React vai desenhar tudo.", TEXT_DIM)],
         "Repare que não existe nenhum texto da tela de login aqui. Toda a "
         "página é montada pelo JavaScript."],
        nota=[("Dica: ", ACCENT, True),
              ("mude o ", TEXT_DIM), ("<title>", C_TAG, True, MONO),
              (" para o nome da tua aplicação — é o texto da aba do "
               "navegador.", TEXT_DIM)],
    )

    # 09 — main.jsx
    slide_codigo(
        prs, passo(), "Arquivos que já existem", "main.jsx",
        "src/main.jsx", COD_MAIN, (5, 11),
        ["É a ponte entre o HTML e o React. Também já vem pronto.",
         [("Ele procura o elemento de id ", TEXT_DIM),
          ('"root"', C_STRING, True, MONO),
          (" e renderiza o ", TEXT_DIM), ("<App />", C_TAG, True, MONO),
          (" ali dentro.", TEXT_DIM)],
         [("A linha ", TEXT_DIM), ('import "./index.css"', C_FUNC, True, MONO),
          (" é o que faz o estilo global valer para a aplicação inteira.",
           TEXT_DIM)]],
    )

    # 10 — index.css
    slide_codigo(
        prs, passo(), "Arquivos que já existem", "index.css",
        "src/index.css", COD_INDEX_CSS, None,
        ["Apague tudo que veio no index.css e deixe só isso.",
         [("O ", TEXT_DIM), ("box-sizing: border-box", C_SELECTOR, True, MONO),
          (" faz o padding entrar na conta da largura — sem ele os campos "
           "estouram o formulário.", TEXT_DIM)],
         [("O ", TEXT_DIM), ("margin: 0", C_SELECTOR, True, MONO),
          (" no body tira a margem branca que o navegador coloca sozinho.",
           TEXT_DIM)]],
        lang="css",
    )

    # 11 — conceito JSX
    slide_conceito(
        prs, passo(), "React", "Trabalhamos em .JSX, não em .JS",
        [("Nós trabalhamos no arquivo ", TEXT_DIM),
         (".JSX", ACCENT, True, MONO),
         (", não ", TEXT_DIM),
         (".JS", TEXT_DIM, False, MONO),
         (".", TEXT_DIM)],
        ['Trabalhando em JSX nós usamos uma "extensão" do HTML dentro do '
         'nosso arquivo JS. Dessa forma temos uma sintaxe um pouco diferente '
         'do JavaScript normal.',
         "Na prática: escrevemos marcação parecida com HTML no meio do código "
         "JavaScript, e o React se encarrega de transformar isso na tela."],
    )

    # 12 — etapa 1: esqueleto
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 1 de 12", "o esqueleto",
        "src/Login.jsx", E1_ESQUELETO, None,
        ["Crie o arquivo Login.jsx dentro da pasta src e comece com isso. "
         "Salve e a palavra Login já aparece no navegador.",
         "Todo componente React é uma função que devolve marcação. O nome "
         "começa com letra MAIÚSCULA — é assim que o React sabe que é um "
         "componente, e não uma tag comum.",
         [("O ", TEXT_DIM), ("export default", C_KEYWORD, True, MONO),
          (" no final libera esse componente para ser usado em outros "
           "arquivos.", TEXT_DIM)]],
    )

    # 13 — etapa 2: imports
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 2 de 12", "import",
        "src/Login.jsx", E2_IMPORT, (0, 1),
        ['O import { useState } from "react" vai até a biblioteca do React e '
         'traz a função useState para dentro do nosso arquivo.',
         [("useState", C_FUNC, True, MONO), (" é o que chamamos de ", TEXT_DIM),
          ("Hook", TEXT, True), (". Ele permite que a função \"lembre\" de "
           "informações entre uma renderização e outra.", TEXT_DIM)],
         "A segunda linha importa o CSS. O arquivo ainda não existe — vamos "
         "criá-lo no passo 3; deixe a linha comentada até lá se preferir."],
        nota=[("Atenção: ", ERR, True),
              ("as chaves em ", TEXT_DIM), ("{ useState }", C_VAR, True, MONO),
              (" são obrigatórias.", TEXT_DIM)],
    )

    # 14 — etapa 3: estados
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 3 de 12", "useState",
        "src/Login.jsx", E3_ESTADO, (1, 3),
        ["Cada estado vem sempre em par: o valor e a função que muda esse "
         "valor. Os nomes são escolha nossa; setAlgumaCoisa é a convenção.",
         [("O que está entre parênteses é o ", TEXT_DIM),
          ("valor inicial", TEXT, True),
          (". Começamos com texto vazio nos três.", TEXT_DIM)],
         [("Nunca mude o valor na mão (", TEXT_DIM),
          ("email = ...", ERR, True, MONO),
          ("). Só a função setEmail avisa o React de que a tela precisa ser "
           "redesenhada.", TEXT_DIM)]],
    )

    # 15 — como o useState funciona
    slide_passos(
        prs, passo(), "Conceito", "Como o estado funciona",
        "Variáveis comuns não avisam a tela quando mudam. O estado avisa — e "
        "é por isso que o React existe. O ciclo é sempre este:",
        [
            [("Você digita no campo e o ", TEXT), ("onChange", C_FUNC, True,
                                                   MONO),
             (" dispara, chamando o setEmail com o texto novo.", TEXT)],
            [("O React guarda o valor e ", TEXT),
             ("executa a função Login de novo", TEXT, True),
             (", do começo ao fim.", TEXT)],
            "A tela é redesenhada já com o valor novo. Sem F5, sem "
            "document.getElementById, sem innerHTML.",
        ],
        marcadores=["01", "02", "03"],
    )

    # 16 — etapa 4: handleSubmit
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 4 de 12",
        "function handleSubmit(e)",
        "src/Login.jsx", E4_HANDLE, (2, 6),
        ["Essa é a nossa primeira função aninhada: ela mora dentro da função "
         "Login. Estranho para quem vem de outras linguagens, normal em React.",
         [("O ", TEXT_DIM), ("e", C_VAR, True, MONO),
          (" é o evento — quem disparou, que tipo foi, qual tecla. Poderíamos "
           "chamá-lo de evento; e é só convenção.", TEXT_DIM)],
         [("e.preventDefault()", C_FUNC, True, MONO),
          (" impede o comportamento padrão do formulário, que é recarregar a "
           "página inteira. Sem ela, tudo se perde ao clicar em Entrar.",
           TEXT_DIM)]],
    )

    # 17 — etapa 5: validacao
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 5 de 12",
        "campos vazios",
        "src/Login.jsx", E5_VAZIO, (3, 6),
        [[("O ", TEXT_DIM), ("!", ERR, True, MONO),
          (" significa \"não\". Então ", TEXT_DIM), ("!email", C_VAR, True,
                                                     MONO),
          (" quer dizer: o e-mail está vazio.", TEXT_DIM)],
         [("O ", TEXT_DIM), ("||", ACCENT, True, MONO),
          (" é o OU: basta um dos dois estar vazio para entrar no if.",
           TEXT_DIM)],
         [("O ", TEXT_DIM), ("return", C_KEYWORD, True, MONO),
          (" sozinho encerra a função ali mesmo — sem ele, o código seguiria "
           "e ainda tentaria conferir a senha.", TEXT_DIM)]],
    )

    # 18 — etapa 6: if/else
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 6 de 12", "if / else",
        "src/Login.jsx", E6_IFELSE, (2, 8),
        ["Aqui está a conferência: se o e-mail é user@email.com E a senha é "
         "123456, o login deu certo.",
         [("O ", TEXT_DIM), ("&&", ACCENT, True, MONO),
          (" é o E: as duas condições precisam ser verdadeiras. E ", TEXT_DIM),
          ("===", ACCENT, True, MONO),
          (" compara valor e tipo — em JavaScript use sempre três iguais.",
           TEXT_DIM)],
         "Deu certo: alert na tela e a mensagem de erro é limpa. Deu errado: "
         "a mensagem de erro recebe o texto e aparece sozinha na tela."],
        nota=[("Isso NÃO é segurança: ", ERR, True),
              ("a senha está escrita no código, que qualquer um lê no "
               "navegador. É exercício de useState.", TEXT_DIM)],
    )

    # 19 — etapa 7: return
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 7 de 12", "return ( )",
        "src/Login.jsx", E7_RETURN, (2, 9),
        ["Fechamos as duas funções e chegamos na parte visual. Daqui para "
         "baixo é praticamente HTML.",
         "O return de uma linha só dispensa parênteses; como o nosso ocupa "
         "várias linhas, os parênteses são obrigatórios.",
         [("Regra de ouro: o return devolve ", TEXT_DIM),
          ("um único elemento", TEXT, True),
          (". Por isso tudo vive dentro da div container.", TEXT_DIM)]],
        nota=[("className, não class: ", ACCENT, True),
              ("class é palavra reservada do JavaScript, então o JSX usa "
               "className.", TEXT_DIM)],
    )

    # 20 — etapa 8: form
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 8 de 12", "form / h2",
        "src/Login.jsx", E8_FORM, (3, 4),
        [[("O ", TEXT_DIM), ("onSubmit={handleSubmit}", C_FUNC, True, MONO),
          (" liga o formulário à nossa função: quem envia o form chama o "
           "handleSubmit.", TEXT_DIM)],
         "Repare nas chaves: em JSX, chaves significam \"aqui dentro é "
         "JavaScript\". Sem elas o React entenderia o texto literal.",
         "O h2 é o título do cartão de login — igualzinho ao HTML."],
    )

    # 21 — etapa 9: input de e-mail
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 9 de 12", "input email",
        "src/Login.jsx", E9_EMAIL, (2, 8),
        ['O type="email" avisa o navegador que o campo é um e-mail — no '
         'celular o teclado já abre com o @.',
         [("O ", TEXT_DIM), ("value={email}", C_FUNC, True, MONO),
          (" e o ", TEXT_DIM), ("onChange", C_FUNC, True, MONO),
          (" juntos fazem o campo virar um componente controlado: o React "
           "passa a ser o dono do que está escrito.", TEXT_DIM)],
         [("A tag fecha em ", TEXT_DIM), ("/>", C_TAG, True, MONO),
          (". Em JSX toda tag precisa ser fechada, inclusive input, img e br.",
           TEXT_DIM)]],
        nota=[("Atenção: ", ERR, True),
              ("value ", TEXT_DIM), ("sem", TEXT, True),
              (" onChange trava o campo — você digita e nada aparece.",
               TEXT_DIM)],
    )

    # 22 — componente controlado
    slide_passos(
        prs, passo(), "Conceito", "O campo controlado, em 3 tempos",
        "Essa ida e volta acontece a cada tecla digitada. Entender esse "
        "circuito é entender React inteiro:",
        [
            [("A tecla é digitada → o ", TEXT), ("onChange", C_FUNC, True,
                                                 MONO),
             (" recebe o evento e lê o ", TEXT),
             ("e.target.value", C_VAR, True, MONO), (".", TEXT)],
            [("O ", TEXT), ("setEmail", C_FUNC, True, MONO),
             (" guarda esse texto no estado e o React redesenha o "
              "componente.", TEXT)],
            [("O campo é redesenhado com ", TEXT),
             ("value={email}", C_FUNC, True, MONO),
             (" — a letra só aparece porque o estado mudou.", TEXT)],
        ],
        marcadores=["01", "02", "03"],
    )

    # 23 — etapa 10: input de senha
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 10 de 12",
        "input senha",
        "src/Login.jsx", E10_SENHA, (2, 8),
        ["É o mesmo campo de antes, trocando três coisas: o type, o "
         "placeholder e o estado usado.",
         [("O ", TEXT_DIM), ('type="password"', C_STRING, True, MONO),
          (" é o que troca as letras por bolinhas na tela.", TEXT_DIM)],
         "Guarde bem esta linha: é exatamente ela que o trabalho da aula pede "
         "para deixar dinâmica."],
    )

    # 24 — etapa 11: erro
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 11 de 12",
        "{erro && …}",
        "src/Login.jsx", E11_ERRO, (2, 2),
        ["Lendo em voz alta:",
         [("se a variável ", TEXT_DIM), ("erro", C_VAR, True, MONO),
          (" tiver algum texto, ENTÃO mostre o parágrafo ", TEXT_DIM),
          ("<p>", C_TAG, True, MONO), (" com a mensagem.", TEXT_DIM)],
         "É o curto-circuito do JavaScript: com o erro vazio, o React "
         "simplesmente não desenha nada — nenhum espaço em branco fica "
         "sobrando na tela."],
    )

    # 25 — etapa 12: botao e export
    slide_codigo(
        prs, passo(), "Passo 2  ·  Login.jsx  ·  etapa 12 de 12", "button",
        "src/Login.jsx", E12_BOTAO, (2, 2),
        [[("Como o botão é ", TEXT_DIM), ('type="submit"', C_STRING, True,
                                          MONO),
          (" e está dentro do form, clicar nele dispara o onSubmit — ou seja, "
           "o nosso handleSubmit.", TEXT_DIM)],
         "Pressionar Enter dentro de um campo faz exatamente a mesma coisa. "
         "Isso é de graça: veio do form.",
         "Confira o fechamento na ordem certa: form, div, o parêntese do "
         "return e a chave da função."],
        nota=[("export default Login", C_FUNC, True, MONO),
              ("\nSem essa linha o App.jsx não consegue importar o "
               "componente.", TEXT_DIM)],
    )

    # 26 e 27 — Login.jsx completo
    slide_duplo(
        prs, passo(), "Passo 2  ·  Login.jsx", "Login.jsx  (1 de 2)",
        "src/Login.jsx", JSX_FULL[0:14], "src/Login.jsx", JSX_FULL[14:28],
        first_a=1,
        rodape=[("Confira o teu arquivo linha a linha: ", TEXT_DIM),
                ("a lógica toda cabe nas primeiras 28 linhas.", TEXT, True)],
    )
    slide_duplo(
        prs, passo(), "Passo 2  ·  Login.jsx", "Login.jsx  (2 de 2)",
        "src/Login.jsx", JSX_FULL[28:42], "src/Login.jsx", JSX_FULL[42:55],
        first_a=29,
        rodape=[("Da linha 26 em diante é a parte visual: ", TEXT_DIM),
                ("um container, um form e quatro filhos dentro dele.", TEXT,
                 True)],
    )

    # 28 — criar o CSS
    slide_passos(
        prs, passo(), "Passo 3  ·  Estilo", "Criando o Login.css",
        "A página deve ser estilizada de acordo com as tuas preferências. "
        "Vamos usar aqui a estética da aula — e para isso o passo a passo é:",
        [
            [("Criar um arquivo novo dentro da pasta ", TEXT),
             ("src", C_TAG, True, MONO), (", ao lado do Login.jsx.", TEXT)],
            [("Nomear ele ", TEXT), ("Login.css", C_FUNC, True, MONO),
             (" — com L maiúsculo, igual ao componente.", TEXT)],
            [("Importar no topo do Login.jsx com ", TEXT),
             ('import "./Login.css";', C_FUNC, True, MONO), (" .", TEXT)],
        ],
        marcadores=["A", "B", "C"],
    )

    # 29 e 30 — Login.css
    slide_duplo(
        prs, passo(), "Passo 3  ·  Estilo", "Login.css  ·  a moldura",
        "src/Login.css", CSS_FULL[0:9], "src/Login.css", CSS_FULL[9:18],
        lang="css", first_a=1,
        rodape=[("O container ocupa a altura toda da tela (", TEXT_DIM),
                ("100vh", C_NUM, True, MONO),
                (") e centraliza o formulário com flex. O form é o cartão "
                 "branco.", TEXT_DIM)],
    )
    slide_duplo(
        prs, passo(), "Passo 3  ·  Estilo", "Login.css  ·  os campos",
        "src/Login.css", CSS_FULL[18:32], "src/Login.css", CSS_FULL[32:45],
        lang="css", first_a=19,
        rodape=[("Sem seletor de classe: ", TEXT_DIM),
                ("input", C_SELECTOR, True, MONO), (" e ", TEXT_DIM),
                ("button", C_SELECTOR, True, MONO),
                (" pegam todos os campos e botões da página — dá para começar "
                 "assim.", TEXT_DIM)],
    )

    # 31 — App.jsx
    slide_codigo(
        prs, passo(), "Passo 4  ·  Montando na tela", "App.jsx",
        "src/App.jsx", COD_APP, (0, 3),
        ["Apague o que o Vite deixou no App.jsx e deixe só isso.",
         [("O ", TEXT_DIM), ("<Login />", C_TAG, True, MONO),
          (" é o nosso componente usado como se fosse uma tag do HTML. Isso "
           "só é possível por causa do export default lá no fim do "
           "Login.jsx.", TEXT_DIM)],
         "Salve tudo, volte ao navegador: a tela de login já deve estar lá, "
         "centralizada."],
        nota=[("Não esqueça: ", ACCENT, True),
              ("o caminho é ", TEXT_DIM), ('"./Login"', C_STRING, True, MONO),
              (" — com ponto e barra, sem a extensão.", TEXT_DIM)],
    )

    # 32 — testando
    slide_passos(
        prs, passo(), "Conferindo", "Testando a tua tela", "",
        [
            "Clique em Entrar com os campos vazios: a mensagem \"Preencha "
            "todos os campos.\" precisa aparecer em vermelho.",
            "Digite qualquer e-mail e senha errados: a mensagem tem que "
            "mudar para \"E-mail ou senha incorretos.\".",
            [("Digite ", TEXT), ("user@email.com", C_STRING, True, MONO),
             (" e ", TEXT), ("123456", C_NUM, True, MONO),
             (": o alerta de sucesso aparece e o erro some.", TEXT)],
            "Aperte Enter dentro de um campo: tem que funcionar igual ao "
            "clique no botão.",
        ],
    )

    # 33 — erros comuns
    slide_cartoes(
        prs, passo(), "Se der errado", "Os erros mais comuns",
        [
            {
                "titulo": "Tela branca",
                "texto": "Quase sempre é erro de digitação no JSX. Abra o "
                         "console do navegador (F12) e leia a primeira linha "
                         "vermelha: ela diz o arquivo e a linha.",
                "comando": ["F12 → Console"],
                "nota": "Tag não fechada e chave a mais são os campeões.",
            },
            {
                "titulo": "Digito e não aparece",
                "texto": "O campo tem value mas está sem o onChange. O React "
                         "manda o valor do estado para a tela e nada devolve "
                         "o que você digitou.",
                "comando": ["value + onChange"],
                "nota": "Os dois sempre andam juntos num campo controlado.",
            },
            {
                "titulo": "A página recarrega",
                "texto": "Faltou o e.preventDefault() na primeira linha do "
                         "handleSubmit. O formulário fez o que ele faz desde "
                         "sempre: recarregar tudo.",
                "comando": ["e.preventDefault();"],
                "nota": "Sinal clássico: a tela pisca e o campo esvazia.",
            },
        ],
    )

    # 34 — mais erros
    slide_passos(
        prs, passo(), "Se der errado", "Mais coisas para conferir", "",
        [
            [("Escreveu ", TEXT), ("class", ERR, True, MONO),
             (" no lugar de ", TEXT), ("className", OK, True, MONO),
             (" — o estilo não aplica e o console avisa.", TEXT)],
            [("Nome do arquivo com outra letra: ", TEXT),
             ("login.jsx", ERR, True, MONO), (" e ", TEXT),
             ("Login.jsx", OK, True, MONO),
             (" são arquivos diferentes para o import.", TEXT)],
            "Esqueceu de salvar o arquivo. O Vite só atualiza a tela depois "
            "do Ctrl+S.",
            [("O terminal parou: o ", TEXT), ("npm run dev", OK, True, MONO),
             (" precisa ficar rodando o tempo todo enquanto você trabalha.",
              TEXT)],
        ],
    )

    # 35 — o trabalho
    slide_passos(
        prs, passo(), "Passo 5  ·  O trabalho", "Mostrar / ocultar senha",
        "A entrega da aula é um botãozinho dentro do campo de senha que "
        "alterna entre esconder e revelar o que foi digitado. São três "
        "mudanças no Login.jsx:",
        [
            [("Um estado novo: ", TEXT), ("verSenha", C_VAR, True, MONO),
             (", que começa em ", TEXT), ("false", C_KEYWORD, True, MONO),
             (".", TEXT)],
            [("O ", TEXT), ("type", C_VAR, True, MONO),
             (" do campo deixa de ser fixo e passa a depender desse estado.",
              TEXT)],
            "Um botão ao lado do campo que inverte o estado a cada clique.",
        ],
        marcadores=["01", "02", "03"],
    )

    # 36 — trabalho etapa 1
    slide_codigo(
        prs, passo(), "Passo 5  ·  O trabalho  ·  1 de 3", "verSenha",
        "src/Login.jsx", T1_VERSENHA, (2, 4),
        ["Mais um estado, junto dos outros três. A diferença é o valor "
         "inicial: em vez de texto vazio, o booleano false.",
         [("false", C_KEYWORD, True, MONO),
          (" significa \"a senha começa escondida\" — que é o comportamento "
           "que a gente quer ao abrir a tela.", TEXT_DIM)],
         "Nome em português e sem abreviação: verSenha diz o que guarda. "
         "Isso conta na correção."],
    )

    # 37 — trabalho etapa 2
    slide_codigo(
        prs, passo(), "Passo 5  ·  O trabalho  ·  2 de 3", "type dinâmico",
        "src/Login.jsx", T2_TYPE, (3, 3),
        [[("Antes o type era o texto fixo ", TEXT_DIM),
          ('"password"', C_STRING, True, MONO),
          (". Agora ele é JavaScript — e por isso vai entre chaves.",
           TEXT_DIM)],
         [("Isso é o ", TEXT_DIM), ("operador ternário", TEXT, True),
          (": condição ? valor se verdadeiro : valor se falso.", TEXT_DIM)],
         "Leia assim: se verSenha for verdadeiro use text (aparece), senão "
         "use password (bolinhas)."],
        nota=[("É uma linha só: ", ACCENT, True),
              ("o ternário substitui um if inteiro dentro do JSX, onde if não "
               "é permitido.", TEXT_DIM)],
    )

    # 38 — trabalho etapa 3
    slide_codigo(
        prs, passo(), "Passo 5  ·  O trabalho  ·  3 de 3", "o botão",
        "src/Login.jsx", T3_BOTAO, (2, 8),
        [[("type=\"button\"", C_STRING, True, MONO),
          (" é obrigatório: sem isso o botão viraria submit e tentaria fazer "
           "login a cada clique.", TEXT_DIM)],
         [("O ", TEXT_DIM), ("!verSenha", C_FUNC, True, MONO),
          (" é o inverso do valor atual — é o que faz um clique alternar "
           "entre os dois estados.", TEXT_DIM)],
         "O texto do botão também é um ternário: ele muda junto, de Mostrar "
         "para Ocultar."],
    )

    # 39 — o campo completo
    slide_codigo(
        prs, passo(), "Passo 5  ·  O trabalho", "o campo pronto",
        "src/Login.jsx", T4_COMPLETO, None,
        ["O input e o botão agora moram juntos dentro de uma div com a classe "
         "campo-senha. É ela que permite encostar o botão dentro do campo "
         "pelo CSS.",
         "Cuidado ao mover o input para dentro da div: a indentação muda, mas "
         "o conteúdo é o mesmo de antes.",
         "Esse bloco inteiro entra no lugar do input de senha da etapa 10."],
    )

    # 40 — CSS do trabalho
    slide_duplo(
        prs, passo(), "Passo 5  ·  O trabalho", "Login.css  ·  o botão",
        "src/Login.css", CSS_TRABALHO[0:16],
        "src/Login.css", CSS_TRABALHO[16:30],
        lang="css", first_a=46,
        rodape=[("Acrescente no final do Login.css. ", TEXT_DIM),
                ("position: absolute", C_SELECTOR, True, MONO),
                (" tira o botão do fluxo e o encaixa dentro do campo; o "
                 "padding-right abre espaço para ele.", TEXT_DIM)],
    )

    # 41 — fechamento
    slide_passos(
        prs, passo(), "Para entregar", "Antes de mandar o trabalho",
        "Envie a pasta do projeto SEM a node_modules — ela é recriada com um "
        "npm install. E lembre do aviso da aula:",
        [
            "Conferir e-mail e senha no front-end é só exercício. Numa "
            "aplicação real, qualquer pessoa lê essa senha no navegador.",
            "A autenticação de verdade acontece no servidor — é o que vamos "
            "construir no back-end em Python (FastAPI) do ponto eletrônico.",
            "Próxima aula: começamos a nova página da aplicação, conforme "
            "conversaremos em sala :)",
        ],
        marcadores=["!", "→", "✓"],
    )

    prs.save(destino)
    print(f"OK: {destino}  ({len(prs.slides)} slides)")
    if len(prs.slides) != TOTAL:
        print(f"AVISO: TOTAL={TOTAL} mas o deck tem {len(prs.slides)} slides "
              "— a régua de progresso vai sair errada.")


if __name__ == "__main__":
    saida = Path(__file__).with_name("Aula04-Primeira-Pagina-React-final.pptx")
    build(str(saida))
