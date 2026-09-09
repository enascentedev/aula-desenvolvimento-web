# -*- coding: utf-8 -*-
"""
Gera a apostila do trabalho da Aula 05 - Navegando entre paginas (rotas).

Segue o mesmo padrao da apostila da Aula 04: A4, capa, sumario, numeracao
de pagina. O trabalho continua o projeto de login da aula anterior.

Uso:
    python gerar_apostila.py

Saida:
    Trabalho-Rotas-Aula05.docx

O sumario e criado como campo do Word: ao abrir o arquivo, clique com o botao
direito sobre ele e escolha "Atualizar campo" para preencher as paginas.
Depois e so exportar em PDF pelo proprio Word.
"""

import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

AUTOR = "Emanuel Nascente"
ESCOLA = "Escola Técnica Mesquita"
DISCIPLINA = "Programação Web 2"
DATA = "Aula 05"

INDIGO = RGBColor(0x4F, 0x46, 0xE5)
CINZA = RGBColor(0x55, 0x55, 0x60)
CINZA_CLARO = RGBColor(0x80, 0x80, 0x8C)
PRETO = RGBColor(0x1A, 0x1A, 0x1A)
VERDE = RGBColor(0x0A, 0x7A, 0x50)
VERMELHO = RGBColor(0xC0, 0x39, 0x3C)

MONO = "Consolas"
SANS = "Segoe UI"


# ----------------------------------------------------------------- BASE ----

def set_cell_bg(cell, hexcolor):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hexcolor)
    cell._tc.get_or_add_tcPr().append(shd)


def add_field(paragraph, instrucao):
    """Insere um campo do Word (usado no sumario e na numeracao de pagina)."""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instrucao
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    for el in (fld_begin, instr, fld_sep, fld_end):
        run._r.append(el)
    return run


def atualizar_campos_ao_abrir(doc):
    """Faz o Word preencher o sumario sozinho na primeira vez que abrir."""
    settings = doc.settings.element
    tag = settings.find(qn("w:updateFields"))
    if tag is None:
        tag = OxmlElement("w:updateFields")
        settings.append(tag)
    tag.set(qn("w:val"), "true")


def estilo_base(doc):
    normal = doc.styles["Normal"]
    normal.font.name = SANS
    normal.font.size = Pt(11)
    normal.font.color.rgb = PRETO
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.25

    for nome, tamanho, cor in (("Heading 1", 20, INDIGO),
                               ("Heading 2", 14, PRETO),
                               ("Heading 3", 12, CINZA)):
        st = doc.styles[nome]
        st.font.name = SANS
        st.font.size = Pt(tamanho)
        st.font.bold = True
        st.font.color.rgb = cor
        st.paragraph_format.space_before = Pt(16 if nome == "Heading 1" else 12)
        st.paragraph_format.space_after = Pt(6)


def rodape_numerado(section):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{DISCIPLINA}  ·  {DATA}  ·  ")
    run.font.size = Pt(8)
    run.font.color.rgb = CINZA_CLARO
    run.font.name = SANS
    campo = add_field(p, "PAGE")
    campo.font.size = Pt(8)
    campo.font.color.rgb = CINZA_CLARO
    campo.font.name = SANS


# ------------------------------------------------------------- BLOCOS ------

def paragrafo(doc, texto, size=11, cor=PRETO, italico=False, negrito=False,
              align=None, space_after=8):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(texto)
    run.font.size = Pt(size)
    run.font.color.rgb = cor
    run.font.italic = italico
    run.font.bold = negrito
    run.font.name = SANS
    return p


def rico(doc, pedacos, size=11, align=None, space_after=8):
    """pedacos = [(texto, negrito, mono, cor), ...]"""
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    for pedaco in pedacos:
        if isinstance(pedaco, str):
            pedaco = (pedaco,)
        txt = pedaco[0]
        negrito = pedaco[1] if len(pedaco) > 1 else False
        mono = pedaco[2] if len(pedaco) > 2 else False
        cor = pedaco[3] if len(pedaco) > 3 else PRETO
        run = p.add_run(txt)
        run.font.size = Pt(size - 0.5 if mono else size)
        run.font.bold = negrito
        run.font.color.rgb = cor
        run.font.name = MONO if mono else SANS
    return p


def lista(doc, itens, marcador="•"):
    for item in itens:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(f"{marcador}  ")
        run.font.color.rgb = INDIGO
        run.font.bold = True
        run.font.name = SANS
        if isinstance(item, str):
            r = p.add_run(item)
            r.font.name = SANS
            r.font.size = Pt(11)
        else:
            for pedaco in item:
                if isinstance(pedaco, str):
                    pedaco = (pedaco,)
                txt = pedaco[0]
                negrito = pedaco[1] if len(pedaco) > 1 else False
                mono = pedaco[2] if len(pedaco) > 2 else False
                r = p.add_run(txt)
                r.font.bold = negrito
                r.font.name = MONO if mono else SANS
                r.font.size = Pt(10.5 if mono else 11)


def codigo(doc, linhas, legenda=None):
    """Bloco de codigo em tabela de uma celula com fundo cinza."""
    tabela = doc.add_table(rows=1, cols=1)
    tabela.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = tabela.cell(0, 0)
    set_cell_bg(cell, "F4F4F7")
    cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    for i, linha in enumerate(linhas):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.1
        run = p.add_run(linha if linha else " ")
        run.font.name = MONO
        run.font.size = Pt(9.5)
        run.font.color.rgb = PRETO
    if legenda:
        paragrafo(doc, legenda, size=9, cor=CINZA_CLARO, italico=True,
                  space_after=12)
    else:
        paragrafo(doc, "", size=4, space_after=4)
    return tabela


def caixa(doc, titulo, texto, cor=INDIGO, fundo="EEF0FE"):
    """Caixa de destaque: dica, atencao, etc."""
    tabela = doc.add_table(rows=1, cols=1)
    cell = tabela.cell(0, 0)
    set_cell_bg(cell, fundo)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(titulo)
    run.font.bold = True
    run.font.size = Pt(10.5)
    run.font.color.rgb = cor
    run.font.name = SANS
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    run = p2.add_run(texto)
    run.font.size = Pt(10.5)
    run.font.name = SANS
    paragrafo(doc, "", size=4, space_after=6)
    return tabela


def h1(doc, texto):
    return doc.add_heading(texto, level=1)


def h2(doc, texto):
    return doc.add_heading(texto, level=2)


def h3(doc, texto):
    return doc.add_heading(texto, level=3)


# --------------------------------------------------------------- CAPA ------

def capa(doc):
    for _ in range(4):
        paragrafo(doc, "", space_after=0)

    paragrafo(doc, DATA.upper(), size=11, cor=INDIGO, negrito=True, space_after=4)
    p = doc.add_paragraph()
    run = p.add_run("TRABALHO\nNAVEGANDO ENTRE PÁGINAS")
    run.font.size = Pt(30)
    run.font.bold = True
    run.font.name = SANS
    run.font.color.rgb = PRETO
    p.paragraph_format.space_after = Pt(10)

    paragrafo(doc, "React Router  ·  Routes  ·  useNavigate  ·  Link",
              size=12, cor=CINZA, space_after=30)

    paragrafo(doc, AUTOR, size=12, negrito=True, space_after=2)
    paragrafo(doc, f"{DISCIPLINA}  ·  {ESCOLA}", size=11, cor=CINZA_CLARO)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def sumario(doc):
    h1(doc, "Sumário")
    p = doc.add_paragraph()
    add_field(p, r'TOC \o "1-2" \h \z \u')
    paragrafo(doc, "", size=9, space_after=4)
    paragrafo(doc,
              "O Word preenche este sumário ao abrir o arquivo. Se ele "
              "aparecer vazio, clique com o botão direito sobre ele e escolha "
              "\"Atualizar campo\".",
              size=9, cor=CINZA_CLARO, italico=True)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ------------------------------------------------------------ CONTEUDO -----

def conteudo(doc):
    # ------------------------------------------------------------------ 1 --
    h1(doc, "O que você vai entregar")
    paragrafo(doc,
              "Na aula 4 construímos a tela de login. Ela valida os campos, "
              "mostra a mensagem de erro e esconde a senha — mas quando o "
              "login dá certo ela só exibe um alerta e a aplicação para ali.")
    paragrafo(doc,
              "Agora esse login vai levar a algum lugar. Neste trabalho você "
              "vai instalar o React Router, transformar o App.jsx no mapa das "
              "telas da aplicação e criar páginas de verdade, cada uma com o "
              "seu endereço na barra do navegador.")

    paragrafo(doc, "São três entregas:", negrito=True, space_after=4)
    lista(doc, [
        [("A navegação funcionando: ao entrar com os dados corretos, o "
          "usuário sai do login e vai para a página "), ("Home", False, True),
         (".")],
        [("Uma "), ("terceira página", True), (" criada por você, com a rota "
          "dela registrada no App.jsx.")],
        [("Links de ida e volta entre as páginas, usando "), ("<Link>", False, True),
         (".")],
    ])

    caixa(doc, "Antes de tudo",
          "Este trabalho continua o projeto da aula 4. Faça uma cópia da pasta "
          "dele (sem a node_modules) e trabalhe na cópia — assim, se algo der "
          "errado, você ainda tem o login funcionando para consultar.")

    # ------------------------------------------------------------------ 2 --
    h1(doc, "Antes de começar")
    paragrafo(doc,
              "Abra a pasta do projeto no terminal e confira se ele ainda "
              "roda. Se a tela de login aparecer no navegador, está tudo "
              "certo para continuar.")
    codigo(doc, [
        "npm install",
        "npm run dev",
    ])

    paragrafo(doc,
              "Com o projeto rodando, pare o servidor (Ctrl + C no terminal) e "
              "instale a biblioteca de rotas:")
    codigo(doc, [
        "npm install react-router-dom",
    ], legenda="Depois de instalar, rode npm run dev de novo.")

    paragrafo(doc,
              "O react-router-dom não vem junto com o React: ele é uma "
              "biblioteca separada, mantida pela mesma comunidade. É por isso "
              "que precisamos instalá-lo à parte.")

    # ------------------------------------------------------------------ 3 --
    h1(doc, "Etapa 1 — Ligando o Router")

    h2(doc, "3.1  O BrowserRouter no main.jsx")
    paragrafo(doc,
              "Uma aplicação React é uma SPA: existe um único index.html, e é "
              "o React que decide qual componente desenhar dentro dele. O "
              "React Router é quem faz essa decisão olhar para a barra de "
              "endereços.")
    paragrafo(doc,
              "Para isso, ele precisa envolver a aplicação inteira. Abra o "
              "arquivo src/main.jsx e deixe-o assim:")
    codigo(doc, [
        'import React from "react";',
        'import ReactDOM from "react-dom/client";',
        'import { BrowserRouter } from "react-router-dom";',
        'import App from "./App";',
        'import "./index.css";',
        "",
        'ReactDOM.createRoot(document.getElementById("root")).render(',
        "  <React.StrictMode>",
        "    <BrowserRouter>",
        "      <App />",
        "    </BrowserRouter>",
        "  </React.StrictMode>",
        ");",
    ], legenda="src/main.jsx")

    paragrafo(doc,
              "Repare que só duas coisas mudaram: a linha do import e o abraço "
              "do <BrowserRouter> em volta do <App />. Nada mais.")

    h2(doc, "3.2  O App.jsx vira o mapa das telas")
    paragrafo(doc,
              "Na aula 4 o App.jsx só chamava o <Login />. Agora ele passa a "
              "listar todas as telas possíveis da aplicação:")
    codigo(doc, [
        'import { Routes, Route } from "react-router-dom";',
        'import Login from "./Login";',
        'import Home from "./Home";',
        "",
        "function App() {",
        "  return (",
        "    <Routes>",
        '      <Route path="/" element={<Login />} />',
        '      <Route path="/home" element={<Home />} />',
        "    </Routes>",
        "  );",
        "}",
        "",
        "export default App;",
    ], legenda="src/App.jsx")

    paragrafo(doc, "Duas peças novas aparecem aqui:", space_after=4)
    lista(doc, [
        [("<Routes>", False, True),
         (" é a caixa que guarda todas as telas. Ele olha a URL atual, escolhe "
          "UMA das rotas de dentro dele e desenha só ela.")],
        [("<Route>", False, True),
         (" é cada linha do mapa. O "), ("path", False, True),
         (" é o endereço; o "), ("element", False, True),
         (" é o componente que deve aparecer quando a URL for aquela.")],
    ])

    paragrafo(doc,
              'O path "/" é a raiz — é o que abre quando você entra no site '
              "sem digitar mais nada. Por isso o login mora nele.")

    caixa(doc, "O Home ainda não existe",
          "Se você salvar agora, a aplicação vai quebrar: estamos importando "
          "um arquivo Home.jsx que ainda não criamos. Isso é normal — a "
          "etapa 3 resolve. Se preferir, crie o arquivo vazio antes de "
          "continuar.")

    # ------------------------------------------------------------------ 4 --
    h1(doc, "Etapa 2 — Saindo do login")

    h2(doc, "4.1  O hook useNavigate")
    paragrafo(doc,
              "O Login.jsx continua igualzinho ao da aula passada: os mesmos "
              "useState, o mesmo handleSubmit, o mesmo botão de mostrar e "
              "ocultar a senha. A única novidade é o useNavigate.")
    paragrafo(doc,
              "Importe o hook no topo do arquivo e chame-o junto dos useState, "
              "dentro da função Login:")
    codigo(doc, [
        'import { useState } from "react";',
        'import { useNavigate } from "react-router-dom";',
        'import "./Login.css";',
        "",
        "function Login() {",
        '  const [email, setEmail] = useState("");',
        '  const [senha, setSenha] = useState("");',
        '  const [erro, setErro] = useState("");',
        "  const [verSenha, setVerSenha] = useState(false);",
        "",
        "  // o navigate troca de pagina pelo codigo",
        "  const navigate = useNavigate();",
    ], legenda="src/Login.jsx — só o topo do arquivo")

    paragrafo(doc,
              "O useNavigate é um Hook, assim como o useState. Como todo Hook, "
              "ele precisa ser chamado no topo do componente — nunca dentro de "
              "um if ou de outra função.")
    paragrafo(doc,
              "Ele não navega sozinho: devolve uma função (a que chamamos de "
              "navigate) que usamos quando quisermos trocar de página.")

    h2(doc, "4.2  Trocando o alert pela navegação")
    paragrafo(doc,
              "Agora vá até o handleSubmit. Aquele if que confere o e-mail e a "
              "senha continua exatamente o mesmo — só a linha do sucesso muda:")
    codigo(doc, [
        '    if (email === "user@email.com" && senha === "123456") {',
        '      setErro("");',
        '      navigate("/home");',
        "    } else {",
        '      setErro("E-mail ou senha incorretos.");',
        "    }",
    ], legenda="src/Login.jsx — dentro do handleSubmit")

    paragrafo(doc,
              'Onde estava alert("Login realizado com sucesso!") agora está o '
              "navigate. Em vez de uma caixinha do navegador, o usuário vai "
              "para outra tela.")
    paragrafo(doc,
              "O texto dentro do navigate é o mesmo path que escrevemos no "
              "Route lá no App.jsx. Se os dois não baterem, nada acontece — "
              "nem erro, nem tela nova.")

    caixa(doc, "Se der errado",
          "Três tropeços são os mais comuns nesta etapa. Confira nesta ordem: "
          "(1) esquecer o <BrowserRouter> no main.jsx — o React reclama que os "
          "componentes de rota estão fora de um Router; "
          "(2) chamar useNavigate() dentro do handleSubmit em vez do topo do "
          "componente; "
          '(3) escrever navigate("home") sem a barra — sem ela o React Router '
          "entende como caminho relativo e não encontra a rota.",
          cor=VERMELHO, fundo="FDEFEF")

    # ------------------------------------------------------------------ 5 --
    h1(doc, "Etapa 3 — A página Home")
    paragrafo(doc,
              "Chegou a hora de criar o lugar para onde o login está levando. "
              "Crie o arquivo src/Home.jsx:")
    codigo(doc, [
        'import { Link } from "react-router-dom";',
        'import "./Home.css";',
        "",
        "function Home() {",
        "  return (",
        '    <div className="container-home">',
        '      <div className="cartao">',
        "        <h2>Ponto Eletrônico</h2>",
        "        <p>Você entrou no sistema.</p>",
        '        <Link to="/">Sair</Link>',
        "      </div>",
        "    </div>",
        "  );",
        "}",
        "",
        "export default Home;",
    ], legenda="src/Home.jsx")

    paragrafo(doc,
              "Repare que a Home não tem nada de especial: é um componente "
              "comum, igual ao Login. Não existe nenhum código nela dizendo "
              '"eu sou a página /home" — quem faz essa ligação é o Route lá no '
              "App.jsx. A página não precisa saber o endereço dela.")

    h2(doc, "5.1  O Link e o <a href>")
    paragrafo(doc,
              "O <Link> é o link do React Router. Ele vira um <a> no "
              "navegador, mas com uma diferença importante: um <a href> comum "
              "RECARREGA a página inteira, e nisso o React sobe do zero e todo "
              "o estado da aplicação se perde.")
    paragrafo(doc,
              "Já o Link troca a tela sem recarregar nada. E repare que ele "
              "usa to, não href.")

    h2(doc, "5.2  O estilo da Home")
    paragrafo(doc,
              "Crie também o src/Home.css. Este é o estilo do exemplo da aula "
              "— use-o como ponto de partida e depois deixe do seu jeito:")
    codigo(doc, [
        ".container-home {",
        "  display: flex;",
        "  justify-content: center;",
        "  align-items: center;",
        "  height: 100vh;",
        "  background: #f0f2f5;",
        "}",
        "",
        ".cartao {",
        "  background: white;",
        "  padding: 2rem;",
        "  border-radius: 8px;",
        "  box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);",
        "  width: 300px;",
        "  text-align: center;",
        "}",
        "",
        ".cartao a {",
        "  display: inline-block;",
        "  margin-top: 1rem;",
        "  color: #4f46e5;",
        "  text-decoration: none;",
        "  font-weight: bold;",
        "}",
    ], legenda="src/Home.css")

    paragrafo(doc,
              "Note que a Home tem o próprio .container-home em vez de "
              "reaproveitar o .container do Login.css. O motivo é prático: se "
              "alguém abrir /home direto no navegador, o Login.jsx nunca é "
              "importado — e o estilo dele não chega junto.")

    paragrafo(doc, "Salve tudo e teste:", negrito=True, space_after=4)
    lista(doc, [
        "Entre com user@email.com e 123456. A URL deve mudar para /home sem a "
        "página piscar.",
        "Clique em Sair. Você volta para o login, também sem recarregar.",
        "Use a seta de voltar do navegador. Ela funciona — e não precisamos "
        "programar nada para isso.",
    ])

    # ------------------------------------------------------------------ 6 --
    h1(doc, "Etapa 4 — A sua rota nova")
    paragrafo(doc,
              "Esta é a parte autoral do trabalho. Agora que você já viu o "
              "caminho inteiro, faça sozinho: crie uma terceira página e "
              "coloque-a no mapa.")
    paragrafo(doc,
              "O assunto da página é escolha sua. Pode ser um Perfil com os "
              "seus dados, um Sobre explicando o sistema, uma tela de Ajuda — "
              "o que fizer sentido para a sua aplicação.")

    paragrafo(doc, "O roteiro é sempre o mesmo:", negrito=True, space_after=4)
    lista(doc, [
        [("Criar o componente, por exemplo "), ("src/Perfil.jsx", False, True),
         (", com o seu conteúdo e o "), ("export default", False, True),
         (" no final.")],
        [("Importar esse componente no App.jsx.")],
        [("Registrar a rota dentro do "), ("<Routes>", False, True),
         (", escolhendo o "), ("path", False, True), (" e o "),
         ("element", False, True), (".")],
        [("Colocar um "), ("<Link>", False, True),
         (" na Home levando até a página nova — e outro nela para voltar.")],
    ])

    paragrafo(doc, "O App.jsx vai ficar parecido com isto:")
    codigo(doc, [
        "    <Routes>",
        '      <Route path="/" element={<Login />} />',
        '      <Route path="/home" element={<Home />} />',
        '      <Route path="/perfil" element={<Perfil />} />',
        "    </Routes>",
    ], legenda="src/App.jsx — com a rota nova")

    paragrafo(doc, "E o Link na Home:")
    codigo(doc, [
        '        <Link to="/perfil">Meu perfil</Link>',
    ], legenda="src/Home.jsx")

    caixa(doc, "Rota não é segurança",
          "Digite /home direto na barra de endereços, sem fazer login: a "
          "página abre. Isso não é um defeito do React Router — é o que "
          "acontece quando a única verificação de acesso está no front-end. "
          "Quem decide se alguém pode ver uma tela é o servidor, e é isso que "
          "vamos construir no back-end em Python com FastAPI do sistema "
          "de ponto eletrônico.")

    # ------------------------------------------------------------------ 7 --
    h1(doc, "Como entregar")
    paragrafo(doc,
              "Compacte a pasta do projeto em um arquivo .zip, mas apague "
              "antes a pasta node_modules — ela tem milhares de arquivos e não "
              "precisa ser enviada. Quem for corrigir roda npm install e ela "
              "volta.")
    paragrafo(doc, "Nomeie o arquivo assim:", space_after=4)
    codigo(doc, [
        "SeuNome-Aula05-Rotas.zip",
    ])
    lista(doc, [
        "Prazo: ____________________",
        "Onde entregar: ____________________",
    ])

    # ------------------------------------------------------------------ 8 --
    h1(doc, "Critérios de avaliação")
    criterios = [
        ("O projeto roda sem erros com npm run dev", "15"),
        ("O BrowserRouter e as rotas estão configurados corretamente", "20"),
        ("O login leva para a Home usando navigate", "25"),
        ("A terceira página existe, tem rota e é alcançada por Link", "25"),
        ("O CSS da página nova é autoral", "15"),
        ("Total", "100"),
    ]
    tabela = doc.add_table(rows=len(criterios) + 1, cols=2)
    tabela.style = "Table Grid"
    cab = tabela.rows[0].cells
    for i, texto in enumerate(("Critério", "Pontos")):
        set_cell_bg(cab[i], "4F46E5")
        run = cab[i].paragraphs[0].add_run(texto)
        run.font.bold = True
        run.font.size = Pt(10.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.name = SANS
    for linha, (criterio, pontos) in enumerate(criterios, start=1):
        celulas = tabela.rows[linha].cells
        ultimo = criterio == "Total"
        if ultimo:
            set_cell_bg(celulas[0], "EEF0FE")
            set_cell_bg(celulas[1], "EEF0FE")
        for celula, texto in zip(celulas, (criterio, pontos)):
            run = celula.paragraphs[0].add_run(texto)
            run.font.size = Pt(10.5)
            run.font.bold = ultimo
            run.font.name = SANS
        celulas[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragrafo(doc, "", size=6, space_after=8)

    # ------------------------------------------------------------------ 9 --
    h1(doc, "Desafio extra (opcional)")
    paragrafo(doc,
              "Se sobrar tempo e vontade, escolha um destes. Nenhum vale nota "
              "— valem aprendizado.")
    lista(doc, [
        [("Criar uma rota "), ('path="*"', False, True),
         (" com uma página de \"endereço não encontrado\", para quando alguém "
          "digitar uma URL que não existe.")],
        [("Levar o e-mail digitado até a Home com "),
         ('navigate("/home", { state: { email } })', False, True),
         (" e lê-lo lá com o hook "), ("useLocation", False, True), (".")],
        [("Trocar o "), ("<Link>", False, True), (" de Sair por um botão que "
          "chama "), ('navigate("/")', False, True),
         (" — e pensar em qual dos dois faz mais sentido ali.")],
        "Criar um menu que aparece em todas as páginas internas, com os links "
        "para cada tela.",
    ])

    paragrafo(doc,
              "Na próxima aula começamos o outro lado da história: o back-end "
              "em Python com FastAPI — quem vai, de verdade, conferir a "
              "senha e decidir quem entra.",
              italico=True, cor=CINZA)


def build(destino):
    doc = Document()
    estilo_base(doc)

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    rodape_numerado(section)

    atualizar_campos_ao_abrir(doc)
    capa(doc)
    sumario(doc)
    conteudo(doc)

    doc.save(destino)
    print(f"OK: {destino}")


if __name__ == "__main__":
    saida = Path(__file__).with_name("Trabalho-Rotas-Aula05.docx")
    build(str(saida))
