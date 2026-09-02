# -*- coding: utf-8 -*-
"""
Gera os slides da Aula 05 - Navegando entre paginas (React Router).

Estilo: dark tech / editor de codigo — o mesmo sistema visual da Aula 04.
Conteudo: continua o projeto de login da aula anterior, agora com rotas
(14 slides).

Uso:
    python gerar_slides.py

Saida:
    Aula05-Navegando-Entre-Paginas.pptx  (editavel no PowerPoint)

Para mudar o visual, edite a secao PALETA / TIPOGRAFIA logo abaixo.
Para mudar o conteudo, edite a secao CONTEUDO DOS SLIDES no final do arquivo.
"""

import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


# ---------------------------------------------------------------- PALETA ----

BG = RGBColor(0x12, 0x12, 0x19)          # fundo
BG_HI = RGBColor(0x1D, 0x1A, 0x2E)       # topo do gradiente de fundo
SURFACE = RGBColor(0x1E, 0x1E, 0x2A)     # cartoes e painel de codigo
SURFACE_2 = RGBColor(0x24, 0x24, 0x32)   # barra de janela do painel
BORDER = RGBColor(0x2E, 0x2E, 0x3E)      # bordas
TEXT = RGBColor(0xE8, 0xE8, 0xF0)        # texto principal
TEXT_DIM = RGBColor(0x9A, 0x9A, 0xB0)    # texto secundario
TEXT_FAINT = RGBColor(0x6A, 0x6A, 0x80)  # texto apagado / codigo fora de foco
ACCENT = RGBColor(0x7C, 0x6C, 0xF5)      # indigo (mesma familia do #4f46e5 do CSS da aula)
ACCENT_DEEP = RGBColor(0x3B, 0x34, 0x70) # indigo escuro (decoracao)
ACCENT_2 = RGBColor(0x4E, 0xC9, 0xF0)    # ciano (tags JSX)
OK = RGBColor(0x4E, 0xC9, 0xA0)          # verde
ERR = RGBColor(0xF0, 0x71, 0x78)         # vermelho
HILITE = RGBColor(0x26, 0x22, 0x4A)      # faixa de destaque atras da linha em foco

# cores de sintaxe (familia VS Code Dark+)
C_KEYWORD = RGBColor(0xC5, 0x86, 0xC0)
C_STRING = RGBColor(0xCE, 0x91, 0x78)
C_FUNC = RGBColor(0xDC, 0xDC, 0xAA)
C_TAG = RGBColor(0x4E, 0xC9, 0xF0)
C_VAR = RGBColor(0x9C, 0xDC, 0xFE)
C_NUM = RGBColor(0xB5, 0xCE, 0xA8)
C_PUNCT = RGBColor(0xC8, 0xC8, 0xD4)
C_COMMENT = RGBColor(0x6A, 0x6A, 0x80)
C_SELECTOR = RGBColor(0xD7, 0xBA, 0x7D)


# ------------------------------------------------------------ TIPOGRAFIA ----

# Display: Bahnschrift condensada (geometrica, tecnica) — nada de Segoe UI
# generico nos titulos. Corpo: Segoe UI, escolhido pela legibilidade no
# projetor. Codigo: Cascadia Mono, a fonte do VS Code / Windows Terminal —
# vem do mundo do proprio assunto da aula.
DISPLAY = "Bahnschrift SemiBold Condensed"
SANS = "Segoe UI"
SANS_LIGHT = "Segoe UI Light"
MONO = "Cascadia Mono"

CODE_SIZE = Pt(13)
CODE_LH = Pt(19)          # altura exata de cada linha de codigo (usada no destaque)
CODE_LH_IN = 19 / 72.0    # a mesma coisa, em polegadas


# ---------------------------------------------------------------- CANVAS ----

W = 13.333   # polegadas
H = 7.5
M = 0.62     # margem lateral
TOTAL_SLIDES = 14


# -------------------------------------------------------- INSTITUCIONAL ----

ESCOLA = "Escola Técnica Mesquita"
PROFESSOR = "Emanuel Nascente"

# O logo e arte preta/vermelha/amarela sobre branco: nao fecha contraste
# sobre o dark do deck, e recortar o fundo apagaria o contorno preto do
# livro. Em vez disso a marca vira um SELO — placa clara com o mesmo raio
# das pilulas do deck — para pertencer ao sistema visual em vez de parecer
# colada nele. De quebra e o unico acento quente (vermelho/amarelo) num
# deck frio de indigo e ciano.
LOGO = Path(__file__).with_name("assets") / "logo-escola-tecnica-mesquita.png"
BRANCO = RGBColor(0xFF, 0xFF, 0xFF)


def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    return prs


def add_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # layout em branco
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = BG
    # camada de profundidade: gradiente diagonal sutil sobre o fundo chapado
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                                Inches(W), Inches(H))
    sp.line.fill.background()
    sp.shadow.inherit = False
    gradient(sp, BG_HI, BG, angle=315)
    return slide


def gradient(shape, cor_a, cor_b, angle=90):
    """Preenche a shape com um gradiente linear de cor_a para cor_b."""
    shape.fill.gradient()
    stops = shape.fill.gradient_stops
    stops[0].color.rgb = cor_a
    stops[0].position = 0.0
    stops[1].color.rgb = cor_b
    stops[1].position = 1.0
    for extra in list(stops)[2:]:
        extra.color.rgb = cor_b
        extra.position = 1.0
    shape.fill.gradient_angle = angle
    return shape


# --------------------------------------------------------------- HELPERS ----

def rect(slide, x, y, w, h, fill=None, line=None, line_w=Pt(1), radius=None,
         shape=MSO_SHAPE.RECTANGLE):
    """Retangulo (opcionalmente arredondado) em polegadas."""
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None:
        sp.adjustments[0] = radius
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = line_w
    sp.shadow.inherit = False
    sp.text_frame.text = ""
    return sp


def oval(slide, x, y, w, h, fill=None, line=None, line_w=Pt(1), rotation=0):
    sp = rect(slide, x, y, w, h, fill=fill, line=line, line_w=line_w,
              shape=MSO_SHAPE.OVAL)
    sp.rotation = rotation
    return sp


def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tf


def style_run(run, text, size, color, font=SANS, bold=False, italic=False):
    run.text = text
    run.font.size = size
    run.font.color.rgb = color
    run.font.name = font
    run.font.bold = bold
    run.font.italic = italic
    return run


def para(tf, text, size, color, font=SANS, bold=False, italic=False,
         align=PP_ALIGN.LEFT, space_before=0, space_after=0, line_spacing=None,
         first=False):
    """Adiciona um paragrafo simples ao text frame."""
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    if line_spacing is not None:
        p.line_spacing = line_spacing
    style_run(p.add_run(), text, size, color, font=font, bold=bold, italic=italic)
    return p


def rich_para(tf, pieces, size, font=SANS, align=PP_ALIGN.LEFT, space_before=0,
              space_after=0, line_spacing=None, first=False):
    """Paragrafo com varios trechos: pieces = [(texto, cor, bold, font_opcional), ...]"""
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    if line_spacing is not None:
        p.line_spacing = line_spacing
    for piece in pieces:
        txt, color = piece[0], piece[1]
        bold = piece[2] if len(piece) > 2 else False
        fnt = piece[3] if len(piece) > 3 else font
        style_run(p.add_run(), txt, size, color, font=fnt, bold=bold)
    return p


def body_block(slide, x, y, w, h, blocks, size=Pt(15), color=TEXT_DIM,
               line_spacing=1.35, space=10):
    """Paragrafos de texto corrido. Cada bloco pode ser str ou lista de trechos."""
    tf = textbox(slide, x, y, w, h)
    for i, block in enumerate(blocks):
        if isinstance(block, str):
            para(tf, block, size, color, line_spacing=line_spacing,
                 space_before=0 if i == 0 else space, first=(i == 0))
        else:
            rich_para(tf, block, size, line_spacing=line_spacing,
                      space_before=0 if i == 0 else space, first=(i == 0))
    return tf


# --------------------------------------------------- REALCE DE SINTAXE ------

KEYWORDS = {
    "import", "from", "function", "const", "let", "var", "return", "if", "else",
    "export", "default", "new", "true", "false", "null", "undefined", "typeof",
}
KNOWN_FUNCS = {
    "useState", "setEmail", "setSenha", "setErro", "setVerSenha", "alert",
    "preventDefault", "handleSubmit", "Login", "App", "console",
    "useNavigate", "navigate", "Home", "BrowserRouter", "Routes", "Route",
    "Link", "createRoot", "getElementById", "render",
}

JS_RE = re.compile(
    r"""(?P<comment>//[^\n]*)
      | (?P<string>"[^"]*"|'[^']*')
      | (?P<arrow>=>)
      | (?P<tag></?[A-Za-z][\w.]*|/>)
      | (?P<num>\b\d+(?:\.\d+)?\b)
      | (?P<ident>[A-Za-z_$][\w$]*)
      | (?P<ws>\s+)
      | (?P<punct>.)""",
    re.VERBOSE,
)


def tokenize_js(line):
    """Devolve [(texto, cor), ...] para uma linha de JS/JSX."""
    if line.strip() in ("...", "…"):
        return [(line, C_COMMENT)]
    out = []
    for m in JS_RE.finditer(line):
        kind = m.lastgroup
        txt = m.group()
        if kind == "comment":
            color = C_COMMENT
        elif kind == "string":
            color = C_STRING
        elif kind == "tag":
            color = C_TAG
        elif kind == "num":
            color = C_NUM
        elif kind == "ident":
            if txt in KEYWORDS:
                color = C_KEYWORD
            elif txt in KNOWN_FUNCS or line[m.end():m.end() + 1] == "(":
                color = C_FUNC
            else:
                color = C_VAR
        elif kind == "ws":
            color = C_PUNCT
        else:  # arrow, punct
            color = C_PUNCT
        out.append((txt, color))
    return out or [(line, C_PUNCT)]


CSS_VALUE_RE = re.compile(r"(#[0-9a-fA-F]{3,8}|\b\d+(?:\.\d+)?\w*\b|[^,\s]+)")


def tokenize_css(line):
    """Devolve [(texto, cor), ...] para uma linha de CSS."""
    stripped = line.strip()
    if not stripped:
        return [(line, C_PUNCT)]
    if stripped.startswith("/*"):
        return [(line, C_COMMENT)]
    if stripped == "}":
        return [(line, C_PUNCT)]
    if stripped.endswith("{"):
        head = line[: line.index("{")]
        return [(head, C_SELECTOR), ("{", C_PUNCT)]
    if ":" in line:
        prop, value = line.split(":", 1)
        out = [(prop, C_VAR), (":", C_PUNCT)]
        pos = 0
        for m in CSS_VALUE_RE.finditer(value):
            if m.start() > pos:
                out.append((value[pos:m.start()], C_PUNCT))
            token = m.group()
            if token.startswith("#") or re.match(r"^\d", token):
                out.append((token, C_NUM))
            else:
                out.append((token, C_STRING))
            pos = m.end()
        if pos < len(value):
            out.append((value[pos:], C_PUNCT))
        return out
    return [(line, C_PUNCT)]


# ------------------------------------------------------ PAINEL DE CODIGO ----

def code_panel(slide, x, y, w, h, filename, lines, highlight=None, lang="js",
               size=CODE_SIZE, line_h=CODE_LH):
    """
    Painel escuro com barra de janela e o codigo dentro.
    highlight = (primeira_linha, ultima_linha) com indice 0, inclusive.
    As linhas fora do destaque ficam apagadas.
    """
    bar_h = 0.42
    pad_x = 0.32
    pad_y = 0.26
    gutter_w = 0.46   # coluna dos numeros de linha, como num editor

    rect(slide, x, y, w, h, fill=SURFACE, line=BORDER,
         radius=0.035, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # barra de janela
    rect(slide, x, y, w, bar_h, fill=SURFACE_2, radius=0.09,
         shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
    rect(slide, x, y + bar_h - 0.02, w, 0.02, fill=BORDER)
    for i, dot in enumerate((ERR, RGBColor(0xE0, 0xB0, 0x60), OK)):
        oval(slide, x + 0.22 + i * 0.22, y + 0.155, 0.11, 0.11, fill=dot)
    tf = textbox(slide, x + 0.95, y + 0.09, w - 1.2, 0.26)
    para(tf, filename, Pt(11), TEXT_DIM, font=MONO, first=True)

    # faixa de destaque
    if highlight is not None:
        a, b = highlight
        top = y + bar_h + pad_y + a * CODE_LH_IN
        alt = (b - a + 1) * CODE_LH_IN + 0.09
        rect(slide, x + 0.02 + gutter_w, top - 0.045,
             w - 0.10 - gutter_w, alt, fill=HILITE)
        rect(slide, x + 0.02 + gutter_w, top - 0.045, 0.035, alt, fill=ACCENT)

    # coluna de numeros de linha
    rect(slide, x + 0.02, y + bar_h, gutter_w, h - bar_h - 0.03, fill=SURFACE_2)
    rect(slide, x + 0.02 + gutter_w, y + bar_h, 0.012, h - bar_h - 0.03,
         fill=BORDER)
    tfn = textbox(slide, x, y + bar_h + pad_y, gutter_w - 0.06, h - bar_h - pad_y)
    for i in range(len(lines)):
        foco = highlight is not None and highlight[0] <= i <= highlight[1]
        p = tfn.paragraphs[0] if i == 0 else tfn.add_paragraph()
        p.alignment = PP_ALIGN.RIGHT
        p.line_spacing = line_h
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        style_run(p.add_run(), str(i + 1), size,
                  ACCENT if foco else TEXT_FAINT, font=MONO, bold=foco)

    tokenizer = tokenize_css if lang == "css" else tokenize_js
    tf = textbox(slide, x + pad_x + gutter_w, y + bar_h + pad_y,
                 w - 2 * pad_x - gutter_w, h - bar_h - 2 * pad_y)
    for i, line in enumerate(lines):
        dim = highlight is not None and not (highlight[0] <= i <= highlight[1])
        if not line.strip():
            pieces = [(" ", C_PUNCT)]
        elif dim:
            pieces = [(line, TEXT_FAINT)]
        else:
            pieces = tokenizer(line)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = line_h
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        for txt, color in pieces:
            style_run(p.add_run(), txt, size, color, font=MONO)
    return tf


# ------------------------------------------------------- CROMO DO SLIDE ----

def eyebrow_pill(slide, texto, y=0.58):
    """Etiqueta da secao dentro de uma pilula com borda accent."""
    largura = 0.30 + len(texto) * 0.082
    rect(slide, M, y, largura, 0.32, fill=HILITE, line=ACCENT_DEEP,
         radius=0.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tf = textbox(slide, M, y, largura, 0.32, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, texto.upper(), Pt(10), ACCENT, bold=True,
         align=PP_ALIGN.CENTER, first=True)
    return y + 0.32


def selo_escola(slide, x, y, size, pad=0.06):
    """Selo institucional: placa clara com o logo centralizado dentro.

    Sem o arquivo do logo desenha so a placa — o deck continua gerando.
    """
    rect(slide, x, y, size, size, fill=BRANCO,
         radius=0.13, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    if LOGO.exists():
        lado = size - 2 * pad
        slide.shapes.add_picture(str(LOGO), Inches(x + pad), Inches(y + pad),
                                 Inches(lado), Inches(lado))
    else:
        print(f"AVISO: logo nao encontrado em {LOGO} — selo saiu vazio.")
    return x + size


def rule(slide, x, y, w_accent=1.15, w_total=3.1):
    """Risco de titulo: barra accent curta sobre um trilho fino."""
    rect(slide, x, y + 0.016, w_total, 0.018, fill=BORDER)
    rect(slide, x, y, w_accent, 0.05, fill=ACCENT)


def chrome(slide, n):
    """Rodape comum: identificacao, numero e regua de progresso."""
    y = H - 0.52
    _ = n
    tf = textbox(slide, M, y, 6.6, 0.26)
    rich_para(tf, [("Programação Web 2", TEXT_FAINT),
                   ("  ·  ", ACCENT),
                   ("Aula 05", TEXT_FAINT),
                   ("  ·  ", ACCENT),
                   (ESCOLA, TEXT_FAINT)], Pt(10), first=True)

    tf = textbox(slide, W - M - 2.0, y, 2.0, 0.26)
    rich_para(tf, [(f"{n:02d}", TEXT_DIM, True), (f" / {TOTAL_SLIDES}", TEXT_FAINT)],
              Pt(10), align=PP_ALIGN.RIGHT, first=True)

    track_w = W - 2 * M
    rect(slide, M, H - 0.16, track_w, 0.022, fill=BORDER)
    rect(slide, M, H - 0.16, track_w * n / TOTAL_SLIDES, 0.022, fill=ACCENT)


def header(slide, eyebrow, title, title_mono=False):
    """Cabecalho padrao: pilula da secao + titulo grande com risco de accent."""
    eyebrow_pill(slide, eyebrow)
    tf = textbox(slide, M, 1.02, W - 2 * M, 0.72)
    para(tf, title, Pt(36) if title_mono else Pt(40), TEXT,
         font=MONO if title_mono else DISPLAY, bold=True, first=True)
    rule(slide, M, 1.72)


# ---------------------------------------------------------------- LAYOUTS ---

def slide_capa(prs, eyebrow, titulo_1, titulo_2, subtitle, author):
    slide = add_slide(prs)

    # halo por tras do atomo, para dar profundidade
    cx, cy = 9.95, 3.85
    oval(slide, cx - 3.2, cy - 3.2, 6.4, 6.4, fill=BG_HI)

    rw, rh = 4.6, 1.72
    for rot in (0, 60, 120):
        oval(slide, cx - rw / 2, cy - rh / 2, rw, rh,
             line=ACCENT_DEEP, line_w=Pt(2.25), rotation=rot)
    oval(slide, cx - 0.26, cy - 0.26, 0.52, 0.52, fill=ACCENT)

    selo_escola(slide, M, 0.62, 1.06)

    eyebrow_pill(slide, eyebrow, y=1.92)

    tf = textbox(slide, M, 2.46, 7.8, 1.9)
    para(tf, titulo_1, Pt(60), TEXT, font=DISPLAY, bold=True,
         line_spacing=1.0, first=True)
    para(tf, titulo_2, Pt(60), ACCENT, font=DISPLAY, bold=True, line_spacing=1.0)

    rule(slide, M, 4.62, w_accent=1.6, w_total=4.2)

    tf = textbox(slide, M, 4.92, 7.4, 0.4)
    para(tf, subtitle, Pt(18), TEXT_DIM, first=True)

    tf = textbox(slide, M, 5.44, 7.4, 0.34)
    para(tf, author, Pt(13), TEXT_FAINT, first=True)

    rect(slide, M, H - 1.02, 1.42, 0.34, fill=HILITE, line=OK,
         radius=0.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tf = textbox(slide, M, H - 1.02, 1.42, 0.34, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, "REACT :D", Pt(11), OK, font=MONO, bold=True,
         align=PP_ALIGN.CENTER, first=True)

    rect(slide, M, H - 0.16, (W - 2 * M) / TOTAL_SLIDES, 0.022, fill=ACCENT)
    return slide


def slide_cartoes(prs, n, eyebrow, title, cards):
    """Tres cartoes numerados lado a lado."""
    slide = add_slide(prs)
    header(slide, eyebrow, title)

    gap = 0.34
    cw = (W - 2 * M - 2 * gap) / 3
    top = 2.06
    ch = 4.5

    for i, card in enumerate(cards):
        x = M + i * (cw + gap)
        rect(slide, x, top, cw, ch, fill=SURFACE, line=BORDER,
             radius=0.028, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(slide, x, top, cw, 0.045, fill=ACCENT)

        tf = textbox(slide, x + 0.34, top + 0.34, cw - 0.68, 0.42)
        para(tf, f"{i + 1:02d}", Pt(26), ACCENT_DEEP, font=MONO, bold=True, first=True)

        tf = textbox(slide, x + 0.34, top + 0.82, cw - 0.68, 0.36)
        para(tf, card["titulo"], Pt(22), TEXT, font=DISPLAY, bold=True, first=True)

        tf = textbox(slide, x + 0.34, top + 1.28, cw - 0.68, 2.0)
        para(tf, card["texto"], Pt(13), TEXT_DIM, line_spacing=1.34, first=True)

        # a caixa de comando comeca sempre na mesma altura nos tres cartoes
        box_top = top + 2.95
        if card.get("comando"):
            cmds = card["comando"]
            box_h = 0.28 + 0.23 * len(cmds)
            rect(slide, x + 0.24, box_top, cw - 0.48, box_h, fill=BG, line=BORDER,
                 radius=0.06, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            tf = textbox(slide, x + 0.38, box_top + 0.12, cw - 0.70, box_h - 0.24)
            for j, cmd in enumerate(cmds):
                # linha comecando com espaco = continuacao do comando anterior
                prefix = ("  ", ACCENT) if cmd.startswith(" ") else ("$ ", ACCENT)
                rich_para(tf, [prefix, (cmd.strip(), OK)], Pt(10),
                          font=MONO, line_spacing=Pt(16), first=(j == 0))
            box_top += box_h

        if card.get("nota"):
            tf = textbox(slide, x + 0.34, box_top + 0.16, cw - 0.68, 0.5)
            para(tf, card["nota"], Pt(10.5), TEXT_FAINT, italic=True,
                 line_spacing=1.2, first=True)

    chrome(slide, n)
    return slide


LABELS_JSX = [("JavaScript", ".js", C_KEYWORD), ("HTML", "<tags>", C_TAG),
              ("JSX", ".jsx", ACCENT)]


def slide_conceito(prs, n, eyebrow, title, destaque, blocos, labels=None):
    """Um conceito, muito respiro.

    A ilustracao da direita e uma soma de tres termos (A + B = C); o ultimo
    vem realcado. Passe `labels` como [(nome, detalhe, cor), ...].
    """
    slide = add_slide(prs)
    header(slide, eyebrow, title)

    rect(slide, M, 2.25, 6.05, 1.2, fill=SURFACE, line=BORDER,
         radius=0.05, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(slide, M, 2.25, 0.055, 1.2, fill=ACCENT)
    tf = textbox(slide, M + 0.42, 2.25, 5.35, 1.2, anchor=MSO_ANCHOR.MIDDLE)
    rich_para(tf, destaque, Pt(19), line_spacing=1.25, first=True)

    body_block(slide, M, 3.85, 6.05, 2.55, blocos, size=Pt(15))

    # ilustracao: A + B = C
    bx = 7.55
    bw = 5.15
    rect(slide, bx, 2.25, bw, 4.2, fill=SURFACE, line=BORDER,
         radius=0.03, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    labels = labels or LABELS_JSX
    row_h = 1.0
    op_h = 0.32
    bloco_h = 3 * row_h + 2 * op_h
    y0 = 2.25 + (4.2 - bloco_h) / 2
    for i, (nome, ext, cor) in enumerate(labels):
        y = y0 + i * (row_h + op_h)
        if i == 2:
            rect(slide, bx + 0.5, y, bw - 1.0, row_h, fill=HILITE,
                 radius=0.1, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        tf = textbox(slide, bx + 0.5, y + 0.12, bw - 1.0, 0.42)
        para(tf, nome, Pt(19), TEXT if i == 2 else TEXT_DIM, bold=(i == 2),
             align=PP_ALIGN.CENTER, first=True)
        tf = textbox(slide, bx + 0.5, y + 0.56, bw - 1.0, 0.32)
        para(tf, ext, Pt(13), cor, font=MONO, align=PP_ALIGN.CENTER, first=True)
        if i < 2:
            tf = textbox(slide, bx + 0.5, y + row_h, bw - 1.0, op_h,
                         anchor=MSO_ANCHOR.MIDDLE)
            para(tf, "+" if i == 0 else "=", Pt(17), ACCENT, font=MONO,
                 bold=True, align=PP_ALIGN.CENTER, first=True)

    chrome(slide, n)
    return slide


AREA_TOP = 1.78          # topo da area de conteudo dos slides de codigo
AREA_H = 4.78            # altura util ate o rodape


def slide_codigo(prs, n, eyebrow, termo, filename, lines, highlight, blocos,
                 lang="js", nota=None):
    """Layout principal: codigo a esquerda, explicacao a direita.

    A altura do painel acompanha o numero de linhas e ele fica centralizado
    verticalmente na area de conteudo — nada de painel meio vazio.
    """
    slide = add_slide(prs)

    eyebrow_pill(slide, eyebrow)
    tf = textbox(slide, M, 1.02, 6.55, 0.66)
    para(tf, termo, Pt(31), TEXT, font=MONO, bold=True, first=True)
    rule(slide, M, 1.72)

    code_h = min(AREA_H, 0.42 + 0.52 + len(lines) * CODE_LH_IN)
    code_y = AREA_TOP + (AREA_H - code_h) / 2
    code_panel(slide, M, code_y, 6.55, code_h, filename, lines,
               highlight=highlight, lang=lang)

    x = 7.55
    w = W - M - x
    rect(slide, x, AREA_TOP, 0.045, AREA_H, fill=BORDER)
    body_block(slide, x + 0.34, AREA_TOP + 0.08, w - 0.34, AREA_H - 1.3,
               blocos, size=Pt(14.5))

    if nota:
        ny = AREA_TOP + AREA_H - 0.98
        rect(slide, x + 0.34, ny, w - 0.34, 0.98, fill=SURFACE, line=BORDER,
             radius=0.05, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        tf = textbox(slide, x + 0.58, ny + 0.17, w - 0.82, 0.66)
        rich_para(tf, nota, Pt(12.5), line_spacing=1.25, first=True)

    chrome(slide, n)
    return slide


def slide_codigo_duplo(prs, n, eyebrow, termo, filename, col_a, col_b, lang="css"):
    """Codigo longo em duas colunas."""
    slide = add_slide(prs)

    eyebrow_pill(slide, eyebrow)
    tf = textbox(slide, M, 1.02, 8.0, 0.62)
    para(tf, termo, Pt(31), TEXT, font=MONO, bold=True, first=True)

    gap = 0.3
    cw = (W - 2 * M - gap) / 2
    code_panel(slide, M, 1.62, cw, 5.05, filename, col_a, lang=lang,
               size=Pt(11.5), line_h=Pt(15.5))
    code_panel(slide, M + cw + gap, 1.62, cw, 5.05, filename, col_b, lang=lang,
               size=Pt(11.5), line_h=Pt(15.5))

    chrome(slide, n)
    return slide


def slide_passos(prs, n, eyebrow, title, intro, passos, marcadores=None,
                 destaque_ultimo=False):
    """Passos numerados (A/B/C ou 01/02/03), um por linha."""
    slide = add_slide(prs)
    header(slide, eyebrow, title)

    y = 1.95
    if intro:
        tf = textbox(slide, M, y, 9.6, 1.05)
        para(tf, intro, Pt(15), TEXT_DIM, line_spacing=1.35, first=True)
        y += 1.20

    marcadores = marcadores or [f"{i + 1:02d}" for i in range(len(passos))]
    row_h = 0.98
    gap = 0.22
    # centraliza o bloco de passos no espaco que sobrou
    bloco_h = len(passos) * (row_h + gap) - gap
    y = y + max(0.0, (6.72 - y - bloco_h) / 2)
    for i, passo in enumerate(passos):
        top = y + i * (row_h + gap)
        realce = destaque_ultimo and i == len(passos) - 1
        rect(slide, M, top, W - 2 * M, row_h,
             fill=HILITE if realce else SURFACE,
             line=ACCENT if realce else BORDER,
             radius=0.09, shape=MSO_SHAPE.ROUNDED_RECTANGLE)

        tf = textbox(slide, M + 0.36, top, 0.9, row_h, anchor=MSO_ANCHOR.MIDDLE)
        para(tf, marcadores[i], Pt(24), ACCENT, font=MONO, bold=True, first=True)

        tw = W - 2 * M - 1.5
        tf = textbox(slide, M + 1.36, top, tw, row_h, anchor=MSO_ANCHOR.MIDDLE)
        if isinstance(passo, str):
            para(tf, passo, Pt(16), TEXT, line_spacing=1.25, first=True)
        else:
            rich_para(tf, passo, Pt(16), line_spacing=1.25, first=True)

    chrome(slide, n)
    return slide


# ------------------------------------------------- CONTEUDO DOS SLIDES ------
#
# Os blocos abaixo sao o codigo que aparece nos paineis. Regras de largura:
# no painel largo cabem ~48 caracteres por linha e ~14 linhas; no painel duplo
# (CSS) cabem ~50 caracteres e ~19 linhas. Linhas mais longas que isso quebram
# e desalinham a numeracao do gutter — quebre a mao, como no if do slide 09.
#
# "  ..." marca continuidade: o arquivo segue, mas nao interessa agora.

CODE_MAIN_TOPO = [
    'import React from "react";',
    'import ReactDOM from "react-dom/client";',
    'import { BrowserRouter } from "react-router-dom";',
    'import App from "./App";',
    'import "./index.css";',
    "",
    "  ...",
]

CODE_MAIN_RENDER = [
    "  ...",
    "",
    "ReactDOM.createRoot(",
    '  document.getElementById("root")',
    ").render(",
    "  <React.StrictMode>",
    "    <BrowserRouter>",
    "      <App />",
    "    </BrowserRouter>",
    "  </React.StrictMode>",
    ");",
]

CODE_APP = [
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
]

CODE_LOGIN_TOPO = [
    'import { useState } from "react";',
    'import { useNavigate } from "react-router-dom";',
    'import "./Login.css";',
    "",
    "function Login() {",
    '  const [email, setEmail] = useState("");',
    '  const [senha, setSenha] = useState("");',
    '  const [erro, setErro] = useState("");',
    "",
    "  // o navigate troca de pagina pelo codigo",
    "  const navigate = useNavigate();",
    "",
    "  ...",
]

CODE_NAVIGATE = [
    "  ...",
    "",
    '    if (email === "user@email.com" &&',
    '        senha === "123456") {',
    '      setErro("");',
    '      navigate("/home");',
    "    } else {",
    '      setErro("E-mail ou senha incorretos.");',
    "    }",
    "  }",
    "",
    "  ...",
]

CODE_HOME = [
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
]

CSS_HOME_A = [
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
    "  box-shadow: 0 2px 10px rgba(0,0,0,0.1);",
    "  width: 300px;",
    "  text-align: center;",
    "}",
]

CSS_HOME_B = [
    ".cartao h2 {",
    "  margin-top: 0;",
    "}",
    "",
    ".cartao a {",
    "  display: inline-block;",
    "  margin-top: 1rem;",
    "  color: #4f46e5;",
    "  text-decoration: none;",
    "  font-weight: bold;",
    "}",
    "",
    ".cartao a:hover {",
    "  text-decoration: underline;",
    "}",
]


def build(destino):
    prs = new_deck()

    # 01 - capa
    slide_capa(
        prs,
        eyebrow="Programação Web 2  ·  Aula 05",
        titulo_1="Navegando entre",
        titulo_2="páginas",
        subtitle="Hoje o nosso login vai levar a algum lugar…",
        author=f"{PROFESSOR}  ·  {ESCOLA}",
    )

    # 02 - tres cartoes
    slide_cartoes(
        prs, 2, "Navegando entre páginas", "Preparando o terreno",
        [
            {
                "titulo": "Instalando",
                "texto": "Na aula passada montamos a tela de login. Vamos "
                         "continuar naquele mesmo projeto: abra a pasta dele no "
                         "terminal e instale a biblioteca de rotas do React.",
                "comando": ["npm install react-router-dom"],
                "nota": "Depois de instalar, rode npm run dev de novo.",
            },
            {
                "titulo": "Criando",
                "texto": "Se o login vai levar para algum lugar, esse lugar "
                         "precisa existir. Vamos criar dois arquivos novos "
                         "dentro da pasta src: a nossa segunda página e o "
                         "estilo dela.",
                "comando": ["Home.jsx", "Home.css"],
            },
            {
                "titulo": "Ligando",
                "texto": "E vamos mexer em dois arquivos que já existem: o "
                         "main.jsx, para ligar o sistema de rotas, e o App.jsx, "
                         "que vira o mapa das telas da aplicação.",
                "comando": ["main.jsx", "App.jsx"],
            },
        ],
    )

    # 03 - conceito SPA
    slide_conceito(
        prs, 3, "React Router", "Uma página só, várias telas",
        [("A página ", TEXT_DIM),
         ("nunca recarrega", TEXT, True),
         (" — o React troca o que aparece na tela.", TEXT_DIM)],
        ["Uma aplicação React é uma SPA (Single Page Application): existe um "
         "único index.html, e é o React que decide qual componente desenhar "
         "dentro dele.",
         "O React Router é quem faz essa decisão olhar para a barra de "
         "endereços. Cada endereço vira uma tela, e mudar de tela não custa um "
         "F5: nada pisca, nada é baixado de novo."],
        labels=[("URL", "/home", C_STRING),
                ("Route", 'path="/home"', C_TAG),
                ("Componente", "<Home />", ACCENT)],
    )

    # 04 - import
    slide_codigo(
        prs, 4, "Ligando o Router  ·  main.jsx", "import",
        "main.jsx", CODE_MAIN_TOPO, (2, 2),
        ["Tudo começa importando o que acabamos de instalar. Repare que o nome "
         "entre chaves é BrowserRouter, mas o pacote se chama react-router-dom.",
         [("Por que ", TEXT_DIM), ("Browser", TEXT, True),
          ("Router? Porque ele usa o recurso de histórico do navegador — o "
           "mesmo que faz as setas de voltar e avançar funcionarem.", TEXT_DIM)],
         "É por isso que, depois de logar, o botão de voltar do navegador leva "
         "de volta para o login. Não precisamos programar nada para isso."],
    )

    # 05 - BrowserRouter
    slide_codigo(
        prs, 5, "Ligando o Router  ·  main.jsx", "BrowserRouter",
        "main.jsx", CODE_MAIN_RENDER, (6, 8),
        ["O BrowserRouter é um componente que ABRAÇA a aplicação inteira. Ele "
         "não desenha nada na tela: só fica de olho na barra de endereços.",
         [("Regra de ouro: tudo que usa rota precisa estar ", TEXT_DIM),
          ("dentro", TEXT, True),
          (" do BrowserRouter. Por isso ele vem aqui no main.jsx, envolvendo "
           "o ", TEXT_DIM),
          ("<App />", C_TAG, True, MONO), (".", TEXT_DIM)],
         "Esquecer esse abraço é o erro mais comum da aula: o React acusa que "
         "os componentes de rota estão fora de um Router."],
    )

    # 06 - Routes
    slide_codigo(
        prs, 6, "O mapa da aplicação  ·  App.jsx", "Routes",
        "App.jsx", CODE_APP, (6, 9),
        ["Na aula passada o App.jsx só chamava o <Login />. Agora ele vira o "
         "mapa da nossa aplicação.",
         [("O ", TEXT_DIM), ("<Routes>", C_TAG, True, MONO),
          (" é a caixa que guarda todas as telas possíveis. Ele olha a URL "
           "atual, escolhe UMA das rotas de dentro dele e desenha só ela.",
           TEXT_DIM)],
         "Uma de cada vez — nunca duas ao mesmo tempo."],
    )

    # 07 - Route
    slide_codigo(
        prs, 7, "O mapa da aplicação  ·  App.jsx", "Route",
        "App.jsx", CODE_APP, (7, 8),
        ["Cada Route é uma linha do mapa, e ela sempre tem duas partes.",
         [("path", C_VAR, True, MONO),
          (" é o endereço que aparece na barra do navegador. ", TEXT_DIM),
          ("element", C_VAR, True, MONO),
          (" é o componente que deve aparecer quando a URL for aquela.",
           TEXT_DIM)],
         'O path "/" é a raiz: é o que abre quando você entra no site sem '
         "digitar mais nada. Por isso o login mora nele."],
        nota=[("Repare na barra: ", TEXT_DIM),
              ('"/home"', C_STRING, True, MONO),
              (" e não ", TEXT_DIM),
              ('"home"', C_STRING, True, MONO),
              (". Sem a barra vira caminho relativo.", TEXT_DIM)],
    )

    # 08 - useNavigate
    slide_codigo(
        prs, 8, "Saindo do login  ·  Login.jsx", "useNavigate",
        "Login.jsx", CODE_LOGIN_TOPO, (9, 10),
        ["O Login.jsx continua igualzinho ao da aula passada: os mesmos "
         "useState, o mesmo handleSubmit, o mesmo botão de mostrar senha.",
         [("A única novidade é o ", TEXT_DIM),
          ("useNavigate", C_FUNC, True, MONO),
          (". Ele é um Hook, assim como o useState — e como todo Hook, "
           "precisa ser chamado no topo do componente, nunca dentro de um if.",
           TEXT_DIM)],
         "Ele não navega sozinho: devolve uma função (a que chamamos de "
         "navigate) que usamos quando quisermos trocar de página."],
    )

    # 09 - navigate
    slide_codigo(
        prs, 9, "Saindo do login  ·  Login.jsx", 'navigate("/home")',
        "Login.jsx", CODE_NAVIGATE, (5, 5),
        ["Lembra desse if da aula passada? Ele é exatamente o mesmo. Só a "
         "linha do sucesso mudou.",
         [("Onde estava ", TEXT_DIM),
          ('alert("Login realizado…")', C_FUNC, True, MONO),
          (" agora está o navigate. Em vez de uma caixinha do navegador, o "
           "usuário vai para outra tela.", TEXT_DIM)],
         "O texto dentro do navigate é o mesmo path que escrevemos no Route lá "
         "no App.jsx. Se os dois não baterem, nada acontece."],
        nota=[("navigate", C_FUNC, True, MONO),
              (" é para navegar por código, depois de uma decisão. Para "
               "navegar por clique existe o Link — próximo slide.", TEXT_DIM)],
    )

    # 10 - Home.jsx
    slide_codigo(
        prs, 10, "A página de destino", "Home.jsx",
        "Home.jsx", CODE_HOME, None,
        ["Esta é a nossa segunda tela. Repare que ela não tem nada de "
         "especial: é um componente comum, igual ao Login.",
         'Não existe nenhum código dizendo "eu sou a página /home". Quem faz '
         "essa ligação é o Route lá no App.jsx — a página não precisa saber o "
         "endereço dela.",
         "Por enquanto ela só dá as boas-vindas. Nas próximas aulas é aqui que "
         "vai nascer o registro de ponto."],
        nota=[("export default Home;", C_FUNC, True, MONO),
              ("\nComo sempre, essa linha no final é o que permite importar o "
               "componente no App.jsx.", TEXT_DIM)],
    )

    # 11 - Link
    slide_codigo(
        prs, 11, "A página de destino  ·  Home.jsx", "Link",
        "Home.jsx", CODE_HOME, (9, 9),
        [[("O ", TEXT_DIM), ("<Link>", C_TAG, True, MONO),
          (" é o link do React Router. Ele vira um ", TEXT_DIM),
          ("<a>", C_TAG, True, MONO),
          (" no navegador, mas com um detalhe importante.", TEXT_DIM)],
         "Um <a href> comum RECARREGA a página inteira: o React sobe do zero e "
         "todo o estado da aplicação se perde.",
         [("Já o Link troca a tela sem recarregar nada. E repare que ele usa ",
           TEXT_DIM), ("to", C_VAR, True, MONO), (", não ", TEXT_DIM),
          ("href", C_VAR, True, MONO), (".", TEXT_DIM)]],
    )

    # 12 - CSS da Home
    slide_codigo_duplo(
        prs, 12, "A página de destino", "Home.css", "Home.css",
        CSS_HOME_A, CSS_HOME_B, lang="css",
    )

    # 13 - recapitulando
    slide_passos(
        prs, 13, "Recapitulando", "O caminho inteiro, em quatro passos", "",
        [
            [("Instalar: ", TEXT),
             ("npm install react-router-dom", C_FUNC, True, MONO)],
            [("Abraçar o App com ", TEXT), ("<BrowserRouter>", C_TAG, True, MONO),
             (" no main.jsx.", TEXT)],
            [("Mapear as telas com ", TEXT), ("<Routes>", C_TAG, True, MONO),
             (" e ", TEXT), ("<Route>", C_TAG, True, MONO),
             (" no App.jsx.", TEXT)],
            [("Navegar: ", TEXT), ("navigate()", C_FUNC, True, MONO),
             (" por código, ", TEXT), ("<Link>", C_TAG, True, MONO),
             (" por clique.", TEXT)],
        ],
    )

    # 14 - trabalho
    slide_passos(
        prs, 14, "E agora?", "O teu trabalho", "",
        [
            [("Criar uma ", TEXT), ("terceira página", TEXT, True),
             (" no projeto — o assunto dela é escolha tua.", TEXT)],
            [("Registrar a rota dela no App.jsx, com o ", TEXT),
             ("path", C_VAR, True, MONO), (" e o ", TEXT),
             ("element", C_VAR, True, MONO), (" que você quiser.", TEXT)],
            [("Colocar um ", TEXT), ("<Link>", C_TAG, True, MONO),
             (" na Home levando até ela — e outro de volta.", TEXT)],
        ],
        destaque_ultimo=False,
    )

    prs.save(destino)
    print(f"OK: {destino}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    saida = Path(__file__).with_name("Aula05-Navegando-Entre-Paginas.pptx")
    build(str(saida))
