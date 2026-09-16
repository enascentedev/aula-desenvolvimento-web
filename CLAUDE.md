# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## O que é este repositório

Material didático da disciplina **Programação Web 2** (Escola Técnica Mesquita), ministrada por **Emanuel Nascente**. Não é um produto de software: é um acervo de aulas, uma pasta por aula (`aula-3/`, `aula-4/`, `aula-5/`…), cada uma contendo slides, apostila e o código-exemplo daquele encontro.

O fio condutor do curso é um **sistema de ponto eletrônico** (React no front, Python + FastAPI + banco no back). Material novo deve manter essa continuidade e ser assinado como Emanuel Nascente — os PDFs antigos de `aula-3/` trazem "Luiz Gomes", o professor anterior.

Toda assinatura institucional usa **Escola Técnica Mesquita** e **Emanuel Nascente**, sem exceção: capa, rodapé de slide e rodapé de apostila. Nas apostilas isso é a constante `ESCOLA`; nos slides, `ESCOLA` e `PROFESSOR`.

O repositório é git (branch `main`), sem suíte de testes. O `README.md` da raiz é a porta de entrada para os alunos — mantenha a tabela de aulas dele em dia quando uma aula nova entrar. `.agents/skills/` e `.claude/` ficam fora do versionamento; a origem das skills baixadas está em `skills-lock.json`.

## Os tipos de artefato

### 1. Material da aula é gerado por script Python

**A fonte de verdade é o `.py`; o `.pptx`/`.docx` é saída.** Nunca edite o binário à mão — altere o script e regenere.

```bash
python aula-4/slides/gerar_slides.py         # → Aula04-Primeira-Pagina-React.pptx (deck da exposição)
python aula-4/slides/gerar_slides_final.py   # → Aula04-Primeira-Pagina-React-final.pptx (guia passo a passo)
python aula-4/apostila/gerar_apostila.py     # → Trabalho-Login-Aula04.docx
python aula-5/slides/gerar_slides.py         # → Aula05-Navegando-Entre-Paginas.pptx
python aula-5/slides/conferir_codigo.py      # confere slides × exemplo-rotas/src
python aula-5/apostila/gerar_apostila.py     # → Trabalho-Rotas-Aula05.docx
python aula-6/slides/gerar_slides.py         # → Aula06-API-Login-FastAPI.pptx
python aula-6/slides/conferir_codigo.py      # confere slides × api-login (main.py, pyproject.toml)
```

Os PDFs dos slides que ficam ao lado do `.pptx` saem do PowerPoint (ver restrições de Office abaixo); regenere-os junto quando o deck mudar.

Dependências: `python-pptx`, `python-docx` (já instaladas globalmente; não há requirements.txt).

`gerar_slides.py` está organizado em camadas, do genérico ao específico — mexa na camada certa:

| Seção | Serve para |
|---|---|
| PALETA / TIPOGRAFIA | mudar o visual (dark tech, família VS Code Dark+) |
| HELPERS | primitivas (`rect`, `oval`, `textbox`, `para`, `rich_para`) |
| REALCE DE SINTAXE | `tokenize_js` / `tokenize_css` — colorização própria, sem lib externa; a aula 6 acrescenta `tokenize_py`, `tokenize_shell`, `tokenize_toml` e `tokenize_arvore`, escolhidos pelo `lang` do painel via `TOKENIZADORES` |
| PAINEL DE CÓDIGO | `code_panel` — janela de editor com gutter de números e faixa de destaque (na aula 4, `first_line` continua a numeração quando uma listagem longa se divide em vários painéis) |
| CROMO / LAYOUTS | `chrome`, `header`, `selo_escola`, `slide_capa`, `slide_cartoes`, `slide_conceito`, `slide_codigo`, `slide_codigo_duplo`, `slide_passos` |
| LAYOUTS DA AULA 06 | `slide_terminal` (comando + saída esperada), `slide_imagem` (print do navegador; `topo=True` para prints baixos e largos), `slide_tabela`, `slide_fluxo` (cliente ↔ servidor), `slide_comparar` |
| CONTEUDO DOS SLIDES + `build()` | o texto de cada slide, em ordem |

Cada aula tem a **sua própria cópia** de `gerar_slides.py` e `gerar_apostila.py` — não há módulo compartilhado entre aulas. Para uma aula nova, copie os geradores da aula mais recente e troque só a camada de conteúdo; **reutilize esses helpers em vez de recomeçar do zero**. Dentro de uma mesma aula, um deck derivado importa os helpers do principal (é o que `aula-4/slides/gerar_slides_final.py` faz com `gerar_slides`, acrescentando `slide_arvore` e `slide_duplo`).

O selo da escola na capa vem de `slides/assets/logo-escola-tecnica-mesquita.png` (constante `LOGO`); sem o arquivo o deck ainda gera, com o selo vazio e um aviso no terminal.

Fontes confirmadas nesta máquina: Bahnschrift SemiBold Condensed (display), Cascadia Mono (código), Segoe UI, Consolas.

A aula 4 tem dois decks: o enxuto, da exposição, e o `-final`, guia de construção para consulta em casa. A partir da aula 5 os slides já são o **guia de construção** do aluno: ele tem que
conseguir digitar o projeto inteiro olhando só para o telão. Quem garante isso é
`aula-5/slides/conferir_codigo.py`, que compara os painéis de código com
`exemplo-rotas/src` nas duas direções — *fidelidade* (o que está no slide existe
no projeto) e *cobertura* (o que está no projeto aparece em algum slide) — e
ainda checa se cada bloco cabe no painel. Rode-o depois de mexer no gerador ou
no exemplo; ele sai com código 1 se achar divergência. Quando um slide mostra de
propósito um estágio anterior do arquivo, declare a linha em
`so_neste_estagio`, com o slide onde isso acontece.

A aula 6 segue a mesma regra no back-end: `aula-6/slides/conferir_codigo.py`
confere os painéis contra `api-login/main.py` e `pyproject.toml`, exige que o
slide do arquivo inteiro bata linha a linha com o `main.py` (em Python a
indentação é código), que a versão do FastAPI citada nos comandos seja a
fixada no `pyproject.toml`, que nenhuma linha passe da largura do painel e que a continuidade dos painéis seja `# ...` (comentário), nunca `...` solto — o aluno que digita tudo o que vê leva `IndentationError`. O slide 24 ensina a ler os painéis (faixa roxa = digitar agora; linhas apagadas = já digitadas): quem redigita o `def login` fica com uma API que devolve 200 para qualquer senha, sem erro nenhum. Os
prints do Swagger em `aula-6/slides/assets/swagger-*.png` foram capturados com
Playwright da API rodando e recortados; se a rota mudar, capture de novo.

`gerar_apostila.py` segue o mesmo espírito para A4 (capa, sumário como campo do Word, numeração de página). O sumário só preenche as páginas quando o usuário abre o `.docx` e escolhe "Atualizar campo".

### 2. Código-exemplo React (`aula-N/exemplo-*/`)

```bash
cd aula-5/exemplo-rotas      # ou aula-4/exemplo-login
npm install
npm run dev                  # Vite em http://localhost:5173
npm run build                # produção em dist/
npm run preview
```

Stack fixa e propositalmente mínima: React 18.3 + Vite 5, JSX, CSS puro por componente, sem gerenciador de estado, sem TypeScript, sem linter. Dependências entram só quando a aula as ensina: `react-router-dom` 6.28 aparece a partir da aula 5 (`exemplo-rotas`); a aula 4 não tem roteador. Versões ficam fixadas no `package.json`, sem `^`. Cada exemplo é um projeto npm independente — não há workspace na raiz.

### 3. Back-end em Python (`aula-6/api-login/`)

```powershell
cd aula-6/api-login
uv sync                      # cria .venv e instala o que o uv.lock pede
.venv\Scripts\activate       # ou, sem ativar: uv run fastapi dev main.py
fastapi dev main.py          # http://127.0.0.1:8000/docs (Swagger)
```

Projeto gerenciado pelo **uv** (0.6+), Python 3.12 fixado em `.python-version`, FastAPI fixado com `==` no `pyproject.toml` (`fastapi[standard]==0.141.1`) — mesma regra do "sem `^`" dos exemplos npm. Pacote entra e sai com `uv add` / `uv remove`; `uv.lock` vai para o Git, `.venv/` não (está no `.gitignore` da raiz). Um arquivo só (`main.py`) com uma rota só (`POST /login`: 200, 401 ou 422 do Pydantic). As linhas do `main.py` são curtas de propósito (`email_certo` / `senha_certa`, dicionário aberto em linhas) para caber no painel do slide — não "simplifique" juntando.

## Restrições de automação Office nesta máquina

- **PowerPoint aceita automação COM.** Use para revisar o resultado antes de entregar: exportar PNG (`$pres.Export(pasta, "PNG", 1400, 788)`) e olhar os slides pega texto estourando caixa e código cortado — coisas que não aparecem inspecionando o XML. Também gera o PDF (`SaveAs(caminho, 32)`).
- **Word NÃO aceita automação COM** — trava indefinidamente ao abrir. Não insista: o `.docx` sai do script e o PDF o Emanuel exporta pelo Word. Matar o WINWORD à força cria o painel de Recuperação de Documento, que trava as tentativas seguintes; se acontecer, limpar `%APPDATA%\Microsoft\Word\*.asd`.

## Convenções ao escrever código de aula

O código dos exemplos é lido em voz alta na aula e copiado pelos alunos — clareza didática vence elegância. Nomes de variáveis, estado e comentários em **português** (`senha`, `verSenha`, `erro`, `handleSubmit`); nomes de API do React ficam como são.

Quando um exemplo faz algo que não se faria em produção — a conferência de e-mail e senha no front-end das aulas 4 e 5 é o caso, assim como a senha escrita no `main.py` da aula 6 (credenciais de teste `user@email.com` / `123456`) — o README do exemplo precisa dizer isso explicitamente e apontar para onde o assunto será retomado de verdade (o back-end em Python com FastAPI).
