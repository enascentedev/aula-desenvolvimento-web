# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## O que é este repositório

Material didático da disciplina **Programação Web 2** (Escola Técnica Mesquita), ministrada por **Emanuel Nascente**. Não é um produto de software: é um acervo de aulas, uma pasta por aula (`aula-3/`, `aula-4/`, `aula-5/`…), cada uma contendo slides, apostila e o código-exemplo daquele encontro.

O fio condutor do curso é um **sistema de ponto eletrônico** (React no front, C# + REST API + banco no back). Material novo deve manter essa continuidade e ser assinado como Emanuel Nascente — os PDFs antigos de `aula-3/` trazem "Luiz Gomes", o professor anterior.

Toda assinatura institucional usa **Escola Técnica Mesquita** e **Emanuel Nascente**, sem exceção: capa, rodapé de slide e rodapé de apostila. Nas apostilas isso é a constante `ESCOLA`; nos slides, `ESCOLA` e `PROFESSOR`.

Não há repositório git inicializado nem suíte de testes.

## Os dois tipos de artefato

### 1. Material da aula é gerado por script Python

**A fonte de verdade é o `.py`; o `.pptx`/`.docx` é saída.** Nunca edite o binário à mão — altere o script e regenere.

```bash
python aula-4/slides/gerar_slides.py       # → Aula04-Primeira-Pagina-React.pptx
python aula-4/apostila/gerar_apostila.py   # → Trabalho-Login-Aula04.docx
```

Dependências: `python-pptx`, `python-docx` (já instaladas globalmente; não há requirements.txt).

`gerar_slides.py` está organizado em camadas, do genérico ao específico — mexa na camada certa:

| Seção | Serve para |
|---|---|
| PALETA / TIPOGRAFIA | mudar o visual (dark tech, família VS Code Dark+) |
| HELPERS | primitivas (`rect`, `oval`, `textbox`, `para`, `rich_para`) |
| REALCE DE SINTAXE | `tokenize_js` / `tokenize_css` — colorização própria, sem lib externa |
| PAINEL DE CÓDIGO | `code_panel` — janela de editor com gutter de números e faixa de destaque |
| CROMO / LAYOUTS | `chrome`, `header`, `slide_capa`, `slide_cartoes`, `slide_conceito`, `slide_codigo`, `slide_codigo_duplo`, `slide_passos` |
| CONTEUDO DOS SLIDES + `build()` | o texto de cada slide, em ordem |

Ao criar material novo, **reutilize esses helpers em vez de recomeçar do zero**. Fontes confirmadas nesta máquina: Bahnschrift SemiBold Condensed (display), Cascadia Mono (código), Segoe UI, Consolas.

`gerar_apostila.py` segue o mesmo espírito para A4 (capa, sumário como campo do Word, numeração de página). O sumário só preenche as páginas quando o usuário abre o `.docx` e escolhe "Atualizar campo".

### 2. Código-exemplo React (`aula-N/exemplo-*/`)

```bash
cd aula-4/exemplo-login
npm install
npm run dev       # Vite em http://localhost:5173
npm run build     # produção em dist/
npm run preview
```

Stack fixa e propositalmente mínima: React 18.3 + Vite 5, JSX, CSS puro por componente, sem roteador, sem gerenciador de estado, sem TypeScript, sem linter. Cada exemplo é um projeto npm independente — não há workspace na raiz.

## Restrições de automação Office nesta máquina

- **PowerPoint aceita automação COM.** Use para revisar o resultado antes de entregar: exportar PNG (`$pres.Export(pasta, "PNG", 1400, 788)`) e olhar os slides pega texto estourando caixa e código cortado — coisas que não aparecem inspecionando o XML. Também gera o PDF (`SaveAs(caminho, 32)`).
- **Word NÃO aceita automação COM** — trava indefinidamente ao abrir. Não insista: o `.docx` sai do script e o PDF o Emanuel exporta pelo Word. Matar o WINWORD à força cria o painel de Recuperação de Documento, que trava as tentativas seguintes; se acontecer, limpar `%APPDATA%\Microsoft\Word\*.asd`.

## Convenções ao escrever código de aula

O código dos exemplos é lido em voz alta na aula e copiado pelos alunos — clareza didática vence elegância. Nomes de variáveis, estado e comentários em **português** (`senha`, `verSenha`, `erro`, `handleSubmit`); nomes de API do React ficam como são.

Quando um exemplo faz algo que não se faria em produção — a conferência de e-mail e senha no front-end da aula 4 é o caso — o README do exemplo precisa dizer isso explicitamente e apontar para onde o assunto será retomado de verdade (o back-end em C#).
