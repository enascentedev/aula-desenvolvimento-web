# Programação Web 2

Material didático da disciplina **Programação Web 2** — Escola Técnica Mesquita.
Professor: **Emanuel Nascente**.

Uma pasta por aula, cada uma com os slides, a apostila do trabalho e o
código-exemplo daquele encontro.

O fio condutor do curso é um **sistema de ponto eletrônico**: React no front-end,
C# + REST API + banco no back-end. Cada aula acrescenta uma peça desse sistema.

## As aulas

| Aula | Assunto | O que tem aqui |
|---|---|---|
| [aula-3](aula-3/) | Cliente-servidor, por que React, instalação do ambiente | PDFs |
| [aula-4](aula-4/) | Primeira página: `Login.jsx`, `useState`, `handleSubmit`, CSS | slides, apostila, [exemplo-login](aula-4/exemplo-login/) |
| [aula-5](aula-5/) | Rotas: `BrowserRouter`, `Routes`/`Route`, `useNavigate`, `Link` | slides, apostila, [exemplo-rotas](aula-5/exemplo-rotas/) |

## Rodando um exemplo

Cada `exemplo-*` é um projeto npm independente — não há workspace na raiz.

```bash
cd aula-5/exemplo-rotas
npm install
npm run dev
```

O Vite mostra o endereço no terminal (normalmente http://localhost:5173).

Credenciais de teste do login: `user@email.com` / `123456`.

> A conferência de e-mail e senha acontece no front-end **apenas como exercício
> didático**. Numa aplicação real isso nunca se faz — a autenticação de verdade
> é o que vamos construir com o back-end em C#.

## Regenerando slides e apostilas

**A fonte de verdade é o script `.py`. O `.pptx` e o `.docx` são saída** — nunca
edite o binário à mão; altere o script e rode de novo.

```bash
python aula-5/slides/gerar_slides.py       # -> Aula05-Navegando-Entre-Paginas.pptx
python aula-5/apostila/gerar_apostila.py   # -> Trabalho-Rotas-Aula05.docx
```

Dependências: `python-pptx` e `python-docx`.

```bash
pip install python-pptx python-docx
```

Os scripts usam as fontes Bahnschrift SemiBold Condensed, Cascadia Mono, Segoe UI
e Consolas. Sem elas instaladas o arquivo ainda é gerado, mas o visual muda.

O sumário da apostila é um campo do Word: ele só mostra os números de página
depois que você abre o `.docx` e escolhe "Atualizar campo".
