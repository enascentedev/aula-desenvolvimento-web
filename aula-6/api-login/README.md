# API de login (Aula 06)

O primeiro pedaço do back-end do ponto eletrônico: uma API em **Python com
FastAPI** com uma rota só, `POST /login`, que confere e-mail e senha **no
servidor** — e não mais no navegador, como nas aulas 4 e 5.

Serve para a demonstração ao vivo da aula e como referência na hora de corrigir
o trabalho. O passo a passo completo, do zero, está nos slides da aula
(`aula-6/slides/Aula06-API-Login-FastAPI.pdf`).

## As ferramentas

| Ferramenta | Para que serve |
|---|---|
| **uv** | gerencia o projeto: instala o Python, cria o venv, adiciona pacotes e trava as versões no `uv.lock` |
| **venv** (`.venv/`) | o ambiente virtual: um Python e os pacotes só deste projeto |
| **FastAPI** | o framework que transforma funções Python em rotas HTTP |
| **Pydantic** | confere se o corpo do pedido tem os campos certos (senão, 422) |
| **Uvicorn** | o servidor que escuta a porta 8000 (o `fastapi dev` liga ele) |

## Como rodar

Precisa do [uv](https://docs.astral.sh/uv/) instalado. O Python 3.12 o próprio
uv baixa, se faltar.

```powershell
cd aula-6/api-login
uv sync                     # cria a .venv e instala o que o uv.lock pede
.venv\Scripts\activate      # PowerShell  (Git Bash: source .venv/Scripts/activate)
fastapi dev main.py
```

Sem ativar o venv também funciona: `uv run fastapi dev main.py`.

Com o servidor ligado, abra o **Swagger** — a documentação que o FastAPI gera
sozinho, onde dá para testar a rota:

http://127.0.0.1:8000/docs

(`/redoc` mostra a mesma documentação em outro formato. A raiz `/` responde
`404`: ela não é rota nossa.)

## A rota

`POST /login`, com o corpo em JSON:

```json
{
  "email": "user@email.com",
  "senha": "123456"
}
```

| Situação | Status | Corpo da resposta |
|---|---|---|
| e-mail e senha certos | `200` | `{"mensagem": "Login realizado com sucesso!", "email": "user@email.com"}` |
| e-mail ou senha errados | `401` | `{"detail": "E-mail ou senha incorretos."}` |
| falta um campo no corpo | `422` | a lista de erros do Pydantic (`"Field required"`) |

## Credenciais de teste

| Campo | Valor |
|---|---|
| E-mail | `user@email.com` |
| Senha | `123456` |

## Onde olhar

- `main.py` — a API inteira: a `app`, a classe `DadosLogin` (o formato do
  corpo) e a rota `login`, que devolve o dicionário (200) ou levanta o
  `HTTPException` (401).
- `pyproject.toml` — a ficha do projeto, com o FastAPI fixado em
  `fastapi[standard]==0.141.1`. Pacote entra e sai com `uv add` / `uv remove`,
  nunca editando à mão.
- `uv.lock` — a versão exata de cada um dos pacotes. Vai para o Git.
- `.python-version` — o Python do projeto (3.12).
- `.venv/` — **não** vai para o Git; `uv sync` recria.

## Aviso importante para a aula

O e-mail e a senha estão **escritos no `main.py`** apenas como exercício
didático. Num sistema de verdade isso nunca se faz: os usuários ficam num
**banco de dados** e a senha é guardada embaralhada (**hash**), nunca como
texto.

A API também ainda não devolve nada que prove o login para os pedidos
seguintes. Esse é o **token** assinado que vai substituir o `"logado": "true"`
que o `localStorage` guardava na aula 5 — e que qualquer um forjava no console.

Os próximos passos do back-end em FastAPI: o `Login.jsx` passar a perguntar
para esta API, depois o banco de dados e o token.
