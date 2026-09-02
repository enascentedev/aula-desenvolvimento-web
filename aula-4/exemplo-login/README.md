# Exemplo — Página de Login (Aula 04)

Gabarito do trabalho da aula 4: a tela de login construída em aula **já com o
botão mostrar/ocultar senha** que os alunos precisam implementar.

Serve para duas coisas: rodar a demonstração ao vivo durante a aula e comparar
com as entregas na hora de corrigir.

## Como rodar

```bash
npm install
npm run dev
```

Abra o endereço que o Vite mostrar no terminal (normalmente
http://localhost:5173).

## Credenciais de teste

| Campo | Valor |
|---|---|
| E-mail | `user@email.com` |
| Senha | `123456` |

## Onde olhar

- `src/Login.jsx` — o componente completo. O estado `verSenha` e o botão
  `.ver-senha` são a parte que o trabalho pede.
- `src/Login.css` — o CSS da aula mais as regras de `.campo-senha` e
  `.ver-senha`.
- `src/App.jsx` — só monta o `<Login />` na tela.

## Aviso importante para a aula

A conferência de e-mail e senha está no front-end **apenas como exercício
didático de `useState`**. Numa aplicação real isso nunca se faz: qualquer
pessoa abre o código-fonte no navegador e lê a senha.

A autenticação de verdade acontece no servidor — é o que vamos construir com o
back-end em C# do sistema de ponto eletrônico.
