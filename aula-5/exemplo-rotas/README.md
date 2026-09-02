# Exemplo — Navegando entre páginas (Aula 05)

O projeto da aula 4 com **React Router**: o login deixa de dar `alert` e passa a
levar o usuário para uma segunda página de verdade, com a URL mudando na barra
de endereços.

Serve para a demonstração ao vivo da aula e como referência na hora de corrigir
o trabalho.

## Como rodar

```bash
npm install
npm run dev
```

Abra o endereço que o Vite mostrar no terminal (normalmente
http://localhost:5173).

## As rotas

| Caminho | Componente | O que é |
|---|---|---|
| `/` | `Login` | a tela da aula 4 |
| `/home` | `Home` | a página nova, para onde o login leva |

## Credenciais de teste

| Campo | Valor |
|---|---|
| E-mail | `user@email.com` |
| Senha | `123456` |

## Onde olhar

- `src/main.jsx` — o `<BrowserRouter>` envolvendo o `<App />`. Sem ele nada de
  rota funciona.
- `src/App.jsx` — o mapa da aplicação: `<Routes>` com um `<Route>` por tela.
- `src/Login.jsx` — o `useNavigate`. Compare com a aula 4: onde havia
  `alert("Login realizado com sucesso!")` agora há `navigate("/home")`.
- `src/Home.jsx` — a página de destino e o `<Link to="/">` de volta.
- `src/Home.css` — o estilo da página nova.

## Aviso importante para a aula

Digite `http://localhost:5173/home` direto no navegador: a página abre mesmo sem
ter feito login. Isso **não é um bug do React Router** — é o que acontece quando
a única verificação de acesso está no front-end.

Rota não é segurança. Quem decide se você pode ver uma tela é o servidor, e é
isso que vamos construir com o back-end em C# do sistema de ponto eletrônico.
