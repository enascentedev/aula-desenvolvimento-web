# Exemplo — Navegando entre páginas (Aula 05)

O projeto da aula 4 com **React Router**: o login deixa de dar `alert` e passa a
levar o usuário para uma segunda página de verdade, com a URL mudando na barra
de endereços — e essa segunda página fica **protegida por um guardião de rota**.

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
| `/home` | `Home` dentro de `RotaPrivada` | só abre com a sessão aberta |
| qualquer outro | `NaoEncontrada` | a página 404 |

## Credenciais de teste

| Campo | Valor |
|---|---|
| E-mail | `user@email.com` |
| Senha | `123456` |

## Onde olhar

- `src/main.jsx` — o `<BrowserRouter>` envolvendo o `<App />`. Sem ele nada de
  rota funciona.
- `src/App.jsx` — o mapa da aplicação: `<Routes>` com um `<Route>` por tela, a
  `/home` embrulhada no guardião e o `path="*"` no fim.
- `src/Login.jsx` — o `useNavigate`. Compare com a aula 4: onde havia
  `alert("Login realizado com sucesso!")` agora há `navigate("/home")`,
  precedido do `localStorage.setItem("logado", "true")` que abre a sessão.
- `src/RotaPrivada.jsx` — o guardião. Não desenha nada: ou devolve o `children`,
  ou redireciona com `<Navigate to="/" replace />`.
- `src/Home.jsx` — a página de destino. O "Sair" é um `<button>`, não um
  `<Link>`: ele precisa apagar a sessão **antes** de navegar.
- `src/NaoEncontrada.jsx` — a 404, e o `<Link to="/">` de volta.
- `src/Home.css` — o estilo da página nova.

## O roteiro da demonstração

1. **O furo.** Com o `RotaPrivada` comentado no `App.jsx`, digite
   `http://localhost:5173/home` direto na barra: a página abre sem login.
2. **O conserto.** Devolva o guardião: o mesmo endereço agora devolve o
   visitante para o login.
3. **O logout de mentira.** Troque o `<button>` da Home por um
   `<Link to="/">Sair</Link>`. Parece funcionar — mas digite `/home` de novo e
   você entra sem senha, porque a sessão nunca foi apagada.
4. **O furo que continua.** Abra o DevTools → Console e digite:

   ```js
   localStorage.setItem("logado", "true");
   ```

   Recarregue `/home`. O guardião abre a porta. Ninguém fez login.

## Aviso importante para a aula

O guardião melhora muito a **experiência**: ninguém cai numa tela que não devia
ver, e o botão de sair passa a significar alguma coisa. Mas o passo 4 acima
mostra por que isso **não é segurança**: a chave que ele confere está no
navegador do próprio visitante, que pode escrevê-la à mão.

Rota não é segurança. Quem decide se você pode ver uma tela — e, principalmente,
quem decide se você pode ver os **dados** dela — é o servidor. É isso que vamos
construir no back-end em **Python com FastAPI**: o servidor confere e-mail e
senha, devolve um token assinado, e passa a exigir esse token em cada pedido de
dado. Aí o `localStorage` deixa de guardar um `"true"` que qualquer um digita e
passa a guardar um token que o front-end não sabe forjar.
