import { Navigate } from "react-router-dom";

// O guardiao de rota. Ele nao desenha nada na tela: so decide se o
// componente de dentro pode aparecer ou se o visitante volta pro login.
function RotaPrivada({ children }) {
  const logado = localStorage.getItem("logado") === "true";

  // replace: troca a entrada atual do historico em vez de empilhar uma nova,
  // assim o botao voltar do navegador nao devolve o usuario para a rota bloqueada
  if (!logado) {
    return <Navigate to="/" replace />;
  }

  // liberado: desenha o que veio dentro do <RotaPrivada> ... </RotaPrivada>
  return children;
}

export default RotaPrivada;
