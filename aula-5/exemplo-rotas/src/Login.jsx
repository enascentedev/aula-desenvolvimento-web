import { useState } from "react";
import { useNavigate } from "react-router-dom";
import "./Login.css";

function Login() {
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [erro, setErro] = useState("");
  // guarda se a senha esta visivel — comeca escondida
  const [verSenha, setVerSenha] = useState(false);

  // o navigate troca de pagina pelo codigo
  const navigate = useNavigate();

  function handleSubmit(e) {
    e.preventDefault();

    if (!email || !senha) {
      setErro("Preencha todos os campos.");
      return;
    }

    if (email === "user@email.com" && senha === "123456") {
      setErro("");
      navigate("/home");
    } else {
      setErro("E-mail ou senha incorretos.");
    }
  }

  return (
    <div className="container">
      <form onSubmit={handleSubmit}>
        <h2>Login</h2>

        <input
          type="email"
          placeholder="Email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
        />

        <div className="campo-senha">
          <input
            type={verSenha ? "text" : "password"}
            placeholder="Senha"
            value={senha}
            onChange={(e) => setSenha(e.target.value)}
          />
          <button
            type="button"
            className="ver-senha"
            onClick={() => setVerSenha(!verSenha)}
          >
            {verSenha ? "Ocultar" : "Mostrar"}
          </button>
        </div>

        {erro && <p className="erro">{erro}</p>}

        <button type="submit">Entrar</button>
      </form>
    </div>
  );
}

export default Login;
