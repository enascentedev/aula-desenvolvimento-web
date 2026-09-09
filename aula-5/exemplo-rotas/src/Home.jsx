import { useNavigate } from "react-router-dom";
import "./Home.css";

function Home() {
  const navigate = useNavigate();

  // sair de verdade: primeiro apaga a sessao, depois volta pro login.
  // um <Link to="/"> sozinho levaria pro login sem desligar nada.
  function sair() {
    localStorage.removeItem("logado");
    navigate("/");
  }

  return (
    <div className="container-home">
      <div className="cartao">
        <h2>Ponto Eletrônico</h2>
        <p>Você entrou no sistema.</p>
        <button className="sair" onClick={sair}>
          Sair
        </button>
      </div>
    </div>
  );
}

export default Home;
