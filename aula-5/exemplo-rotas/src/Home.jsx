import { Link } from "react-router-dom";
import "./Home.css";

function Home() {
  return (
    <div className="container-home">
      <div className="cartao">
        <h2>Ponto Eletrônico</h2>
        <p>Você entrou no sistema.</p>
        <Link to="/">Sair</Link>
      </div>
    </div>
  );
}

export default Home;
