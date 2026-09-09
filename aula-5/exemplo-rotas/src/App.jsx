import { Routes, Route } from "react-router-dom";
import Login from "./Login";
import Home from "./Home";
import RotaPrivada from "./RotaPrivada";
import NaoEncontrada from "./NaoEncontrada";

function App() {
  return (
    <Routes>
      <Route path="/" element={<Login />} />

      {/* a Home so aparece se o guardiao deixar */}
      <Route path="/home" element={
        <RotaPrivada><Home /></RotaPrivada>
      } />

      {/* o * pega qualquer endereco que nao bateu com os de cima */}
      <Route path="*" element={<NaoEncontrada />} />
    </Routes>
  );
}

export default App;
