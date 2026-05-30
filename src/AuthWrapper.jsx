import React, { useState } from "react";
import BoutiqueApp from "./App.jsx";

const PASSWORD = "boutique2024"; // ← CAMBIA ESTO POR TU CONTRASEÑA

const G = "#c9a96e";

export default function AuthWrapper() {
  const [input, setInput]       = useState("");
  const [error, setError]       = useState(false);
  const [shake, setShake]       = useState(false);
  const [unlocked, setUnlocked] = useState(
    () => sessionStorage.getItem("boutique_auth") === "1"
  );

  const handleLogin = () => {
    if (input === PASSWORD) {
      sessionStorage.setItem("boutique_auth", "1");
      setUnlocked(true);
    } else {
      setError(true);
      setShake(true);
      setInput("");
      setTimeout(() => setShake(false), 500);
    }
  };

  const handleKey = (e) => {
    if (e.key === "Enter") handleLogin();
  };

  if (unlocked) return <BoutiqueApp />;

  return (
    <div style={{
      display: "flex", alignItems: "center", justifyContent: "center",
      height: "100vh", background: "#0b0b0e",
      fontFamily: "'Segoe UI', sans-serif",
    }}>
      <div style={{
        background: "#13131a", borderRadius: "20px",
        padding: "48px 40px", width: "340px",
        border: "1px solid #2a2a35", textAlign: "center",
        animation: shake ? "shake 0.4s ease" : "none",
      }}>
        <div style={{ fontSize: "40px", marginBottom: "12px" }}>🛍️</div>
        <div style={{ color: G, fontWeight: "800", fontSize: "22px", letterSpacing: "4px", marginBottom: "6px" }}>
          BOUTIQUE
        </div>
        <div style={{ color: "#444", fontSize: "11px", letterSpacing: "2px", marginBottom: "36px" }}>
          GESTIÓN COMERCIAL
        </div>

        <div style={{ position: "relative", marginBottom: "14px" }}>
          <input
            type="password"
            placeholder="Contraseña"
            value={input}
            onChange={e => { setInput(e.target.value); setError(false); }}
            onKeyDown={handleKey}
            autoFocus
            style={{
              width: "100%", padding: "14px 18px",
              background: "#0b0b0e", color: "#e0e0e0",
              border: `1px solid ${error ? "#f87171" : "#2a2a35"}`,
              borderRadius: "10px", fontSize: "15px",
              outline: "none", textAlign: "center",
              letterSpacing: "4px", boxSizing: "border-box",
              transition: "border-color 0.2s",
            }}
          />
        </div>

        {error && (
          <div style={{ color: "#f87171", fontSize: "12px", marginBottom: "14px" }}>
            Contraseña incorrecta
          </div>
        )}

        <button
          onClick={handleLogin}
          style={{
            width: "100%", padding: "14px",
            background: G, color: "#000",
            border: "none", borderRadius: "10px",
            fontWeight: "700", fontSize: "14px",
            cursor: "pointer", letterSpacing: "1px",
          }}>
          ENTRAR
        </button>

        <div style={{ color: "#2a2a35", fontSize: "11px", marginTop: "28px" }}>
          Acceso restringido
        </div>
      </div>

      <style>{`
        @keyframes shake {
          0%,100% { transform: translateX(0); }
          20%      { transform: translateX(-10px); }
          40%      { transform: translateX(10px); }
          60%      { transform: translateX(-8px); }
          80%      { transform: translateX(8px); }
        }
      `}</style>
    </div>
  );
}