import React, { useState } from "react";
import { createRoot } from "react-dom/client";
import { OrganizacoesApp } from "./portal-frontend.jsx";

// App próprio "Organizações & DFT" (organograma + painel de unidade + relatório).
// Mesmo arquivo-fonte do Catálogo (portal-frontend.jsx), mas monta o export
// nomeado OrganizacoesApp em vez do PortalEntregas.
//
// O relatório de unidade (abas Diagnóstico/Alavancas/Carreiras/Decisão) roda
// dentro do <iframe srcDoc> do PainelUnidade, que lê window.__RELATORIO_HTML__.
// Como o PainelUnidade é usado por ESTE app, a injeção precisa acontecer aqui
// também — importado como TEXTO puro via loader do esbuild (.html=text),
// ANTES do render.
import relatorioHtml from "./relatorio-unidade-v4.html";

if (typeof window !== "undefined") {
  window.__RELATORIO_HTML__ = relatorioHtml;
}

// Barreira simples de acesso — NÃO é segurança de verdade (senha fica visível
// no código-fonte compilado), só evita que quem tem o link caia direto na
// tela. Uma vez acertada, fica lembrada neste navegador (localStorage).
const SENHA_ORGANOGRAMA = "12345";
const CHAVE_STORAGE = "organograma_autenticado";

function PortaoSenha({ children }) {
  const [liberado, setLiberado] = useState(
    typeof window !== "undefined" &&
      window.localStorage.getItem(CHAVE_STORAGE) === "ok"
  );
  const [senha, setSenha] = useState("");
  const [erro, setErro] = useState(false);

  if (liberado) return children;

  function entrar(e) {
    e.preventDefault();
    if (senha === SENHA_ORGANOGRAMA) {
      window.localStorage.setItem(CHAVE_STORAGE, "ok");
      setLiberado(true);
    } else {
      setErro(true);
    }
  }

  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: "#F4F6F9",
        fontFamily: "Raleway, sans-serif",
      }}
    >
      <form
        onSubmit={entrar}
        style={{
          background: "#fff",
          padding: 32,
          borderRadius: 12,
          boxShadow: "0 2px 12px rgba(0,0,0,0.08)",
          width: 280,
        }}
      >
        <h2 style={{ margin: "0 0 16px", fontSize: 18 }}>Organizações & DFT</h2>
        <p style={{ margin: "0 0 16px", fontSize: 13, color: "#666" }}>
          Área restrita — digite a senha para continuar.
        </p>
        <input
          type="password"
          autoFocus
          value={senha}
          onChange={(e) => {
            setSenha(e.target.value);
            setErro(false);
          }}
          style={{
            width: "100%",
            padding: "8px 10px",
            borderRadius: 6,
            border: erro ? "1px solid #d33" : "1px solid #ccc",
            boxSizing: "border-box",
            marginBottom: 8,
          }}
        />
        {erro && (
          <p style={{ color: "#d33", fontSize: 12, margin: "0 0 8px" }}>
            Senha incorreta.
          </p>
        )}
        <button
          type="submit"
          style={{
            width: "100%",
            padding: "8px 10px",
            borderRadius: 6,
            border: "none",
            background: "#2563eb",
            color: "#fff",
            fontWeight: 600,
            cursor: "pointer",
          }}
        >
          Entrar
        </button>
      </form>
    </div>
  );
}

const raiz = document.getElementById("root");
createRoot(raiz).render(
  <PortaoSenha>
    <OrganizacoesApp />
  </PortaoSenha>
);
