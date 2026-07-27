import React from "react";
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

const raiz = document.getElementById("root");
createRoot(raiz).render(<OrganizacoesApp />);
