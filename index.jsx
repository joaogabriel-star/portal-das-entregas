import React from "react";
import { createRoot } from "react-dom/client";
import PortalEntregas from "./portal-frontend.jsx";

// Relatório de unidade (abas Diagnóstico/Alavancas/Decisão), documento HTML
// autocontido que roda dentro do <iframe srcDoc> do painel de unidade.
// Importado como TEXTO puro via loader do esbuild (.html=text) — vira parte
// do bundle, sem precisar escapar backtick/${/</script> à mão. Isso substitui
// o IIFE colado manualmente no fim do bundle.js (ver GUIA-frontend-portal.md,
// seção 5). A atribuição acontece ANTES do render, então o painel de unidade
// já encontra window.__RELATORIO_HTML__ preenchido quando for renderizado.
import relatorioHtml from "./relatorio-unidade-v4.html";

if (typeof window !== "undefined") {
  window.__RELATORIO_HTML__ = relatorioHtml;
}

const raiz = document.getElementById("root");
createRoot(raiz).render(<PortalEntregas />);
