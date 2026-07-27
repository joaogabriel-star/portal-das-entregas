# Portal das Entregas — pacote unificado (fonte única)

Este é o **modelo publicável** para o desenvolvimento seguinte. A fonte de verdade
do frontend é o `portal-frontend.jsx` — **sempre parta daqui**, não reconstrua do zero.

## O que este pacote resolve

Unifica, num só arquivo-fonte, os **dois apps** que o portal serve:

| App | Export em `portal-frontend.jsx` | Entry | Shell HTML | Bundle |
|---|---|---|---|---|
| **Catálogo de Serviços** | `PortalEntregas` (default) | `index.jsx` | `index.html` | `bundle.js` |
| **Organizações & DFT** | `OrganizacoesApp` (nomeado) | `organizacoes.jsx` | `organizacoes.html` | `organizacoes-bundle.js` |

O app **Organizações & DFT** contém tudo o que estava em produção: o painel
"Organizações" (cobertura do DFT por unidade, com as lentes Status do DFT /
Necessidade de pessoal / Cobertura do DFT / Qualidade do catálogo), e o
detalhamento ao clicar numa unidade — a **lista de entregas** e o
**relatório completo** (abas Diagnóstico / Alavancas / Carreiras / Decisão).

Os dois apps se linkam entre si:
- Catálogo → botão de organizações leva a `./organizacoes.html`
- Organizações → "Catálogo de Serviços →" volta para `./index.html`

## O relatório de unidade

O `relatorio-unidade-v4.html` é embutido nos **dois** bundles via loader
`.html=text` do esbuild e injetado em `window.__RELATORIO_HTML__` antes do render
(em `index.jsx` e em `organizacoes.jsx`). O `PainelUnidade` o roda dentro de um
`<iframe srcDoc>`. Sem essa injeção, o relatório fica em branco.

## Rodar / buildar

```bash
npm install
npm run build      # gera dist/ com os DOIS bundles + os dois .html
npm run dev        # build + servidor estático local sobre dist/
```

Scripts individuais, se precisar: `npm run build:catalogo`, `npm run build:organizacoes`.

## Estrutura

```
portal-frontend.jsx        fonte única (Catálogo + Organizações num arquivo)
relatorio-unidade-v4.html  relatório da unidade (embutido no build)
index.jsx / index.html     entry + shell do Catálogo
organizacoes.jsx / .html   entry + shell do Organizações & DFT
package.json               dependências + build dos dois bundles
dist/                      GERADO pelo build (não versionar) — já incluso pronto
```

## Deploy no Netlify

**Opção A — arrastar `dist/`:** rode `npm run build` e arraste a pasta `dist/`
inteira (os dois bundles + os dois .html) na aba Deploys.

**Opção B — repo conectado (recomendado):** build command `npm run build`,
publish directory `dist`. Cada `git push` na `main` gera build + deploy.

Backend permanece em `pedrobaena/portal-backend` (Render):
`https://portal-backend-bftz.onrender.com`.
