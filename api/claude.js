// api/claude.js
// Equivalente da netlify/functions/claude.js, para deploys na Vercel.
// Proxy seguro para a API do Claude — a chave NUNCA vai para o navegador,
// fica na variável de ambiente ANTHROPIC_API_KEY (painel da Vercel:
// Project Settings → Environment Variables).
//
// O vercel.json redireciona "/.netlify/functions/claude" para esta função,
// então o front-end (chamarIA em portal-frontend.jsx) não precisa saber
// se está rodando no Netlify ou na Vercel — chama sempre o mesmo caminho.

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).send("Method not allowed");
    return;
  }
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: "ANTHROPIC_API_KEY não configurada" });
    return;
  }
  try {
    const body = typeof req.body === "string" ? req.body : JSON.stringify(req.body);
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body,
    });
    const text = await r.text();
    res.status(r.status).setHeader("Content-Type", "application/json").send(text);
  } catch (err) {
    res.status(502).json({ error: String(err) });
  }
};
