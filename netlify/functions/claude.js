// netlify/functions/claude.js
// Proxy seguro para a API do Claude. A chave NUNCA vai para o navegador:
// fica na variável de ambiente ANTHROPIC_API_KEY (configurada no painel do Netlify).
//
// O portal, em produção, deve chamar "/.netlify/functions/claude" em vez de
// "https://api.anthropic.com/v1/messages". Veja COMO-LIGAR-IA.txt.

export default async (request) => {
  if (request.method !== "POST") {
    return new Response("Method not allowed", { status: 405 });
  }
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return new Response(JSON.stringify({ error: "ANTHROPIC_API_KEY não configurada" }), {
      status: 500, headers: { "Content-Type": "application/json" },
    });
  }
  try {
    const body = await request.text(); // repassa o corpo exatamente como veio do portal
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
    return new Response(text, {
      status: r.status,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: String(err) }), {
      status: 502, headers: { "Content-Type": "application/json" },
    });
  }
};
