// /api/claude.js — Vercel serverless function (proxy seguro para a API da Anthropic)
// Esta função roda no servidor da Vercel; a API key NUNCA é exposta ao navegador.

export default async function handler(req, res) {
  // CORS básico (mesmo domínio, mas evita problemas)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed. Use POST." });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({
      error: "ANTHROPIC_API_KEY não configurada no ambiente da Vercel.",
    });
  }

  try {
    const { prompt, max_tokens = 2000 } = req.body || {};

    if (!prompt || typeof prompt !== "string") {
      return res.status(400).json({ error: "Campo 'prompt' é obrigatório (string)." });
    }

    const anthropicResponse = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: Math.min(max_tokens, 8000),
        messages: [{ role: "user", content: prompt }],
      }),
    });

    const data = await anthropicResponse.json();

    if (!anthropicResponse.ok) {
      return res.status(anthropicResponse.status).json(data);
    }

    return res.status(200).json(data);
  } catch (err) {
    console.error("Erro no proxy:", err);
    return res.status(500).json({ error: err.message });
  }
}
