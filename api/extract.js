/**
 * Función serverless (Vercel): proxy seguro hacia la API de Anthropic.
 * La API key vive SOLO aquí, como variable de entorno del servidor
 * (ANTHROPIC_API_KEY). Nunca se expone al navegador.
 */
export const config = { api: { bodyParser: { sizeLimit: "4.5mb" } } };

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Método no permitido" });
  }
  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: "ANTHROPIC_API_KEY no configurada en Vercel" });
  }

  const { messages, content } = req.body || {};
  const msgs = messages || (content ? [{ role: "user", content }] : null);
  if (!msgs) return res.status(400).json({ error: "Cuerpo inválido: se espera messages o content" });

  try {
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": process.env.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: process.env.CLAUDE_MODEL || "claude-sonnet-4-6",
        max_tokens: 1000,
        messages: msgs,
      }),
    });
    const data = await r.json();
    return res.status(r.status).json(data);
  } catch (e) {
    return res.status(502).json({ error: "Fallo al contactar la API de Anthropic: " + e.message });
  }
}
