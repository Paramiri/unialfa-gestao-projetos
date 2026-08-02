// Edge Function: send-notification
// Envia e-mails transacionais via Resend para as notificacoes de transicao de estado
// do Sistema de Gestao de Projetos UNIALFA (Gate 1, Gate 2, SMP, Canvas, TAP, TEP).
//
// Segredo necessario (definido via `supabase secrets set RESEND_API_KEY=...`):
//   RESEND_API_KEY
//
// Verificacao de JWT permanece ligada (padrao do Supabase): só aceita chamadas
// autenticadas com a anon key ou um token de sessao valido, igual ao resto da API REST.

const RESEND_API_KEY = Deno.env.get("RESEND_API_KEY");
const FROM_ADDRESS = "UNIALFA - Gerência de Projetos <notificacoes@sistemas.alfa.br>";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  if (!RESEND_API_KEY) {
    return json({ error: "RESEND_API_KEY não configurada nos secrets da function" }, 500);
  }

  let payload: { to?: unknown; subject?: unknown; html?: unknown };
  try {
    payload = await req.json();
  } catch {
    return json({ error: "Corpo da requisição não é um JSON válido" }, 400);
  }

  const rawTo = Array.isArray(payload.to) ? payload.to : [payload.to];
  const recipients = Array.from(
    new Set(rawTo.filter((e): e is string => typeof e === "string" && e.includes("@")))
  );
  const subject = typeof payload.subject === "string" ? payload.subject.trim() : "";
  const html = typeof payload.html === "string" ? payload.html : "";

  if (!recipients.length || !subject || !html) {
    return json({ error: "Parâmetros inválidos — 'to' (com ao menos 1 e-mail), 'subject' e 'html' são obrigatórios" }, 400);
  }

  try {
    const r = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${RESEND_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ from: FROM_ADDRESS, to: recipients, subject, html }),
    });
    const result = await r.json();
    if (!r.ok) {
      console.error("Resend retornou erro:", result);
      return json({ error: result }, 502);
    }
    return json({ ok: true, id: result?.id });
  } catch (e) {
    console.error("Falha ao chamar a API do Resend:", e);
    return json({ error: String(e) }, 500);
  }
});
