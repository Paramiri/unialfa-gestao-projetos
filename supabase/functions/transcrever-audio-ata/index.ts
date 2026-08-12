// Edge Function: transcrever-audio-ata
// Recebe um pedaço (chunk) de áudio de reunião, enviado pelo navegador em
// multipart/form-data, e usa a API de transcrição da OpenAI (gpt-4o-transcribe)
// para converter em texto. O front-end (Ata de Reunião) chama esta função uma
// vez por pedaço de ~4 minutos (o áudio é dividido no navegador antes do envio,
// por causa do limite de 25 min / 25 MB por chamada da OpenAI) e concatena o
// texto de todos os pedaços na caixa de transcrição, para o usuário revisar
// antes de rodar a análise por IA (Edge Function analisar-transcricao-ata).
//
// Esta function so transcreve audio para texto - nao analisa nem preenche
// nada sozinha, e nao salva nada no banco.
//
// Segredo necessario (definido via `supabase secrets set OPENAI_API_KEY=valor`):
//   OPENAI_API_KEY - chave da API da OpenAI (platform.openai.com)
// SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY sao injetados automaticamente pelo
// runtime das Edge Functions.
//
// Verificacao de JWT permanece ligada (padrao do Supabase) - so aceita
// chamadas autenticadas. Alem disso, esta funcao confere no banco (nao so na
// interface) se a importacao de audio esta habilitada e se o papel de quem
// chamou esta na lista permitida - Admin sempre passa quando habilitada. Essa
// regra e INDEPENDENTE da regra de importacao de transcricao em texto
// (analisar-transcricao-ata) - chaves de configuracao separadas, a pedido do
// Admin, para poder liberar uma sem liberar a outra.

const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY");
const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

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

async function papelPermitidoAudio(authHeader: string): Promise<{ ok: boolean; motivo?: string }> {
  if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
    return { ok: false, motivo: "Configuração do servidor incompleta" };
  }
  try {
    const userResp = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
      headers: { apikey: SUPABASE_SERVICE_ROLE_KEY, Authorization: authHeader },
    });
    if (!userResp.ok) return { ok: false, motivo: "Sessão inválida" };
    const user = await userResp.json();
    const email = user?.email;
    if (!email) return { ok: false, motivo: "Sessão inválida" };

    const perfilResp = await fetch(
      `${SUPABASE_URL}/rest/v1/perfis?email=eq.${encodeURIComponent(email)}&select=papel`,
      { headers: { apikey: SUPABASE_SERVICE_ROLE_KEY, Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}` } }
    );
    const perfilRows = perfilResp.ok ? await perfilResp.json() : [];
    const papel = perfilRows[0]?.papel;
    if (!papel) return { ok: false, motivo: "Perfil não encontrado" };

    const cfgResp = await fetch(
      `${SUPABASE_URL}/rest/v1/configuracoes?chave=in.(permitir_importar_audio_ata,papeis_importar_audio_ata)&select=chave,valor`,
      { headers: { apikey: SUPABASE_SERVICE_ROLE_KEY, Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}` } }
    );
    const cfgRows = cfgResp.ok ? await cfgResp.json() : [];
    const habilitado = cfgRows.find((r: any) => r.chave === "permitir_importar_audio_ata")?.valor === true;
    // Interruptor mestre vale para todo mundo, inclusive Admin - controle de
    // custo/disponibilidade separado da regra de importacao de texto.
    if (!habilitado) return { ok: false, motivo: "A importação de áudio por IA está desativada" };
    if (papel === "admin") return { ok: true };

    const papeis = cfgRows.find((r: any) => r.chave === "papeis_importar_audio_ata")?.valor;
    if (!Array.isArray(papeis) || !papeis.includes(papel)) {
      return { ok: false, motivo: "Seu papel não tem permissão para importar áudio" };
    }
    return { ok: true };
  } catch (e) {
    console.error("Falha ao checar permissão:", e);
    return { ok: false, motivo: "Falha ao checar permissão" };
  }
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  const authHeader = req.headers.get("Authorization");
  if (!authHeader) return json({ error: "Não autenticado" }, 401);

  const permissao = await papelPermitidoAudio(authHeader);
  if (!permissao.ok) return json({ error: permissao.motivo || "Sem permissão" }, 403);

  if (!OPENAI_API_KEY) {
    return json({ error: "OPENAI_API_KEY não configurada nos secrets da function" }, 500);
  }

  let form: FormData;
  try {
    form = await req.formData();
  } catch {
    return json({ error: "Corpo da requisição não é multipart/form-data válido" }, 400);
  }

  const audio = form.get("audio");
  if (!(audio instanceof File)) {
    return json({ error: "Envie o pedaço de áudio no campo 'audio'" }, 400);
  }
  if (audio.size > 24 * 1024 * 1024) {
    return json({ error: "Pedaço de áudio muito grande (máximo 24 MB)" }, 400);
  }

  try {
    const openaiForm = new FormData();
    openaiForm.append("file", audio, audio.name || "audio.wav");
    openaiForm.append("model", "gpt-4o-transcribe");
    openaiForm.append("language", "pt");
    openaiForm.append("response_format", "json");

    const resp = await fetch("https://api.openai.com/v1/audio/transcriptions", {
      method: "POST",
      headers: { Authorization: `Bearer ${OPENAI_API_KEY}` },
      body: openaiForm,
    });

    if (!resp.ok) {
      const errBody = await resp.text();
      console.error("OpenAI API retornou erro:", resp.status, errBody);
      return json({ error: "Falha ao transcrever este pedaço de áudio (erro na API de IA)" }, 502);
    }

    const data = await resp.json();
    const texto = typeof data.text === "string" ? data.text : "";

    return json({ ok: true, texto });
  } catch (e) {
    console.error("Falha ao chamar a API da OpenAI:", e);
    return json({ error: "Falha ao transcrever este pedaço de áudio" }, 502);
  }
});
