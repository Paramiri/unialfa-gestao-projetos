// Edge Function: analisar-transcricao-ata
// Recebe a transcricao de uma reuniao colada pelo usuario no formulario de
// Ata de Reuniao e usa a API da Claude (Anthropic) para extrair pauta, data,
// horarios, participantes, resumo, encaminhamentos e entraves, devolvendo
// JSON estruturado para o formulario preencher automaticamente. O usuario
// sempre revisa e completa antes de salvar - esta funcao so sugere, nunca
// salva nada sozinha.
//
// Segredo necessario (definido via `supabase secrets set ANTHROPIC_API_KEY=valor`):
//   ANTHROPIC_API_KEY - chave da API da Claude (console.anthropic.com)
// SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY sao injetados automaticamente pelo
// runtime das Edge Functions.
//
// Verificacao de JWT permanece ligada (padrao do Supabase) - so aceita
// chamadas autenticadas. Alem disso, esta funcao confere no banco (nao so na
// interface) se a importacao esta habilitada (tabela configuracoes) e se o
// papel de quem chamou esta na lista permitida - Admin sempre passa. Isso
// evita que alguem acione a analise (que tem custo real por chamada) so
// escondendo o botao na interface.

const ANTHROPIC_API_KEY = Deno.env.get("ANTHROPIC_API_KEY");
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

const SYSTEM_PROMPT = `Você é um assistente que analisa transcrições de reuniões de projetos e extrai dados estruturados para uma Ata de Reunião, em português do Brasil.

A partir da transcrição fornecida pelo usuário, extraia:
- pauta: título/assunto da reunião (curto, uma linha)
- data: data da reunião no formato YYYY-MM-DD, SE mencionada na transcrição (senão null)
- horaIni: horário de início HH:MM, SE mencionado (senão null)
- horaFim: horário de encerramento HH:MM, SE mencionado (senão null)
- participantes: array de objetos {nome, depto} — uma entrada por pessoa identificada na transcrição
- descricao: resumo executivo dos principais pontos discutidos (2 a 5 parágrafos, português claro e objetivo)
- encaminhamentos: array de objetos {encam, resp, prazo} — ações/próximos passos definidos, com responsável e prazo quando mencionados (prazo em YYYY-MM-DD se houver data explícita, senão texto livre como "próxima semana")
- entraves: array de objetos {entrave} — impedimentos, bloqueios ou problemas levantados na reunião (retorne [] se não houver nenhum)

Regras importantes:
- Só preencha um campo se a informação estiver realmente presente na transcrição — nunca invente nomes, datas ou compromissos.
- Se não houver informação suficiente para um campo, retorne null (campos de texto) ou [] (listas).
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"pauta":"...","data":"...","horaIni":"...","horaFim":"...","participantes":[...],"descricao":"...","encaminhamentos":[...],"entraves":[...]}`;

async function papelPermitidoNoBanco(authHeader: string): Promise<{ ok: boolean; motivo?: string }> {
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
      `${SUPABASE_URL}/rest/v1/configuracoes?chave=in.(permitir_importar_transcricao_ata,papeis_importar_transcricao_ata)&select=chave,valor`,
      { headers: { apikey: SUPABASE_SERVICE_ROLE_KEY, Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}` } }
    );
    const cfgRows = cfgResp.ok ? await cfgResp.json() : [];
    const habilitado = cfgRows.find((r: any) => r.chave === "permitir_importar_transcricao_ata")?.valor === true;
    // O interruptor mestre vale para todo mundo, inclusive Admin - é um controle de
    // custo/disponibilidade, não uma lista de acesso. A lista de papeis, essa sim,
    // sempre libera Admin implicitamente (mesma convenção do Painel Executivo/relatórios).
    if (!habilitado) return { ok: false, motivo: "A importação de transcrição por IA está desativada" };
    if (papel === "admin") return { ok: true };

    const papeis = cfgRows.find((r: any) => r.chave === "papeis_importar_transcricao_ata")?.valor;
    if (!Array.isArray(papeis) || !papeis.includes(papel)) {
      return { ok: false, motivo: "Seu papel não tem permissão para importar transcrição" };
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

  const permissao = await papelPermitidoNoBanco(authHeader);
  if (!permissao.ok) return json({ error: permissao.motivo || "Sem permissão" }, 403);

  if (!ANTHROPIC_API_KEY) {
    return json({ error: "ANTHROPIC_API_KEY não configurada nos secrets da function" }, 500);
  }

  let payload: { transcricao?: unknown };
  try {
    payload = await req.json();
  } catch {
    return json({ error: "Corpo da requisição não é um JSON válido" }, 400);
  }

  const transcricao = typeof payload.transcricao === "string" ? payload.transcricao.trim() : "";
  if (!transcricao) return json({ error: "Envie o campo 'transcricao' com o texto da reunião" }, 400);
  if (transcricao.length > 150000) {
    return json({ error: "Transcrição muito longa (máximo de ~150.000 caracteres)" }, 400);
  }

  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: 8000,
        system: SYSTEM_PROMPT,
        messages: [{ role: "user", content: `Transcrição da reunião:\n\n${transcricao}` }],
      }),
    });

    if (!resp.ok) {
      const errBody = await resp.text();
      console.error("Anthropic API retornou erro:", resp.status, errBody);
      return json({ error: "Falha ao analisar a transcrição (erro na API de IA)" }, 502);
    }

    const data = await resp.json();

    if (data.stop_reason === "max_tokens") {
      console.error("Resposta da IA cortada por atingir o limite de tokens.");
      return json({ error: "A transcrição é longa demais para ser analisada de uma vez — tente dividir em partes menores" }, 502);
    }

    const text = (data.content || [])
      .filter((b: any) => b.type === "text")
      .map((b: any) => b.text)
      .join("");
    const clean = text.replace(/```json|```/g, "").trim();

    let parsed: unknown;
    try {
      parsed = JSON.parse(clean);
    } catch (parseErr) {
      console.error("Resposta da IA nao era JSON valido:", clean.slice(0, 500));
      return json({ error: "A IA não conseguiu estruturar os dados dessa transcrição — tente novamente" }, 502);
    }

    return json({ ok: true, dados: parsed });
  } catch (e) {
    console.error("Falha ao chamar a API da Claude:", e);
    return json({ error: "Falha ao analisar a transcrição" }, 502);
  }
});
