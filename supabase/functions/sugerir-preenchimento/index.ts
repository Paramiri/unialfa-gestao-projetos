// Edge Function: sugerir-preenchimento
// Le os documentos ja registrados de um projeto (Solicitacao de Demanda,
// Canvas de Projeto, TAP, Planejamento e Atas de Reuniao) e usa a API da
// Claude (Anthropic) para sugerir o preenchimento dos campos do proximo
// formulario da esteira (Fase 1: Canvas e TAP; Fase 2: Planejamento e EAP).
// O usuario sempre revisa antes de salvar - esta funcao so sugere, nunca
// salva nada sozinha. Mesma arquitetura da funcao ja em producao
// `analisar-transcricao-ata`, generalizada para ler o historico do
// projeto em vez de uma transcricao colada.
//
// Segredo necessario (ja configurado nesta base para analisar-transcricao-ata):
//   ANTHROPIC_API_KEY - chave da API da Claude (console.anthropic.com)
// SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY sao injetados automaticamente.
//
// Verificacao de JWT permanece ligada. Alem disso, esta funcao confere no
// banco: (1) interruptor mestre + papel permitido em `configuracoes`
// (Admin sempre passa), e (2) se quem chamou faz parte da equipe do
// projeto informado (`projeto_equipe`) - para que ninguem gaste uma
// chamada de API (custo real) sugerindo preenchimento de projeto alheio.

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

function svcHeaders(extra?: Record<string, string>) {
  return { apikey: SUPABASE_SERVICE_ROLE_KEY!, Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`, ...(extra || {}) };
}

const FORMULARIOS = ["canvas", "tap", "planejamento", "eap"] as const;
type Formulario = typeof FORMULARIOS[number];

const SYSTEM_PROMPTS: Record<Formulario, string> = {
  canvas: `Você é um assistente que ajuda a preencher o Canvas de Projeto do sistema de gestão de projetos da UNIALFA, em português do Brasil, a partir de documentos já registrados do mesmo projeto (Solicitação de Demanda e, quando houver, Atas de Reunião).

O Canvas segue o modelo Por quê / O quê / Quem / Como / Quando e quanto, com estes campos:
- justificativas: as dores que originam o projeto (2 a 4 frases)
- objetivo: o objetivo do projeto (1 a 2 frases, formato SMART quando possível)
- beneficios: benefícios esperados
- produto: o produto/resultado do projeto (1 a 2 frases)
- parceiros: parceiros/áreas envolvidas no projeto
- entregas: grupos de entrega macro
- restricoes: restrições do projeto
- riscos: principais riscos identificáveis
- custos: estimativa de custos por grupo de entrega, somente se os documentos derem base para isso

Regras importantes:
- Baseie-se SOMENTE nos documentos fornecidos — nunca invente informação que não esteja neles.
- Se não houver informação suficiente para um campo, retorne null para ele — não escreva texto genérico de preenchimento.
- Escreva no mesmo tom dos documentos de origem: direto, objetivo, português formal de gestão de projetos.
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"justificativas":"...","objetivo":"...","beneficios":"...","produto":"...","parceiros":"...","entregas":"...","restricoes":"...","riscos":"...","custos":"..."}`,

  tap: `Você é um assistente que ajuda a preencher o TAP (Termo de Abertura de Projeto) do sistema de gestão de projetos da UNIALFA, em português do Brasil, a partir de documentos já registrados do mesmo projeto (Solicitação de Demanda, Canvas de Projeto e, quando houver, Atas de Reunião).

Extraia/redija:
- justificativa: justificativa do projeto
- objetivos: objetivos do projeto
- publico: público-alvo
- beneficios: benefícios esperados
- exclusoes: exclusões — o que fica fora do escopo
- premissas: premissas assumidas
- restricoes: restrições
- criterios: critérios de aceitação
- riscos: array com até 5 riscos, cada item {"risco":"...","status":"Aberto","resp":""} — status sempre "Aberto" para riscos novos; "resp" vazio se os documentos não indicarem um responsável
- cronograma: array com até 6 marcos de entrega macro, cada item {"marco":"...","resp":"","ini":null,"fim":null,"custo":""} — datas em YYYY-MM-DD SOMENTE se houver menção explícita, senão null
- custos: array com até 6 itens, cada item {"item":"...","espec":"","unid":"","qtd":"","valor":""} — deixe campos vazios quando o documento não der o detalhe, nunca invente números
- interessadas: array de partes interessadas já nomeadas nos documentos, cada item {"nome":"","unidade":""}

Regras importantes:
- Baseie-se SOMENTE nos documentos fornecidos — nunca invente nomes, valores, datas ou responsáveis que não estejam neles.
- Campo de texto sem informação suficiente: retorne null. Tabela sem item identificável: retorne [].
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"justificativa":"...","objetivos":"...","publico":"...","beneficios":"...","exclusoes":"...","premissas":"...","restricoes":"...","criterios":"...","riscos":[...],"cronograma":[...],"custos":[...],"interessadas":[...]}`,

  planejamento: `Você é um assistente que ajuda a preencher o dossiê de Planejamento e Desenvolvimento de Projeto do sistema de gestão de projetos da UNIALFA, em português do Brasil, a partir de documentos já registrados do mesmo projeto (TAP, Canvas de Projeto e, quando houver, Atas de Reunião).

Extraia/redija, para a aba Pré-projeto:
- produtos: produtos impactados pelo projeto
- contexto: contexto e problemática que originou o projeto
- objGeral: objetivo geral do projeto
- objEspec: objetivos específicos do projeto
- escIncluido: o que está incluído no escopo
- escExcluido: o que fica fora do escopo
- entregaveis: entregáveis do projeto
- premissas: premissas assumidas — parta das já descritas no TAP, detalhando se possível
- restricoes: restrições — parta das já descritas no TAP, detalhando se possível
- partes: partes interessadas — parta das já descritas no TAP, detalhando se possível

Para a aba Viabilidade:
- beneficios: benefícios do projeto — parta dos benefícios esperados já descritos no TAP, detalhando se possível

Para a aba Cronograma:
- cronograma: array com até 8 marcos do projeto, cada item {"marco":"...","resp":"","duracao":"","entrega":null} — baseie-se no cronograma de entregas macro do TAP, se houver; "entrega" é uma data YYYY-MM-DD SOMENTE se houver menção explícita, senão null

Regras importantes:
- Baseie-se SOMENTE nos documentos fornecidos — nunca invente informação que não esteja neles.
- Campo de texto sem informação suficiente: retorne null. Tabela sem item identificável: retorne [].
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"produtos":"...","contexto":"...","objGeral":"...","objEspec":"...","escIncluido":"...","escExcluido":"...","entregaveis":"...","premissas":"...","restricoes":"...","partes":"...","beneficios":"...","cronograma":[...]}`,

  eap: `Você é um assistente que ajuda a esboçar a EAP (Estrutura Analítica de Projeto) do sistema de gestão de projetos da UNIALFA, em português do Brasil, a partir de documentos já registrados do mesmo projeto (Planejamento e Desenvolvimento de Projeto, TAP e, quando houver, Atas de Reunião).

A EAP é uma árvore de 3 níveis: pacotes de trabalho (nível 1) → entregas (nível 2) → atividades (nível 3). Sua tarefa é sugerir SOMENTE os nomes dos pacotes de trabalho (nível 1) e das entregas (nível 2) dentro de cada pacote — NUNCA o nível 3 (atividades), que é detalhado depois por quem executa o trabalho.

Baseie-se principalmente nas saídas/entregáveis já descritos no Planejamento e nos entregáveis do TAP, agrupando entregas relacionadas sob um pacote de trabalho comum.

Regras importantes:
- Baseie-se SOMENTE nos documentos fornecidos — nunca invente pacotes ou entregas sem base neles.
- Sugira no máximo 4 pacotes de trabalho, cada um com no máximo 6 entregas.
- Se não houver informação suficiente para uma estrutura confiável, retorne um array vazio.
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"pacotes":[{"nome":"...","entregas":["...","..."]}]}`,
};

async function papelEquipePermitido(authHeader: string, projetoId: string): Promise<{ ok: boolean; motivo?: string }> {
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
    const userId = user?.id;
    if (!email || !userId) return { ok: false, motivo: "Sessão inválida" };

    const perfilResp = await fetch(
      `${SUPABASE_URL}/rest/v1/perfis?email=eq.${encodeURIComponent(email)}&select=papel`,
      { headers: svcHeaders() }
    );
    const perfilRows = perfilResp.ok ? await perfilResp.json() : [];
    const papel = perfilRows[0]?.papel;
    if (!papel) return { ok: false, motivo: "Perfil não encontrado" };
    const isAdmin = papel === "admin";

    const cfgResp = await fetch(
      `${SUPABASE_URL}/rest/v1/configuracoes?chave=in.(permitir_assistente_preenchimento,papeis_assistente_preenchimento)&select=chave,valor`,
      { headers: svcHeaders() }
    );
    const cfgRows = cfgResp.ok ? await cfgResp.json() : [];
    const habilitado = cfgRows.find((r: any) => r.chave === "permitir_assistente_preenchimento")?.valor === true;
    if (!habilitado) return { ok: false, motivo: "O assistente de preenchimento por IA está desativado" };
    if (!isAdmin) {
      const papeis = cfgRows.find((r: any) => r.chave === "papeis_assistente_preenchimento")?.valor;
      if (!Array.isArray(papeis) || !papeis.includes(papel)) {
        return { ok: false, motivo: "Seu papel não tem permissão para usar o assistente de preenchimento" };
      }
    }

    if (!isAdmin) {
      const equipeResp = await fetch(
        `${SUPABASE_URL}/rest/v1/projeto_equipe?projeto_id=eq.${encodeURIComponent(projetoId)}&usuario_id=eq.${encodeURIComponent(userId)}&select=usuario_id`,
        { headers: svcHeaders() }
      );
      const equipeRows = equipeResp.ok ? await equipeResp.json() : [];
      if (!equipeRows.length) return { ok: false, motivo: "Você não faz parte da equipe deste projeto" };
    }

    return { ok: true };
  } catch (e) {
    console.error("Falha ao checar permissão:", e);
    return { ok: false, motivo: "Falha ao checar permissão" };
  }
}

async function fetchDocsByPrefix(prefix: string, projetoId: string): Promise<any[]> {
  const r = await fetch(
    `${SUPABASE_URL}/rest/v1/kv_store?key=like.${encodeURIComponent(prefix)}*&select=value`,
    { headers: svcHeaders() }
  );
  if (!r.ok) return [];
  const rows = await r.json();
  const out: any[] = [];
  for (const row of rows) {
    try {
      const parsed = JSON.parse(row.value);
      if (parsed && parsed.projetoId === projetoId) out.push(parsed);
    } catch { /* ignora registro malformado */ }
  }
  return out;
}

async function fetchDemanda(solicitacaoId: string | null): Promise<any | null> {
  if (!solicitacaoId) return null;
  const r = await fetch(
    `${SUPABASE_URL}/rest/v1/kv_store?key=eq.${encodeURIComponent("demanda:" + solicitacaoId)}&select=value`,
    { headers: svcHeaders() }
  );
  if (!r.ok) return null;
  const rows = await r.json();
  if (!rows.length) return null;
  try { return JSON.parse(rows[0].value); } catch { return null; }
}

function bloco(titulo: string, campos: Array<[string, string | null | undefined]>): string | null {
  const linhas = campos
    .filter(([, v]) => v != null && String(v).trim() !== "")
    .map(([label, v]) => `${label}: ${v}`);
  if (!linhas.length) return null;
  return `${titulo}\n${linhas.join("\n")}`;
}

function formatDemanda(d: any): string | null {
  return bloco(`SOLICITAÇÃO DE DEMANDA (protocolo ${d.protocolo || "—"})`, [
    ["Projeto", d.projeto],
    ["Solicitante", d.solicitante],
    ["Departamento/unidade", d.departamento],
    ["Justificativa/necessidade", d.justificativa],
    ["Objetivo/resultado esperado", d.objetivo],
    ["Escopo", d.escopo],
    ["Prazo desejado", d.prazo],
    ["Prioridade", d.prioridade],
    ["Orçamento estimado", d.orcamento],
    ["Partes interessadas", d.partes],
  ]);
}

function formatCanvas(c: any): string | null {
  return bloco(`CANVAS DE PROJETO (protocolo ${c.protocolo || "—"})`, [
    ["Justificativas", c.justificativas],
    ["Objetivo", c.objetivo],
    ["Benefícios", c.beneficios],
    ["Produto", c.produto],
    ["Parceiros", c.parceiros],
    ["Grupo de entregas", c.entregas],
    ["Restrições", c.restricoes],
    ["Riscos", c.riscos],
    ["Custos", c.custos],
  ]);
}

function formatTap(t: any): string | null {
  const cronogramaResumo = Array.isArray(t.cronograma) && t.cronograma.length
    ? t.cronograma.map((c: any) => `- ${c.marco || ""}${c.resp ? ` (responsável: ${c.resp})` : ""}${c.fim ? ` (término: ${c.fim})` : ""}`).join("\n")
    : "";
  return bloco(`TAP - TERMO DE ABERTURA DE PROJETO (protocolo ${t.protocolo || "—"})`, [
    ["Alinhamento estratégico", t.alinhamento],
    ["Programa vinculado", t.programa],
    ["Justificativa", t.justificativa],
    ["Objetivos", t.objetivos],
    ["Público-alvo", t.publico],
    ["Benefícios esperados", t.beneficios],
    ["Exclusões (fora do escopo)", t.exclusoes],
    ["Premissas", t.premissas],
    ["Restrições", t.restricoes],
    ["Critérios de aceitação", t.criterios],
    ["Cronograma de entregas macro", cronogramaResumo || null],
  ]);
}

function formatPlanejamento(p: any): string | null {
  const saidasResumo = Array.isArray(p.saidas) && p.saidas.length
    ? p.saidas.map((s: any) => `- ${s.item || ""}`).filter((l: string) => l !== "- ").join("\n")
    : "";
  return bloco(`PLANEJAMENTO E DESENVOLVIMENTO DE PROJETO (protocolo ${p.protocolo || "—"})`, [
    ["Produtos impactados", p.produtos],
    ["Contexto e problemática", p.contexto],
    ["Objetivo geral", p.objGeral],
    ["Objetivos específicos", p.objEspec],
    ["Escopo incluído", p.escIncluido],
    ["Escopo excluído", p.escExcluido],
    ["Entregáveis", p.entregaveis],
    ["Benefícios do projeto", p.beneficios],
    ["Saídas / entregáveis detalhados", saidasResumo || null],
  ]);
}

function formatAta(a: any): string | null {
  const encaminhamentos = Array.isArray(a.saidas) && a.saidas.length
    ? a.saidas.map((s: any) => `- ${s.encam || ""}${s.resp ? ` (responsável: ${s.resp})` : ""}${s.prazo ? ` (prazo: ${s.prazo})` : ""}`).join("\n")
    : "";
  const entraves = Array.isArray(a.entraves) && a.entraves.length
    ? a.entraves.map((e: any) => `- ${e.entrave || ""}`).join("\n")
    : "";
  return bloco(`ATA DE REUNIÃO (protocolo ${a.protocolo || "—"}${a.data ? `, ${a.data}` : ""})`, [
    ["Pauta", a.pauta],
    ["Resumo", a.descricao],
    ["Encaminhamentos", encaminhamentos || null],
    ["Entraves", entraves || null],
  ]);
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  const authHeader = req.headers.get("Authorization");
  if (!authHeader) return json({ error: "Não autenticado" }, 401);

  let payload: { projetoId?: unknown; formulario?: unknown };
  try {
    payload = await req.json();
  } catch {
    return json({ error: "Corpo da requisição não é um JSON válido" }, 400);
  }

  const projetoId = typeof payload.projetoId === "string" ? payload.projetoId.trim() : "";
  const formulario = typeof payload.formulario === "string" ? payload.formulario.trim() : "";
  if (!projetoId) return json({ error: "Envie o campo 'projetoId'" }, 400);
  if (!FORMULARIOS.includes(formulario as Formulario)) {
    return json({ error: "Campo 'formulario' inválido — use 'canvas', 'tap', 'planejamento' ou 'eap'" }, 400);
  }

  const permissao = await papelEquipePermitido(authHeader, projetoId);
  if (!permissao.ok) return json({ error: permissao.motivo || "Sem permissão" }, 403);

  if (!ANTHROPIC_API_KEY) {
    return json({ error: "ANTHROPIC_API_KEY não configurada nos secrets da function" }, 500);
  }

  try {
    const projetoResp = await fetch(
      `${SUPABASE_URL}/rest/v1/projetos?id=eq.${encodeURIComponent(projetoId)}&select=id,nome,solicitacao_id`,
      { headers: svcHeaders() }
    );
    const projetoRows = projetoResp.ok ? await projetoResp.json() : [];
    const projeto = projetoRows[0];
    if (!projeto) return json({ error: "Projeto não encontrado" }, 404);

    // Quais documentos anteriores cada formulario-alvo le, seguindo a esteira documental
    const FONTES_POR_FORMULARIO: Record<Formulario, { demanda?: boolean; canvas?: boolean; tap?: boolean; planejamento?: boolean; atas?: boolean }> = {
      canvas: { demanda: true, atas: true },
      tap: { demanda: true, canvas: true, atas: true },
      planejamento: { canvas: true, tap: true, atas: true },
      eap: { tap: true, planejamento: true, atas: true },
    };
    const fontesConfig = FONTES_POR_FORMULARIO[formulario as Formulario];

    const demanda = fontesConfig.demanda ? await fetchDemanda(projeto.solicitacao_id || null) : null;
    const canvas = fontesConfig.canvas ? (await fetchDocsByPrefix("canvas:", projetoId))[0] || null : null;
    const tap = fontesConfig.tap ? (await fetchDocsByPrefix("tap:", projetoId))[0] || null : null;
    const planejamento = fontesConfig.planejamento ? (await fetchDocsByPrefix("plan:", projetoId))[0] || null : null;
    const atas = fontesConfig.atas ? await fetchDocsByPrefix("ata:", projetoId) : [];

    const blocos: string[] = [];
    const fontes: Array<{ tipo: string; protocolo: string }> = [];
    if (demanda) {
      const b = formatDemanda(demanda);
      if (b) { blocos.push(b); fontes.push({ tipo: "Solicitação de Demanda", protocolo: demanda.protocolo || "—" }); }
    }
    if (canvas) {
      const b = formatCanvas(canvas);
      if (b) { blocos.push(b); fontes.push({ tipo: "Canvas de Projeto", protocolo: canvas.protocolo || "—" }); }
    }
    if (tap) {
      const b = formatTap(tap);
      if (b) { blocos.push(b); fontes.push({ tipo: "TAP", protocolo: tap.protocolo || "—" }); }
    }
    if (planejamento) {
      const b = formatPlanejamento(planejamento);
      if (b) { blocos.push(b); fontes.push({ tipo: "Planejamento e Desenvolvimento", protocolo: planejamento.protocolo || "—" }); }
    }
    for (const ata of atas) {
      const b = formatAta(ata);
      if (b) { blocos.push(b); fontes.push({ tipo: "Ata de Reunião", protocolo: ata.protocolo || "—" }); }
    }

    if (!blocos.length) {
      return json({ error: "Nenhum documento anterior com conteúdo foi encontrado para este projeto ainda" }, 422);
    }

    let contexto = blocos.join("\n\n");
    if (contexto.length > 60000) contexto = contexto.slice(0, 60000) + "\n\n[...conteúdo adicional omitido por limite de tamanho...]";

    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: 4000,
        system: SYSTEM_PROMPTS[formulario as Formulario],
        messages: [{ role: "user", content: `Documentos do projeto "${projeto.nome || ""}":\n\n${contexto}` }],
      }),
    });

    if (!resp.ok) {
      const errBody = await resp.text();
      console.error("Anthropic API retornou erro:", resp.status, errBody);
      return json({ error: "Falha ao gerar a sugestão (erro na API de IA)" }, 502);
    }

    const data = await resp.json();
    if (data.stop_reason === "max_tokens") {
      console.error("Resposta da IA cortada por atingir o limite de tokens.");
      return json({ error: "O contexto do projeto é longo demais para sugerir de uma vez" }, 502);
    }

    const text = (data.content || [])
      .filter((b: any) => b.type === "text")
      .map((b: any) => b.text)
      .join("");
    const clean = text.replace(/```json|```/g, "").trim();

    let parsed: unknown;
    try {
      parsed = JSON.parse(clean);
    } catch {
      console.error("Resposta da IA nao era JSON valido:", clean.slice(0, 500));
      return json({ error: "A IA não conseguiu estruturar a sugestão — tente novamente" }, 502);
    }

    return json({ ok: true, dados: parsed, fontes });
  } catch (e) {
    console.error("Falha ao gerar sugestão de preenchimento:", e);
    return json({ error: "Falha ao gerar a sugestão" }, 502);
  }
});
