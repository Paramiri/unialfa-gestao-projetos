// Edge Function: verificar-prazos-cronograma
// Roda 1x por dia via pg_cron (ver supabase/migrations/<data>_prazos_cronograma_cron.sql),
// sem depender de ninguem abrir o sistema. Varre o cronograma (marcos) de cada projeto
// ainda ativo (Planejamento e Desenvolvimento, chave `plan:` no kv_store) e dispara aviso
// por e-mail/push (reaproveitando a function send-notification) para os marcos que estao
// entrando na janela do prazo, sem Conclusao real preenchida:
//   - 5 dias antes do Termino previsto (faixa de 3 a 5 dias, tolerante a uma execucao perdida)
//   - 2 dias antes (faixa de 1 a 2 dias)
//   - no dia do vencimento
//   - atrasado (repete a cada 7 dias enquanto continuar sem Conclusao real)
//
// Destinatarios de cada aviso:
//   - os e-mails vinculados no campo "Responsavel (conta)" do marco (respConta —
//     pode ter mais de uma pessoa), se preenchido;
//   - + toda conta com papel 'gerente_projetos' que estiver na equipe daquele projeto
//     (projeto_equipe + perfis.papel — nao usa o campo de texto livre "Gerente do projeto").
// Se nenhum dos dois existir, o marco fica sem destinatario e nenhum e-mail e enviado
// (o indicador visual do Painel a Vista continua mostrando o atraso normalmente).
//
// Projetos com TEP ja registrado (kv_store tipo='tep') sao tratados como encerrados e
// ignorados, mesmo que o Planejamento correspondente ainda tenha marcos sem Conclusao real.
//
// Controle de "ja avisado" (evita repetir o mesmo aviso todo dia dentro da mesma janela):
// gravado no proprio kv_store, prefixo `alertaprazo:<marcoId>` (marcoId e o `id` estavel
// gerado no formulario 04 para cada marco — ver novoId() em
// "04 - planejamento-desenvolvimento-projeto.html"; marcos antigos sem esse id usam um
// identificador de fallback baseado em posicao, best-effort).
//
// Cada e-mail traz um link direto para o Meu Painel (20 - meu-painel.html) — a mesma tela
// mostra os itens de quem estiver logado, entao o link e generico, sem parametro por pessoa;
// quem abrir precisa logar com a propria conta para ver o que esta atribuido a ela. A URL
// muda com o ambiente (producao vs treino) verificando se SUPABASE_URL contem o ref de
// producao, mesmo padrao ja usado em reset-treino/index.ts (PRODUCAO_REF).
//
// Cada e-mail tambem termina indicando a Gerencia de Projetos como alternativa de contato,
// caso a pessoa prefira ajustar o prazo diretamente em vez de seguir só pelo aviso automatico
// — o nome do Gerente vem do campo "Gerente do projeto" (rec.gestor) do proprio Planejamento,
// nao e fixo.
//
// Segredos necessarios: nenhum novo — reaproveita SUPABASE_URL e
// SUPABASE_SERVICE_ROLE_KEY (injetados automaticamente pelo runtime das Edge Functions)
// para ler kv_store/perfis/projeto_equipe direto (sem RLS) e chamar a send-notification.
//
// Verificacao de JWT permanece ligada (padrao do Supabase, igual as demais functions) — o
// cron chama esta function com a anon key (publica, ja embutida em todo o front-end),
// exatamente como o restante da API REST.

const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
const PRODUCAO_REF = "fiarntunpqteopwjkhjg";
const SITE_URL = SUPABASE_URL && SUPABASE_URL.includes(PRODUCAO_REF)
  ? "https://gestaoprojetos.alfa.br"
  : "https://paramiri.github.io/unialfa-gestao-projetos-treino";
const MEU_PAINEL_URL = `${SITE_URL}/20%20-%20meu-painel.html`;

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
  return {
    apikey: SUPABASE_SERVICE_ROLE_KEY!,
    Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
    "Content-Type": "application/json",
    ...(extra || {}),
  };
}

function esc(s: unknown): string {
  return String(s ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c] as string)
  );
}

// "Hoje" em America/Sao_Paulo (UTC-3, sem horario de verao desde 2019) — o cron roda as
// 11:00 UTC (08:00 BRT), horario bem longe da virada de meia-noite local.
function hojeDataBR(): string {
  return new Date(Date.now() - 3 * 3600 * 1000).toISOString().slice(0, 10);
}
function diasEntre(dataA: string, dataB: string): number {
  const a = Date.parse(dataA + "T00:00:00Z");
  const b = Date.parse(dataB + "T00:00:00Z");
  return Math.round((a - b) / 86400000);
}
// Responsável (conta) pode ter mais de uma pessoa, guardado como string separada
// por vírgula (mesmo formato usado nos formulários) — um e-mail só, sem vírgula,
// continua funcionando normalmente (formato anterior, sem precisar migrar dado).
function parseRespConta(v?: string): string[] {
  return String(v || "").split(",").map((s) => s.trim()).filter(Boolean);
}

type Marco = {
  id?: string;
  marco?: string;
  resp?: string;
  respConta?: string;
  ini?: string;
  fim?: string;
  concl?: string;
};
type PlanoRec = {
  id?: string;
  nomeProjeto?: string;
  protocolo?: string;
  projetoId?: string;
  gestor?: string;
  cronograma?: Marco[];
};

type Bucket = "antes5" | "antes2" | "nodia" | "atrasado";
const BUCKET_LABEL: Record<Bucket, string> = {
  antes5: "vence em até 5 dias",
  antes2: "vence em até 2 dias",
  nodia: "vence hoje",
  atrasado: "está atrasado",
};

function bucketParaDias(dias: number): Bucket | null {
  if (dias < 0) return "atrasado";
  if (dias === 0) return "nodia";
  if (dias >= 1 && dias <= 2) return "antes2";
  if (dias >= 3 && dias <= 5) return "antes5";
  return null;
}

async function restGet(path: string) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, { headers: svcHeaders() });
  if (!r.ok) throw new Error(`GET ${path} -> ${r.status}: ${await r.text()}`);
  return r.json();
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }
  if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
    return json({ error: "SUPABASE_URL/SUPABASE_SERVICE_ROLE_KEY não disponíveis" }, 500);
  }

  const hoje = hojeDataBR();
  const resultado = { avaliados: 0, avisosEnviados: 0, erros: [] as string[] };

  try {
    // Projetos ja encerrados (TEP registrado) — ficam de fora.
    const teps: { projeto_id: string | null }[] = await restGet(
      "kv_store?tipo=eq.tep&select=projeto_id"
    );
    const projetosEncerrados = new Set(teps.map((t) => t.projeto_id).filter(Boolean));

    // Contas com papel Gerente de Projetos, e a equipe de cada projeto — para resolver o
    // segundo destinatario (GP da equipe do projeto) sem uma consulta por projeto.
    const gps: { id: string; email: string }[] = await restGet(
      "perfis?papel=eq.gerente_projetos&select=id,email"
    );
    const idsGP = new Set(gps.map((g) => g.id));
    const equipe: { projeto_id: string; usuario_id: string; usuario_email: string }[] =
      await restGet("projeto_equipe?select=projeto_id,usuario_id,usuario_email");
    const gpsPorProjeto = new Map<string, Set<string>>();
    for (const m of equipe) {
      if (!idsGP.has(m.usuario_id) || !m.usuario_email) continue;
      if (!gpsPorProjeto.has(m.projeto_id)) gpsPorProjeto.set(m.projeto_id, new Set());
      gpsPorProjeto.get(m.projeto_id)!.add(m.usuario_email);
    }

    // Todos os Planejamentos (cronogramas).
    const planos: { key: string; value: string }[] = await restGet(
      "kv_store?tipo=eq.plan&select=key,value"
    );

    for (const row of planos) {
      let rec: PlanoRec;
      try {
        rec = JSON.parse(row.value);
      } catch {
        continue;
      }
      if (!rec.projetoId || projetosEncerrados.has(rec.projetoId)) continue;
      const marcos = Array.isArray(rec.cronograma) ? rec.cronograma : [];
      const planoIdSuffix = row.key.replace(/^plan:/, "");

      for (let idx = 0; idx < marcos.length; idx++) {
        const marco = marcos[idx];
        if (!marco.marco || !marco.fim || marco.concl) continue;
        resultado.avaliados++;
        const dias = diasEntre(marco.fim, hoje);
        const bucket = bucketParaDias(dias);
        if (!bucket) continue;

        const marcoId = marco.id || `noid_${planoIdSuffix}_${idx}`;
        const alertaKey = `alertaprazo:${marcoId}`;

        let jaAvisado: { buckets?: Partial<Record<Bucket, string>> } = {};
        try {
          const existentes = await restGet(
            `kv_store?key=eq.${encodeURIComponent(alertaKey)}&select=value`
          );
          if (existentes[0]) jaAvisado = JSON.parse(existentes[0].value) || {};
        } catch {
          // sem registro anterior — segue como se nunca tivesse avisado
        }
        const buckets = jaAvisado.buckets || {};
        const ultimoEnvioBucket = buckets[bucket];

        let deveAvisar: boolean;
        if (bucket === "atrasado") {
          deveAvisar = !ultimoEnvioBucket || diasEntre(hoje, ultimoEnvioBucket.slice(0, 10)) >= 7;
        } else {
          deveAvisar = !ultimoEnvioBucket;
        }
        if (!deveAvisar) continue;

        const destinatarios = new Set<string>();
        for (const email of parseRespConta(marco.respConta)) destinatarios.add(email);
        for (const email of gpsPorProjeto.get(rec.projetoId) || []) destinatarios.add(email);
        if (!destinatarios.size) continue; // nada a enviar, mas nao marca como avisado

        const projetoNome = esc(rec.nomeProjeto || "Projeto sem nome");
        const marcoNome = esc(marco.marco);
        const dataFim = new Date(marco.fim + "T00:00:00Z").toLocaleDateString("pt-BR", {
          timeZone: "UTC",
        });
        const situacao =
          bucket === "atrasado"
            ? `está <b>atrasado</b> — término previsto era ${dataFim}`
            : bucket === "nodia"
            ? `vence <b>hoje</b> (${dataFim})`
            : `vence em <b>${dias} dia${dias === 1 ? "" : "s"}</b> (${dataFim})`;
        const subject = `Prazo do marco "${marco.marco}" ${
          bucket === "atrasado" ? "atrasado" : "se aproximando"
        } — ${rec.nomeProjeto || rec.protocolo || ""}`;
        const html =
          `<p>O marco <b>${marcoNome}</b> do cronograma do projeto <b>${projetoNome}</b>` +
          (rec.protocolo ? ` (${esc(rec.protocolo)})` : "") +
          ` ${situacao}.</p>` +
          (marco.resp ? `<p><b>Responsável (cronograma):</b> ${esc(marco.resp)}</p>` : "") +
          `<p>Avaliação automática do Painel de Prazos.</p>` +
          `<p><a href="${MEU_PAINEL_URL}">Abrir o Meu Painel</a> para ver e concluir este e os demais itens atribuídos a você.</p>` +
          `<p><b>Procurar a Gerência de Projetos${rec.gestor ? ` - Gerente - ${esc(rec.gestor)}` : ""}. Se houver necessidade de ajuste.</b></p>`;

        try {
          const r = await fetch(`${SUPABASE_URL}/functions/v1/send-notification`, {
            method: "POST",
            headers: svcHeaders(),
            body: JSON.stringify({ to: Array.from(destinatarios), subject, html }),
          });
          if (!r.ok) throw new Error(await r.text());
          resultado.avisosEnviados++;
        } catch (e) {
          resultado.erros.push(`Falha ao notificar marco ${marcoId}: ${String(e)}`);
          continue; // não marca como avisado — tenta de novo na próxima execução
        }

        buckets[bucket] = new Date().toISOString();
        try {
          await fetch(`${SUPABASE_URL}/rest/v1/kv_store?on_conflict=key`, {
            method: "POST",
            headers: svcHeaders({ Prefer: "resolution=merge-duplicates" }),
            body: JSON.stringify({
              key: alertaKey,
              value: JSON.stringify({ buckets }),
              updated_at: new Date().toISOString(),
            }),
          });
        } catch (e) {
          resultado.erros.push(`Falha ao gravar controle de aviso ${marcoId}: ${String(e)}`);
        }
      }
    }
  } catch (e) {
    return json({ error: String(e), ...resultado }, 500);
  }

  return json({ ok: true, hoje, ...resultado });
});
