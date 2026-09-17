// Edge Function: verificar-prazos-cronograma
// Roda 1x por dia via pg_cron (ver supabase/migrations/<data>_prazos_cronograma_cron.sql),
// sem depender de ninguem abrir o sistema. Varre o cronograma (marcos) de cada projeto
// ainda ativo (Planejamento e Desenvolvimento, chave `plan:` no kv_store) e dispara dois
// tipos de aviso por e-mail/push (reaproveitando a function send-notification), sem
// Conclusao real preenchida:
//
// 1) Aviso de prazo (existente):
//   - N1 dias antes do Termino previsto (padrao 5, faixa entre N2+1 e N1 dias)
//   - N2 dias antes (padrao 2, faixa de 1 a N2 dias)
//   - no dia do vencimento
//   - atrasado (repete a cada N3 dias enquanto continuar sem Conclusao real — padrao 7)
//
// 2) Aviso de silencio (novo — fecha o ciclo do campo Execucao, ver 04/20/19): marco cujo
//   Inicio previsto ja chegou e que esta ha N4 dias ou mais sem nenhuma mudanca no campo
//   Execucao (comparando com execStatusAtualizadoEm, ou com o proprio Inicio previsto se o
//   campo nunca foi preenchido — padrao N4 = 5) — repete a cada N5 dias enquanto continuar
//   parado (padrao N5 = 5). Existe porque o aviso de prazo acima so olha a data final: um
//   marco com o Termino previsto ainda longe (ex.: 44 dias) podia ficar semanas parado, sem
//   Execucao preenchida, sem ninguem ser avisado. Nao dispara no mesmo dia em que o aviso de
//   prazo ja foi enviado para o mesmo marco, para nao mandar dois e-mails de uma vez.
//
// N1..N5 sao parametrizaveis em Administracao > Configuracoes > Regras de aviso automatico
// (tabela `configuracoes`, chaves aviso_prazo_dias_antes1/antes2/repeticao_atrasado e
// aviso_silencio_dias_sem_atualizar/repeticao — ver migracao
// 20260917160000_config_avisos_prazo_execucao.sql). Lidos a cada execucao no inicio da
// function; se uma chave nao existir ou tiver valor invalido, cai no mesmo padrao que era
// fixo no codigo antes desta migracao — a function nunca fica sem rodar por configuracao
// ausente.
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
  execStatus?: string;
  execStatusAtualizadoEm?: string;
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

// dAntes2/dAntes1 vêm de Administração > Configurações — normaliza a ordem (min/max) para
// nunca abrir um buraco na faixa de dias se alguém configurar dAntes2 maior que dAntes1.
function bucketParaDias(dias: number, dAntes2: number, dAntes1: number): Bucket | null {
  if (dias < 0) return "atrasado";
  if (dias === 0) return "nodia";
  const antes2 = Math.min(dAntes2, dAntes1);
  const antes1 = Math.max(dAntes2, dAntes1);
  if (dias >= 1 && dias <= antes2) return "antes2";
  if (dias > antes2 && dias <= antes1) return "antes5";
  return null;
}

async function restGet(path: string) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, { headers: svcHeaders() });
  if (!r.ok) throw new Error(`GET ${path} -> ${r.status}: ${await r.text()}`);
  return r.json();
}

// Lê um numero configuravel em Administracao > Configuracoes (tabela `configuracoes`),
// caindo no padrao se a chave nao existir, nao ser numero, ou nao for positiva.
async function getConfigDias(chave: string, padrao: number): Promise<number> {
  try {
    const rows: { valor: unknown }[] = await restGet(
      `configuracoes?chave=eq.${chave}&select=valor`
    );
    const v = rows[0]?.valor;
    return typeof v === "number" && v > 0 ? v : padrao;
  } catch {
    return padrao;
  }
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
    // Dias configuraveis das duas regras de disparo (Administracao > Configuracoes).
    const [diasAntes1, diasAntes2, diasRepeticaoAtrasado, diasSilencioLimite, diasSilencioRepeticao] =
      await Promise.all([
        getConfigDias("aviso_prazo_dias_antes1", 5),
        getConfigDias("aviso_prazo_dias_antes2", 2),
        getConfigDias("aviso_prazo_dias_repeticao_atrasado", 7),
        getConfigDias("aviso_silencio_dias_sem_atualizar", 5),
        getConfigDias("aviso_silencio_dias_repeticao", 5),
      ]);

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

        const marcoId = marco.id || `noid_${planoIdSuffix}_${idx}`;

        const destinatarios = new Set<string>();
        for (const email of parseRespConta(marco.respConta)) destinatarios.add(email);
        for (const email of gpsPorProjeto.get(rec.projetoId) || []) destinatarios.add(email);
        if (!destinatarios.size) continue; // nada a enviar, mas nao marca como avisado

        const alertaKey = `alertaprazo:${marcoId}`;
        let jaAvisado: { buckets?: Partial<Record<Bucket, string>> & { silencio?: string } } = {};
        try {
          const existentes = await restGet(
            `kv_store?key=eq.${encodeURIComponent(alertaKey)}&select=value`
          );
          if (existentes[0]) jaAvisado = JSON.parse(existentes[0].value) || {};
        } catch {
          // sem registro anterior — segue como se nunca tivesse avisado
        }
        const buckets = jaAvisado.buckets || {};
        let mudouBuckets = false;

        const projetoNome = esc(rec.nomeProjeto || "Projeto sem nome");
        const marcoNome = esc(marco.marco);
        const respHtml = marco.resp ? `<p><b>Responsável (cronograma):</b> ${esc(marco.resp)}</p>` : "";
        const rodapeGerente =
          `<p><b>Procurar a Gerência de Projetos${rec.gestor ? ` - Gerente - ${esc(rec.gestor)}` : ""}. Se houver necessidade de ajuste.</b></p>`;

        // --- 1) Aviso de prazo ---
        let avisouPrazoAgora = false;
        const dias = diasEntre(marco.fim, hoje);
        const bucket = bucketParaDias(dias, diasAntes2, diasAntes1);
        if (bucket) {
          const ultimoEnvioBucket = buckets[bucket];
          const deveAvisar =
            bucket === "atrasado"
              ? !ultimoEnvioBucket ||
                diasEntre(hoje, ultimoEnvioBucket.slice(0, 10)) >= diasRepeticaoAtrasado
              : !ultimoEnvioBucket;
          if (deveAvisar) {
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
              respHtml +
              `<p>Avaliação automática do Painel de Prazos.</p>` +
              `<p><a href="${MEU_PAINEL_URL}">Abrir o Meu Painel</a> para ver e concluir este e os demais itens atribuídos a você.</p>` +
              rodapeGerente;
            try {
              const r = await fetch(`${SUPABASE_URL}/functions/v1/send-notification`, {
                method: "POST",
                headers: svcHeaders(),
                body: JSON.stringify({ to: Array.from(destinatarios), subject, html }),
              });
              if (!r.ok) throw new Error(await r.text());
              buckets[bucket] = new Date().toISOString();
              mudouBuckets = true;
              avisouPrazoAgora = true;
              resultado.avisosEnviados++;
            } catch (e) {
              resultado.erros.push(`Falha ao notificar marco ${marcoId} (prazo): ${String(e)}`);
            }
          }
        }

        // --- 2) Aviso de silêncio (Execução parada há 5+ dias, marco já iniciado) ---
        if (!avisouPrazoAgora && marco.ini && diasEntre(hoje, marco.ini) >= 0) {
          const referencia = marco.execStatusAtualizadoEm
            ? marco.execStatusAtualizadoEm.slice(0, 10)
            : marco.ini;
          const diasSemAtualizar = diasEntre(hoje, referencia);
          if (diasSemAtualizar >= diasSilencioLimite) {
            const ultimoSilencio = buckets.silencio;
            const deveAvisarSilencio =
              !ultimoSilencio ||
              diasEntre(hoje, ultimoSilencio.slice(0, 10)) >= diasSilencioRepeticao;
            if (deveAvisarSilencio) {
              const subject = `Marco "${marco.marco}" sem atualização de execução — ${
                rec.nomeProjeto || rec.protocolo || ""
              }`;
              const html =
                `<p>O marco <b>${marcoNome}</b> do cronograma do projeto <b>${projetoNome}</b>` +
                (rec.protocolo ? ` (${esc(rec.protocolo)})` : "") +
                ` já começou e está sem nenhuma atualização de <b>Execução</b> há <b>${diasSemAtualizar} dias</b>.</p>` +
                respHtml +
                `<p>Avaliação automática do Painel de Prazos.</p>` +
                `<p><a href="${MEU_PAINEL_URL}">Abrir o Meu Painel</a> para informar o status de execução deste e dos demais itens atribuídos a você.</p>` +
                rodapeGerente;
              try {
                const r = await fetch(`${SUPABASE_URL}/functions/v1/send-notification`, {
                  method: "POST",
                  headers: svcHeaders(),
                  body: JSON.stringify({ to: Array.from(destinatarios), subject, html }),
                });
                if (!r.ok) throw new Error(await r.text());
                buckets.silencio = new Date().toISOString();
                mudouBuckets = true;
                resultado.avisosEnviados++;
              } catch (e) {
                resultado.erros.push(`Falha ao notificar marco ${marcoId} (silêncio): ${String(e)}`);
              }
            }
          }
        }

        if (mudouBuckets) {
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
    }
  } catch (e) {
    return json({ error: String(e), ...resultado }, 500);
  }

  return json({ ok: true, hoje, ...resultado });
});
