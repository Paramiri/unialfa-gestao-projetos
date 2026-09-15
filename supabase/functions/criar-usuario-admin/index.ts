// Edge Function: criar-usuario-admin
// Acionada pelo botão "+ Adicionar usuário" da Administração
// (13 - administracao-usuarios.html), visível só para quem tem papel
// 'admin'. Substitui o antigo pré-cadastro em `perfis_pendentes`
// (que só virava conta de verdade no primeiro login da pessoa): agora a
// conta é criada de verdade no Supabase Auth na hora, via Admin API, para
// que a pessoa já possa ser adicionada à equipe de um projeto
// imediatamente, sem esperar ela logar.
//
// Quando a pessoa eventualmente fizer o primeiro login (link mágico ou
// Microsoft), o e-mail já bate com uma conta existente — o Supabase Auth
// autentica direto nela (sem duplicar), e ensureProfile() (definido em
// cada formulário) encontra o perfil já criado e não mexe em nada, o
// mesmo caminho já usado para qualquer conta que já tem perfil.
//
// Segredos necessários: nenhum novo — reaproveita SUPABASE_URL e
// SUPABASE_SERVICE_ROLE_KEY (injetados automaticamente pelo runtime das
// Edge Functions) para chamar a Auth Admin API e gravar em public.perfis.
//
// Verificação de JWT permanece ligada (padrão do Supabase) — além disso,
// esta function confere explicitamente que quem chamou tem papel 'admin'
// em public.perfis (mesmo padrão já usado em reset-treino), já que criar
// uma conta de verdade é uma ação sensível.

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
  return {
    apikey: SUPABASE_SERVICE_ROLE_KEY!,
    Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
    "Content-Type": "application/json",
    ...(extra || {}),
  };
}

// Pagina toda a lista de usuários do Auth e procura pelo e-mail exato — mais lento que um
// filtro no servidor, mas confiável, e o volume de contas desta ferramenta (dezenas, não
// milhares) torna isso barato o suficiente. Limite de segurança de 20 páginas (2000 contas).
async function buscarUsuarioIdPorEmail(email: string): Promise<string | null> {
  const alvo = email.toLowerCase();
  for (let page = 1; page <= 20; page++) {
    const r = await fetch(`${SUPABASE_URL}/auth/v1/admin/users?page=${page}&per_page=100`, {
      headers: svcHeaders(),
    });
    if (!r.ok) return null;
    const data = await r.json();
    const usuarios = Array.isArray(data) ? data : data?.users || [];
    const achado = usuarios.find((u: { email?: string }) => (u.email || "").toLowerCase() === alvo);
    if (achado) return achado.id;
    if (usuarios.length < 100) break;
  }
  return null;
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }
  if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
    return json({ error: "SUPABASE_URL/SUPABASE_SERVICE_ROLE_KEY não disponíveis" }, 500);
  }

  const authHeader = req.headers.get("Authorization") || "";
  const token = authHeader.replace(/^Bearer\s+/i, "").trim();
  if (!token) return json({ error: "Não autenticado." }, 401);

  const userResp = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
    headers: { apikey: SUPABASE_SERVICE_ROLE_KEY, Authorization: `Bearer ${token}` },
  });
  if (!userResp.ok) return json({ error: "Sessão inválida ou expirada." }, 401);
  const caller = await userResp.json();
  if (!caller?.id) return json({ error: "Não foi possível identificar o usuário." }, 401);

  const perfilResp = await fetch(`${SUPABASE_URL}/rest/v1/perfis?id=eq.${caller.id}&select=papel`, {
    headers: svcHeaders(),
  });
  const perfilRows = perfilResp.ok ? await perfilResp.json() : [];
  if (perfilRows?.[0]?.papel !== "admin") {
    return json({ error: "Apenas administradores (PMO) podem executar esta ação." }, 403);
  }

  let payload: { email?: unknown; nome?: unknown; telefone?: unknown; papel?: unknown };
  try {
    payload = await req.json();
  } catch {
    return json({ error: "Corpo da requisição não é um JSON válido" }, 400);
  }
  const email = typeof payload.email === "string" ? payload.email.trim().toLowerCase() : "";
  const nome = typeof payload.nome === "string" && payload.nome.trim() ? payload.nome.trim() : null;
  const telefone = typeof payload.telefone === "string" && payload.telefone.trim() ? payload.telefone.trim() : null;
  const PAPEIS_VALIDOS = ["solicitante", "gerente_projetos", "gestor_responsavel", "dono_negocio", "alta_gestao", "admin"];
  const papel = typeof payload.papel === "string" && PAPEIS_VALIDOS.includes(payload.papel) ? payload.papel : "solicitante";
  if (!email || !email.includes("@")) {
    return json({ error: "Informe um e-mail válido." }, 400);
  }

  const existente = await fetch(`${SUPABASE_URL}/rest/v1/perfis?email=eq.${encodeURIComponent(email)}&select=id`, {
    headers: svcHeaders(),
  }).then((r) => (r.ok ? r.json() : []));
  if (existente?.length) {
    return json({ error: "Este e-mail já tem uma conta ativa." }, 409);
  }

  // Cria a conta de verdade no Supabase Auth (Admin API) — email_confirm:true porque
  // é o Admin do sistema vouching pela pessoa, sem exigir confirmação por e-mail.
  let userId: string | null = null;
  const criarResp = await fetch(`${SUPABASE_URL}/auth/v1/admin/users`, {
    method: "POST",
    headers: svcHeaders(),
    body: JSON.stringify({ email, email_confirm: true, user_metadata: nome ? { name: nome } : {} }),
  });
  if (criarResp.ok) {
    const novo = await criarResp.json();
    userId = novo?.id || null;
  } else {
    const erro = await criarResp.json().catch(() => ({}));
    const jaExiste = criarResp.status === 422 || /already.*registered|already.*exists/i.test(JSON.stringify(erro));
    if (!jaExiste) {
      return json({ error: "Falha ao criar a conta: " + JSON.stringify(erro) }, 500);
    }
    // A pessoa já teve conta antes (ex.: removida da lista sem apagar a conta em si) —
    // reaproveita a conta existente no Auth em vez de tentar duplicar. O parâmetro
    // ?email= do endpoint de listagem não filtra de forma confiável (já observado
    // devolvendo o primeiro usuário da lista, ignorando o filtro) — por isso a busca
    // aqui pagina a lista inteira e compara o e-mail exatamente, em vez de confiar nele.
    userId = await buscarUsuarioIdPorEmail(email);
    if (!userId) {
      return json({ error: "E-mail já tem uma conta no Auth, mas não foi possível recuperá-la: " + JSON.stringify(erro) }, 500);
    }
  }

  const perfilCriado = await fetch(`${SUPABASE_URL}/rest/v1/perfis`, {
    method: "POST",
    headers: svcHeaders({ Prefer: "return=representation" }),
    body: JSON.stringify({ id: userId, email, nome, telefone, papel }),
  });
  if (!perfilCriado.ok) {
    return json({ error: "Conta criada no Auth, mas falhou ao gravar o perfil: " + (await perfilCriado.text()) }, 500);
  }
  const perfilRows2 = await perfilCriado.json();

  // Se havia um pré-cadastro antigo (perfis_pendentes) com o mesmo e-mail, remove —
  // a conta real já foi criada com os mesmos dados, não faz mais sentido manter os dois.
  await fetch(`${SUPABASE_URL}/rest/v1/perfis_pendentes?email=eq.${encodeURIComponent(email)}`, {
    method: "DELETE",
    headers: svcHeaders(),
  }).catch(() => {});

  return json({ ok: true, usuario: perfilRows2?.[0] || { id: userId, email, nome, telefone, papel } });
});
