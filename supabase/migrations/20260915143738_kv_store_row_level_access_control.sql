-- ============================================================
-- Migração: colunas geradas + RLS real no kv_store
-- Espelha, direto no banco, as regras já aplicadas no JS de cada
-- formulário: criador/Gerente de Projetos/Admin edita, Gerente de
-- Projetos/Admin exclui, e (nos 7 formulários com essa restrição)
-- equipe do projeto continua sendo exigida pra gravar.
--
-- Aplicada em treino (uuxvdulunrwppbmofyux) e produção
-- (fiarntunpqteopwjkhjg) em 15/09/2026, testada de ponta a ponta
-- via API direta (bypassando o JS do app) com contas reais de
-- treino antes de ir para produção.
--
-- Idempotente (pode ser reaplicada sem erro): ADD COLUMN IF NOT
-- EXISTS / DROP POLICY IF EXISTS + recriação.
-- ============================================================

-- ---- funções auxiliares (SECURITY DEFINER, mesmo padrão de is_admin()) ----

CREATE OR REPLACE FUNCTION public.safe_uuid(t text) RETURNS uuid
LANGUAGE plpgsql IMMUTABLE AS $$
BEGIN
  RETURN t::uuid;
EXCEPTION WHEN others THEN
  RETURN NULL;
END;
$$;

CREATE OR REPLACE FUNCTION public.meu_papel() RETURNS text
LANGUAGE sql STABLE SECURITY DEFINER SET search_path TO 'public' AS $$
  SELECT papel FROM perfis WHERE id = auth.uid();
$$;

CREATE OR REPLACE FUNCTION public.sou_membro_projeto(p_projeto_id uuid) RETURNS boolean
LANGUAGE sql STABLE SECURITY DEFINER SET search_path TO 'public' AS $$
  SELECT p_projeto_id IS NOT NULL AND EXISTS(
    SELECT 1 FROM projeto_equipe WHERE projeto_id = p_projeto_id AND usuario_id = auth.uid()
  );
$$;

-- ---- colunas geradas (derivadas de key/value, sem precisar mudar o app) ----

ALTER TABLE kv_store ADD COLUMN IF NOT EXISTS tipo text
  GENERATED ALWAYS AS (split_part(key, ':', 1)) STORED;

ALTER TABLE kv_store ADD COLUMN IF NOT EXISTS created_by text
  GENERATED ALWAYS AS (
    CASE WHEN value ~ '^[[:space:]]*\{' THEN value::jsonb->>'criadoPorEmail' ELSE NULL END
  ) STORED;

ALTER TABLE kv_store ADD COLUMN IF NOT EXISTS projeto_id uuid
  GENERATED ALWAYS AS (
    CASE WHEN value ~ '^[[:space:]]*\{' THEN safe_uuid(value::jsonb->>'projetoId') ELSE NULL END
  ) STORED;

CREATE INDEX IF NOT EXISTS kv_store_tipo_idx ON kv_store (tipo);
CREATE INDEX IF NOT EXISTS kv_store_projeto_id_idx ON kv_store (projeto_id);

-- ---- políticas RLS ----

DROP POLICY IF EXISTS kv_select ON kv_store;
DROP POLICY IF EXISTS kv_insert ON kv_store;
DROP POLICY IF EXISTS kv_update ON kv_store;
DROP POLICY IF EXISTS kv_delete ON kv_store;

-- leitura: sem mudança (continua aberta pra qualquer autenticado — ver
-- item "quem pode só VER um registro" como possível melhoria futura)
CREATE POLICY kv_select ON kv_store FOR SELECT TO authenticated USING (true);

-- criação: sem mudança, exceto equipe do projeto nos 7 formulários que já exigem isso
-- (canvas, tap, plan, eap, smp, tep, rla)
CREATE POLICY kv_insert ON kv_store FOR INSERT TO authenticated WITH CHECK (
  tipo NOT IN ('canvas','tap','plan','eap','smp','tep','rla')
  OR sou_membro_projeto(projeto_id)
  OR meu_papel() = 'admin'
);

-- edição: quem criou, Gerente de Projetos ou Admin; e (nos 7 de equipe)
-- também precisa ser da equipe do projeto pra gravar
CREATE POLICY kv_update ON kv_store FOR UPDATE TO authenticated USING (
  tipo NOT IN ('ata','canvas','demanda','eap','plan','rla','smp','tap','tep')
  OR created_by = auth.email()
  OR meu_papel() IN ('admin','gerente_projetos')
) WITH CHECK (
  tipo NOT IN ('canvas','tap','plan','eap','smp','tep','rla')
  OR sou_membro_projeto(projeto_id)
  OR meu_papel() = 'admin'
);

-- exclusão: só Gerente de Projetos ou Admin (sem exigir equipe, igual ao JS)
CREATE POLICY kv_delete ON kv_store FOR DELETE TO authenticated USING (
  tipo NOT IN ('ata','canvas','demanda','eap','plan','rla','smp','tap','tep')
  OR meu_papel() IN ('admin','gerente_projetos')
);

-- Fora do escopo desta migração (deliberado, ver Regras de Acesso e
-- Permissões, seção 5.4):
--   - plancom_data / relsit_data / relent_data: continuam sem trava de
--     linha no banco (o Gate 2 do Relatório de Entregas e o controle de
--     papel do Plano de Comunicação seguem só no JS).
--   - Gate 1 (Solicitação de Demanda) e Gate 2 (TAP) continuam admin-only
--     só no JS, não no RLS — a política acima libera edição pra
--     criador/GP/Admin de forma geral, não campo a campo.
--   - lock:*, relsit_hist:*, relent_hist:*: sempre livres pra qualquer
--     autenticado (edição simultânea e histórico de versões).
