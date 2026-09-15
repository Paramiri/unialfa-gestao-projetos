-- ============================================================
-- Migração: função listar_emails_usuarios()
--
-- Necessária para o campo "Responsável (conta)" de Entraves e
-- Encaminhamentos no Relatório de Situação — diferente do mesmo campo
-- no cronograma do Planejamento e no Registro de riscos do TAP (que
-- listam só a equipe do projeto, já visível via projeto_equipe), Entraves
-- e Encaminhamentos não têm projeto vinculado, então o seletor precisa
-- listar todas as contas do sistema.
--
-- A política de SELECT de `perfis` (perfis_select) só libera a própria
-- linha para quem não é Admin ((auth.uid() = id) OR is_admin()) — de
-- propósito, para não expor nome/telefone/papel de todo mundo a
-- qualquer usuário autenticado. Por isso esta função devolve só o
-- e-mail (nada de nome/telefone/papel), via SECURITY DEFINER, mesmo
-- padrão já usado em meu_papel()/sou_membro_projeto() (ver migração
-- 20260915143738).
--
-- Aplicada em treino (uuxvdulunrwppbmofyux) e produção
-- (fiarntunpqteopwjkhjg) em 16/09/2026.
--
-- Idempotente (CREATE OR REPLACE).
-- ============================================================

CREATE OR REPLACE FUNCTION public.listar_emails_usuarios()
RETURNS TABLE(email text)
LANGUAGE sql STABLE SECURITY DEFINER SET search_path TO 'public' AS $$
  SELECT perfis.email FROM perfis ORDER BY perfis.email;
$$;
