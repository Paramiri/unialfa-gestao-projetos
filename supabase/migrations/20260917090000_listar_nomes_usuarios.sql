-- ============================================================
-- Migração: listar_emails_usuarios() passa a devolver também o nome
--
-- Contexto: o seletor de "Responsável (conta)" (Riscos do TAP,
-- Entradas do Planejamento, Entraves/Encaminhamentos do Relatório de
-- Situação, todos com suporte a múltiplas pessoas desde 16/09/2026)
-- mostrava só o e-mail de cada conta na lista de marcação - a
-- interface pediu para mostrar o nome, caindo para o e-mail só
-- quando a conta não tiver nome cadastrado.
--
-- A politica de SELECT de `perfis` (perfis_select) so libera a
-- propria linha para quem nao e Admin ((auth.uid() = id) OR
-- is_admin()), entao um SELECT direto na tabela nao serve para
-- montar essa lista. listar_emails_usuarios() ja contornava isso
-- (SECURITY DEFINER) devolvendo so o e-mail, de proposito, para nao
-- expor nome/telefone/papel de todo mundo. Como a superficie exposta
-- por essa funcao ja e ampla (qualquer usuario autenticado, para
-- montar o seletor de "todas as contas" nos Entraves/Encaminhamentos),
-- e nome ja aparece como texto livre em varios lugares do sistema
-- (Responsavel, Gerente do projeto, Solicitante), adicionar a coluna
-- nome aqui nao muda a natureza da informacao exposta - continua sem
-- telefone nem papel, que sao os dados mais sensiveis do perfil.
--
-- Precisa recriar a funcao (DROP + CREATE) porque o tipo de retorno
-- muda (nao da so para REPLACE). Aplicada em treino
-- (uuxvdulunrwppbmofyux) e producao (fiarntunpqteopwjkhjg) em
-- 17/09/2026.
-- ============================================================

DROP FUNCTION IF EXISTS public.listar_emails_usuarios();

CREATE OR REPLACE FUNCTION public.listar_emails_usuarios()
RETURNS TABLE(email text, nome text)
LANGUAGE sql STABLE SECURITY DEFINER SET search_path TO 'public' AS $$
  SELECT perfis.email, perfis.nome FROM perfis ORDER BY perfis.email;
$$;
