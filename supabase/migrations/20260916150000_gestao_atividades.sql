-- ============================================================
-- Migração: nova página "Gestão de Atividades" (Admin/PMO)
--
-- Página de gestão, restrita por papel (mesmo padrão do Painel
-- Executivo/Painel de Prazos), que agrega — por conta responsável —
-- os itens pendentes das 5 categorias já usadas pelo Meu Painel
-- (marcos, tarefas, riscos, entradas, entraves/encaminhamentos),
-- para Admin/PMO enxergar a carga de trabalho de toda a equipe.
--
-- Os dados de atividade em si (kv_store) já são de leitura ampla
-- para qualquer autenticado (é assim que o próprio Meu Painel lê
-- projetos que não são do usuário logado) — não precisa de função
-- nova para isso. O que falta é a lista de nome/papel de todos os
-- usuários, hoje bloqueada pela política perfis_select
-- ((auth.uid() = id) OR is_admin()) para quem não é Admin.
--
-- listar_perfis_gestao_atividades() é a exceção deliberada: só
-- devolve linhas quando quem chama tem papel Admin OU está na lista
-- configurável 'papeis_gestao_atividades' (tabela configuracoes) —
-- a checagem fica dentro da função (SECURITY DEFINER), não só na
-- tela, para não depender só da interface esconder o link.
--
-- Também insere a linha de configuração 'papeis_gestao_atividades'
-- (começa vazia — só Admin — até o Admin liberar outros papéis em
-- Administração > Configurações) e corrige uma lacuna encontrada de
-- passagem: a chave 'papeis_painel_prazos' nunca tinha sido inserida
-- na tabela configuracoes, então o Admin não conseguia salvar quem
-- mais pode ver o Painel de Prazos pela interface (o PATCH da tela
-- não cria linha nova, só atualiza uma já existente).
-- ============================================================

INSERT INTO public.configuracoes (chave, valor)
VALUES
  ('papeis_gestao_atividades', '[]'::jsonb),
  ('papeis_painel_prazos', '[]'::jsonb)
ON CONFLICT (chave) DO NOTHING;

CREATE OR REPLACE FUNCTION public.pode_ver_gestao_atividades()
RETURNS boolean
LANGUAGE sql STABLE SECURITY DEFINER SET search_path TO 'public' AS $$
  SELECT is_admin() OR EXISTS (
    SELECT 1
    FROM public.configuracoes c
    JOIN public.perfis p ON p.id = auth.uid()
    WHERE c.chave = 'papeis_gestao_atividades'
      AND c.valor ? p.papel
  );
$$;

CREATE OR REPLACE FUNCTION public.listar_perfis_gestao_atividades()
RETURNS TABLE(id uuid, email text, nome text, papel text)
LANGUAGE sql STABLE SECURITY DEFINER SET search_path TO 'public' AS $$
  SELECT perfis.id, perfis.email, perfis.nome, perfis.papel
  FROM public.perfis
  WHERE public.pode_ver_gestao_atividades()
  ORDER BY perfis.nome NULLS LAST, perfis.email;
$$;
