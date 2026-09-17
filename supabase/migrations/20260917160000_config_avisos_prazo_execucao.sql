-- ============================================================
-- Migração: parametriza os dias das duas regras de disparo automático
-- (Edge Function verificar-prazos-cronograma)
--
-- Antes, os números de dias de cada regra eram fixos no código da
-- function:
--   - Aviso de prazo: 5 dias antes / 2 dias antes / repete a cada
--     7 dias enquanto atrasado.
--   - Aviso de silêncio: dispara com 5+ dias sem nenhuma mudança no
--     campo Execução do marco, repete a cada 5 dias enquanto
--     continuar parado.
--
-- Esta migração insere as 5 chaves correspondentes na tabela
-- `configuracoes` (mesmo padrão já usado pelas metas do Relatório de
-- Resultados — meta_marcos_no_prazo etc.), com os mesmos valores que
-- já estavam fixos no código, para editar em Administração >
-- Configurações > Regras de aviso automático sem mudar nada de
-- comportamento até o Admin decidir alterar. A Edge Function lê essas
-- chaves a cada execução, caindo no mesmo valor padrão se a linha
-- não existir ou tiver um valor inválido — nunca fica sem rodar por
-- causa de uma configuração ausente ou zerada.
-- ============================================================

INSERT INTO public.configuracoes (chave, valor)
VALUES
  ('aviso_prazo_dias_antes1', '5'::jsonb),
  ('aviso_prazo_dias_antes2', '2'::jsonb),
  ('aviso_prazo_dias_repeticao_atrasado', '7'::jsonb),
  ('aviso_silencio_dias_sem_atualizar', '5'::jsonb),
  ('aviso_silencio_dias_repeticao', '5'::jsonb)
ON CONFLICT (chave) DO NOTHING;
