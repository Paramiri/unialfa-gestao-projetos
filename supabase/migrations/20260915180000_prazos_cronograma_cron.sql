-- ============================================================
-- Migração: cron diário do Painel de Prazos (Edge Function
-- verificar-prazos-cronograma)
--
-- Habilita pg_cron + pg_net (nenhum dos dois estava habilitado antes) e
-- agenda uma chamada diária (11:00 UTC = 08:00 horário de Brasília, sem
-- horário de verão desde 2019) à Edge Function verificar-prazos-cronograma,
-- que varre os cronogramas e dispara os avisos de prazo.
--
-- A chamada usa a anon key do projeto (chave pública, já embutida em todo
-- o front-end) só para passar pelo gate de autenticação padrão das Edge
-- Functions — a function em si usa sua própria SUPABASE_SERVICE_ROLE_KEY
-- (injetada automaticamente pelo runtime) para o acesso privilegiado ao
-- banco. Nenhum segredo novo é armazenado nesta migração.
--
-- Aplicada em treino (uuxvdulunrwppbmofyux) e produção (fiarntunpqteopwjkhjg)
-- em 15/09/2026 — a URL/anon key abaixo são as de PRODUÇÃO; a versão
-- aplicada em treino usa a URL e a anon key de treino (mesmo conteúdo,
-- só esses dois valores trocados — não versionada à parte porque não há
-- segredo real envolvido, só o endpoint/projeto de destino).
--
-- Idempotente: reaplicar não duplica o job (cron.unschedule antes de
-- cron.schedule) nem falha se as extensões já estiverem habilitadas.
-- ============================================================

create extension if not exists pg_cron;
create extension if not exists pg_net;

do $$
begin
  if exists (select 1 from cron.job where jobname = 'verificar-prazos-cronograma-diario') then
    perform cron.unschedule('verificar-prazos-cronograma-diario');
  end if;
end $$;

select cron.schedule(
  'verificar-prazos-cronograma-diario',
  '0 11 * * *',
  $cron$
  select net.http_post(
    url := 'https://fiarntunpqteopwjkhjg.supabase.co/functions/v1/verificar-prazos-cronograma',
    headers := jsonb_build_object(
      'Content-Type', 'application/json',
      'apikey', 'sb_publishable_RO-UPexCYhZ0rVZiIYWunA_a7YqodnQ',
      'Authorization', 'Bearer sb_publishable_RO-UPexCYhZ0rVZiIYWunA_a7YqodnQ'
    ),
    body := '{}'::jsonb
  );
  $cron$
);
