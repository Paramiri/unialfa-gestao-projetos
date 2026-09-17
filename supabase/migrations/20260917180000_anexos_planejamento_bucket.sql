-- Bucket de Storage para os anexos reais (upload de arquivo) da aba Viabilidade
-- do Planejamento e Desenvolvimento de Projeto (04) — mesmo padrao ja usado por
-- anexos-demanda (01) e anexos-smp (06): privado, 15 MB por arquivo, qualquer
-- usuario autenticado pode enviar/ver/remover (sem restricao extra por equipe,
-- espelhando exatamente a politica dos dois buckets existentes).

insert into storage.buckets (id, name, public, file_size_limit)
values ('anexos-planejamento', 'anexos-planejamento', false, 15728640)
on conflict (id) do nothing;

DROP POLICY IF EXISTS anexos_planejamento_insert ON storage.objects;
DROP POLICY IF EXISTS anexos_planejamento_select ON storage.objects;
DROP POLICY IF EXISTS anexos_planejamento_delete ON storage.objects;

CREATE POLICY anexos_planejamento_insert ON storage.objects FOR INSERT TO authenticated
  WITH CHECK (bucket_id = 'anexos-planejamento');
CREATE POLICY anexos_planejamento_select ON storage.objects FOR SELECT TO authenticated
  USING (bucket_id = 'anexos-planejamento');
CREATE POLICY anexos_planejamento_delete ON storage.objects FOR DELETE TO authenticated
  USING (bucket_id = 'anexos-planejamento');
