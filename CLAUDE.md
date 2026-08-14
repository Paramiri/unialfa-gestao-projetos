# UNIALFA — Sistema de Gestão de Projetos

Site estático (HTML/CSS/JS) publicado no GitHub Pages, com Supabase como backend (auth, tabelas via REST). Repositório: `https://github.com/Paramiri/unialfa-gestao-projetos`. Ao contrário do que uma versão antiga desta nota dizia, a maioria das tabelas **tem RLS habilitado** (`perfis`, `perfis_pendentes`, `projeto_equipe`, `projeto_historico`, `projetos`, `registro_historico`, `validador_avaliacoes`, `configuracoes`, `kv_store`, `push_subscriptions`) — confirmado direto no banco em 10/08/2026, não presuma "sem RLS" sem checar (`select relname, relrowsecurity from pg_class ...`). A maior parte das políticas usa a função `is_admin()` (`SECURITY DEFINER`, evita o [[feedback_rls_recursion_bug|bug de recursão]] de subconsultar a própria tabela) para liberar ações restritas a Admin; tabelas com dado por usuário (ex.: `push_subscriptions`, `perfis_pendentes`) também restringem por `auth.email() = <coluna>`. Regras de negócio mais finas (ex.: quem pode editar um registro específico) continuam aplicadas no JS de cada formulário, não no banco.

## Documentação oficial de regras de acesso

`Regras de Acesso e Permissoes - Sistema UNIALFA.docx` (na raiz do repositório) é o **documento oficial** e versionado de todas as regras de acesso e permissão do sistema: login obrigatório, acesso sem login (convidado), papéis de usuário, gates de aprovação (Gate 1/Gate 2) e restrição por equipe de projeto.

**Regra permanente:** sempre que uma regra de acesso for incluída, alterada ou removida no código — por exemplo, um novo formulário passa a exigir login/equipe, uma opção de "sem login" é adicionada, um novo papel é criado, uma trava de gate muda — este documento deve ser atualizado no mesmo commit (ou logo em seguida) para continuar refletindo o estado real do código. Isso vale tanto para mudanças feitas por mim quanto pelo usuário.

Como atualizar:
1. Editar o conteúdo em `gerar-regras-acesso.ps1` (raiz do repo) — cada seção é montada com as funções auxiliares `H1`/`H2`/`P`/`Bul`/tabela.
2. Rodar o script via PowerShell: `powershell -File "gerar-regras-acesso.ps1"` — ele usa automação COM do Microsoft Word (`New-Object -ComObject Word.Application`) para gerar o `.docx` diretamente na raiz do repo, sobrescrevendo o anterior.
3. Conferir o resultado (ex.: exportar para PDF via `$doc.SaveAs2($pdfPath, 17)` e ler o PDF) antes de commitar.
4. Commitar o `.docx` e o `.ps1` junto com a mudança de código que motivou a atualização.

Motivo de usar Word COM em vez do pacote `docx` (Node.js) da skill padrão: este ambiente não tem Node.js, pandoc, nem LibreOffice instalados — apenas o Microsoft Word está disponível localmente.

## Manual de Uso da ferramenta

`Manual de Uso - Ferramenta de Gestao de Projetos.docx` (na raiz do repositório) é o **manual oficial** e versionado de uso do sistema, com prints de tela e passo a passo de cada formulário, papéis, gates e relatórios.

**Regra permanente:** sempre que uma funcionalidade for incluída, alterada ou removida no código — novo formulário, novo campo relevante, nova regra de negócio (ex.: registro de riscos, sinalização automática, prioridade da EAP), mudança de fluxo — este manual deve ser atualizado no mesmo commit (ou logo em seguida) para continuar refletindo o estado real do sistema. Isso vale tanto para mudanças feitas por mim quanto pelo usuário.

Como atualizar:
1. Editar o conteúdo em `gerar-manual-uso.ps1` (raiz do repo) — inclui capa (com changelog de versão), TOC automático, seções por formulário com tabelas de campos e prints.
2. Rodar o script via PowerShell: `powershell -File "gerar-manual-uso.ps1" -ImgDir "<pasta com os prints>"` — **atenção**: `-ImgDir` precisa apontar para a pasta que já contém os prints existentes (não uma pasta vazia/nova), senão o regenerado perde todas as imagens antigas silenciosamente. Se novos prints forem necessários, tirar antes e colocar nessa mesma pasta.
3. Conferir o resultado (exportar para PDF via `$doc.SaveAs2($pdfPath, 17)` e ler o PDF) antes de commitar — comparar a contagem de páginas com a versão anterior como sinal rápido de que nenhuma imagem foi perdida.
4. Commitar o `.docx` e o `.ps1` junto com a mudança de código que motivou a atualização.

Se essa atualização for delegada a um agente, a verificação do passo 3 (conferir o PDF de fato, não só o texto) deve ser feita antes de reportar a tarefa como concluída — uma checagem só de texto (ex. `pdftotext`) não detecta perda de imagens.

## Central de Ajuda (in-app)

`15 - central-ajuda.html` é uma **terceira superfície de documentação**, independente dos dois `.docx` acima — é a versão navegável (com busca e menu lateral) do Manual de Uso, publicada como página do próprio site, e é a que os usuários realmente abrem a partir do ícone `?` em cada formulário. Uma mudança de funcionalidade documentada no Manual de Uso ou nas Regras de Acesso e esquecida aqui deixa o `?` do sistema desatualizado mesmo com os `.docx` corretos — já aconteceu uma vez (notificações push).

**Regra permanente:** sempre que o Manual de Uso ou as Regras de Acesso forem atualizados por uma mudança de funcionalidade, checar se `15 - central-ajuda.html` também precisa de uma seção nova/ajustada (adicionar `<a>` no menu lateral `#sbNav` + `<section class="sec" id="...">` correspondente) — no mesmo commit da mudança de código, junto com os `.docx`.

Diferente de `app.css`/`app.js`, essa página não é referenciada por `<link>`/`<script>` num único lugar, e sim por um link `?` (`class="doc-help"`/`"head-help"`) em cada um dos 12 formulários + o item de menu em `app.js` — todos com cache-busting `?v=N` (ex.: `href="15 - central-ajuda.html?v=1#passo-1"`). Sempre que o conteúdo de `15 - central-ajuda.html` mudar, incrementar o `N` em **todos** esses links (não só no arquivo em si), senão o GitHub Pages/CDN pode continuar servindo a versão em cache por até 10 minutos.

## Ambiente de Treinamento (sincronização)

`Paramiri/unialfa-gestao-projetos-treino` é um repositório-espelho, publicado separadamente no GitHub Pages (`https://paramiri.github.io/unialfa-gestao-projetos-treino/`), usado como ambiente de treinamento/demonstração: mesmo código, backend Supabase totalmente separado (projeto `unialfa-treinamento`, ref `uuxvdulunrwppbmofyux`), selecionado em runtime pelo hostname (`IS_TREINO = location.href.indexOf('treino')>-1`, já presente em todos os arquivos HTML). Scripts de seed/reset dos dados de exemplo desse ambiente ficam versionados em `treino/` neste repositório; o botão que roda esse reset a partir da Administração (produção) chama a Edge Function `reset-treino`.

**Regra permanente:** sempre que um arquivo do site de produção for alterado — HTML, `app.js`/`app.css`, ou uma Edge Function em `supabase/functions/` — replicar a mesma mudança no repositório-espelho, no mesmo commit (ou logo em seguida) da mudança em produção, para o ambiente de treinamento continuar refletindo a mesma versão do sistema. Isso vale tanto para mudanças feitas por mim quanto pelo usuário, exatamente como já vale para o Manual de Uso e a Central de Ajuda.

Como sincronizar:
1. Copiar os arquivos alterados de `unialfa-gestao-projetos` (produção) para a working copy local do repositório-espelho, preservando as poucas divergências propositais dele: sem arquivo `CNAME` (usa o endereço padrão do GitHub Pages) e sem a senha das contas de treino em texto puro no `README.md`.
2. Commitar e dar push no repositório-espelho (mesmo padrão de commit em português + `Co-Authored-By`).
3. Confirmar a publicação fazendo polling no GitHub Pages do espelho (`https://paramiri.github.io/unialfa-gestao-projetos-treino/...`) até o conteúdo novo aparecer — mesmo procedimento do "Padrão de deploy" abaixo, aplicado ao repositório-espelho.
4. Se a mudança envolver uma Edge Function nova ou alterada, reimplantá-la também no projeto de treinamento (`supabase functions deploy <nome> --project-ref uuxvdulunrwppbmofyux`) — sem copiar nenhum secret real de e-mail/IA (`RESEND_API_KEY`, `VAPID_*`, `ANTHROPIC_API_KEY`, `OPENAI_API_KEY`) para lá, propositalmente, para o ambiente de treino nunca poder disparar e-mail real, push real nem chamada de IA paga.

Se a mudança alterar o formato dos dados salvos por algum formulário (novo campo no objeto salvo no `kv_store`, nova coluna em `projetos`/`projeto_equipe`/`projeto_historico`), também avaliar se `treino/treino_seed.sql` precisa de ajuste para continuar gerando dados de exemplo válidos.

## Padrão de deploy

`git add` dos arquivos específicos → commit em português com `Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>` → `git push` → confirmar publicação fazendo polling no GitHub Pages (`https://paramiri.github.io/unialfa-gestao-projetos/...`) até o conteúdo novo aparecer.

`app.css`/`app.js` são referenciados em `index.html` com cache-busting `?v=N` — incrementar o número sempre que esses dois arquivos mudarem.
