# UNIALFA — Sistema de Gestão de Projetos

Site estático (HTML/CSS/JS) publicado no GitHub Pages, com Supabase como backend (auth, tabelas via REST). Repositório: `https://github.com/Paramiri/unialfa-gestao-projetos`. Todas as regras de acesso são aplicadas no JS de cada formulário — não há políticas RLS documentadas no Supabase.

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

## Padrão de deploy

`git add` dos arquivos específicos → commit em português com `Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>` → `git push` → confirmar publicação fazendo polling no GitHub Pages (`https://paramiri.github.io/unialfa-gestao-projetos/...`) até o conteúdo novo aparecer.

`app.css`/`app.js` são referenciados em `index.html` com cache-busting `?v=N` — incrementar o número sempre que esses dois arquivos mudarem.
