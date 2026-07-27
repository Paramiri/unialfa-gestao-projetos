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

## Padrão de deploy

`git add` dos arquivos específicos → commit em português com `Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>` → `git push` → confirmar publicação fazendo polling no GitHub Pages (`https://paramiri.github.io/unialfa-gestao-projetos/...`) até o conteúdo novo aparecer.

`app.css`/`app.js` são referenciados em `index.html` com cache-busting `?v=N` — incrementar o número sempre que esses dois arquivos mudarem.
