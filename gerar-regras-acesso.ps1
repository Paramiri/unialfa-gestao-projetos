param([string]$OutPath = (Join-Path $PSScriptRoot "Regras de Acesso e Permissoes - Sistema UNIALFA.docx"))
$ErrorActionPreference = "Stop"

# Gera "Regras de Acesso e Permissoes - Sistema UNIALFA.docx" via automacao COM do Microsoft Word
# (Node.js/pandoc/LibreOffice nao estao disponiveis neste ambiente).
# Ver CLAUDE.md - secao "Documentacao oficial de regras de acesso" para quando reexecutar este script.

function RGB($r,$g,$b) { return [int]($r + ($g*256) + ($b*65536)) }
$colRed   = RGB 0xB9 0x1D 0x2E
$colInk   = RGB 0x1A 0x1A 0x1A
$colMuted = RGB 0x6A 0x6A 0x70
$colWhite = RGB 0xFF 0xFF 0xFF

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()
$sel = $word.Selection

function P($text, $size=11, $bold=$false, $italic=$false, $color=$colInk, $align="left", $spaceAfter=8) {
  $sel.Font.Size = $size
  $sel.Font.Bold = $bold
  $sel.Font.Italic = $italic
  $sel.Font.Color = $color
  $sel.ParagraphFormat.Alignment = if($align -eq "center"){1}else{0}
  $sel.ParagraphFormat.SpaceAfter = $spaceAfter
  $sel.ParagraphFormat.LineSpacing = 14
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Font.Bold = $false
  $sel.Font.Italic = $false
}

function H1($text) {
  $sel.Font.Size = 16
  $sel.Font.Bold = $true
  $sel.Font.Color = $colInk
  $sel.ParagraphFormat.Alignment = 0
  $sel.ParagraphFormat.SpaceBefore = 18
  $sel.ParagraphFormat.SpaceAfter = 8
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 1  # bottom border (wdBorderBottom=3, wdLineStyleSingle=1)
  $sel.ParagraphFormat.Borders.Item(3).Color = RGB 0xE4 0xE4 0xE7
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 0
  $sel.Font.Bold = $false
}

function H2($text) {
  $sel.Font.Size = 13
  $sel.Font.Bold = $true
  $sel.Font.Color = $colRed
  $sel.ParagraphFormat.Alignment = 0
  $sel.ParagraphFormat.SpaceBefore = 12
  $sel.ParagraphFormat.SpaceAfter = 6
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Font.Bold = $false
  $sel.Font.Color = $colInk
}

function Bul($text, $level=0) {
  $sel.Font.Size = 11
  $sel.Font.Bold = $false
  $sel.Font.Color = $colInk
  $sel.ParagraphFormat.SpaceAfter = 4
  $sel.ParagraphFormat.LineSpacing = 13
  if ($level -eq 0) { $sel.ParagraphFormat.LeftIndent = $word.CentimetersToPoints(0.6) }
  else { $sel.ParagraphFormat.LeftIndent = $word.CentimetersToPoints(1.2) }
  $sel.Range.ListFormat.ApplyBulletDefault()
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Range.ListFormat.RemoveNumbers()
  $sel.ParagraphFormat.LeftIndent = 0
}

function HR() {
  $sel.ParagraphFormat.SpaceBefore = 6
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 1
  $sel.ParagraphFormat.Borders.Item(3).Color = RGB 0xD1 0xD5 0xDB
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 0
}

# ---- Capa ----
P "UNIALFA - GERENCIA DE PROJETOS" 10 $true $false $colRed "left" 4
$sel.Font.Size = 26; $sel.Font.Bold = $true; $sel.Font.Color = $colInk
$sel.ParagraphFormat.SpaceAfter = 4
$sel.TypeText("Regras de Acesso e Permissoes")
$sel.TypeParagraph()
$sel.Font.Bold = $false
P "Sistema de Gestao de Projetos" 14 $false $true $colMuted "left" 20

P "Este documento descreve, de forma completa, todas as regras de acesso e permissao implementadas no sistema de gestao de projetos da UNIALFA (login, acesso sem login, papeis de usuario, gates de aprovacao e restricao por equipe de projeto), conforme o estado atual do codigo. Este e o documento oficial de referencia: qualquer inclusao, alteracao ou remocao de regra de acesso no sistema deve ser refletida aqui." 11 $false $false $colInk "left" 16

HR

# ---- 1 ----
H1 "1. Login obrigatorio (regra geral)"
P "Todos os 14 formularios do sistema exigem login antes de carregar ou salvar qualquer dado. O login pode ser feito de duas formas:"
Bul "Link magico por e-mail - o usuario informa o e-mail e recebe um link de acesso, sem senha."
Bul "Microsoft (SSO) - `"Entrar com Microsoft - UNIALFA`", usando a conta institucional."
P "A unica pagina aberta sem exigir login e a pagina inicial (index.html, mapa de diretrizes) - ela mostra o conteudo institucional livremente e so exibe a barra `"Conectado como...`" caso ja exista uma sessao ativa."

HR

# ---- 2 ----
H1 "2. Acesso sem login (convidado)"
P "Dois formularios oferecem uma opcao de envio sem necessidade de login, para facilitar o registro por pessoas que nao tem (ou nao querem usar) uma conta institucional:"
Bul "Solicitacao de Demanda"
Bul "Ata de Reuniao"
P "Nenhum outro formulario do sistema tem essa opcao."

H2 "2.1 Como e ativado"
P "Cada um dos dois formularios tem um interruptor independente, controlado exclusivamente pelo Admin, na aba `"Configuracoes`" da pagina de Administracao. Ou seja, e possivel ativar o `"sem login`" so na Solicitacao de Demanda, so na Ata, nas duas, ou em nenhuma - sao chaves separadas."

H2 "2.2 O que o convidado pode fazer"
P "Quando a opcao esta ativada, aparece o botao `"Continuar sem login`" na tela de entrada. A pessoa informa nome completo e e-mail e pode:"
Bul "Preencher e enviar um registro novo (uma nova solicitacao ou uma nova ata)."

H2 "2.3 O que o convidado NAO pode fazer"
P "O acesso sem login e somente para criacao. Um usuario sem login nao consegue, em nenhuma hipotese:"
Bul "Editar um registro ja existente - inclusive um que ele mesmo tenha criado."
Bul "Excluir um registro existente."
Bul "Alterar o status de um registro (ex.: aprovar, mudar etapa) - aplicavel a Ata de Reuniao."
P "Essas acoes ficam bloqueadas de duas formas: os botoes `"Editar dados`" e `"Excluir`" somem da tela para quem esta sem login, e, mesmo que a acao seja tentada por outro caminho, o sistema recusa e mostra um aviso explicando que so e permitido criar novos registros."

H2 "2.4 Identificacao do registro"
P "Todo registro criado sem login recebe um selo `"Sem login`" na listagem e no detalhe, junto com o nome e e-mail informados pela pessoa. Se depois um usuario autenticado normalmente abrir esse mesmo registro e salvar uma edicao, o selo `"Sem login`" e removido - o registro passa a valer como editado por um usuario identificado."

HR

# ---- 3 ----
H1 "3. Papeis de usuario"
P "Cada pessoa que faz login recebe um papel, usado para liberar ou restringir acoes especificas no sistema:"
Bul "Solicitante - papel padrao, atribuido automaticamente a todo novo usuario no primeiro login."
Bul "Gerente de Projetos"
Bul "Gestor Responsavel"
Bul "Dono do Negocio"
Bul "Alta Gestao"
Bul "Admin - papel de administracao do sistema, atribuido manualmente por quem ja e Admin."

H2 "3.1 Acoes exclusivas de Admin"
P "Somente usuarios com papel Admin podem:"
Bul "Acessar a pagina de Administracao (qualquer outro papel ve a mensagem `"Acesso restrito`")."
Bul "Alterar o papel de outros usuarios."
Bul "Adicionar ou remover membros da equipe de um projeto."
Bul "Ativar/desativar as opcoes de `"sem login`" (Solicitacao de Demanda e Ata de Reuniao)."
Bul "Aprovar ou reprovar o Gate 1 na Solicitacao de Demanda (ver secao 4)."
P "Importante: um Admin sempre e considerado `"membro`" de qualquer equipe de projeto automaticamente - nao precisa ser adicionado manualmente para poder editar registros vinculados a um projeto (ver secao 5)."

HR

# ---- 4 ----
H1 "4. Gates de aprovacao"
P "O sistema tem dois pontos de decisao formal (gates) no ciclo de vida de um projeto, ambos aparecendo no fluxo de diretrizes da pagina inicial:"

H2 "4.1 Gate 1 - Triagem"
P "Ocorre logo apos o registro da demanda (D01.4/D01.5). Decide se a demanda entra ou nao no portfolio de projetos."
Bul "Quem decide: Admin (atuando como Gestor Responsavel no sistema)."
Bul "Onde: no status da Solicitacao de Demanda - so o Admin consegue alterar o campo de status; qualquer outro papel ve o campo travado, com um aviso explicando que so o PMO/Admin pode aprovar ou reprovar."

H2 "4.2 Gate 2 - Pactuacao"
P "Ocorre ao final do Planejamento (D02), antes do inicio da Execucao (D03). E o gate mais critico: autoriza formalmente o inicio da execucao do projeto."
Bul "Quem decide: Dono do Negocio, perante a Alta Gestao (registrado institucionalmente; o sistema hoje nao implementa uma trava tecnica separada para este gate especificamente, alem do controle de papeis ja descrito)."

HR

# ---- 5 ----
H1 "5. Restricao por equipe do projeto"
P "Varios formularios exigem vincular o registro a um `"Projeto vinculado`" (um projeto ja Aprovado no Gate 1). Nesses formularios, apenas quem faz parte da equipe daquele projeto especifico - ou um Admin - pode criar ou editar um registro vinculado a ele."

H2 "5.1 Formularios com essa restricao"
P "Aplicada nos 7 formularios que usam o conceito de `"equipe por projeto`":"
Bul "Canvas de Projeto"
Bul "TAP - Termo de Abertura de Projeto"
Bul "Planejamento e Desenvolvimento de Projeto"
Bul "EAP - Estrutura Analitica de Projeto"
Bul "SMP - Solicitacao de Mudanca de Projeto"
Bul "TEP - Termo de Encerramento de Projeto"
Bul "RLA - Registro de Licoes Aprendidas"

H2 "5.2 Como funciona"
P "Ao selecionar um projeto no campo `"Projeto vinculado`":"
Bul "Se a pessoa nao for da equipe daquele projeto (e nao for Admin), aparece um aviso na tela e o botao de enviar/registrar fica desabilitado."
Bul "Mesmo que o botao seja habilitado por algum outro meio, o envio e bloqueado no momento de salvar, com a mensagem `"Voce nao faz parte da equipe deste projeto.`""
Bul "A equipe de cada projeto e definida pelo Admin, na pagina de Administracao, aba Equipes."

H2 "5.3 Formularios sem essa restricao"
P "Os demais formularios nao usam o conceito de equipe por projeto - qualquer usuario autenticado pode criar/editar/excluir registros neles, sujeito apenas as regras de papel e (quando aplicavel) as regras especificas ja descritas nas secoes 2 e 4:"
Bul "Solicitacao de Demanda (tem o Gate 1 e a opcao de sem login, descritos acima)."
Bul "Ata de Reuniao (tem a opcao de sem login, descrita acima)."
Bul "Plano de Comunicacao de Projeto"
Bul "Relatorio de Situacao de Projetos"
Bul "Relatorio de Entregas e Beneficios"

HR

# ---- 6: tabela resumo ----
H1 "6. Resumo por formulario"

$rows = @(
  @("Formulario","Requer login","Restrito por equipe","Aprovacao especial"),
  @("Solicitacao de Demanda","Sim (ou sem login, se ativado)","Nao","Gate 1 (Admin)"),
  @("Canvas de Projeto","Sim","Sim","-"),
  @("TAP","Sim","Sim","-"),
  @("Planejamento e Desenvolvimento","Sim","Sim","-"),
  @("EAP","Sim","Sim","-"),
  @("SMP","Sim","Sim","-"),
  @("Ata de Reuniao","Sim (ou sem login, se ativado)","Nao","-"),
  @("Plano de Comunicacao","Sim","Nao","-"),
  @("TEP","Sim","Sim","-"),
  @("RLA","Sim","Sim","-"),
  @("Relatorio de Situacao","Sim","Nao","-"),
  @("Relatorio de Entregas","Sim","Nao","-"),
  @("Administracao","Sim (so Admin acessa)","-","-"),
  @("Validador de Projetos","Sim","Nao","-")
)

$nRows = $rows.Count
$nCols = 4
$tableRange = $sel.Range
$table = $doc.Tables.Add($tableRange, $nRows, $nCols)
$table.Borders.Enable = $true
$table.Borders.InsideLineStyle = 1
$table.Borders.OutsideLineStyle = 1
$table.Borders.InsideColor = RGB 0xD1 0xD5 0xDB
$table.Borders.OutsideColor = RGB 0xD1 0xD5 0xDB

for ($r=0; $r -lt $nRows; $r++) {
  for ($c=0; $c -lt $nCols; $c++) {
    $cell = $table.Cell($r+1, $c+1)
    $cell.Range.Text = $rows[$r][$c]
    $cell.Range.Font.Size = 9.5
    if ($r -eq 0) {
      $cell.Range.Font.Bold = $true
      $cell.Range.Font.Color = $colWhite
      $cell.Shading.BackgroundPatternColor = $colInk
    } else {
      $cell.Range.Font.Bold = $false
      $cell.Range.Font.Color = $colInk
    }
  }
}
$table.Columns.Item(1).Width = $word.CentimetersToPoints(5.2)
$table.Columns.Item(2).Width = $word.CentimetersToPoints(4.2)
$table.Columns.Item(3).Width = $word.CentimetersToPoints(3.6)
$table.Columns.Item(4).Width = $word.CentimetersToPoints(3.0)

$sel.EndKey(6) | Out-Null  # wdStory
$sel.TypeParagraph()
P "Documento gerado a partir do estado atual do codigo do sistema." 9 $false $true $colMuted "left" 0

if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
$doc.SaveAs2($OutPath, 16)
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output "SAVED: $OutPath"
