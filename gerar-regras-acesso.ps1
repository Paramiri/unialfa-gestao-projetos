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
P "A maioria das 15 paginas do sistema exige login antes de carregar ou salvar qualquer dado. O login pode ser feito de duas formas:"
Bul "Link magico por e-mail - o usuario informa o e-mail e recebe um link de acesso, sem senha."
Bul "Microsoft (SSO) - `"Entrar com Microsoft - UNIALFA`", usando a conta institucional."
P "As paginas abertas sem exigir login sao a pagina inicial (index.html, mapa de diretrizes) e o Validador de Projetos (ferramenta de apoio a decisao) - ambas mostram o conteudo livremente e so exibem a barra `"Conectado como...`" caso ja exista uma sessao ativa. Diferente da Solicitacao de Demanda e da Ata de Reuniao (secao 2), o acesso sem login do Validador nao depende de nenhuma ativacao pelo Admin - e sempre aberto."

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

H2 "2.5 Caso particular: Validador de Projetos"
P "O Validador de Projetos e mais aberto que os dois formularios acima: o quadro de conexoes, o simulador `"e se?`" e o assistente de decisao (8 perguntas, com veredito) funcionam por inteiro sem nenhum login - nao ha selo, nao ha admin para ativar, e nao ha bloqueio de nenhuma acao dentro da propria ferramenta."
Bul "Login so e pedido para uma funcionalidade especifica: vincular a avaliacao a um `"Projeto vinculado`" e salvar o veredito no historico daquele projeto."
Bul "Sem login, essa area do painel mostra um aviso com um botao `"Entrar`" - a pessoa pode logar a qualquer momento sem perder as respostas ja dadas no assistente."
Bul "Depois de logada, a pessoa ve o campo de projeto normalmente, como qualquer outro usuario autenticado."

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
Bul "Acessar a pagina de Administracao e o Painel Executivo (qualquer outro papel ve a mensagem `"Acesso restrito`" em ambas)."
Bul "Alterar o papel de outros usuarios."
Bul "Adicionar ou remover membros da equipe de um projeto."
Bul "Ativar/desativar as opcoes de `"sem login`" (Solicitacao de Demanda e Ata de Reuniao)."
Bul "Aprovar ou reprovar o Gate 1 na Solicitacao de Demanda (ver secao 4)."
Bul "Pactuar ou reabrir o Gate 2 no Relatorio de Entregas e Beneficios (ver secao 4)."
Bul "Pre-cadastrar uma pessoa que ainda nao fez login, informando nome, telefone, e-mail e papel (ver 3.2)."
P "Importante: um Admin sempre e considerado `"membro`" de qualquer equipe de projeto automaticamente - nao precisa ser adicionado manualmente para poder editar registros vinculados a um projeto (ver secao 5)."

H2 "3.2 Pre-cadastro de usuario (antes do primeiro login)"
P "Normalmente uma pessoa so aparece na aba `"Usuarios`" da Administracao depois de fazer login pela primeira vez (o perfil e criado automaticamente, com papel `"Solicitante`"). O pre-cadastro permite ao Admin adiantar esse processo:"
Bul "Na aba Usuarios, o Admin preenche nome, telefone (opcional), e-mail e papel e clica em `"+ Adicionar usuario`"."
Bul "A pessoa aparece na lista com o selo `"Pendente - 1o login`", com nome/telefone/papel ja editaveis pelo Admin mesmo antes de ela logar."
Bul "Quando essa pessoa faz o primeiro login (link magico ou Microsoft), o sistema aplica automaticamente o nome, telefone e papel definidos no pre-cadastro ao perfil recem-criado, e o pre-cadastro pendente e removido."
Bul "Se o Admin nao quiser mais aguardar aquele pre-cadastro, pode remove-lo a qualquer momento pelo botao `"Remover`" - a pessoa continua podendo logar normalmente depois, so que sem os dados pre-preenchidos (entra como `"Solicitante`", papel padrao)."
P "Controle de acesso ao pre-cadastro (Row Level Security no Supabase, tabela `perfis_pendentes`): somente Admin pode criar, editar ou alterar o papel de um pre-cadastro. A propria pessoa so enxerga e pode remover o pre-cadastro que corresponde ao seu proprio e-mail - e exatamente essa permissao restrita que permite o autopreenchimento no momento do primeiro login, sem abrir a tabela para qualquer usuario autenticado."

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
Bul "Quem decide: Dono do Negocio, perante a Alta Gestao."
Bul "Onde: no Relatorio de Entregas e Beneficios (FORALF12), aba Editar dados - campo Status (Pendente de pactuacao / Pactuado). So o Admin consegue alterar esse campo; qualquer outro papel ve o controle travado, com um aviso explicando que so o PMO/Admin pode pactuar."
Bul "Ao marcar `"Pactuado`", o sistema preenche automaticamente a Data de pactuacao (hoje) e o Aprovador (nome de quem esta logado), ambos editaveis pelo Admin."
Bul "Enquanto o status estiver `"Pactuado`", todos os demais campos do relatorio (ficha do programa, indicadores, projetos vinculados) ficam bloqueados para edicao - inclusive para o Admin - ate que o Gate 2 seja reaberto (status voltar para `"Pendente de pactuacao`")."

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
P "O Relatorio de Situacao e o Relatorio de Entregas tem, cada um, um campo opcional `"Projeto vinculado (Gate 1)`" em cada projeto listado - diferente do conceito desta secao, ele nao aplica nenhuma restricao de edicao por equipe. Serve apenas para ligar aquele item ao historico compartilhado do projeto (a mesma trilha de auditoria usada pelos 7 formularios com restricao por equipe), visivel pelo botao `"Ver historico`". Deixar sem selecionar mantem o comportamento anterior: qualquer usuario autenticado continua podendo criar/editar esses registros livremente."

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
  @("Relatorio de Entregas","Sim","Nao","Gate 2 (Admin)"),
  @("Administracao","Sim (so Admin acessa)","-","-"),
  @("Painel Executivo","Sim (so Admin acessa)","-","-"),
  @("Validador de Projetos","Nao (so p/ vincular projeto)","Nao","-")
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
