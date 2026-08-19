param(
  [string]$OutPath = (Join-Path $PSScriptRoot "Manual de Uso - Ferramenta de Gestao de Projetos.docx"),
  [Parameter(Mandatory=$true)][string]$ImgDir
)
$ErrorActionPreference = "Stop"

function RGB($r,$g,$b) { return [int]($r + ($g*256) + ($b*65536)) }
$colRed   = RGB 0xB9 0x1D 0x2E
$colInk   = RGB 0x1A 0x1A 0x1A
$colMuted = RGB 0x6A 0x6A 0x70
$colWhite = RGB 0xFF 0xFF 0xFF
$colGood  = RGB 0x05 0x96 0x69

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()
$sel = $word.Selection

$sec = $doc.Sections.Item(1)
$sec.PageSetup.PageWidth = $word.CentimetersToPoints(21.59)
$sec.PageSetup.PageHeight = $word.CentimetersToPoints(27.94)

$footer = $sec.Footers.Item(1)
$footer.Range.Font.Size = 8.5
$footer.Range.Font.Color = $colMuted
$footer.Range.Text = "UNIALFA - Manual de Uso da Ferramenta de Gestao de Projetos | Pagina "
$footer.Range.Collapse(0) | Out-Null
$footer.Range.Fields.Add($footer.Range, 33) | Out-Null  # wdFieldPage = 33

function P($text, $size=11, $bold=$false, $italic=$false, $color=$colInk, $align="left", $spaceAfter=8) {
  $sel.Style = $doc.Styles.Item(-1)
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = $size
  $sel.Font.Bold = $bold
  $sel.Font.Italic = $italic
  $sel.Font.Color = $color
  $sel.ParagraphFormat.Alignment = if($align -eq "center"){1}else{0}
  $sel.ParagraphFormat.SpaceAfter = $spaceAfter
  $sel.ParagraphFormat.LineSpacing = 14
  $sel.ParagraphFormat.LeftIndent = 0
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Font.Bold = $false
  $sel.Font.Italic = $false
}

function H1($text) {
  $sel.Style = $doc.Styles.Item(-2)
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 17
  $sel.Font.Bold = $true
  $sel.Font.Color = $colInk
  $sel.Font.Italic = $false
  $sel.ParagraphFormat.Alignment = 0
  $sel.ParagraphFormat.SpaceBefore = 4
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 1
  $sel.ParagraphFormat.Borders.Item(3).Color = RGB 0xE4 0xE4 0xE7
  $sel.ParagraphFormat.PageBreakBefore = $true
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 0
  $sel.ParagraphFormat.PageBreakBefore = $false
  $sel.Style = $doc.Styles.Item(-1)
}

function H2($text) {
  $sel.Style = $doc.Styles.Item(-3)
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 13.5
  $sel.Font.Bold = $true
  $sel.Font.Color = $colRed
  $sel.Font.Italic = $false
  $sel.ParagraphFormat.Alignment = 0
  $sel.ParagraphFormat.SpaceBefore = 16
  $sel.ParagraphFormat.SpaceAfter = 6
  $sel.ParagraphFormat.PageBreakBefore = $false
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item(-1)
  $sel.Font.Color = $colInk
}

function H3($text) {
  $sel.Style = $doc.Styles.Item(-1)
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 12
  $sel.Font.Bold = $true
  $sel.Font.Color = $colInk
  $sel.ParagraphFormat.SpaceBefore = 12
  $sel.ParagraphFormat.SpaceAfter = 5
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Font.Bold = $false
}

function Bul($text) {
  $sel.Style = $doc.Styles.Item(-1)
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 11
  $sel.Font.Bold = $false
  $sel.Font.Color = $colInk
  $sel.ParagraphFormat.SpaceAfter = 4
  $sel.ParagraphFormat.LineSpacing = 13
  $sel.ParagraphFormat.LeftIndent = $word.CentimetersToPoints(0.6)
  $sel.Range.ListFormat.ApplyBulletDefault()
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Range.ListFormat.RemoveNumbers()
  $sel.ParagraphFormat.LeftIndent = 0
}

function Nota($text){
  $sel.Style = $doc.Styles.Item(-1)
  $sel.ParagraphFormat.LeftIndent = $word.CentimetersToPoints(0.5)
  $sel.ParagraphFormat.Borders.Item(1).LineStyle = 1
  $sel.ParagraphFormat.Borders.Item(1).Color = $colRed
  $sel.ParagraphFormat.Borders.Item(1).LineWidth = 6
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 10.5
  $sel.Font.Italic = $true
  $sel.Font.Color = $colMuted
  $sel.ParagraphFormat.SpaceBefore = 8
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Font.Italic = $false
  $sel.ParagraphFormat.Borders.Item(1).LineStyle = 0
  $sel.ParagraphFormat.LeftIndent = 0
}

function Exemplo($text){
  $sel.Style = $doc.Styles.Item(-1)
  $sel.ParagraphFormat.LeftIndent = $word.CentimetersToPoints(0.5)
  $sel.ParagraphFormat.Borders.Item(1).LineStyle = 1
  $sel.ParagraphFormat.Borders.Item(1).Color = $colGood
  $sel.ParagraphFormat.Borders.Item(1).LineWidth = 6
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 10.5
  $sel.Font.Color = $colInk
  $sel.ParagraphFormat.SpaceBefore = 8
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.Font.Bold = $true
  $sel.Font.Color = $colGood
  $sel.TypeText("EXEMPLO - ")
  $sel.Font.Bold = $false
  $sel.Font.Color = $colInk
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Borders.Item(1).LineStyle = 0
  $sel.ParagraphFormat.LeftIndent = 0
}

function Img($filename, $caption, $widthIn=5.6){
  $path = Join-Path $ImgDir $filename
  if (-not (Test-Path $path)) { Write-Warning "Imagem nao encontrada: $path"; return }
  $sel.Style = $doc.Styles.Item(-1)
  $sel.ParagraphFormat.Alignment = 1
  $sel.ParagraphFormat.SpaceBefore = 6
  $sel.ParagraphFormat.SpaceAfter = 2
  $shape = $sel.InlineShapes.AddPicture($path)
  $ratio = $shape.Height / $shape.Width
  $shape.Width = $word.InchesToPoints($widthIn)
  $shape.Height = $shape.Width * $ratio
  $shape.Line.Visible = $true
  $shape.Line.ForeColor.RGB = RGB 0xD1 0xD5 0xDB
  $shape.Line.Weight = 0.75
  $sel.TypeParagraph()
  if ($caption) {
    $sel.Font.Name = "Montserrat"
    $sel.Font.Size = 9.5
    $sel.Font.Italic = $true
    $sel.Font.Color = $colMuted
    $sel.ParagraphFormat.SpaceAfter = 14
    $sel.TypeText($caption)
    $sel.TypeParagraph()
    $sel.Font.Italic = $false
  }
  $sel.ParagraphFormat.Alignment = 0
}

function TableSimple($rows, $colWidthsCm){
  $nRows = $rows.Count; $nCols = $rows[0].Count
  $range = $sel.Range
  $table = $doc.Tables.Add($range, $nRows, $nCols)
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
      $cell.Range.Font.Name = "Montserrat"
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
  for ($c=0; $c -lt $nCols; $c++) { $table.Columns.Item($c+1).Width = $word.CentimetersToPoints($colWidthsCm[$c]) }
  $sel.EndKey(6) | Out-Null
  $sel.Style = $doc.Styles.Item(-1)
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.TypeParagraph()
}

# ============================================================
# CAPA
# ============================================================
P "UNIALFA - GERENCIA DE PROJETOS" 10 $true $false $colRed "left" 4
$sel.Style = $doc.Styles.Item(-1)
$sel.Font.Name="Montserrat"; $sel.Font.Size = 28; $sel.Font.Bold = $true; $sel.Font.Color = $colInk
$sel.ParagraphFormat.SpaceAfter = 4
$sel.TypeText("MANUAL DE USO")
$sel.TypeParagraph()
$sel.Font.Bold = $false
P "Ferramenta de Gestao de Projetos - UNIALFA" 15 $false $false $colInk "left" 20
P "Guia passo a passo: do mapa de diretrizes e da Solicitacao de Demanda a geracao dos Relatorios de Situacao e de Entregas e Beneficios - incluindo login, papeis de usuario, gates de aprovacao, restricao por equipe e o Validador de Projetos." 12 $false $true $colMuted "left" 30
P "UNIALFA - Gerencia de Projetos" 11 $false $false $colMuted "left" 2
P "Grupo Jose Alves" 11 $false $false $colMuted "left" 2
P "Versao 2.25 - 19 de agosto de 2026 (substitui a versao 2.24 de 19/08/2026)" 11 $false $false $colMuted "left" 2

$sel.InsertBreak(7) | Out-Null

# ============================================================
# SUMARIO (TOC automatico)
# ============================================================
$sel.Style = $doc.Styles.Item(-1)
$sel.Font.Name="Montserrat"; $sel.Font.Size = 16; $sel.Font.Bold = $true; $sel.Font.Color = $colInk
$sel.ParagraphFormat.SpaceAfter = 10
$sel.TypeText("SUMARIO")
$sel.TypeParagraph()
$sel.Font.Bold = $false
$tocRange = $sel.Range
$toc = $doc.TablesOfContents.Add($tocRange, $true, 1, 2)
$sel.EndKey(6) | Out-Null
$sel.InsertBreak(7) | Out-Null

# ============================================================
# 1. INTRODUCAO
# ============================================================
H1 "1. Introducao"

H2 "1.1 Sobre este manual"
P "Este manual explica, passo a passo, como usar as ferramentas eletronicas que dao suporte a gestao de projetos da UNIALFA. Ele cobre a sequencia completa de uso: do login e do mapa de diretrizes, passando pelo registro de uma nova demanda, ate a geracao dos dois relatorios de fechamento - o Relatorio de Situacao de Projetos (FORALF11) e o Relatorio de Entregas e Beneficios (FORALF12)."
P "Esta e a versao 2.0 do manual. Em relacao a versao 1.0 (15/07/2026), foram adicionadas as secoes sobre login obrigatorio, acesso sem login, papeis de usuario, os dois gates de aprovacao, restricao por equipe do projeto, o Validador de Projetos e a pagina de Administracao - alem de capturas de tela reais de cada ferramenta."
P "Esta e a versao 2.1 do manual. Em relacao a versao 2.0 (28/07/2026), foram adicionadas as secoes sobre o registro vivo de riscos e as sugestoes de riscos recorrentes no TAP (Passo 3), a sinalizacao automatica de risco de atraso no Relatorio de Situacao (secao 4.1) e a prioridade leve (P0/P1/P2) das atividades da EAP (Passo 5)."
P "Esta e a versao 2.2 do manual. Em relacao a versao 2.1 (29/07/2026), foi adicionada a implementacao tecnica do Gate 2 - Pactuacao no Relatorio de Entregas e Beneficios (secao 4.2): status controlado pelo Admin, data e aprovador preenchidos automaticamente, e trava de edicao do relatorio enquanto pactuado."
P "Esta e a versao 2.3 do manual. Em relacao a versao 2.2 (29/07/2026), foi adicionada a Gestao de Entraves e Encaminhamentos no Relatorio de Situacao (secao 4.1): dois cartoes de registro (Descricao, Responsavel, Prazo, Status Aberto/Resolvido/Cancelado), exibidos lado a lado no Painel e incluidos no CSV exportado."
P "Esta e a versao 2.4 do manual. Em relacao a versao 2.3 (29/07/2026), foi adicionada a propagacao automatica da decisao da SMP (Passo 7) para o historico compartilhado do projeto vinculado: toda vez que uma SMP e aprovada ou nao aprovada, o evento fica visivel em `Ver historico` em qualquer formulario ligado aquele projeto, sem alterar o Status de Gate 1."
P "Esta e a versao 2.5 do manual. Em relacao a versao 2.4 (29/07/2026), foi adicionado o campo opcional `Projeto vinculado (Gate 1)` a cada projeto listado no Relatorio de Situacao (secao 4.1) e no Relatorio de Entregas e Beneficios (secao 4.2): a selecao preenche o nome automaticamente e libera o `Ver historico` compartilhado, sem aplicar a restricao por equipe usada nos outros 7 formularios vinculados."
P "Esta e a versao 2.6 do manual. Em relacao a versao 2.5 (29/07/2026), foi adicionado o botao `Importar do TAP` na secao Indicadores do projeto do Relatorio de Entregas e Beneficios (secao 4.2): reaproveita os indicadores de resultado ja preenchidos no TAP do mesmo Projeto vinculado, evitando redigitacao, com protecao contra indicadores duplicados."
P "Esta e a versao 2.7 do manual. Em relacao a versao 2.6 (30/07/2026), foi adicionado o versionamento historico ao Relatorio de Situacao (secao 4.1) e ao Relatorio de Entregas e Beneficios (secao 4.2): cada `Salvar e ver painel` arquiva uma copia completa dos dados, consultavel pelo botao `Historico de versoes` no rodape do Painel, com visualizacao somente leitura de versoes antigas e trava de edicao ate voltar a versao atual."
P "Esta e a versao 2.8 do manual. Em relacao a versao 2.7 (01/08/2026), foi adicionado o cartao `Resultados alcancados` ao Relatorio de Situacao (secao 4.1): registra indicadores, entregas e beneficios de fato alcancados frente ao planejado, atendendo o Reporte de Resultados (D06.5) sem exigir um 13o formulario separado. A tabela de artefatos (secao 8) e o texto sobre cobertura das Diretrizes (secao 1.4) foram atualizados de acordo."
P "Esta e a versao 2.9 do manual. Em relacao a versao 2.8 (01/08/2026), foi adicionada a nova secao 2.8 - Notificacoes por e-mail em transicoes de estado: seis transicoes (Gate 1, Canvas aprovado, TAP aprovado, decisao da SMP, registro de TEP e Gate 2 pactuado) agora disparam automaticamente um e-mail via Resend para quem criou o registro, a equipe do projeto vinculado e os Admins."
P "Esta e a versao 2.10 do manual. Em relacao a versao 2.9 (01/08/2026), foi adicionada a nova secao 6.4 - Painel Executivo: uma 15a pagina, restrita a Admin, que agrega em tempo real os 10 formularios de registro e os 2 relatorios numa unica tela (gates de aprovacao, entraves/encaminhamentos/resultados, contagem por formulario e portfolio de projetos), sem exigir nenhuma mudanca no banco de dados."
P "Esta e a versao 2.11 do manual. Em relacao a versao 2.10 (03/08/2026), foi adicionada a nova secao 2.9 - Lembretes por notificacao push no celular: um segundo canal de aviso, alem do e-mail, ativado por cada usuario individualmente na pagina inicial. Dispara nos mesmos seis momentos de decisao ja descritos na secao 2.8."
P "Esta e a versao 2.12 do manual. Em relacao a versao 2.11 (04/08/2026), foi adicionado um setimo momento de notificacao (secao 2.8): ao registrar uma nova Solicitacao de Demanda, os Admins agora recebem um aviso imediato (e-mail e push) de que ha um Gate 1 - Triagem pendente, em vez de depender de alguem checar a lista manualmente."
P "Esta e a versao 2.13 do manual. Em relacao a versao 2.12 (06/08/2026), foi adicionado o campo obrigatorio Prioridade (Critica/Alta/Media/Baixa) a Solicitacao de Demanda (Passo 1): exibido como selo colorido na lista de demandas cadastradas e no detalhe do registro, ajudando o Admin a priorizar a triagem do Gate 1."
P "Esta e a versao 2.14 do manual. Em relacao a versao 2.13 (06/08/2026), foi adicionado o pre-cadastro de usuario e o campo Telefone a aba Usuarios da Administracao (secao 6.1 e 3.2 das Regras de Acesso): o Admin pode informar nome, telefone, e-mail e papel de uma pessoa antes do primeiro login dela, e os dados sao aplicados automaticamente ao perfil assim que ela loga pela primeira vez."
P "Esta e a versao 2.15 do manual. Em relacao a versao 2.14 (10/08/2026), foi adicionado o botao Remover a aba Usuarios da Administracao (secao 6.1 e 3.3 das Regras de Acesso): o Admin so consegue excluir um usuario que nao tenha nenhum vinculo com projetos ou registros no sistema, checado automaticamente antes de confirmar a remocao."
P "Esta e a versao 2.16 do manual. Em relacao a versao 2.15 (10/08/2026), foi adicionada a restricao de acesso por papel ao Painel Executivo, ao Relatorio de Situacao e ao Relatorio de Entregas (secao 6.3 e 3.4 das Regras de Acesso): o Admin marca, por caixas de selecao, quais papeis podem ver cada uma dessas tres paginas, podendo marcar mais de um papel por pagina."
P "Esta e a versao 2.17 do manual. Em relacao a versao 2.16 (11/08/2026), foi adicionada a importacao de transcricao por IA na Ata de Reuniao (secao 2.10, Passo 6 e 6.3 - substitui o antigo painel `Preencher com Read AI`, que nunca funcionou em producao): o usuario cola a transcricao de uma reuniao e a IA sugere pauta, participantes, resumo, encaminhamentos e entraves para revisao, com interruptor geral e lista de papeis permitidos configuraveis pelo Admin."
P "Esta e a versao 2.18 do manual. Em relacao a versao 2.17 (12/08/2026), a importacao de transcricao por IA (secao 2.10) passou a aceitar tambem arquivo `.txt`, `.docx` ou `.pdf` anexado, alem de colar o texto diretamente - o sistema extrai o texto do arquivo automaticamente para revisao antes de analisar."
P "Esta e a versao 2.19 do manual. Em relacao a versao 2.18 (12/08/2026), foi adicionada a importacao de audio por IA na Ata de Reuniao (secao 2.11): alem de colar/anexar a transcricao em texto, o usuario pode anexar a propria gravacao da reuniao (.mp3, .m4a, .aac, .wav, .ogg...), dividida e transcrita automaticamente em pedacos pelo sistema - com interruptor geral e lista de papeis permitidos proprios, independentes dos da importacao por texto, configuraveis pelo Admin (secao 6.3)."
P "Esta e a versao 2.20 do manual. Em relacao a versao 2.19 (12/08/2026), os antigos botoes `Exportar CSV` e `Imprimir` da Ata de Reuniao (Passo 6), que agiam sobre a lista inteira e nao geravam um documento util, foram substituidos por `Imprimir` e `Exportar Word` no painel de uma ata especifica: ambos geram o documento no formato oficial do FORALF00340 (cabecalho, unidade, pauta, participantes, descricao, saidas e entraves)."
P "Esta e a versao 2.21 do manual. Em relacao a versao 2.20 (13/08/2026), o mesmo ajuste foi estendido aos demais formularios que ainda tinham os antigos botoes `Exportar CSV` e `Imprimir` sem funcionalidade real: Solicitacao de Demanda (Passo 1), Canvas de Projeto (Passo 2), TAP (Passo 3), Planejamento e Desenvolvimento (Passo 4), EAP (Passo 5), SMP (Passo 7), TEP (Passo 8), RLA (Passo 9), Relatorio de Situacao (secao 4.1) e Relatorio de Entregas e Beneficios (secao 4.2). Em todos, `Imprimir` e `Exportar Word` agora geram o documento no formato oficial do respectivo FORALF (ou, no caso da EAP, um layout padrao do sistema, ja que este artefato nao tem FORALF proprio)."
P "Esta e a versao 2.22 do manual. Em relacao a versao 2.21 (13/08/2026), o Plano de Comunicacao de Projeto (Passo 10) ganhou as 4 colunas que faltavam em relacao ao documento oficial FORALF00308 - Quando Comunicar, Com Quem se Comunicar, Como Comunicar e Quem Comunica, alem das ja existentes Tipo de Comunicacao e O que Comunicar - e passou a ter a aba `Editar dados` restrita por papel (padrao PMO/Admin, configuravel em Administracao > Configuracoes, secao 6.3), mantendo a visualizacao e impressao livres para qualquer usuario autenticado."
P "Esta e a versao 2.23 do manual. Em relacao a versao 2.22 (13/08/2026), foi adicionada a nova secao 6.5 - Ambiente de Treino: um ambiente de treinamento/demonstracao totalmente separado da producao, com 7 projetos ficticios cobrindo o ciclo completo, e uma quarta aba na Administracao (visivel so em producao) com um botao que reseta e repopula esse ambiente com um clique, restrito a Admin (secao 3.8 das Regras de Acesso)."
P "Esta e a versao 2.24 do manual. Em relacao a versao 2.23 (14/08/2026), foi adicionada a nova secao 2.12 - Assistente de preenchimento por IA (Canvas e TAP): um botao `Sugerir com IA` que le a Solicitacao de Demanda, o Canvas (no caso do TAP) e as Atas de Reuniao do projeto vinculado para sugerir o preenchimento dos campos ainda vazios, com a mesma logica de interruptor geral e lista de papeis liberados ja usada na importacao de transcricao (secao 2.10)."
P "Esta e a versao 2.25 do manual. Em relacao a versao 2.24 (19/08/2026), foi adicionada a nova secao 2.13 - Importacao de documento por IA (Solicitacao de Demanda): um painel `Importar documento preenchido` que le o formulario oficial FORALF00339 preenchido a mao em Word ou PDF (colado ou anexado) e sugere o preenchimento dos campos do formulario online, com a mesma logica de interruptor geral e lista de papeis liberados das demais importacoes por IA - sem restricao por equipe de projeto, ja que e o primeiro formulario da esteira."
P "O manual nao substitui as Diretrizes para a Gestao de Projetos da UNIALFA (documento institucional que define o framework D01 a D07) nem o documento Regras de Acesso e Permissoes (que detalha cada regra de controle de acesso); ele e o guia operacional de como usar cada ferramenta na pratica."

H2 "1.2 Visao geral da ferramenta"
P "A ferramenta e composta por 15 paginas eletronicas independentes:"
Bul "12 formularios de registro (codigo FORALF), cada um correspondendo a um artefato institucional - da Solicitacao de Demanda ao Relatorio de Entregas e Beneficios."
Bul "1 ferramenta de apoio a decisao, o Validador de Projetos, que nao gera registros de artefato mas ajuda a avaliar projetos."
Bul "1 pagina de Administracao, restrita a Admin, para gerenciar usuarios, equipes e configuracoes do sistema."
Bul "1 Painel Executivo, tambem restrito a Admin, com a visao agregada de todo o sistema (secao 6.4)."
P "Cada formulario de registro tem duas partes: uma tela de preenchimento (`+Novo...`) e uma tela de consulta (`...cadastrados`), onde ficam listados todos os registros ja salvos, com opcao de abrir, editar o status ou excluir cada um."
P "O Mapa de Diretrizes (pagina inicial) apresenta visualmente as 7 diretrizes (D01 a D07) e mostra, para cada etapa do processo, quais ferramentas usar. E o ponto de partida recomendado para quem quer entender o processo antes de preencher os formularios - e a unica pagina, junto com o Validador de Projetos, que nao exige login."
Img "00_home.png" "Mapa de diretrizes - pagina inicial. Nao exige login." 5.8

H2 "1.3 Como acessar e fazer login"
P "Todos os 12 formularios de registro e a pagina de Administracao exigem login antes de carregar ou salvar qualquer dado. Ao abrir qualquer um deles sem sessao ativa, aparece a tela de entrada:"
Img "01_login_gate.png" "Tela de login, exibida ao abrir qualquer formulario sem sessao ativa." 4.6
P "O login pode ser feito de duas formas, sem necessidade de senha:"
Bul "Link magico por e-mail - digite seu e-mail e clique em `Enviar link de acesso`. Um link chega no seu e-mail; clique nele e a pagina de origem atualiza sozinha."
Bul "Microsoft (SSO) - clique em `Entrar com Microsoft - UNIALFA` e autentique com a sua conta institucional."
Nota "No primeiro login, cada pessoa recebe automaticamente o papel Solicitante. Para obter outro papel (ex.: Gerente de Projetos, Admin), peca a um Admin para altera-lo em Administracao > Usuarios (secao 6)."
P "Depois de logado, uma barra preta no topo da pagina mostra `Conectado como [seu e-mail]` e o seu papel atual, com um botao `Sair`."

H2 "1.4 Como a ferramenta se relaciona com as Diretrizes (D01-D07)"
P "Os 12 formularios cobrem os 13 artefatos previstos nas Diretrizes. O Reporte de Resultados (D06.5) nao tem um formulario proprio - e atendido pela secao `Resultados alcancados` do Relatorio de Situacao (FORALF11, secao 4.1)."
$rows = @(
  @("Passo","Formulario","Diretriz / Estrategia"),
  @("1","Solicitacao de Demanda (FORALF00339)","D01.1 - Recebimento e Registro de Demandas"),
  @("-","Gate 1 - Triagem","D01.4/D01.5 - aprovacao do Gestor Responsavel (Admin)"),
  @("2","Canvas de Projeto (FORALF00344)","D01.6 - Elaboracao do Canvas de Projeto"),
  @("3","TAP - Termo de Abertura (FORALF00338)","D02.1 - Desenvolvimento do TAP"),
  @("4","Planejamento e Desenvolvimento (FORALF00325)","D02.2 a D02.10"),
  @("5","EAP - Estrutura Analitica do Projeto","D02.9 - Desenvolvimento da EAP"),
  @("-","Gate 2 - Pactuacao","D06.1/D06.2 - autorizacao da Alta Gestao"),
  @("6","Ata de Reuniao (FORALF00340)","D01 a D07 - transversal, qualquer interacao formal"),
  @("7","SMP - Solicitacao de Mudanca (FORALF00343)","D04.4 - Controle Integrado de Mudancas"),
  @("8","TEP - Termo de Encerramento (FORALF00341)","D05.1.9 - Emissao do Termo de Encerramento"),
  @("9","RLA - Registro de Licoes Aprendidas (FORALF00342)","D05.1.6 - Conducao de Licoes Aprendidas"),
  @("10","Plano de Comunicacao de Projeto (FORALF00308)","D03.5 - Planejamento das Comunicacoes"),
  @("11","Relatorio de Situacao de Projetos (FORALF11)","D04.3 - Elaboracao de Relatorios"),
  @("12","Relatorio de Entregas e Beneficios (FORALF12)","D06.1/D06.2 - Governanca e Tomada de Decisao")
)
TableSimple $rows @(1.5,7.5,7.0)

# ============================================================
# 2. ANTES DE COMECAR
# ============================================================
H1 "2. Antes de comecar"

H2 "2.1 Onde encontrar as ferramentas"
P "O Mapa de Diretrizes (pagina inicial) e o ponto central: o menu lateral lista os 12 formularios, o Validador de Projetos e a Administracao. A tabela no Anexo (secao 8) traz o nome de cada artefato e seu codigo, para referencia rapida."

H2 "2.2 Convencoes usadas neste manual"
Bul "Campos marcados com asterisco (*) nos formularios sao obrigatorios. O sistema nao deixa registrar enquanto algum campo obrigatorio estiver vazio - ele destaca o campo em vermelho e rola a tela ate ele."
Bul "Ao salvar, cada formulario gera um numero de protocolo automatico no formato CODIGO-ANO-SEQUENCIAL (ex.: FORALF00339-2026-001). Guarde esse numero para localizar o registro depois."
Bul "Todo registro tem um status (ex.: Pendente de aprovacao, Aprovado, Reprovado), alteravel a partir da tela de consulta - em alguns formularios, so o Admin pode altera-lo (ver secao 2.4)."
Bul "A tela de consulta de cada formulario tem um campo de busca. Ao abrir um registro especifico, o rodape do painel lateral traz os botoes `Imprimir` e `Exportar Word`, que geram o documento no formato oficial do respectivo FORALF, pronto para impressao ou para abrir no Word."

H2 "2.3 Papeis de usuario e o que cada um pode fazer"
P "Cada pessoa que faz login recebe um papel, usado para liberar ou restringir acoes especificas no sistema. Papeis sao atribuidos e alterados por um Admin, em Administracao > Usuarios."
$rows2 = @(
  @("Papel","Observacao"),
  @("Solicitante","Papel padrao, atribuido automaticamente a todo novo usuario no primeiro login."),
  @("Gerente de Projetos","Sem restricoes adicionais alem das descritas neste manual."),
  @("Gestor Responsavel","Sem restricoes adicionais alem das descritas neste manual."),
  @("Dono do Negocio","Sem restricoes adicionais alem das descritas neste manual."),
  @("Alta Gestao","Sem restricoes adicionais alem das descritas neste manual."),
  @("Admin (PMO/Admin)","Papel de administracao do sistema - unico que pode executar as acoes da lista abaixo.")
)
TableSimple $rows2 @(5.0,11.0)
P "Acoes exclusivas de Admin:"
Bul "Acessar a pagina de Administracao (qualquer outro papel ve `Acesso restrito`)."
Bul "Alterar o papel de outros usuarios."
Bul "Adicionar ou remover membros da equipe de um projeto."
Bul "Ativar ou desativar as opcoes de `sem login` da Solicitacao de Demanda e da Ata de Reuniao."
Bul "Aprovar ou reprovar o Gate 1 na Solicitacao de Demanda."
Nota "Um Admin e sempre considerado membro de qualquer equipe de projeto automaticamente - nao precisa ser adicionado manualmente para editar registros vinculados a um projeto (ver secao 2.5)."

H2 "2.4 Os dois gates de aprovacao"
P "O sistema tem dois pontos de decisao formal no ciclo de vida de um projeto:"
Bul "Gate 1 - Triagem: ocorre logo apos o registro da demanda. So o Admin consegue alterar o status da Solicitacao de Demanda para Aprovada ou Reprovada; qualquer outro papel ve o campo travado."
Bul "Gate 2 - Pactuacao: ocorre ao final do planejamento, quando o Relatorio de Entregas e Beneficios e apresentado e pactuado com o Dono do Negocio perante a Alta Gestao (ver secao 4.2)."
Img "14_f01_gate1.png" "Detalhe de uma Solicitacao de Demanda - o campo Status, no rodape, so pode ser alterado pelo Admin (Gate 1)." 5.4

H2 "2.5 Restricao por equipe do projeto"
P "Sete dos doze formularios exigem vincular o registro a um `Projeto vinculado` (um projeto ja Aprovado no Gate 1): Canvas, TAP, Planejamento e Desenvolvimento, EAP, SMP, TEP e RLA. Nesses formularios, apenas quem faz parte da equipe daquele projeto - ou um Admin - pode criar ou editar um registro vinculado a ele."
Img "17_f03_novo.png" "TAP com um projeto vinculado selecionado: unidade, gerente e status sao preenchidos automaticamente a partir do projeto." 5.6
Bul "Se a pessoa nao for da equipe daquele projeto (e nao for Admin), aparece um aviso na tela e o botao de enviar/registrar fica desabilitado."
Bul "A equipe de cada projeto e definida por um Admin, em Administracao > Equipes (secao 6)."
Nota "Os outros 5 formularios (Solicitacao de Demanda, Ata de Reuniao, Plano de Comunicacao, Relatorio de Situacao e Relatorio de Entregas) nao usam esse conceito de forma obrigatoria - qualquer usuario autenticado pode criar/editar registros neles, sem restricao por equipe. O Relatorio de Situacao e o Relatorio de Entregas ganharam, no entanto, um campo opcional `Projeto vinculado (Gate 1)` em cada projeto listado (secoes 4.1 e 4.2) - ele so serve para ligar aquele projeto ao historico compartilhado, sem aplicar a trava de equipe."

H2 "2.6 Acesso sem login - Solicitacao de Demanda e Ata de Reuniao"
P "Dois formularios podem, opcionalmente, ser abertos por um Admin para aceitar registros sem login: Solicitacao de Demanda e Ata de Reuniao. Cada um tem um interruptor independente em Administracao > Configuracoes (secao 6) - podem estar ligados, desligados ou so um dos dois."
Img "02_sem_login_gate.png" "Quando a opcao esta ativa, a tela de login ganha um terceiro caminho: Continuar sem login." 4.6
Img "03_sem_login_dados.png" "A pessoa sem login informa apenas nome completo e e-mail para registrar." 4.6
Bul "Quem usa esse caminho so consegue criar um registro novo - nunca editar, excluir ou alterar status, mesmo em um registro que ela mesma criou."
Bul "O registro criado sem login recebe um selo `Sem login` na listagem, junto com o nome e e-mail informados."

H2 "2.7 Onde os dados ficam armazenados"
P "Todos os 12 formularios de registro gravam os dados em um banco de dados real (PostgreSQL, hospedado no Supabase), associados a conta autenticada de quem salvou o registro (ou ao nome/e-mail informado, no caso de registro sem login). Nao ha mais uma chave unica compartilhada por todos: cada sessao usa o token da propria conta, e o acesso as ferramentas segue as regras de login, papel e equipe descritas neste capitulo."
P "Cada formulario ainda grava seus registros de forma independente - nao existe herenca automatica de informacoes entre eles. Ao preencher o TAP logo depois do Canvas, por exemplo, e preciso informar novamente o nome do projeto, a unidade e o gerente (ou selecionar o mesmo `Projeto vinculado`, quando o formulario tiver esse campo). Preencha os dados de identificacao da mesma forma em todos os formularios de um mesmo projeto, para manter a rastreabilidade entre eles."

H2 "2.8 Notificacoes por e-mail em transicoes de estado"
P "Sete momentos do sistema disparam automaticamente um e-mail de notificacao, enviado via Resend por uma Supabase Edge Function (a chave da API fica guardada no servidor, nunca no navegador):"
Bul "Registro de nova Solicitacao de Demanda: assim que a demanda e salva, antes de qualquer decisao - avisa os Admins que ha um Gate 1 - Triagem pendente de analise."
Bul "Gate 1 - Solicitacao de Demanda: quando o Admin muda o status para `Aprovada` ou `Reprovada`."
Bul "Canvas de Projeto: quando o status muda para `Aprovado`."
Bul "TAP: quando o status muda para `Aprovado`."
Bul "SMP: quando a decisao e finalizada como `Aprovada` ou `Nao aprovada`."
Bul "TEP: ao registrar um novo Termo de Encerramento."
Bul "Gate 2 - Relatorio de Entregas: quando o status muda para `Pactuado`."
P "Nos seis momentos de decisao (todos exceto o registro de nova demanda), os destinatarios sao: quem criou o registro (e-mail capturado no login, ou informado no registro sem login), todos os e-mails cadastrados na equipe do Projeto vinculado (Administracao > Equipes) e todos os usuarios com papel Admin. Ja o aviso de registro de nova demanda vai somente para os Admins, ja que ainda nao ha Gate 1 aprovado nem equipe designada nesse momento. Se o envio falhar por qualquer motivo (ex.: instabilidade do provedor de e-mail), o registro e salvo normalmente mesmo assim - a notificacao nunca bloqueia o fluxo de trabalho."

H2 "2.9 Lembretes por notificacao push no celular"
P "Alem do e-mail (secao 2.8), o sistema oferece um segundo canal de aviso: uma notificacao push, que aparece no celular ou no computador mesmo com o navegador fechado, igual a uma notificacao de aplicativo. E opcional e ativada individualmente por cada usuario - nao existe um interruptor geral do Admin para isso."
Bul "Onde ativar: na pagina inicial (mapa de diretrizes), logado, aparece o botao `Ativar lembretes no celular` na barra preta do topo."
Bul "Ao tocar no botao, o navegador pede permissao para mostrar notificacoes - aceite para concluir a ativacao. O botao muda para `Lembretes ativos`."
Bul "Para desativar, basta tocar novamente no botao ja ativo."
Bul "Dispara nos mesmos sete momentos da secao 2.8 (registro de nova demanda, Gate 1, Canvas aprovado, TAP aprovado, decisao da SMP, registro de TEP e Gate 2 pactuado), para os mesmos destinatarios - mas so quem tiver ativado o push naquele navegador/celular recebe a notificacao."
Nota "No iPhone/iPad, a Apple so libera notificacao push para sites adicionados a Tela de Inicio. Antes de tocar em `Ativar lembretes no celular`, va em Compartilhar > Adicionar a Tela de Inicio no Safari, abra o sistema pelo icone criado, e so entao ative - num navegador comum (aba do Safari) o botao explica essa limitacao em vez de ativar direto. Em Android e computador nao ha essa exigencia."

H2 "2.10 Importacao de transcricao por IA (Ata de Reuniao)"
P "A Ata de Reuniao pode, opcionalmente, preencher pauta, data/horario, participantes, resumo, encaminhamentos e entraves automaticamente a partir da transcricao de uma reuniao - o texto e analisado por IA (Claude, via uma Supabase Edge Function; a chave da API fica guardada no servidor, nunca no navegador) e os campos sugeridos ficam disponiveis para revisao antes de salvar. O usuario nunca perde o controle: nada e enviado ao Supabase automaticamente, so o preenchimento do formulario."
P "O texto pode ser fornecido de duas formas: colando diretamente na caixa de texto, ou anexando um arquivo `.txt`, `.docx` (Word) ou `.pdf` pelo botao `Anexar arquivo` - o sistema extrai o texto do arquivo e preenche a caixa automaticamente para revisao antes de clicar em `Analisar e preencher`. Limite de 15 MB por arquivo."
Nota "PDFs escaneados (so imagem, sem texto real por tras) nao sao suportados - a extracao depende do PDF ter texto selecionavel. Nesse caso, copie e cole o texto manualmente."
P "Diferente do envio de e-mail e do push (secoes 2.8 e 2.9), essa funcionalidade tem um interruptor geral do Admin, desligado por padrao - cada analise tem um custo real de API. Em Administracao > Configuracoes (secao 6.3), o Admin:"
Bul "Liga ou desliga a importacao para o sistema inteiro - desligada, ninguem tem acesso ao painel, nem o proprio Admin."
Bul "Com a importacao ligada, marca quais papeis podem usa-la (pode marcar mais de um) - PMO/Admin sempre tem acesso quando a importacao esta ligada, mesmo sem estar marcado na lista."
Nota "Quem nao tem permissao simplesmente nao ve o painel `Importar transcricao da reuniao` na Ata - nao ha mensagem de erro, o painel so aparece para quem pode usa-lo."

H2 "2.11 Importacao de audio por IA (Ata de Reuniao)"
P "Alem de colar/anexar a transcricao em texto, a Ata de Reuniao pode transcrever a propria gravacao da reuniao: no painel `Importar transcricao da reuniao`, o botao `Anexar audio` aceita `.mp3`, `.m4a`, `.aac`, `.wav`, `.ogg` e outros formatos comuns. O sistema divide o audio automaticamente em pedacos de cerca de 8 minutos direto no navegador (por causa de um limite tecnico da API de transcricao) e transcreve cada pedaco por IA (OpenAI, via uma segunda Supabase Edge Function; a chave fica guardada no servidor, igual a da Claude). O texto de cada pedaco vai sendo concatenado na caixa de transcricao, junto com uma lista de progresso (`Parte 1 de N`, `Parte 2 de N`...) - reunioes longas podem levar alguns minutos para processar."
Bul "Se algum pedaco falhar (ex.: instabilidade de rede), aparece um botao `Tentar novamente` especifico para aquele pedaco - os pedacos ja transcritos com sucesso nao precisam ser refeitos."
Bul "Depois de transcrito, o texto fica disponivel na mesma caixa da importacao por texto, para revisar e so entao clicar em `Analisar e preencher` normalmente."
P "A importacao de audio tem uma regra de habilitacao SEPARADA da importacao de texto (secao 2.10) - dois interruptores independentes em Administracao > Configuracoes (secao 6.3), cada um desligado por padrao. Na pratica, o botao `Anexar audio` so aparece para quem tem as duas permissoes ativas ao mesmo tempo (importar transcricao E importar audio), ja que o audio so serve para gerar o texto que depois passa pela mesma analise."
Nota "PDFs e audio muito longos consomem mais tempo de processamento; nao ha limite de duracao da reuniao em si, mas o navegador precisa decodificar o arquivo inteiro antes de dividir em pedacos - arquivos muito grandes (varias horas) podem demorar mais para começar a transcrever."

H2 "2.12 Assistente de preenchimento por IA (Canvas e TAP)"
P "O Canvas de Projeto (Passo 2) e o TAP (Passo 3) tem um botao `Sugerir com IA`, que aparece assim que um Projeto vinculado e selecionado. Diferente da importacao de transcricao (secao 2.10), que le um texto colado pelo usuario, este assistente le os proprios documentos ja registrados do projeto - a Solicitacao de Demanda e, no caso do TAP, tambem o Canvas - alem de todas as Atas de Reuniao vinculadas a ele, e usa IA (Claude, pela mesma Supabase Edge Function em espirito da secao 2.10) para sugerir o preenchimento dos campos ainda vazios."
Bul "So preenche o que estiver vazio: um campo ja digitado pelo usuario nunca e sobrescrito pela sugestao."
Bul "Cada campo preenchido pela IA fica com um selo `IA` e destaque visual ate ser editado - assim fica claro, campo a campo, o que veio de sugestao e o que foi escrito pela pessoa."
Bul "Uma linha no rodape do bloco lista os documentos usados como base (ex.: `Sugestao baseada em: Solicitacao de Demanda FORALF00339-2026-001, Canvas de Projeto FORALF00344-2026-002`)."
Bul "No TAP, alem dos campos de texto, a sugestao tambem propoe linhas iniciais para as tabelas de riscos, cronograma de entregas macro, custos e partes interessadas - sempre que a tabela ainda estiver vazia."
P "Mesma logica de custo e permissao da importacao de transcricao (secao 2.10): interruptor geral do Admin, desligado por padrao, e lista de papeis liberados em Administracao > Configuracoes (secao 6.3) - PMO/Admin sempre tem acesso quando o assistente esta ligado. Alem disso, so quem faz parte da equipe do projeto selecionado (secao 2.5) pode usar o botao."
Nota "Como qualquer sugestao de IA, o resultado pode estar incompleto ou impreciso quando os documentos anteriores tambem estiverem - revise sempre antes de clicar em `Registrar`."

H2 "2.13 Importacao de documento por IA (Solicitacao de Demanda)"
P "A Solicitacao de Demanda (Passo 1) tem um painel `Importar documento preenchido`, no topo do formulario, para quem ja preencheu o formulario oficial FORALF00339 em Word ou PDF fora do sistema e agora precisa transcrever esses dados para a tela. O texto pode ser colado diretamente, ou o arquivo (`.txt`, `.docx` ou `.pdf`) anexado pelo botao `Anexar arquivo` - o sistema extrai o texto automaticamente para revisao antes de clicar em `Analisar e preencher`. Limite de 15 MB por arquivo."
Bul "A IA identifica nome do projeto, solicitante, departamento, justificativa, objetivo, escopo, prazo, orcamento, partes interessadas e anexos, e preenche so os campos ainda vazios do formulario - exatamente como as demais importacoes por IA (secoes 2.10 e 2.12)."
Bul "Departamento e Prioridade so sao preenchidos se o valor identificado pela IA corresponder exatamente a uma das opcoes da lista suspensa - caso contrario, o campo fica em branco para preenchimento manual."
Nota "PDFs escaneados (so imagem, sem texto real por tras) nao sao suportados - a extracao depende do PDF ter texto selecionavel. Nesse caso, copie e cole o texto manualmente."
P "Mesma logica de custo e permissao das demais importacoes por IA: interruptor geral do Admin, desligado por padrao, e lista de papeis liberados em Administracao > Configuracoes (secao 6.3). Diferente das secoes 2.10 e 2.12, nao ha restricao por equipe de projeto aqui - a Solicitacao de Demanda e o primeiro formulario, preenchida antes de existir um projeto aprovado - mas o acesso sem login (secao 2.6) nunca tem acesso a este painel, mesmo com a importacao habilitada."

# ============================================================
H1 "3. Passo a passo do fluxo de um projeto"
P "Esta secao percorre os 12 formularios de registro na ordem em que normalmente sao usados ao longo da vida de um projeto. Os Passos 1 a 5 e 8 a 9 seguem a esteira sequencial do projeto; os Passos 6, 7 e 10 sao de uso recorrente ou condicional; os relatorios finais (Passos 11 e 12) sao detalhados na secao 4."

H2 "Passo 1 - Solicitacao de Demanda (FORALF00339)"
P "Quando usar: no inicio de tudo, para registrar formalmente uma ideia, necessidade ou problema institucional antes de qualquer outra acao. Pode ser preenchida com login normal ou, se a opcao estiver ativa, sem login (ver secao 2.6)."
$r1 = @(
  @("Campo","Obrig.","Descricao"),
  @("Nome do projeto","Sim","Nome que vai identificar a iniciativa em todos os registros futuros"),
  @("Solicitante","Sim","Nome de quem esta pedindo o projeto"),
  @("Departamento / unidade","Sim","UNIALFA, FADISP, Colegio Alfa, TLA, Controladoria, TI ou Outro"),
  @("Justificativa / necessidade","Sim","Por que o projeto e necessario"),
  @("Objetivo / resultado esperado","Sim","O que deve ser entregue ao final"),
  @("Escopo","Sim","O que esta incluido e o que nao esta"),
  @("Prazo desejado","Sim","Data limite desejada pelo solicitante"),
  @("Prioridade","Sim","Critica, Alta, Media ou Baixa - urgencia da demanda, ajuda o Admin a priorizar a triagem do Gate 1"),
  @("Orcamento estimado","Nao","Valor aproximado, se ja houver"),
  @("Partes interessadas","Sim","Areas, equipes ou pessoas impactadas"),
  @("Anexos","Nao","Nomes de documentos ou links de apoio")
)
TableSimple $r1 @(5.0,1.8,9.2)
Img "12_f01_novo.png" "Tela de novo registro da Solicitacao de Demanda." 5.6
Img "13_f01_lista.png" "Demandas cadastradas, com protocolo, projeto, prazo e status." 5.6
Exemplo "`Automatizacao de servicos` - solicitante Hudson Lucas Aleixo, unidade Relacionamento, justificativa: reduzir o tempo de espera dos alunos e desafogar o atendimento presencial/manual da secretaria, oferecendo disponibilidade 24/7 para solicitacoes basicas."
P "Como salvar: clique em `Registrar solicitacao`. O sistema gera o protocolo e leva voce para `Demandas cadastradas`. Clique em qualquer linha da tabela para abrir o registro e revisar os dados. Ao salvar, todos os Admins recebem um aviso automatico (e-mail e, se ativado, push) de que ha um Gate 1 pendente de analise (secao 2.8)."
Nota "E neste ponto que ocorre o Gate 1 - Triagem: so um Admin pode mudar o status para Aprovada ou Reprovada (secao 2.4). So avance para o Passo 2 depois que o status estiver Aprovada. Essa mudanca de status dispara um e-mail de notificacao (secao 2.8)."
Bul "No rodape do painel de detalhes de uma demanda, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do FORALF00339, pronto para impressao ou para abrir no Word."
Bul "Assistente de preenchimento por IA: o painel `Importar documento preenchido`, no topo do formulario, extrai os dados de uma copia do formulario oficial ja preenchida em Word ou PDF fora do sistema (secao 2.13)."

H2 "Passo 2 - Canvas de Projeto (FORALF00344)"
P "Quando usar: assim que a demanda for aprovada. O Canvas e o primeiro documento estruturado do projeto - reune, em uma unica tela, a motivacao, o produto, os parceiros, as entregas, os riscos e os custos. Exige selecionar um `Projeto vinculado` ja Aprovado no Gate 1, e so membros da equipe daquele projeto (ou Admin) podem registrar (secao 2.5)."
$r2 = @(
  @("Campo","Obrig.","Descricao"),
  @("Nome / Unidade / Gerente / Data","Sim","Identificacao do projeto"),
  @("Justificativas","Sim","Por que estamos aqui - dores e problemas nao atendidos"),
  @("Objetivo","Sim","Objetivo SMART do projeto"),
  @("Beneficios","Nao","O que a organizacao ganha depois da implantacao"),
  @("Produto","Sim","O que sera construido ou entregue"),
  @("Parceiros","Nao","Pessoas ou entidades que apoiam, sem se subordinar ao gerente"),
  @("Grupo de entregas","Sim","Componentes do produto, que depois viram a EAP"),
  @("Restricoes","Nao","Limitacoes que reduzem a liberdade de decisao da equipe"),
  @("Riscos","Sim","Eventos incertos que podem afetar os objetivos"),
  @("Custos","Sim","Estimativa de gastos por grupo de entrega")
)
TableSimple $r2 @(5.0,1.8,9.2)
Img "15_f02_novo.png" "Tela de novo Canvas, com o seletor de Projeto vinculado no topo." 5.6
Img "16_f02_lista.png" "Canvas cadastrados." 5.6
P "Como salvar: clique em `Registrar canvas`. O status inicial e Pendente de aprovacao; altere para Aprovado ao validar o Canvas com o Dono do Negocio e o Gerente de Projetos. Ao marcar Aprovado, um e-mail de notificacao e disparado automaticamente (secao 2.8)."
Bul "No rodape do painel de detalhes de um canvas, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do FORALF00344, pronto para impressao ou para abrir no Word."
Bul "Assistente de preenchimento por IA: apos selecionar o Projeto vinculado, o botao `Sugerir com IA` le a Solicitacao de Demanda e as Atas de Reuniao do projeto e sugere o preenchimento dos campos ainda vazios (secao 2.12)."

H2 "Passo 3 - TAP, Termo de Abertura de Projeto (FORALF00338)"
P "Quando usar: depois do Canvas aprovado, para autorizar formalmente a existencia do projeto. Tambem exige `Projeto vinculado` e segue a restricao por equipe (secao 2.5) - a figura da secao 2.5 mostra este formulario com um projeto ja selecionado."
$r3 = @(
  @("Campo","Obrig.","Descricao"),
  @("Identificacao (nome, unidade, gerente, data)","Sim","Cabecalho do TAP"),
  @("Alinhamento estrategico / Programa vinculado","Nao","Relacao com outras acoes da Holding, se houver"),
  @("Justificativa","Sim","Problema ou oportunidade que motiva o projeto"),
  @("Objetivos","Sim","O que o projeto entrega"),
  @("Publico-alvo / Beneficios / Exclusoes","Nao","Detalhamento do escopo - inclui o que fica de fora"),
  @("Premissas / Restricoes / Criterios","Nao","Condicionantes do projeto"),
  @("Registro de riscos (tabela)","Nao","Registro vivo: risco, status, responsavel e data da ultima revisao"),
  @("Cronograma de entregas macro (tabela)","Nao","Marco, responsavel, inicio e termino previstos, custo"),
  @("Custo total estimado (tabela)","Nao","Soma automatica a medida que as linhas sao preenchidas"),
  @("Partes interessadas e Equipe (tabelas)","Nao","Nome e unidade de cada pessoa"),
  @("Indicadores de resultado (tabela)","Nao","Valor inicial e valor final esperado, com datas")
)
TableSimple $r3 @(5.6,1.6,8.8)
P "Registro de riscos (secao 9 do formulario): diferente dos demais campos do TAP, o registro de riscos e vivo - pode ser reaberto e atualizado a qualquer momento do projeto pelo botao `Editar dados`, sem se limitar ao preenchimento inicial. Cada linha da tabela tem Risco, Status (Aberto, Monitorando, Mitigado, Materializado ou Encerrado, exibido como selo colorido), Responsavel e Ultima revisao (data)."
P "Sugestoes de riscos recorrentes: ao selecionar um Projeto vinculado, o formulario analisa automaticamente o campo Entraves de todas as Atas de Reuniao ja registradas para aquele projeto e destaca, logo abaixo da tabela de riscos, qualquer entrave que se repita em 2 ou mais atas diferentes - com a contagem de atas e um botao `+ Adicionar ao registro` que insere a sugestao direto na tabela (status Aberto, revisao com a data de hoje). E uma regra automatica de comparacao de texto, sem chamar nenhuma IA - ela so aproveita entraves ja registrados nas atas, inclusive os que tiverem sido preenchidos via `Importar transcricao da reuniao` (Passo 6, secao 2.10)."
Img "37_f03_riscos.png" "Registro de riscos preenchido, com selos de status coloridos e a caixa de sugestoes de riscos recorrentes vindas das atas do projeto." 5.6
Img "18_f03_lista.png" "TAPs cadastrados." 5.6
P "Como salvar: clique em `Registrar TAP`. Um projeto sem TAP aprovado nao deve avancar para a execucao. Ao marcar Aprovado, um e-mail de notificacao e disparado automaticamente (secao 2.8)."
Bul "No rodape do painel de detalhes de um TAP, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do FORALF00338 (incluindo o registro de riscos e as tabelas de cronograma, custos, partes interessadas, equipe e indicadores), pronto para impressao ou para abrir no Word."
Bul "Assistente de preenchimento por IA: apos selecionar o Projeto vinculado, o botao `Sugerir com IA` le a Solicitacao de Demanda, o Canvas e as Atas de Reuniao do projeto e sugere o preenchimento dos campos de texto ainda vazios, alem de rascunhos iniciais para as tabelas de riscos, cronograma, custos e partes interessadas (secao 2.12) - diferente da sugestao de riscos recorrentes acima, que e uma regra automatica sem IA."

H2 "Passo 4 - Planejamento e Desenvolvimento de Projeto (FORALF00325)"
P "Quando usar: logo apos o TAP, para detalhar o projeto em profundidade. E o dossie mais extenso da ferramenta, organizado em 8 abas internas, navegaveis pela barra de estagios no topo."
$r4 = @(
  @("Aba","Conteudo"),
  @("1. Pre-projeto","Produtos impactados, escopo, contexto, objetivos, entregaveis, premissas, restricoes, partes interessadas, macro fases"),
  @("2. Viabilidade","Situacao atual, proposta de mudanca, beneficios, alternativas, impactos, plano de implementacao, conclusao"),
  @("3. Cronograma","Tabela de marcos: nome, responsavel, duracao, data de entrega"),
  @("4. Entradas","Tabela de requisitos de entrada, com 7 categorias pre-preenchidas, para completar com descricao, fonte, responsavel e status"),
  @("5. Saidas","Tabela de entregaveis ligados a um requisito de entrada, com criterios de aceitacao"),
  @("6. Verificacao","Tabela de verificacao de conformidade de cada saida"),
  @("7. Validacao","Tabela de validacao do uso pretendido, com metodo e resultado"),
  @("8. Analise critica","Tabela de pontos de controle por etapa do projeto")
)
TableSimple $r4 @(3.0,13.0)
Img "19_f04_novo.png" "Capa de identificacao do dossie, com a barra das 8 abas no rodape visivel." 5.6
Img "20_f04_lista.png" "Projetos cadastrados (estado vazio, antes do primeiro dossie ser salvo)." 5.6
P "Como salvar: preencha a capa (nome do projeto, unidade e gestor sao obrigatorios) e navegue pelas 8 abas. Clique em `Registrar dossie`. E este documento que alimenta, junto com o Canvas, a elaboracao do Relatorio de Entregas e Beneficios (Passo 12)."
Bul "No rodape do painel de detalhes de um dossie, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do FORALF00325, com as 8 abas em sequencia, pronto para impressao ou para abrir no Word."

H2 "Passo 5 - EAP, Estrutura Analitica do Projeto"
P "Quando usar: em paralelo ao planejamento, para decompor visualmente o escopo em pacotes de trabalho."
Img "21_f05_novo.png" "Construtor hierarquico de 3 niveis: Pacote de trabalho, Entrega e Atividade." 5.6
Img "22_f05_lista.png" "EAPs cadastradas (estado vazio)." 5.6
Bul "Informe o nome do projeto (obrigatorio), a unidade e o gerente."
Bul "Use o construtor hierarquico de 3 niveis: `+ Adicionar pacote de trabalho`, depois `+ Adicionar entrega` dentro do pacote, depois `+ Atividade` dentro da entrega."
Bul "Clique em `Visualizar arvore` a qualquer momento para conferir o diagrama antes de salvar."
Bul "Cada Atividade (nivel 3) pode receber, de forma opcional, uma prioridade: `P0 - Critica`, `P1 - Alta` ou `P2 - Normal` (ou nenhuma), exibida como selo colorido no construtor, na `Visualizar arvore` e no painel de detalhes. Serve para repriorizar o trabalho no dia a dia sem precisar abrir uma SMP - Solicitacao de Mudanca de Projeto (Passo 7). EAPs que ja existiam sem prioridade definida continuam funcionando normalmente."
Img "39_f05_prioridade.png" "Atividades de nivel 3 com os tres niveis de prioridade (P0, P1 e P2), exibidos como selo colorido ao lado de cada atividade." 5.6
P "Clique em `Registrar EAP` para gerar o protocolo."
Bul "No rodape do painel de detalhes de uma EAP, os botoes `Imprimir` e `Exportar Word` geram um documento com a arvore hierarquica completa (pacotes, entregas e atividades, com a prioridade de cada uma), pronto para impressao ou para abrir no Word. Este artefato nao tem um FORALF oficial - o layout segue o padrao visual dos demais documentos gerados pelo sistema."

H2 "Passo 6 - Ata de Reuniao (FORALF00340) - uso recorrente"
P "Quando usar: a qualquer momento do projeto, para registrar formalmente qualquer reuniao - nao faz parte da esteira sequencial, fica sempre disponivel. Pode ser preenchida com login normal ou, se a opcao estiver ativa, sem login (ver secao 2.6)."
Img "25_f07_novo.png" "Tela de nova ata, com o painel `Importar transcricao da reuniao` para preenchimento automatico por IA." 5.6
Img "26_f07_lista.png" "Atas cadastradas (estado vazio)." 5.6
Bul "Selecione a(s) unidade(s) envolvidas, preencha Pauta e Projeto (obrigatorios), participantes, resumo, encaminhamentos e entraves."
Bul "Opcional, quando habilitado pelo Admin: em `Importar transcricao da reuniao`, cole o texto, anexe um arquivo `.txt`/`.docx`/`.pdf`, ou anexe o audio da gravacao (`.mp3`/`.m4a`/`.aac`/`.wav`/`.ogg`) para transcricao automatica (secao 2.11), depois clique em `Analisar e preencher` - a IA sugere pauta, data/horario (se mencionados), participantes, resumo, encaminhamentos e entraves. Revise sempre os dados sugeridos antes de registrar (secao 2.10)."
P "Clique em `Registrar ata` para gerar o protocolo."
Bul "Na lista `Atas cadastradas`, clique numa ata para abri-la e usar os botoes `Imprimir` e `Exportar Word`, no rodape do painel lateral: ambos geram o documento no formato oficial do FORALF00340 (cabecalho, unidade, pauta, participantes, descricao, saidas e entraves), pronto para impressao ou para abrir no Word."

H2 "Passo 7 - SMP, Solicitacao de Mudanca de Projeto (FORALF00343) - uso condicional"
P "Quando usar: sempre que for necessario alterar escopo, cronograma, custo ou qualidade de um projeto ja em andamento. Nenhuma mudanca deve ser feita sem passar por este formulario."
Img "23_f06_novo.png" "Tela de nova SMP." 5.6
Img "24_f06_lista.png" "SMPs cadastradas (estado vazio)." 5.6
Bul "Identifique o projeto e a mudanca (titulo e solicitante sao obrigatorios); descreva a mudanca, os beneficios e o impacto de nao implementa-la."
Bul "Preencha a analise de impactos nas 8 dimensoes: objetivo, cronograma, escopo, custo, alinhamento estrategico, qualidade, riscos e outros impactos."
Bul "Marque a decisao: Aprovada, Nao aprovada ou Pendente de avaliacao, com justificativa."
P "Clique em `Registrar SMP` - o status do registro acompanha automaticamente a decisao marcada."
Bul "No rodape do painel de detalhes de uma SMP, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do FORALF00343 (incluindo a analise das 8 dimensoes de impacto), pronto para impressao ou para abrir no Word."
Nota "Propagacao para o projeto vinculado: sempre que a decisao da SMP e finalizada como `Aprovada` ou `Nao aprovada` (seja ao registrar/editar a SMP ou ao trocar o Status na tela de detalhes), o sistema grava automaticamente um evento no historico compartilhado do projeto vinculado (o mesmo historico acessivel pelo botao `Ver historico` em qualquer formulario com `Projeto vinculado` - secao 2.5). Isso torna as decisoes de mudanca visiveis para quem estiver no Canvas, TAP, Planejamento, EAP, TEP ou RLA daquele projeto, sem precisar abrir a SMP. Essa propagacao apenas registra o evento no historico - ela nao altera o Status do projeto usado no Gate 1 (Aprovado/Reprovado). A mesma decisao (Aprovada ou Nao aprovada) tambem dispara um e-mail de notificacao (secao 2.8)."

H2 "Passo 8 - TEP, Termo de Encerramento de Projeto (FORALF00341)"
P "Quando usar: ao encerrar o projeto, seja por conclusao, paralisacao ou cancelamento."
Img "28_f09_novo.png" "Tela de novo TEP." 5.6
Img "29_f09_lista.png" "TEPs cadastrados (estado vazio)." 5.6
Bul "Preencha a identificacao e o programa vinculado, se houver; selecione o tipo de encerramento (Concluido, Paralisado ou Cancelado)."
Bul "Se Paralisado ou Cancelado, o campo Justificativa aparece automaticamente."
Bul "Registre entregas de resultados, atividades encerradas, o link da pasta do projeto e a analise de efetividade."
P "Clique em `Registrar TEP` para gerar o protocolo. O registro de um novo TEP dispara um e-mail de notificacao (secao 2.8)."
Bul "No rodape do painel de detalhes de um TEP, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do FORALF00341, pronto para impressao ou para abrir no Word."

H2 "Passo 9 - RLA, Registro de Licoes Aprendidas (FORALF00342)"
P "Quando usar: junto com o TEP, ao final do projeto (ou tambem em pontos intermediarios), para capturar o aprendizado organizacional."
Img "30_f10_novo.png" "Tela de novo RLA." 5.6
Img "31_f10_lista.png" "RLAs cadastrados (estado vazio)." 5.6
Bul "Responda as perguntas abertas dos blocos Visao geral, Destaques, Desafios e Pos-projeto."
Bul "Nos blocos de avaliacao estruturada, marque Sim/Nao/Parcial/N-A para cada afirmacao - cada bloco calcula automaticamente um placar percentual."
P "Clique em `Registrar RLA`. O score global fica salvo junto com o registro."
Bul "No rodape do painel de detalhes de um RLA, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do FORALF00342 (incluindo os quatro blocos de avaliacao estruturada com o placar de cada um), pronto para impressao ou para abrir no Word."

H2 "Passo 10 - Plano de Comunicacao de Projeto (FORALF00308)"
P "Quando usar: para planejar o que sera comunicado, a quem e em qual momento do projeto. Este artefato tem uma ferramenta eletronica propria, organizada como uma tabela de referencia com abas Painel (visualizacao) e Editar dados."
Img "27_f08_painel.png" "Painel do Plano de Comunicacao: cada linha da tabela e um tipo de comunicacao do projeto." 5.6
$r10 = @(
  @("Coluna","Descricao"),
  @("Tipo de Comunicacao","Nome do momento/evento de comunicacao (ex.: Status semanal, Go/No-Go)"),
  @("O que Comunicar","Conteudo tratado nessa comunicacao"),
  @("Quando Comunicar","Periodicidade ou gatilho (ex.: Semanal, Em ate 2h do evento)"),
  @("Com Quem se Comunicar","Publico-alvo da comunicacao"),
  @("Como Comunicar","Canal e formato (ex.: Reuniao + ata, E-mail resumo + dashboard)"),
  @("Quem Comunica","Responsavel por comunicar (ex.: GP, GP / PMO)")
)
TableSimple $r10 @(4.5,11.5)
P "As 15 linhas padrao do plano (Comeco do projeto, Diario de equipe, Status semanal, Report executivo, Comite de Mudancas, Gestao de Riscos, Comunicacao de incidentes, Entregas e marcos, Integracao com stakeholders externos, Comunicacao de mudancas organizacionais, Treinamentos, Homologacao/UAT, Go/No-Go, Pos-Lancamento e Licoes aprendidas & Encerramento) vem preenchidas com o conteudo do documento oficial FORALF00308 nas 6 colunas - use `Editar dados` para ajustar aos seus projetos."
Nota "Edicao restrita por papel: diferente dos demais formularios, so PMO/Admin pode usar `Editar dados` e salvar alteracoes no Plano de Comunicacao - controlavel pelo Admin em Administracao > Configuracoes (secao 6.3), mesmo esquema de lista de papeis usado no Painel Executivo e nos dois relatorios. Qualquer usuario autenticado continua podendo abrir a aba Painel, visualizar todas as colunas e usar o botao Imprimir; quem nao tem permissao de edicao ve a aba `Editar dados` escondida e um aviso no lugar dela."

# ============================================================
# 4. RELATORIOS FINAIS
# ============================================================
H1 "4. Gerando os relatorios finais"
P "Esta secao detalha os dois relatorios que fecham o ciclo de uso da ferramenta: o Relatorio de Situacao de Projetos (visao de controle do portfolio) e o Relatorio de Entregas e Beneficios (ficha executiva de programa e projetos, base do Gate 2)."

H2 "4.1 Relatorio de Situacao de Projetos (FORALF11)"
P "Visao consolidada de todos os projetos do portfolio: status, percentual de execucao, marcos e pontos de atencao. Tem tres abas: Painel, Editar dados e Importar Project."
Img "32_f11_painel.png" "Painel do Relatorio de Situacao (estado vazio, antes do primeiro preenchimento)." 5.6
Img "33_f11_editar.png" "Aba Editar dados: cabecalho do relatorio e botao para adicionar projetos/marcos." 5.6
Img "34_f11_importar.png" "Aba Importar Project: arraste uma exportacao do MS Project (Excel ou CSV) para preencher em lote." 5.6
Bul "Preencha o cabecalho (mes/ano de referencia, responsavel, previsao financeira, entregas planejadas)."
Bul "Adicione cada projeto do portfolio, com status, % execucao, datas e os campos `Merece atencao` e `Merece destaque`."
Nota "Projeto vinculado (Gate 1) - opcional: cada linha de projeto pode, opcionalmente, ser associada a um Projeto ja Aprovado no Gate 1 (o mesmo cadastro compartilhado usado pelo Canvas, TAP, Planejamento, EAP, SMP, TEP e RLA - secao 2.5). Ao selecionar um projeto na lista, o campo Nome do projeto e preenchido automaticamente (se ainda estiver vazio) e aparece o botao `Ver historico`, que mostra as alteracoes desse projeto vindas de qualquer outro formulario vinculado a ele. Deixar em `Nenhum` mantem o comportamento anterior - o projeto do Relatorio de Situacao continua sendo texto livre, sem vinculo."
Img "42_f11_projeto_vinculado.png" "Projeto vinculado selecionado, com o nome preenchido automaticamente e o historico compartilhado exibido." 5.6
Bul "Alternativa mais rapida: use `Importar Project` - o sistema detecta automaticamente as colunas de nome, % concluido, inicio, termino e responsavel."
Bul "Clique em `Salvar e ver painel`. No rodape, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do FORALF11 com o retrato completo do portfolio (visao geral dos projetos, entraves e encaminhamentos, resultados alcancados e o resumo de atencao/destaque por programa), pronto para impressao ou para abrir no Word. Se uma versao historica estiver sendo visualizada (Nota abaixo), os dois botoes exportam essa versao, nao a atual."
Nota "Sinalizacao automatica: a tabela de projetos do Painel ganhou a coluna `Sinalizacao` e o quadro de indicadores ganhou o card `Em risco de atraso`. E um calculo automatico, independente do campo Status manual: compara o % Execucao informado com o % que seria esperado pelo tempo ja decorrido entre o Inicio previsto e o Termino previsto do projeto. Se a defasagem for de 15 pontos percentuais ou mais, aparece `Risco de atraso`; se o Termino previsto ja passou e o projeto nao esta em 100%, aparece `Prazo vencido`. Projetos Concluido, Cancelado ou Paralisado ficam fora desse calculo. O resultado tambem sai no documento exportado."
Img "38_f11_sinalizacao.png" "Painel com o card `Em risco de atraso` e a coluna Sinalizacao mostrando os selos Risco de atraso e Prazo vencido." 5.6
Bul "Na aba Editar dados, alem dos projetos, ha dois cartoes independentes: `Entraves` e `Encaminhamentos`. Cada item tem Descricao, Responsavel, Prazo e Status (Aberto / Resolvido / Cancelado)."
Nota "Gestao de Entraves e Encaminhamentos: entraves sao bloqueios que impedem o avanco do portfolio (ex.: dependencia externa, aprovacao pendente); encaminhamentos sao os proximos passos e decisoes acordadas para desbloquear ou avancar o trabalho. Os dois sao listas simples, independentes dos projetos cadastrados - continuam visiveis no Painel mesmo que nenhum projeto tenha sido informado ainda. No Painel, aparecem lado a lado na secao `Entraves e encaminhamentos`, com o Status de cada item destacado por selo colorido, e tambem saem no documento exportado."
Img "41_f11_entraves_encaminhamentos.png" "Secao Entraves e encaminhamentos no Painel, com itens Aberto e Resolvido lado a lado." 5.6
Bul "Ainda em Editar dados, o cartao `Resultados alcancados` registra o desempenho real do portfolio: Descricao, Tipo (Indicador / Entrega / Beneficio), Planejado e Alcancado."
Nota "Reporte de Resultados (D06.5): este cartao e a base do momento em que o Gerente de Projetos e o Dono do Negocio apresentam as instancias de decisao no que o projeto foi alcancado - nao so o que foi planejado, mas indicadores, entregas e beneficios de fato realizados, lado a lado com a meta original. E uma lista simples, independente dos projetos cadastrados, com o mesmo comportamento de Entraves e Encaminhamentos: continua visivel no Painel mesmo sem nenhum projeto informado, e sai no documento exportado. Junto com os Entraves e Encaminhamentos (D06.6/D06.7), completa os tres artefatos de governanca da Diretriz D06 sem exigir um 13o formulario separado - o Reporte de Resultados e atendido dentro do proprio Relatorio de Situacao."
Img "47_f11_resultados_alcancados.png" "Secao Resultados alcancados no Painel, com selos coloridos por Indicador, Entrega e Beneficio." 5.6
Bul "No rodape do Painel, o botao `Historico de versoes` mostra todas as vezes que este relatorio foi salvo, com data/hora e quem salvou."
Nota "Versionamento: toda vez que o relatorio e salvo (`Salvar e ver painel`), o sistema arquiva uma copia completa dos dados daquele momento, alem de atualizar a versao atual. Isso preserva as referencias mensais anteriores, que antes eram sobrescritas a cada novo salvamento. Ao clicar em `Visualizar` numa versao antiga, o Painel passa a mostrar aquele instantaneo, com um aviso amarelo no topo (`Visualizando versao... - somente leitura`) e um botao `Voltar a versao atual`. Enquanto uma versao antiga esta sendo visualizada, as abas `Editar dados` e `Importar Project` ficam bloqueadas - e preciso voltar a versao atual antes de editar."
Img "46_f11_versao_historica_banner.png" "Aviso de versao historica (somente leitura), com o botao para voltar a versao atual." 5.6
Nota "Este relatorio nao tem aprovacao/status formal - e uma ferramenta viva de acompanhamento, atualizada sempre que a situacao dos projetos mudar."

H2 "4.2 Relatorio de Entregas e Beneficios (FORALF12)"
P "O relatorio mais importante do ponto de vista de governanca: e sobre ele que ocorre o Gate 2 - Pactuacao, quando o Dono do Negocio assume formalmente, perante a Alta Gestao, o compromisso de que todo o trabalho planejado sera executado."
Img "35_f12_painel.png" "Painel do Relatorio de Entregas e Beneficios (estado vazio)." 5.6
Img "36_f12_editar.png" "Aba Editar dados: Ficha do Programa, o primeiro bloco a preencher." 5.6
Bul "Preencha a Ficha do Programa (codigo, nome, unidade, responsavel, justificativa, objetivo, alinhamento estrategico)."
Bul "Cadastre indicadores e valores estimados do programa, depois cada projeto vinculado (com suas proprias entregas, indicadores e valores)."
Nota "Projeto vinculado (Gate 1) - opcional: assim como no Relatorio de Situacao (secao 4.1), cada projeto cadastrado aqui pode ser associado a um Projeto ja Aprovado no Gate 1. A selecao preenche automaticamente o Nome do projeto (se ainda estiver vazio) e libera o botao `Ver historico`, mostrando as alteracoes feitas nesse projeto em qualquer outro formulario vinculado a ele - incluindo decisoes de SMP (Passo 7)."
Img "43_f12_projeto_vinculado.png" "Projeto vinculado selecionado no Relatorio de Entregas, com nome e historico preenchidos." 5.6
Nota "Importar indicadores do TAP: quando um Projeto vinculado e selecionado, a secao `Indicadores do projeto` ganha o botao `Importar do TAP`. Ele busca o(s) TAP(s) registrados para aquele mesmo projeto e traz os indicadores de resultado ja preenchidos la (secao 9 do TAP - Passo 3), evitando digitar tudo de novo: o Valor inicial do TAP vira o Valor atual aqui, e o Valor final vira a Meta. Indicadores com o mesmo nome que ja estiverem na lista nao sao duplicados - o botao pode ser clicado varias vezes com seguranca. Se nenhum TAP for encontrado para o projeto selecionado, aparece um aviso."
Img "44_f12_importar_tap.png" "Indicadores do projeto apos importar do TAP - Valor atual e Meta preenchidos automaticamente." 5.6
Bul "Clique em `Salvar e ver painel`. No rodape, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do FORALF12 (ficha do programa, indicadores, valores estimados e a ficha de cada projeto com suas entregas), pronto para levar a reuniao de pactuacao ou abrir no Word."
Nota "Gate 2 - Pactuacao: no topo da aba Editar dados ha um cartao `Gate 2 - Pactuacao` com Status (Pendente de pactuacao / Pactuado), Data de pactuacao e Aprovador. Assim como o Gate 1 da Solicitacao de Demanda, so o Admin consegue alterar o Status - qualquer outro papel ve o controle travado com um aviso. Ao marcar `Pactuado`, a Data e o Aprovador sao preenchidos automaticamente (editaveis). Enquanto o status estiver `Pactuado`, todos os demais campos do relatorio ficam bloqueados para edicao, inclusive para o Admin, ate que o Gate 2 seja reaberto (status volte para `Pendente de pactuacao`)."
Img "40_f12_gate2.png" "Cartao Gate 2 - Pactuacao preenchido, com o relatorio travado apos a pactuacao." 5.6
Nota "So depois do Gate 2 pactuado o projeto deve avancar para a execucao. Um projeto que comeca a ser executado sem essa pactuacao corre o risco de ser questionado ou desautorizado em momentos criticos. Ao marcar Pactuado, um e-mail de notificacao e disparado automaticamente (secao 2.8)."
Bul "No rodape do Painel, o botao `Historico de versoes` lista todas as vezes que este relatorio foi salvo."
Nota "Versionamento: assim como no Relatorio de Situacao (secao 4.1), toda vez que este relatorio e salvo, uma copia completa fica arquivada com data/hora e quem salvou - util para reconstituir o programa como ele estava em uma pactuacao (Gate 2) anterior. Clicar em `Visualizar` numa versao antiga mostra o mesmo aviso de somente leitura e bloqueia a edicao ate voltar a versao atual."
Img "45_f12_historico_versoes.png" "Lista de versoes salvas do Relatorio de Entregas, com data, programa e quem salvou." 5.6

# ============================================================
# 5. VALIDADOR DE PROJETOS
# ============================================================
H1 "5. Validador de Projetos"
P "O Validador de Projetos e uma ferramenta de apoio a decisao sobre Fatores Criticos de Sucesso (FCS) e Beneficios em instituicoes de ensino. Diferente dos 12 formularios de registro, ele nao gera um artefato FORALF - e um painel de reflexao e diagnostico."
P "E a unica ferramenta, alem do Mapa de Diretrizes, que nao exige login. Todo o conteudo funciona livremente; login so e pedido para vincular uma avaliacao a um projeto e salvar o veredito no historico daquele projeto (secao 2.6 do documento Regras de Acesso detalha esse caso)."
H2 "5.1 Quadro de conexoes"
P "Mostra a matriz analitica entre 7 fatores criticos de sucesso (F1-F7) e 5 dimensoes de beneficio (B1-B5). Clique em um fator ou dimensao para ver as conexoes."
Img "05_validador_quadro.png" "Quadro de conexoes com o fator F1 selecionado - o painel a direita mostra os beneficios que ele habilita." 5.8
H2 "5.2 Simulador - e se...?"
P "Ajuste o quanto cada fator esta presente no seu projeto (0-100%) e veja o potencial estimado de cada dimensao de beneficio recalcular em tempo real."
Img "06_validador_simulador.png" "Simulador de fatores x beneficios. Abaixo dele, o aviso de login para quem quiser vincular a um projeto." 5.8
H2 "5.3 Assistente de decisao em 8 perguntas"
P "Responda a oito perguntas do semaforo de decisao, uma de cada vez. Ao final, o painel devolve um veredito com a regra aplicada."
Img "07_validador_assistente.png" "Primeira pergunta do assistente de decisao." 5.8
Img "08_validador_veredito.png" "Veredito final, com o placar de respostas verdes/amarelas/vermelhas." 5.8
Nota "Regra aplicada: qualquer resposta vermelha leva a encerrar/repactuar; tres ou mais amarelas levam a replanejar; caso contrario, o veredito e concluir."

# ============================================================
# 6. ADMINISTRACAO DO SISTEMA
# ============================================================
H1 "6. Administracao e Painel Executivo"
P "Duas paginas independentes, listadas no menu lateral em `Gestao`: a Administracao (gerencia usuarios, equipes e configuracoes, sempre restrita a Admin - secao 2.3) e o Painel Executivo (visao consolidada e agregada de todo o sistema, com acesso configuravel por papel - secao 6.3)."
H2 "6.1 Usuarios"
P "Lista todos os usuarios que ja fizeram login pelo menos uma vez, com nome, telefone (editavel diretamente na lista) e um seletor para alterar o papel."
Img "09_admin_usuarios.png" "Aba Usuarios: nome, telefone, papel atual e seletor de alteracao de papel por usuario." 5.8
P "No topo da lista, o Admin pode pre-cadastrar uma pessoa que ainda nao fez login, preenchendo nome, telefone (opcional), e-mail e papel e clicando em `+ Adicionar usuario`. Enquanto a pessoa nao faz o primeiro login, ela aparece na lista com o selo `Pendente - 1o login`, e o Admin ainda pode ajustar telefone e papel ou remover o pre-cadastro pelo botao `Remover`."
Nota "Quando a pessoa pre-cadastrada faz o primeiro login, o sistema aplica automaticamente o nome, telefone e papel definidos pelo Admin ao perfil recem-criado - ela nao entra mais como `Solicitante` por padrao, e o pre-cadastro pendente desaparece da lista."
P "Cada usuario ja cadastrado (menos o proprio Admin logado) tem um botao `Remover` na ultima coluna. Ao clicar, o sistema checa na hora se a pessoa faz parte da equipe de algum projeto, criou algum projeto, ou tem registro criado em seu nome em algum dos 9 formularios que gravam essa informacao. Se algum desses vinculos existir, a remocao e recusada e o Admin ve exatamente qual foi o motivo; caso contrario, pede confirmacao final antes de excluir o perfil."
Nota "Remover aqui apaga so o perfil (papel, nome, telefone) do sistema - nao impede a pessoa de logar de novo depois; se logar, um perfil novo e criado do zero, com o papel padrao Solicitante. Os tres formularios que sao documentos unicos e compartilhados (Plano de Comunicacao, Relatorio de Situacao e Relatorio de Entregas) nao entram na checagem automatica, por nao terem o conceito de `quem criou o registro`."
H2 "6.2 Equipes"
P "Define quem faz parte da equipe de cada projeto - a base da restricao por equipe descrita na secao 2.5. Selecione um projeto, escolha um papel e clique em `+ Adicionar a equipe`."
Img "10_admin_equipes.png" "Aba Equipes, com um projeto selecionado e um membro ja cadastrado." 5.8
H2 "6.3 Configuracoes"
P "Controla os dois interruptores de acesso sem login (Solicitacao de Demanda e Ata de Reuniao), descritos na secao 2.6."
Img "11_admin_config.png" "Aba Configuracoes: os dois interruptores de acesso sem login, hoje desativados." 5.8
P "Logo abaixo, tres grupos de caixas de selecao controlam quem pode ver o Painel Executivo, o Relatorio de Situacao de Projetos e o Relatorio de Entregas e Beneficios: o Admin marca um ou mais papeis por pagina, e quem tiver um papel fora da lista marcada ve `Acesso restrito` ao tentar abrir aquela pagina. Por padrao, o Painel Executivo comeca marcado so para Admin (mesmo comportamento de antes), e os dois relatorios comecam com todos os papeis marcados (acesso livre, tambem o comportamento de antes) - a restricao so entra em vigor se o Admin desmarcar algum papel."
P "Um quarto grupo, `Quem pode editar o Plano de Comunicacao de Projeto`, controla so a aba `Editar dados` desse formulario (Passo 10) - diferente dos tres grupos acima, que controlam a pagina inteira. Comeca marcado apenas para PMO/Admin; qualquer papel fora da lista continua vendo e imprimindo o Painel normalmente, so nao consegue editar."
Nota "O papel Admin sempre tem acesso a essas quatro configuracoes, mesmo que fique desmarcado por engano - isso evita que o proprio Admin perca o acesso e fique sem como reverter a configuracao."
P "Por ultimo, dois interruptores independentes controlam a importacao por IA na Ata de Reuniao - cada um desligado por padrao, ja que cada uso tem custo real de API, e cada um vale para todo mundo sem excecao, inclusive Admin: um para a importacao de transcricao em texto (secao 2.10) e outro, separado, para a importacao de audio (secao 2.11). Cada interruptor tem seu proprio grupo de papeis logo abaixo, com Admin sempre incluido quando ligado. Como a importacao de audio so serve para gerar o texto que passa pela mesma analise, o botao `Anexar audio` na Ata so aparece para quem tem as duas permissoes ativas ao mesmo tempo."
H2 "6.4 Painel Executivo"
P "Agrega em tempo real dados de todos os 10 formularios de registro e dos 2 relatorios - sem exigir nenhuma mudanca no banco de dados, apenas consultando o que ja esta salvo. Pensada para dar ao PMO uma visao geral do sistema inteiro em uma unica tela, sem precisar abrir cada formulario individualmente. Por padrao so o Admin ve esta pagina, mas isso pode ser ampliado para outros papeis em Administracao > Configuracoes (secao 6.3)."
Bul "Gates de aprovacao: contagem de Solicitacoes de Demanda por status (Gate 1) e o status atual do Relatorio de Entregas (Gate 2 - Pactuacao)."
Bul "Governanca (dados do Relatorio de Situacao - secao 4.1): quantidade de Entraves e Encaminhamentos ainda `Aberto`, e quantos Resultados alcancados ja foram registrados."
Bul "Registros por formulario: uma linha por artefato (Canvas, TAP, Planejamento, EAP, SMP, Ata de Reuniao, TEP, RLA, alem da propria Solicitacao de Demanda), com o total e a contagem por status de cada um."
Bul "Portfolio de projetos: contagem de projetos Aprovados, Reprovados e em avaliacao no Gate 1, a partir do cadastro compartilhado (secao 2.5)."
Img "48_painel_executivo.png" "Painel Executivo: Gates de aprovacao e indicadores de governanca do Relatorio de Situacao." 5.8
Img "49_painel_executivo_tabela.png" "Registros por formulario e portfolio de projetos, com contagem por status." 5.8
Nota "Por reunir dados de varios formularios numa unica tela, o carregamento faz uma consulta para cada um dos 10 formularios de registro - pode levar alguns segundos a mais que os outros formularios da ferramenta. Clique em `Atualizar` para recarregar os numeros mais recentes a qualquer momento."
H2 "6.5 Ambiente de Treino"
P "O sistema tem um ambiente de treinamento/demonstracao totalmente separado da producao - mesma ferramenta, mesmo schema de dados, mas um banco proprio, populado com 7 projetos ficticios cobrindo o ciclo completo (de uma Solicitacao de Demanda ainda pendente ate um projeto totalmente encerrado com TEP e RLA). Serve para apresentacoes, estudo e capacitacao sem tocar em nenhum dado real. Veja o roteiro `Como rodar o Ambiente de Treinamento` para o endereco, as contas de acesso e o passo a passo completo."
P "Quando a Administracao e acessada em producao, aparece uma quarta aba, `Ambiente de Treino`, com um unico botao: `Resetar ambiente de treinamento`. Ao clicar (e confirmar o aviso), o sistema apaga todos os projetos e registros de formulario atualmente no ambiente de treinamento e recria os 7 projetos de exemplo do zero - util para deixar o ambiente limpo antes de uma nova turma, sem precisar rodar nenhum comando manualmente."
Bul "A aba so aparece quando a Administracao e acessada pelo endereco de producao - quem acessa a Administracao pelo proprio ambiente de treinamento nao ve essa opcao."
Bul "Nao afeta nenhum dado de producao em nenhuma hipotese - a acao roda numa Edge Function que so tem permissao de escrita no banco do ambiente de treinamento."
Bul "As 6 contas fixas de treinamento e as configuracoes do ambiente de treino nao sao apagadas pelo reset - so os projetos e os registros de formulario."
Nota "Assim como as outras acoes exclusivas de Admin, a Edge Function confere o papel do usuario direto no banco de producao antes de fazer qualquer alteracao - nao basta ver o botao na tela."

# ============================================================
# 7. PERGUNTAS FREQUENTES
# ============================================================
H1 "7. Perguntas frequentes"
H3 "Como corrijo um registro que ja salvei?"
P "Va ate a aba de registros cadastrados do formulario correspondente, clique na linha da tabela para abrir o painel de detalhes e clique em `Editar dados` (ou altere o status pelo seletor no rodape). Para os relatorios FORALF11 e FORALF12, volte para `Editar dados`, ajuste os campos e clique novamente em `Salvar e ver painel`."
H3 "Como excluo um registro?"
P "Abra o registro na tela de detalhes e clique em `Excluir`, no rodape. A acao pede confirmacao e nao pode ser desfeita depois de confirmada."
H3 "Por que o botao de registrar/salvar esta desabilitado?"
P "Nos 7 formularios com restricao por equipe (secao 2.5), o botao fica desabilitado se voce nao for membro da equipe do projeto selecionado - ou nao tiver selecionado um projeto vinculado ainda. Peca a um Admin para adiciona-lo a equipe em Administracao > Equipes."
H3 "Por que nao consigo abrir um formulario?"
P "Se aparecer a tela de login (secao 1.3), voce precisa se autenticar. Se depois de logado aparecer `Acesso restrito`, essa pagina (Administracao) e exclusiva para o papel Admin."
H3 "Os dados ficam salvos onde?"
P "Em um banco de dados real (Supabase/PostgreSQL), associados a sua conta. Nao e necessario salvar manualmente em outro lugar, mas recomenda-se exportar periodicamente em CSV os registros importantes, como backup."
H3 "Preciso preencher tudo na mesma sessao?"
P "Nao. Cada formulario pode ser preenchido e salvo em etapas diferentes - o registro fica disponivel na aba de consulta para ser retomado quando for necessario."

# ============================================================
# 8. ANEXO
# ============================================================
H1 "8. Anexo - Tabela de artefatos"
$r8 = @(
  @("#","Artefato","Codigo","Diretriz","Equipe?"),
  @("1","Solicitacao de Demanda","FORALF00339","D01.1","Nao"),
  @("2","Canvas de Projeto","FORALF00344","D01.6","Sim"),
  @("3","TAP - Termo de Abertura","FORALF00338","D02.1","Sim"),
  @("4","Planejamento e Desenvolvimento","FORALF00325","D02.2-D02.10","Sim"),
  @("5","EAP - Estrutura Analitica","(sem codigo)","D02.9","Sim"),
  @("6","Ata de Reuniao","FORALF00340","D01-D07 (transversal)","Nao"),
  @("7","SMP - Solicitacao de Mudanca","FORALF00343","D04.4","Sim"),
  @("8","TEP - Termo de Encerramento","FORALF00341","D05.1.9","Sim"),
  @("9","RLA - Licoes Aprendidas","FORALF00342","D05.1.6","Sim"),
  @("10","Plano de Comunicacao","FORALF00308","D03.5","Nao"),
  @("11","Relatorio de Situacao","FORALF11","D04.3 / D06.5-D06.7","Nao"),
  @("12","Relatorio de Entregas e Beneficios","FORALF12","D06.1/D06.2","Nao")
)
TableSimple $r8 @(1.0,5.2,3.2,4.6,2.0)
P "Documento gerado a partir do estado atual do codigo do sistema, com capturas de tela reais coletadas em 28 e 29/07/2026." 9 $false $true $colMuted "left" 0

# ============================================================
# FINALIZAR
# ============================================================
$doc.TablesOfContents.Item(1).Update()
$doc.Fields.Update()

if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
$doc.SaveAs2($OutPath, 16)
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output "SAVED: $OutPath"
