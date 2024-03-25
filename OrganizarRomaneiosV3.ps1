# Criado por Luan Carvalho em 05.03.2024 às 18:30
# Nome do robô: Nara
# Função do robô: Renomeia arquivo do romaneio, envia uma cópia para o Escala e organiza o arquivo localmente.

$OutputEncoding = [System.Text.Encoding]::UTF8 # Exibir a saída de texto em UTF-8 com BOM

# Parametros gerais
#$PastaProcurar = "C:\Users\luan.carvalho\Desktop\TesteOrganizarRomaneios"
#$PastaProcurar = "C:\Users\luan.carvalho\CONTLOG LOG. E COM. EXTERIOR\PCM - 2.EXPEDIÇÃO\ROMANEIOS"
$PastaProcurar = $( Get-Location ).Path
$EnviarParaEscala = $true
$TituloRobo = "Nara – Robô organizadora de romaneios"

# Definindo título da janela
$TituloAntigo = $Host.UI.RawUI.WindowTitle
$Host.UI.RawUI.WindowTitle = $TituloRobo
Write-Host "█ $( $TituloRobo )`n"
# Configurando barra de progresso
try { $PSStyle.Progress.View = 'Minimal' } catch { Write-Host " " }
# Definindo notificação
Add-Type -AssemblyName System.Windows.Forms
$Notificacao = New-Object System.Windows.Forms.NotifyIcon
$Notificacao.Icon = [System.Drawing.SystemIcons]::Information
$Notificacao.Visible = $true
# Preparação
$Pessoa = ConvertFrom-Json -InputObject '["xWXpnMZEGdQ09OVExPR1xhZG1pbmlzdHJhdG9y","OaNCXlFJkIcGFzczEyM0BST09UIUAjJA=="]'
$PessoaDoCeu = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $([Text.Encoding]::Utf8.GetString([Convert]::FromBase64String($Pessoa[0][10..$Pessoa[0].Length] -join ""))),$(ConvertTo-SecureString -String $([Text.Encoding]::Utf8.GetString([Convert]::FromBase64String($Pessoa[1][10..$Pessoa[1].Length] -join ""))) -AsPlainText -Force)
$SessaoDaPessoaDoCeu = New-PSSession -Credential $PessoaDoCeu
# Instalando módulos, se necessário
$ModulosNecessarios = @("ImagePlayground")
$ModulosNecessarios | ForEach-Object -Begin { $ProgressoId0 = 0; $AtividadeId0 = "Instalando módulos necessários"; $ProgId0 = 1 } -End { Write-Progress -Activity $AtividadeId0 -PercentComplete 100 -Id $ProgId0 -Completed } -Process {
    # Exibindo progresso
    $ProgressoId0++
    Write-Progress -Status "$( $ProgressoId0 ) / $( $ModulosNecessarios.Count ): $( $_ )" -Activity $AtividadeId0 -PercentComplete $( $( $ProgressoId0 / $ModulosNecessarios.Count ) * 100 ) -Id $ProgId0
    # Módulos
    if ( $null -eq $( Get-Module -Name $( $_ ) ).Version.Major ) {
        Invoke-Command -Session $SessaoDaPessoaDoCeu -ScriptBlock {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force > $null
            Install-Module -Name $using:_ -Force -Confirm:$false
        }
        Import-Module -Name $_
    }
}
for ( $i = 0; $i -lt $ModulosNecessarios.Count; $i++ ) {
}

#  #  FLUXO
$EstProcessosBemSucedidos = 0
$ObjetoFileSystem = New-Object -ComObject Scripting.FileSystemObject
# Obter caminho do programa Ghostscript
$CaminhoConversor = "$( $PSScriptRoot )\Ghostscript\gswin64c.exe"
# Obter caminho do programa Magick
$CaminhoRecortador = "$( $PSScriptRoot )\magick\magick.exe"
# Ajustes no nome do cliente
$DicionarioDeAjustes = @(@('.', ''), @('/', '_'), @('\', '_'), @('?', '_'), @('*', '_'), @('"', '_'), @('<', '_'), @('>', '_'), @('|', '_'), @(':', '_'))
# Criar pasta de apoio, se necessário
if ( $( Test-Path -LiteralPath "$( $env:USERPROFILE )\OrganizarRomaneio\" ) -eq $false ) {
    New-Item "$( $env:USERPROFILE )\OrganizarRomaneio\" -ItemType Directory -ErrorAction SilentlyContinue > $null
}
$CaminhoCurtoSaida = $( $ObjetoFileSystem.GetFolder( "$( $env:USERPROFILE )\OrganizarRomaneio\") ).ShortPath
Get-Item -Path "$( $CaminhoCurtoSaida )\*" | Remove-Item -Force
# Encontrando PDFs
$PdfNaPasta = Get-ChildItem -Path $PastaProcurar -Filter "*.pdf" -File
$PdfNaPasta | ForEach-Object -Begin { $ProgressoId1 = 0; $AtividadeId1 = "Lendo PDF"; $ProgId1 = 1 } -End { Write-Progress -Activity $AtividadeId1 -PercentComplete 100 -Id $ProgId1 -Completed } -Process {
    # Exibindo progresso
    $ProgressoId1++
    Write-Progress -Status "$( $ProgressoId1 ) / $( $pdfNaPasta.Count ): $( $_.Name )" -Activity $AtividadeId1 -PercentComplete $( $( $ProgressoId1 / $pdfNaPasta.Count ) * 100 ) -Id $ProgId1
    # Obter caminho curto do PDF
    $CaminhoCurtoPdf = $( $ObjetoFileSystem.GetFile($_.FullName) ).ShortPath
    # Converter PDF para imagem
    $ProcessoConverter = Start-Process -FilePath $CaminhoConversor -ArgumentList @("-dBATCH", "-dNOPAUSE", "-sDEVICE=jpeg", "-r300", "-sOutputFile=$($CaminhoCurtoSaida)\saida%d.jpg", $CaminhoCurtoPdf) -Wait -WindowStyle Hidden
    # Rotacionar imagens, se necessário.
    $JpgNaPastaDeSaida = Get-ChildItem -Path $CaminhoCurtoSaida -Filter "*.jpg"
    $JpgNaPastaDeSaida | ForEach-Object -Begin { $ProgressoId2b = 0; $AtividadeId2b = "Rotacionando imagem para modo paisagem"; $ProgId2b = 2 } -End { Write-Progress -Activity $AtividadeId2b -PercentComplete 100 -Id $ProgId2b -Completed } -Process {
        # Exibindo progresso
        $ProgressoId2b++
        Write-Progress -Status "$( $ProgressoId2b ) / $( $JpgNaPastaDeSaida.Count )" -Activity $AtividadeId2b -PercentComplete $( $( $ProgressoId2b / $JpgNaPastaDeSaida.Count ) * 100 ) -Id $ProgId2b
        # Verificar se está no modo paisagem
        $CaminhoCurtoArquivoRotacionar = $ObjetoFileSystem.GetFile( $_.FullName ).ShortPath
        if ( $( Start-Process -FilePath $CaminhoRecortador -ArgumentList @('identify', '-format "%[fx:(w>=h)?1:0]"', $CaminhoCurtoArquivoRotacionar ) -Wait -NoNewWindow ) -eq 0 ) {
            $ProcessoRotacionar = Start-Process -FilePath $CaminhoRecortador -ArgumentList @('convert', '-rotate -90', $CaminhoCurtoArquivoRotacionar, $CaminhoCurtoArquivoRotacionar ) -Wait -WindowStyle Hidden
        }
    }
    # Recortando imagens
    $JpgNaPastaDeSaida = Get-ChildItem -Path $CaminhoCurtoSaida -Filter "*.jpg"
    $JpgNaPastaDeSaida | ForEach-Object -Begin { $ProgressoId2 = 0; $AtividadeId2 = "Dividindo imagens do PDF"; $ProgId2 = 2 } -End { Write-Progress -Activity $AtividadeId2 -PercentComplete 100 -Id $ProgId2 -Completed } -Process {
        # Exibindo progresso
        $ProgressoId2++
        Write-Progress -Status "$( $ProgressoId2 ) / $( $JpgNaPastaDeSaida.Count )" -Activity $AtividadeId2 -PercentComplete $( $( $ProgressoId2 / $JpgNaPastaDeSaida.Count ) * 100 ) -Id $ProgId2
        # Recortar
        $ProcessoRecorte = Start-Process -FilePath $CaminhoRecortador -ArgumentList @($( $ObjetoFileSystem.GetFile( $_.FullName ) ).ShortPath, "-crop 700x620", "$($CaminhoCurtoSaida)\saida_recorte-$( $_.BaseName )-%03d.png") -Wait -WindowStyle Hidden
    }
    # Percorrendo recortes
    $CodigosEncontrados = @()
    $PedacosDaImagem = Get-Item -Path "$( $CaminhoCurtoSaida )\*" -Include @("saida_recorte*-004.png", "saida_recorte*-018.png")
    $PedacosDaImagem | ForEach-Object -Begin { $ProgressoId3 = 0; $AtividadeId3 = "Procurando códigos de barra"; $ProgId3 = 3 } -End { Write-Progress -Activity $AtividadeId3 -PercentComplete 100 -Id $ProgId3 -Completed } -Process {
        # Exibindo progresso
        $ProgressoId3++
        Write-Progress -Status "$( $ProgressoId3 ) / $( $PedacosDaImagem.Count )" -Activity $AtividadeId3 -PercentComplete $( $( $ProgressoId3 / $PedacosDaImagem.Count ) * 100 ) -Id $ProgId3
        # Escaneando
        $ImagemScaneada = Get-ImageBarCode -FilePath $_.FullName
        if ( $ImagemScaneada.Status -eq "Found" -and $( $ImagemScaneada.Value ).Length -le 6 -and $CodigosEncontrados -notcontains $ImagemScaneada.Value ) {
            $CodigosEncontrados += $ImagemScaneada.Value
        }
    }
    # Renomeando
    if ( $CodigosEncontrados.Count -gt 0 ) {
        # Se houver códigos encontrados
        $NovoNomePdf = "$( $CodigosEncontrados -join "-" ).pdf"
        $NovoNomePdf2 = "$( $CodigosEncontrados -join "-" )_$( Get-Date -Format "yyyyMMddHHmmssfffffff" ).pdf"
        $CaminhoPasta = $ObjetoFileSystem.GetParentFolderName($_.FullName)
        if ( $( Test-Path -Path "$( $CaminhoPasta )\$( $NovoNomePdf )" ) ) {
            Rename-Item -Path $CaminhoCurtoPdf -NewName $NovoNomePdf2 -Force
            $CaminhoArquivoRenomeado = "$( $CaminhoPasta )\$( $NovoNomePdf2 )"
        }
        else {
            Rename-Item -Path $CaminhoCurtoPdf -NewName $NovoNomePdf -Force
            $CaminhoArquivoRenomeado = "$( $CaminhoPasta )\$( $NovoNomePdf )"
        }
        # Escalasoft
        if ( $EnviarParaEscala ) {
            # Se chave estiver ativada, enviar romaneio para escalasoft
            # Converter arquivo para base64
            $ArquivoBytes = [System.IO.File]::ReadAllBytes($CaminhoArquivoRenomeado)
            $ArquivoBase64 = [Convert]::ToBase64String($ArquivoBytes)
            # Enviar para cada número de romaneio encontrado
            $CodigosEncontrados | ForEach-Object -Begin { $ProgressoId4 = 0; $AtividadeId4 = "Escala/Organizar"; $ProgId4 = 4 } -End { Write-Progress -Activity $AtividadeId4 -PercentComplete 100 -Id $ProgId4 -Completed } -Process {
                # Exibindo progresso
                $ProgressoId4++
                Write-Progress -Activity "Carga $_" -Status "$( $ProgressoId4 ) / $( $CodigosEncontrados.Count ): Enviando para Escala" -PercentComplete $( $( $ProgressoId4 / $CodigosEncontrados.Count ) * 100 ) -Id $ProgId4
                # Corpo da requisição
                $CorpoObjeto = @{
                    Carga   = $_
                    Arquivo = @{
                        Nome     = $NovoNomePdf
                        Conteudo = $ArquivoBase64
                    }
                }
                $CorpoJson = ConvertTo-Json -InputObject $CorpoObjeto -Depth 10 -Compress
                $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
                $headers.Add("Content-Type", "application/json")
                $RespostaEscala = Invoke-WebRequest "http://10.20.0.113:8000/api/cargadescarga/AnexarRomaneio?Carga=$( $_ )" -Method 'POST' -Headers $headers -Body $CorpoJson -UserAgent "LuanCarvalho/$( Get-Date -Format "yyyy.MM.dd" ) RoboNara/3.0" # -OutFile "$PastaProcurar\out.json"
                $RespostaEscalaJson = ConvertFrom-Json -InputObject $( [system.Text.Encoding]::UTF8.GetString( $RespostaEscala.Content ) ) -Depth 10
                # Coletar cliente, mes e ano da carga
                $RECliente = $RespostaEscalaJson.Cliente
                $REAno = Get-Date -Date $( $RespostaEscalaJson.Data ) -Format "yyyy"
                $REMes = Get-Date -Date $( $RespostaEscalaJson.Data ) -Format "MM.MMMM"
                # Ajuste na String com nome do cliente
                foreach ( $Substituicao in $DicionarioDeAjustes ) {
                    $RECliente = $RECliente.Replace($Substituicao[0], $Substituicao[1])
                }
                # Exibir progresso atualizado
                Write-Progress -Activity "Carga $_" -Status "$( $ProgressoId4 ) / $( $CodigosEncontrados.Count ): Organizando localmente" -PercentComplete $( $( $ProgressoId4 / $CodigosEncontrados.Count ) * 100 ) -Id $ProgId4
                # Criar pasta apropriada, se não houver
                $CaminhoApropriado = "$( $PastaProcurar )\$( $REAno )\$( $RECliente )\$( $REMes )"
                if ( $( Test-Path -Path $CaminhoApropriado) -eq $false ) {
                    $NovoCaminho = New-Item -Path $CaminhoApropriado -ItemType Directory
                }
                $ParametrosCopiaOuMove = @{
                    Path        = $CaminhoArquivoRenomeado
                    Destination = "$( $CaminhoApropriado )\$( $NovoNomePdf )"
                }
                # Copiar PDF para a pasta apropriada, caso seja encontrado mais de um código em um mesmo PDF. Mover PDF para a pasta apropriada caso esteja processando o ultimo código ou seja o único código no PDF.
                if ( $( Test-Path -Path "$( $CaminhoApropriado )\$( $NovoNomePdf )" ) ) {
                    $ParametrosCopiaOuMove.Destination = "$( $CaminhoApropriado )\$( $NovoNomePdf2 )"
                }
                if ( $( $CodigosEncontrados.IndexOf($_) + 1 ) -eq $CodigosEncontrados.Count ) {
                    # Mover PDF para a pasta
                    $ItemMovido = Move-Item @ParametrosCopiaOuMove
                }
                else {
                    # Copiar PDF para a pasta
                    $ItemMovido = Copy-Item @ParametrosCopiaOuMove
                }
            }
        }
        $EstProcessosBemSucedidos++
    }
    # Limpando pasta de saída e variavel de códigos
    Get-Item -Path "$( $CaminhoCurtoSaida )\*" | Remove-Item -Force
    $CodigosEncontrados.Clear()
}
if ( $PdfNaPasta.Count -ge 1 ) {
    $TextoNotificacao = "FINALIZADO! $( [Math]::Floor( $( 100 / $PdfNaPasta.Count ) * $EstProcessosBemSucedidos) )% processado com sucesso ($( $EstProcessosBemSucedidos )/$( $PdfNaPasta.Count ))"
    $Notificacao.ShowBalloonTip(5000, $TituloRobo, $TextoNotificacao, [System.Windows.Forms.ToolTipIcon]::Info)
    Write-Host $TextoNotificacao
}
else {
    $TextoNotificacao = "NADA FEITO! Não há arquivos para processar. Caminho de busca: $( $PastaProcurar )"
    $Notificacao.ShowBalloonTip(5000, $TituloRobo, $TextoNotificacao, [System.Windows.Forms.ToolTipIcon]::Error)
    Write-Host $TextoNotificacao
}

$Host.UI.RawUI.WindowTitle = $TituloAntigo
<#
    V1:
        Projeto inicial;
    V2:
        Adicionado barra de progresso;
        Otimizações no uso de recursos;
        Enviar arquvos para Escalasoft;
    V3:
        Otimizado informações das barras de progresso;
        Rotaciona a imagem do PDF se ela não estiver em modo paisagem;
        Adicionado notificações de conclusão no Windows;
        User Agent da requisição para Escala personalizado;
#>