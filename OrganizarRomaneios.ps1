# Criado por Luan Carvalho em 05.03.2024 às 18:30

$OutputEncoding = [System.Text.Encoding]::UTF8 # Exibir a saída de texto em UTF-8 com BOM

# Parametros gerais  ▀▄
$PSStyle.Progress.View = 'Minimal'
#$PastaProcurar = "C:\Users\luan.carvalho\Desktop\TesteOrganizarRomaneios"
$PastaProcurar = "\\ad02\PCM"
$EnviarParaEscala = $false

# Instalando módulos 
$ModulosNecessarios = @("ImagePlayground")
for ( $i = 0; $i -lt $ModulosNecessarios.Count; $i++ ) {
    if ( $null -eq $( Get-Module -Name $( $ModulosNecessarios[$i] ) ).Version.Major ) {
        Install-Module -Name $( $ModulosNecessarios[$i] )
    }
}

#  #  FLUXO
$EstProcessosBemSucedidos = 0
$ObjetoFileSystem = New-Object -ComObject Scripting.FileSystemObject
$CaminhoCurtoSaida = $( $ObjetoFileSystem.GetFolder( "$( $env:USERPROFILE )\OrganizarRomaneio\") ).ShortPath
Get-Item -Path "$( $CaminhoCurtoSaida )\*" | Remove-Item -Force
# Obter caminho do programa Ghostscript
$CaminhoConversor = "$( $PSScriptRoot )\Ghostscript\gswin64c.exe"
# Obter caminho do programa Magick
$CaminhoRecortador = "$( $PSScriptRoot )\magick\magick.exe"
# Ajustes no nome do cliente
$DicionarioDeAjustes = @(
    @('.', ''), @('/', '_'), @('\', '_'), @('?', '_'), @('*', '_'), @('"', '_'), @('<', '_'), @('>', '_'), @('|', '_'), @(':', '_')
)
# Criar pasta de apoio, se necessário
if ( $( Test-Path -LiteralPath "$( $env:USERPROFILE )\OrganizarRomaneio\" ) -eq $false ) {
    New-Item "$( $env:USERPROFILE )\OrganizarRomaneio\" -ItemType Directory -ErrorAction SilentlyContinue > $null
}
# Encontrando PDFs
$PdfNaPasta = Get-ChildItem -Path $PastaProcurar -Filter "*.pdf"
$PdfNaPasta | ForEach-Object -Begin { $ProgressoId1 = 0; $AtividadeId1 = "Lendo PDF"; $ProgId1 = 1 } -End { Write-Progress -Activity $AtividadeId1 -PercentComplete 100 -Id $ProgId1 -Completed } -Process {
    # Exibindo progresso
    $ProgressoId1++
    Write-Progress -Status $_.Name -Activity $AtividadeId1 -PercentComplete $( $( $ProgressoId1 / $pdfNaPasta.Count ) * 100 ) -Id $ProgId1
    # Obter caminho curto do PDF
    $CaminhoCurtoPdf = $( $ObjetoFileSystem.GetFile($_.FullName) ).ShortPath
    # Converter PDF para imagem
    $ProcessoConverter = Start-Process -FilePath $CaminhoConversor -ArgumentList @("-dBATCH", "-dNOPAUSE", "-sDEVICE=jpeg", "-r300", "-sOutputFile=$($CaminhoCurtoSaida)\saida%d.jpg", $CaminhoCurtoPdf) -Wait -WindowStyle Hidden
    # Recortando imagens
    $CodigosEncontrados = @()
    $JpgNaPastaDeSaida = Get-ChildItem -Path $CaminhoCurtoSaida -Filter "*.jpg"
    $JpgNaPastaDeSaida | ForEach-Object -Begin { $ProgressoId2 = 0; $AtividadeId2 = "Dividindo imagens do PDF"; $ProgId2 = 2 } -End { Write-Progress -Activity $AtividadeId2 -PercentComplete 100 -Id $ProgId2 -Completed } -Process {
        # Exibindo progresso
        $ProgressoId2++
        Write-Progress -Status $_.BaseName -Activity $AtividadeId2 -PercentComplete $( $( $ProgressoId2 / $JpgNaPastaDeSaida.Count ) * 100 ) -Id $ProgId2
        $ProcessoRecorte = Start-Process -FilePath $CaminhoRecortador -ArgumentList @($( $ObjetoFileSystem.GetFile( $_.FullName ) ).ShortPath, "-crop 700x620", "$($CaminhoCurtoSaida)\saida_recorte-$( $_.BaseName )-%03d.png") -Wait -WindowStyle Hidden
    }
    # Percorrendo recortes
    $PedacosDaImagem = Get-Item -Path "$($CaminhoCurtoSaida)\saida_recorte*.png"
    $PedacosDaImagem | ForEach-Object -Begin { $ProgressoId3 = 0; $AtividadeId3 = "Procurando códigos de barra"; $ProgId3 = 3 } -End { Write-Progress -Activity $AtividadeId3 -PercentComplete 100 -Id $ProgId3 -Completed } -Process {
        # Exibindo progresso
        $ProgressoId3++
        Write-Progress -Status $_.Name -Activity $AtividadeId3 -PercentComplete $( $( $ProgressoId3 / $PedacosDaImagem.Count ) * 100 ) -Id $ProgId3
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
        } else {
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
                Write-Progress -Activity "Carga $_" -Status "Enviando para Escala" -PercentComplete $( $( $ProgressoId4 / $CodigosEncontrados.Count ) * 100 ) -Id $ProgId4
                # Corpo da requisição
                $CorpoObjeto = @{
                    Carga = $_
                    Arquivo = @{
                        Nome = $NovoNomePdf
                        Conteudo = $ArquivoBase64
                    }
                }
                $CorpoJson = ConvertTo-Json -InputObject $CorpoObjeto -Depth 10 -Compress
                $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
                $headers.Add("Content-Type", "application/json")
                $RespostaEscala = Invoke-WebRequest "http://10.20.0.113:8000/api/cargadescarga/AnexarRomaneio?Carga=$( $_ )" -Method 'POST' -Headers $headers -Body $CorpoJson # -OutFile "$PastaProcurar\out.json"
                $RespostaEscalaJson = ConvertFrom-Json -InputObject $( [system.Text.Encoding]::UTF8.GetString( $RespostaEscala.Content ) ) -Depth 10
                #$TesteRespostaEscala = '{"Carga":103453,"Cliente":"PITARON - ITAJAI/SC","Data":"31/12/2024 14:32:00"}'
                #$RespostaEscalaJson = ConvertFrom-Json -InputObject $TesteRespostaEscala -Depth 10
                # Coletar cliente, mes e ano da carga
                $RECliente = $RespostaEscalaJson.Cliente
                $REAno = Get-Date -Date $( $RespostaEscalaJson.Data ) -Format "yyyy"
                $REMes = Get-Date -Date $( $RespostaEscalaJson.Data ) -Format "MM.MMMM"
                # Ajuste na String com nome do cliente
                foreach ( $Substituicao in $DicionarioDeAjustes ) {
                    $RECliente = $RECliente.Replace($Substituicao[0], $Substituicao[1])
                }
                # Exibir progresso atualizado
                Write-Progress -Activity "Carga $_" -Status "Organizando localmente" -PercentComplete $( $( $ProgressoId4 / $CodigosEncontrados.Count ) * 100 ) -Id $ProgId4
                # Criar pasta apropriada, se não houver
                $CaminhoApropriado = "$( $PastaProcurar )\$( $REAno )\$( $RECliente )\$( $REMes )"
                if ( $( Test-Path -Path $CaminhoApropriado) -eq $false ) {
                    $NovoCaminho = New-Item -Path $CaminhoApropriado -ItemType Directory
                }
                $ParametrosCopiaOuMove = @{
                    Path = $CaminhoArquivoRenomeado
                    Destination = "$( $CaminhoApropriado )\$( $NovoNomePdf )"
                }
                # Copiar PDF para a pasta apropriada, caso seja encontrado mais de um código em um mesmo PDF. Mover PDF para a pasta apropriada caso esteja processando o ultimo código ou seja o único código no PDF.
                if ( $( Test-Path -Path "$( $CaminhoApropriado )\$( $NovoNomePdf )" ) ) {
                    $ParametrosCopiaOuMove.Destination = "$( $CaminhoApropriado )\$( $NovoNomePdf2 )"
                }
                if ( $( $CodigosEncontrados.IndexOf($_) + 1 ) -eq $CodigosEncontrados.Count ) {
                    # Mover PDF para a pasta
                    $ItemMovido = Move-Item @ParametrosCopiaOuMove
                } else {
                    # Copiar PDF para a pasta
                    $ItemMovido = Copy-Item @ParametrosCopiaOuMove
                }
            }
        }
        $EstProcessosBemSucedidos++
    }
    # Limpando pasta de saída
    #Get-Item -Path "$( $CaminhoCurtoSaida )\*" | Remove-Item -Force
}
Write-Host "RESULTADO: $( [Math]::Floor( $( 100 / $PdfNaPasta.Count ) * $EstProcessosBemSucedidos) )% processado com sucesso ($( $EstProcessosBemSucedidos )/$( $PdfNaPasta.Count ))"