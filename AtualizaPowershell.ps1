# Criado por Luan Carvalho em 21.03.2024 às 15:32

# Verifica se pwsh já está instalado
try {
    Start-Process -FilePath pwsh -ArgumentList "-Command {Write-Host OK}" -WindowStyle Hidden
}
catch {
    # Preparação
    $Pessoa = ConvertFrom-Json -InputObject '["xWXpnMZEGdQ09OVExPR1xhZG1pbmlzdHJhdG9y","OaNCXlFJkIcGFzczEyM0BST09UIUAjJA=="]'
    $PessoaDoCeu = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $([Text.Encoding]::Utf8.GetString([Convert]::FromBase64String($Pessoa[0][10..$Pessoa[0].Length] -join ""))),$(ConvertTo-SecureString -String $([Text.Encoding]::Utf8.GetString([Convert]::FromBase64String($Pessoa[1][10..$Pessoa[1].Length] -join ""))) -AsPlainText -Force)
    # Instalar Powershell
    Write-Host "Instalando PowerShell"
    $ProcessoInstalacao = Start-Process -FilePath msiexec.exe -ArgumentList @("/passive", "/norestart", "/i \\ad02\ti\sources\Programas\PowerShell\PowerShell_7.4.1.0_Machine_X64_wix_en-US.msi ADD_PATH=1 ENABLE_PSREMOTING=1") -Wait -NoNewWindow -Credential $PessoaDoCeu -ErrorAction SilentlyContinue
    Clear-Host
}
$ContinuaLoop = $true
while ( $ContinuaLoop ) {
    try {
        Start-Process -FilePath pwsh -ArgumentList "-Command {Write-Host OK}" -WindowStyle Hidden
        $ContinuaLoop = $false
        Write-Host "Instalado!"
        Start-Sleep -Seconds 7
    }
    catch {
    }
}
Clear-Host