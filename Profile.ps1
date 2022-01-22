Set-PSReadLineOption -HistorySaveStyle SaveNothing

$ChocolateyProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"

function prompt{
    Write-Host "$(get-location)" -ForegroundColor Blue 
    return "❯ "
}

function la{ Get-ChildItem | Format-Wide -Property Name -AutoSize }
function lf{ Get-ChildItem -File | Format-Wide -Property Name -AutoSize }
function ld{ Get-ChildItem -Directory | Format-Wide -Property Name -AutoSize }
