$ErrorActionPreference = "Stop"

# PowerShell https://github.com/PowerShell/PowerShell
Write-Host "> Installing PowerShell ..."
winget install -e --id Microsoft.PowerShell

# Git https://git-scm.com/download/win
Write-Host "> Installing Git ..."
winget install -e --id Git.Git

# gsudo https://github.com/gerardog/gsudo
Write-Host "> Installing gsudo ..."
winget install -e --id gerardog.gsudo

# Oh My Posh (via Microsoft Store) https://github.com/jandedobbeleer/oh-my-posh
Write-Host "> Installing Oh My Posh ..."
winget install -e --id XP8K0HKJFRXGCK

# Terminal-Icons https://github.com/devblackops/Terminal-Icons
Write-Host "> Installing Terminal-Icons ..."
Install-Module Terminal-Icons -Repository PSGallery -Force

# posh-git https://github.com/dahlbyk/posh-git
Write-Host "> Installing posh-git ..."
Install-Module posh-git -Force

# # PSReadLine https://github.com/PowerShell/PSReadLine
# Write-Host "> Installing PSReadLine ..."
# Install-Module PSReadLine -Scope CurrentUser -AllowPrerelease -Force

# z https://github.com/badmotorfinger/z
Write-Host "> Installing z ..."
Install-Module z -Force

# Configure PowerShell profile
Write-Host "> Configuring PowerShell profile ..."
$ProfileScript = {
    Import-Module posh-git
    Import-Module Terminal-Icons

    oh-my-posh init pwsh --config "$env:POSH_THEMES_PATH\paradox.omp.json" | Invoke-Expression

    Set-PSReadLineOption -PredictionSource History
    Set-PSReadLineOption -PredictionViewStyle ListView

    Set-Alias grep findstr
    Set-Alias which gcm
}
$ProfileScript.ToString().Trim() -replace "    ", "" | Out-File $PROFILE

# Font: Hack NF (https://github.com/ryanoasis/nerd-fonts)
Write-Host "> Installing font: Hack NF ..."
$TargetAsset = "Hack.zip"
$FontFileFilter = "HackNerdFont-*.ttf"
# download fonts archive
$Release = Invoke-RestMethod "https://api.github.com/repos/ryanoasis/nerd-fonts/releases/latest"
$Asset = $Release.assets | Where-Object { $_.name -eq $TargetAsset }
$TempFile = New-TemporaryFile
Invoke-RestMethod $Asset.browser_download_url -OutFile $TempFile
# extract fonts archive
$TempFolder = New-Item -ItemType Directory -Path (Join-Path $TempFile.DirectoryName (New-Guid))
Expand-Archive $TempFile $TempFolder
# install fonts
$ShellApplication = New-Object -ComObject Shell.Application
$FontFiles = $ShellApplication.Namespace($TempFolder.FullName).Items()
$FontFiles.Filter(0x40, $FontFileFilter)
$FontsFolder = $ShellApplication.Namespace(0x14)
$FontsFolder.CopyHere($FontFiles)

Write-Host "> Done!"
