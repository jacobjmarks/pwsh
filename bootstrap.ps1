$ErrorActionPreference = "Stop"

Push-Location (Split-Path $MyInvocation.MyCommand.Path)

try {
    # PowerShell https://github.com/PowerShell/PowerShell
    Write-Host "> Installing PowerShell ..."
    winget install -e --id Microsoft.PowerShell

    # Git https://git-scm.com/download/win
    Write-Host "> Installing Git ..."
    winget install -e --id Git.Git

    # Oh My Posh (via Microsoft Store) https://github.com/jandedobbeleer/oh-my-posh
    Write-Host "> Installing Oh My Posh ..."
    winget install -e --id XP8K0HKJFRXGCK

    # Oh My Posh - Theme
    Write-Host "> Downloading Oh My Posh theme ..."
    Invoke-WebRequest  | Out-File "$env:POSH_THEMES_PATH\paradox.omp.json"

    # Terminal-Icons https://github.com/devblackops/Terminal-Icons
    Write-Host "> Installing Terminal-Icons ..."
    Install-Module Terminal-Icons -Repository PSGallery

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
    Get-Content .\profile.ps1 | Out-File $PROFILE

    Write-Host "> Done!"
}
finally {
    Pop-Location
}
