<#
    .SYNOPSIS
        Windows PowerShell Core bootstrapper
#>

param(
    # Oh My Posh theme to configure for use
    [string] $Theme = "paradox",

    # Nerd Font to install
    [string] $NerdFont = "Hack",

    # Skip installation of fonts
    [switch] $NoFonts = $false
)

$ErrorActionPreference = "Stop"

$NonTerminatingErrorCount = 0

$Steps = @(
    @{
        Descriptor  = "Installing Windows Terminal"
        Metadata    = @{ Source = "https://github.com/microsoft/terminal" }
        Skip        = $true
        ScriptBlock = {
            winget install -e --id Microsoft.WindowsTerminal
            Write-Host "> $((winget list -e --id Microsoft.WindowsTerminal).Split([Environment]::NewLine) | Select-Object -Last 1)"
        }
    }
    @{
        Descriptor  = "Installing PowerShell"
        Metadata    = @{ Source = "https://github.com/PowerShell/PowerShell" }
        Skip        = $true
        ScriptBlock = {
            winget install -e --id Microsoft.PowerShell
            Write-Host "> $((winget list -e --id Microsoft.PowerShell).Split([Environment]::NewLine) | Select-Object -Last 1)"
        }
    }
    @{
        Descriptor  = "Installing Git"
        Metadata    = @{ Source = "https://git-scm.com/download/win" }
        Skip        = $true
        ScriptBlock = {
            winget install -e --id Git.Git
            Write-Host "> $((winget list -e --id Git.Git).Split([Environment]::NewLine) | Select-Object -Last 1)"
        }
    }
    @{
        Descriptor  = "Installing gsudo"
        Metadata    = @{ Source = "https://github.com/gerardog/gsudo" }
        Skip        = $true
        ScriptBlock = {
            winget install -e --id gerardog.gsudo
            Write-Host "> $((winget list -e --id gerardog.gsudo).Split([Environment]::NewLine) | Select-Object -Last 1)"
        }
    }
    @{
        Descriptor  = "Installing Oh My Posh"
        Metadata    = @{ Source = "https://github.com/jandedobbeleer/oh-my-posh" }
        Skip        = $true
        ScriptBlock = {
            winget install -e --id XP8K0HKJFRXGCK # via Microsoft Store
            Write-Host "> $((winget list -e --id XP8K0HKJFRXGCK).Split([Environment]::NewLine) | Select-Object -Last 1)"
        }
    }
    @{
        Descriptor  = "Installing Terminal-Icons"
        Metadata    = @{ Source = "https://github.com/devblackops/Terminal-Icons" }
        Skip        = $true
        ScriptBlock = {
            Install-Module Terminal-Icons -Repository PSGallery
            Write-Host "> Version: $((Get-Module Terminal-Icons).Version)"
        }
    }
    @{
        Descriptor  = "Installing posh-git"
        Metadata    = @{ Source = "https://github.com/dahlbyk/posh-git" }
        Skip        = $true
        ScriptBlock = {
            Install-Module posh-git
            Write-Host "> Version: $((Get-Module posh-git).Version)"
        }
    }
    @{
        Descriptor  = "Installing z"
        Metadata    = @{ Source = "https://github.com/badmotorfinger/z" }
        Skip        = $true
        ScriptBlock = {
            Install-Module z
            Write-Host "> Version: $((Get-Module z).Version)"
        }
    }
    @{
        Descriptor  = "Configuring PowerShell profile"
        Skip        = $false
        ScriptBlock = {
            $ThemeFile = [System.IO.FileInfo](Join-Path $env:POSH_THEMES_PATH "$Theme.omp.json")
            if (-Not $ThemeFile.Exists) {
                $AvailableThemes = Get-ChildItem $env:POSH_THEMES_PATH -Filter "*.omp.json" `
                | ForEach-Object { $_.Name -replace ".omp.json", "" } | Join-String -Separator ", "
                Write-Error -ErrorAction 'Continue' "Oh My Posh theme '$Theme' not found. Expected one of: $AvailableThemes"
                $Script:NonTerminatingErrorCount++
                return
            }

            $ProfileScript = {
                Import-Module posh-git
                Import-Module Terminal-Icons

                oh-my-posh init pwsh --config "$env:POSH_THEMES_PATH\{{THEME_FILE_NAME}}" | Invoke-Expression

                Set-PSReadLineOption -PredictionSource History
                Set-PSReadLineOption -PredictionViewStyle ListView
            }.ToString() -replace "{{THEME_FILE_NAME}}", $ThemeFile.Name

            if ((Test-Path $PROFILE) -and (Get-Content $PROFILE).Trim().Length -gt 0) {
                Write-Warning "Your PowerShell profile is not empty.`n$PROFILE"
                if ((Read-Host "> Overwrite? [yN]").Trim().ToLower() -ne 'y') {
                    Write-Warning "Skipping step."
                    return
                }
            }

            $ProfileScript.Trim() -replace "                ", "" | Out-File $PROFILE
            Write-Host "> Updated: $PROFILE"
        }
    }
    @{
        Descriptor  = "Installing font: $NerdFont Nerd Font"
        Metadata    = @{ Source = "https://github.com/ryanoasis/nerd-fonts" }
        Skip        = $false
        ScriptBlock = {
            if ($NoFonts) {
                Write-Warning "Skipping step."
                return
            }

            $TargetRelease = "tags/v2.3.3"
            $TargetAssetName = "$NerdFont.zip"
            $FontFileFilter = "* Nerd Font Complete Windows Compatible.ttf"

            # download fonts archive
            $Release = Invoke-RestMethod "https://api.github.com/repos/ryanoasis/nerd-fonts/releases/$TargetRelease"
            $Asset = $Release.assets | Where-Object { $_.name -eq $TargetAssetName }
            if (-Not $Asset) {
                $AvailableFonts = $Release.assets | Where-Object { $_.name -like "*.zip" } `
                | ForEach-Object { $_.name -replace ".zip", "" } | Join-String -Separator ", "
                Write-Error -ErrorAction 'Continue' "Nerd Font '$NerdFont' not found. Expected one of: $AvailableFonts"
                $Script:NonTerminatingErrorCount++
                return
            }
            $TempFile = New-TemporaryFile
            Invoke-RestMethod $Asset.browser_download_url -OutFile $TempFile

            # extract fonts archive
            $TempFolder = New-Item -ItemType Directory -Path (Join-Path $TempFile.DirectoryName (New-Guid))
            Expand-Archive $TempFile $TempFolder

            # install fonts
            $ShellApplication = New-Object -ComObject Shell.Application
            $FontFiles = $ShellApplication.Namespace($TempFolder.FullName).Items()
            $FontFiles.Filter(0x40, $FontFileFilter)
            if ($FontFiles.Count -eq 0) {
                throw "No files found within asset '$TargetAssetName' matching the filter '$FontFileFilter'"
            }
            $FontsFolder = $ShellApplication.Namespace(0x14)
            $FontsFolder.CopyHere($FontFiles)
        }
    }
)

Write-Host "> Bootstrapping ..."

for ($i = 0; $i -lt $Steps.Length; $i++) {
    $Step = $Steps[$i]
    if ($Step.Skip) { continue; }
    Write-Host "> [$('{0:d2}' -f ($i + 1))/$('{0:d2}' -f $Steps.Length)] $($Step.Descriptor) ..."
    Invoke-Command $Step.ScriptBlock
}

if ($NonTerminatingErrorCount -gt 0) {
    Write-Warning ("Bootstrapping completed with $NonTerminatingErrorCount non-terminating errors. " `
            + "It is recommended you resolve these errors and re-run the script.")
}
else {
    Write-Host "`n> Done!"
}
