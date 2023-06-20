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

function WithPwsh {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]
        $Command
    )

    if ($PSEdition -ne 'Core') {
        pwsh -NoProfile -c $Command
    }
    else {
        Invoke-Command -NoNewScope $Command
    }
}

$NonTerminatingErrorCount = 0

$Steps = @(
    @{
        Descriptor  = "Installing Windows Terminal"
        Metadata    = @{ Source = "https://github.com/microsoft/terminal" }
        ScriptBlock = {
            winget install -e --id Microsoft.WindowsTerminal
            Write-Host "> $((winget list -e --id Microsoft.WindowsTerminal).Split([Environment]::NewLine) | Select-Object -Last 1)"
        }
    }
    @{
        Descriptor  = "Installing PowerShell"
        Metadata    = @{ Source = "https://github.com/PowerShell/PowerShell" }
        ScriptBlock = {
            winget install -e --id Microsoft.PowerShell
            Write-Host "> $((winget list -e --id Microsoft.PowerShell).Split([Environment]::NewLine) | Select-Object -Last 1)"
        }
    }
    @{
        Descriptor  = "Installing Git"
        Metadata    = @{ Source = "https://git-scm.com/download/win" }
        ScriptBlock = {
            winget install -e --id Git.Git
            Write-Host "> $((winget list -e --id Git.Git).Split([Environment]::NewLine) | Select-Object -Last 1)"
        }
    }
    @{
        Descriptor  = "Installing gsudo"
        Metadata    = @{ Source = "https://github.com/gerardog/gsudo" }
        ScriptBlock = {
            winget install -e --id gerardog.gsudo
            Write-Host "> $((winget list -e --id gerardog.gsudo).Split([Environment]::NewLine) | Select-Object -Last 1)"
        }
    }
    @{
        Descriptor  = "Installing Oh My Posh"
        Metadata    = @{ Source = "https://github.com/jandedobbeleer/oh-my-posh" }
        ScriptBlock = {
            winget install -e --id XP8K0HKJFRXGCK # via Microsoft Store
            Write-Host "> $((winget list -e --id XP8K0HKJFRXGCK).Split([Environment]::NewLine) | Select-Object -Last 1)"
        }
    }
    @{
        Descriptor  = "Installing Terminal-Icons"
        Metadata    = @{ Source = "https://github.com/devblackops/Terminal-Icons" }
        ScriptBlock = {
            WithPwsh {
                Install-Module Terminal-Icons -Repository PSGallery
                Write-Host "> Version: $((Get-InstalledModule Terminal-Icons).Version)"
            }
        }
    }
    @{
        Descriptor  = "Installing posh-git"
        Metadata    = @{ Source = "https://github.com/dahlbyk/posh-git" }
        ScriptBlock = {
            WithPwsh {
                Install-Module posh-git
                Write-Host "> Version: $((Get-InstalledModule posh-git).Version)"
            }
        }
    }
    @{
        Descriptor  = "Installing z"
        Metadata    = @{ Source = "https://github.com/badmotorfinger/z" }
        ScriptBlock = {
            WithPwsh {
                Install-Module z
                Write-Host "> Version: $((Get-InstalledModule z).Version)"
            }
        }
    }
    @{
        Descriptor  = "Configuring PowerShell profile"
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

            $PwshProfile = WithPwsh { $PROFILE }

            if ((Test-Path $PwshProfile) -and (Get-Content $PwshProfile).Trim().Length -gt 0) {
                Write-Warning "Your PowerShell profile is not empty.`n$PwshProfile"
                if ((Read-Host "> Overwrite? [yN]").Trim().ToLower() -ne 'y') {
                    Write-Warning "Skipping step."
                    return
                }
            }

            $ProfileScript.Trim() -replace "                ", "" | Out-File $PwshProfile
            Write-Host "> Updated: $PwshProfile"
        }
    }
    @{
        Descriptor  = "Installing font: $NerdFont Nerd Font"
        Metadata    = @{ Source = "https://github.com/ryanoasis/nerd-fonts" }
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

            $TempDirectory = [System.IO.Path]::GetTempPath()

            $TempFile = New-Item -ItemType File -Path $TempDirectory -Name "$(New-Guid).zip"
            Invoke-RestMethod $Asset.browser_download_url -OutFile $TempFile

            # extract fonts archive
            $TempFolder = New-Item -ItemType Directory -Path $TempDirectory -Name (New-Guid)
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
