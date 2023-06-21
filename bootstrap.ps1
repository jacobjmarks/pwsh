<#
    .SYNOPSIS
        Windows PowerShell Core bootstrapper
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]

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

function With-Pwsh {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock] $Command
    )

    if ($PSEdition -ne 'Core') {
        pwsh -NoProfile -c $Command
    }
    else {
        Invoke-Command -NoNewScope $Command
    }
}

function RefreshPath {
    $env:Path = @(
        [System.Environment]::GetEnvironmentVariable("Path", "Machine")
        [System.Environment]::GetEnvironmentVariable("Path", "User")
    ) -match '.' -join ';'
}

$Steps = @(
    @{
        Descriptor  = "Installing Windows Terminal"
        ScriptBlock = {
            winget install -e --id Microsoft.WindowsTerminal
        }
    }
    @{
        Descriptor  = "Installing PowerShell"
        ScriptBlock = {
            winget install -e --id Microsoft.PowerShell
            RefreshPath
        }
    }
    @{
        Descriptor  = "Installing Git"
        ScriptBlock = {
            winget install -e --id Git.Git
        }
    }
    @{
        Descriptor  = "Installing gsudo"
        ScriptBlock = {
            winget install -e --id gerardog.gsudo
        }
    }
    @{
        Descriptor  = "Installing Oh My Posh"
        ScriptBlock = {
            winget install -e --id XP8K0HKJFRXGCK # via Microsoft Store
        }
    }
    @{
        Descriptor  = "Installing Terminal-Icons"
        ScriptBlock = {
            With-Pwsh { Install-Module Terminal-Icons }
        }
    }
    @{
        Descriptor  = "Installing posh-git"
        ScriptBlock = {
            With-Pwsh { Install-Module posh-git }
        }
    }
    @{
        Descriptor  = "Installing z"
        ScriptBlock = {
            With-Pwsh { Install-Module z }
        }
    }
    @{
        Descriptor  = "Configuring PowerShell profile"
        ScriptBlock = {
            $PoshThemesPath = [System.Environment]::GetEnvironmentVariable("POSH_THEMES_PATH", "User")
            $ThemeFile = [System.IO.FileInfo](Join-Path $PoshThemesPath "$Theme.omp.json")
            if (-Not $ThemeFile.Exists) {
                $AvailableThemes = Get-ChildItem $PoshThemesPath -Filter "*.omp.json" `
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

            $PwshProfile = With-Pwsh { $PROFILE }

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
