$ErrorActionPreference = "Stop"

$Steps = @(
    @{
        Descriptor  = "Installing Windows Terminal"
        Metadata    = @{ Source = "https://github.com/microsoft/terminal" }
        ScriptBlock = {
            winget install -e --id Microsoft.WindowsTerminal
        }
    }
    @{
        Descriptor  = "Installing PowerShell"
        Metadata    = @{ Source = "https://github.com/PowerShell/PowerShell" }
        ScriptBlock = {
            winget install -e --id Microsoft.PowerShell
        }
    }
    @{
        Descriptor  = "Installing Git"
        Metadata    = @{ Source = "https://git-scm.com/download/win" }
        ScriptBlock = {
            winget install -e --id Git.Git
        }
    }
    @{
        Descriptor  = "Installing gsudo"
        Metadata    = @{ Source = "https://github.com/gerardog/gsudo" }
        ScriptBlock = {
            winget install -e --id gerardog.gsudo
        }
    }
    @{
        Descriptor  = "Installing Oh My Posh"
        Metadata    = @{ Source = "https://github.com/jandedobbeleer/oh-my-posh" }
        ScriptBlock = {
            winget install -e --id XP8K0HKJFRXGCK # via Microsoft Store
        }
    }
    @{
        Descriptor  = "Installing Terminal-Icons"
        Metadata    = @{ Source = "https://github.com/devblackops/Terminal-Icons" }
        ScriptBlock = {
            Install-Module Terminal-Icons -Repository PSGallery -Force
        }
    }
    @{
        Descriptor  = "Installing posh-git"
        Metadata    = @{ Source = "https://github.com/dahlbyk/posh-git" }
        ScriptBlock = {
            Install-Module posh-git -Force
        }
    }
    @{
        Descriptor  = "Installing z"
        Metadata    = @{ Source = "https://github.com/badmotorfinger/z" }
        ScriptBlock = {
            Install-Module z -Force
        }
    }
    @{
        Descriptor  = "Configuring PowerShell profile"
        ScriptBlock = {
            $ProfileScript = {
                Import-Module posh-git
                Import-Module Terminal-Icons
        
                oh-my-posh init pwsh --config "$env:POSH_THEMES_PATH\paradox.omp.json" | Invoke-Expression
        
                Set-PSReadLineOption -PredictionSource History
                Set-PSReadLineOption -PredictionViewStyle ListView
        
                Set-Alias grep findstr
                Set-Alias which gcm
            }

            if ((Test-Path $PROFILE) -and (Get-Content $PROFILE).Trim().Length -gt 0) {
                Write-Warning "Your PowerShell profile is not empty!`n$PROFILE"
                if ((Read-Host "> Overwrite? [yN]").Trim().ToLower() -ne 'y') {
                    Write-Warning "Skipping step."
                    return
                }
            }

            $ProfileScript.ToString().Trim() -replace "                ", "" | Out-File $PROFILE
        }
    }
    @{
        Descriptor  = "Installing font: Hack Nerd Font"
        Metadata    = @{ Source = "https://github.com/ryanoasis/nerd-fonts" }
        ScriptBlock = {
            $TargetRelease = "tags/v2.3.3"
            $TargetAssetName = "Hack.zip"
            $FontFileFilter = "* Nerd Font Complete Windows Compatible.ttf"

            # download fonts archive
            $Release = Invoke-RestMethod "https://api.github.com/repos/ryanoasis/nerd-fonts/releases/$TargetRelease"
            $Asset = $Release.assets | Where-Object { $_.name -eq $TargetAssetName }
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
        }
    }
)

Write-Host "> Bootstrapping ..."

for ($i = 0; $i -lt $Steps.Length; $i++) {
    $Step = $Steps[$i]
    Write-Host "> [$('{0:d2}' -f ($i + 1))/$('{0:d2}' -f $Steps.Length)] $($Step.Descriptor) ..."
    Invoke-Command $Step.ScriptBlock
}

Write-Host "> Done!"
