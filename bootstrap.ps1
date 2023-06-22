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

function Get-FontFamilyName {
    param (
        [Parameter(Mandatory)]
        [string] $File
    )

    Add-Type -AssemblyName System.Drawing
    $FontCollection = [System.Drawing.Text.PrivateFontCollection]::new()
    try {
        $FontCollection.AddFontFile($File)
        $FontCollection.Families.Name
    }
    finally {
        $FontCollection.Dispose()
    }
}

function Update-TerminalSettings {
    Write-Host "> Checking Windows Terminal settings ..."

    $TerminalSettingsFile = [System.IO.FileInfo](Join-Path $env:LOCALAPPDATA "Packages/Microsoft.WindowsTerminal_8wekyb3d8bbwe/LocalState/settings.json")
    $TerminalSettings = Get-Content $TerminalSettingsFile.FullName | ConvertFrom-Json

    $PowerShellCoreProfile = $TerminalSettings.profiles.list | Where-Object { $_.source -eq "Windows.Terminal.PowershellCore" }

    $Manifest = @(
        @{
            SettingPath  = "defaultProfile"
            CurrentValue = $TerminalSettings.defaultProfile
            DesiredValue = $PowerShellCoreProfile.guid
        }
        @{
            SettingPath  = "profiles.defaults.colorScheme"
            CurrentValue = $TerminalSettings.profiles.defaults.colorScheme
            DesiredValue = "One Half Dark"
        }
        @{
            SettingPath  = "profiles.defaults.font.face"
            CurrentValue = $TerminalSettings.profiles.defaults.font.face
            DesiredValue = "$NerdFontFamilyName"
        }
        @{
            SettingPath  = "profiles.defaults.opacity"
            CurrentValue = $TerminalSettings.profiles.defaults.opacity
            DesiredValue = 75
        }
        @{
            SettingPath  = "profiles.defaults.useAcrylic"
            CurrentValue = $TerminalSettings.profiles.defaults.useAcrylic
            DesiredValue = $true
        }
        @{
            SettingPath  = "useAcrylicInTabRow"
            CurrentValue = $TerminalSettings.useAcrylicInTabRow
            DesiredValue = $true
        }
    )

    $DiscrepantSettings = $Manifest | Where-Object { $_.CurrentValue -ne $_.DesiredValue }
    if (-Not $DiscrepantSettings) {
        return
    }

    foreach ($Entry in $DiscrepantSettings) {
        Write-Host "    $($Entry.SettingPath) = $(if($Entry.CurrentValue){$Entry.CurrentValue}else{"NULL"}) -> $($Entry.DesiredValue)"
    }

    $ShouldProceed = Read-Host "> Update Windows Terminal settings as above? (recommended) [yN]"
    if ($ShouldProceed.Trim().ToLower() -ne 'y') {
        return
    }

    foreach ($Entry in $DiscrepantSettings) {
        $Segments = $Entry.SettingPath -split "\."

        # hydrate
        $Ptr = $TerminalSettings
        foreach ($Segment in ($Segments | Select-Object -SkipLast 1)) {
            if (-Not $Ptr."$Segment") {
                $Ptr | Add-Member -NotePropertyName $Segment -NotePropertyValue ([PSCustomObject]@{})
            }
            $Ptr = $Ptr."$Segment"
        }

        # set desired value
        $Ptr | Add-Member -NotePropertyName ($Segments | Select-Object -Last 1) -NotePropertyValue $Entry.DesiredValue
    }

    $TerminalSettings | ConvertTo-Json -Depth 100 | Out-File $TerminalSettingsFile -Encoding utf8
    Write-Host "> Updated: $TerminalSettingsFile"
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

            $ProfileScript.Trim() -replace "                ", "" | Out-File $PwshProfile -Encoding utf8
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
            $script:NerdFontFamilyName = Get-FontFamilyName ($FontFiles | Select-Object -First 1).Path
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
    Update-TerminalSettings
    Write-Host "> Bootstrapping Complete!"
}
