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

if (-Not (Get-Command winget -ErrorAction Ignore)) {
    Write-Error "This script requires the Windows Package Manager (winget). Please refer to https://learn.microsoft.com/en-us/windows/package-manager/winget/#install-winget for installation options."
}

$NonTerminatingErrorCount = 0
$WindowsTerminalSettingsFile = (Join-Path $env:LOCALAPPDATA "Packages/Microsoft.WindowsTerminal_8wekyb3d8bbwe/LocalState/settings.json")

function Update-Path {
    $env:Path = @(
        [System.Environment]::GetEnvironmentVariable("Path", "Machine")
        [System.Environment]::GetEnvironmentVariable("Path", "User")
    ) -match '.' -join ';'
}

function Invoke-WithCore {
    param (
        [Parameter(Mandatory)]
        [ScriptBlock] $ScriptBlock
    )

    if ($PSEdition -ne 'Core') {
        pwsh -NoProfile -c $ScriptBlock
        if ($LASTEXITCODE -ne 0) { exit }
    }
    else {
        Invoke-Command -NoNewScope $ScriptBlock
    }
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

function Initialize-TerminalSettings {
    if (Test-Path $WindowsTerminalSettingsFile) {
        return
    }

    Write-Host "> Initializing Windows Terminal settings ..."

    wt --version # Starts a new Terminal process

    # Wait for the Terminal process to initialise its settings file, then terminate
    $LatestWindowsTerminalProcess = @(Get-Process -Name WindowsTerminal -ErrorAction Ignore) | Sort-Object -Property StartTime | Select-Object -Last 1
    while (-Not (Test-Path $WindowsTerminalSettingsFile)) {
        Start-Sleep -Milliseconds 50
    }
    $LatestWindowsTerminalProcess | Stop-Process
}

function Update-TerminalSettings {
    Initialize-TerminalSettings

    Write-Host "> Checking Windows Terminal settings ..."
    $TerminalSettings = Get-Content $WindowsTerminalSettingsFile | ConvertFrom-Json

    $Manifest = @(
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

    $CoreTerminalProfile = $TerminalSettings.profiles.list | Where-Object { $_.source -eq "Windows.Terminal.PowershellCore" } | Select-Object -First 1
    if ($CoreTerminalProfile) {
        $Manifest += @{
            SettingPath  = "defaultProfile"
            CurrentValue = $TerminalSettings.defaultProfile
            DesiredValue = $CoreTerminalProfile.guid
        }
    }
    else {
        Write-Warning "Skipping Terminal setting 'defaultProfile'; Could not find PowerShell Core profile"
    }

    $Manifest = ($Manifest | Sort-Object -Property SettingPath)

    $DiscrepantSettings = $Manifest | Where-Object { $_.CurrentValue -ne $_.DesiredValue }
    if (-Not $DiscrepantSettings) {
        return
    }

    foreach ($Entry in $DiscrepantSettings) {
        Write-Host "    $($Entry.SettingPath) = $(if($Entry.CurrentValue){$Entry.CurrentValue}else{"NULL"}) -> $($Entry.DesiredValue)"
    }

    if ($Host.UI.PromptForChoice($null, "Update Windows Terminal settings as above? (recommended)", @("&Yes", "&No"), 1) -ne 0) {
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
        $Ptr | Add-Member -NotePropertyName ($Segments | Select-Object -Last 1) -NotePropertyValue $Entry.DesiredValue -Force
    }

    $TerminalSettings | ConvertTo-Json -Depth 100 | Out-File $WindowsTerminalSettingsFile -Encoding utf8
    Write-Host "> Updated: $WindowsTerminalSettingsFile"
}

$Steps = @(
    @{
        Descriptor  = "Installing Windows Terminal"
        ScriptBlock = {
            winget install -e --id Microsoft.WindowsTerminal
            Update-Path
        }
    }
    @{
        Descriptor  = "Installing PowerShell"
        ScriptBlock = {
            winget install -e --id Microsoft.PowerShell
            Update-Path
        }
    }
    @{
        Descriptor  = "Installing Git"
        ScriptBlock = {
            winget install -e --id Git.Git
            Update-Path
        }
    }
    @{
        Descriptor  = "Installing gsudo"
        ScriptBlock = {
            winget install -e --id gerardog.gsudo
            Update-Path
        }
    }
    @{
        Descriptor  = "Installing Oh My Posh"
        ScriptBlock = {
            winget install -e --id XP8K0HKJFRXGCK # via Microsoft Store
            Update-Path
        }
    }
    @{
        Descriptor  = "Installing Terminal-Icons"
        ScriptBlock = {
            Invoke-WithCore { Install-Module Terminal-Icons }
        }
    }
    @{
        Descriptor  = "Installing posh-git"
        ScriptBlock = {
            Invoke-WithCore { Install-Module posh-git }
        }
    }
    @{
        Descriptor  = "Installing z"
        ScriptBlock = {
            Invoke-WithCore { Install-Module z }
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

            $PwshProfile = Invoke-WithCore { $PROFILE }

            if ((Test-Path $PwshProfile) -and (Get-Content $PwshProfile).Trim().Length -gt 0) {
                Write-Warning "Your PowerShell profile is not empty.`n$PwshProfile"
                if ($Host.UI.PromptForChoice($null, "Overwrite?", @("&Yes", "&No"), 1) -ne 0) {
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

            # get GitHub asset
            $Release = Invoke-RestMethod "https://api.github.com/repos/ryanoasis/nerd-fonts/releases/$TargetRelease"
            $TargetAsset = $Release.assets | Where-Object { $_.name -eq $TargetAssetName }
            if (-Not $TargetAsset) {
                $AvailableFonts = $Release.assets | Where-Object { $_.name -like "*.zip" } `
                | ForEach-Object { $_.name -replace ".zip", "" } | Join-String -Separator ", "
                Write-Error -ErrorAction 'Continue' "Nerd Font '$NerdFont' not found. Expected one of: $AvailableFonts"
                $Script:NonTerminatingErrorCount++
                return
            }

            # download fonts archive
            $TempDirectory = [System.IO.Path]::GetTempPath()
            $AssetArchive = (Join-Path $TempDirectory "nerd-fonts-asset-$($TargetAsset.id).zip")

            if (-Not (Test-Path $AssetArchive) -or (Get-Item $AssetArchive).Length -ne $TargetAsset.size) {
                New-Item -ItemType File -Path $AssetArchive | Out-Null
                Invoke-RestMethod $TargetAsset.browser_download_url -OutFile $AssetArchive
            }

            # extract fonts archive
            $ExtractFolder = (Join-Path $TempDirectory ([System.IO.Path]::GetFileNameWithoutExtension($AssetArchive)))
            if (Test-Path $ExtractFolder) { Remove-Item -Force -Recurse $ExtractFolder }
            New-Item -ItemType Directory -Path $ExtractFolder | Out-Null
            Expand-Archive $AssetArchive $ExtractFolder

            # install fonts
            $ShellApplication = New-Object -ComObject Shell.Application
            $FontFiles = $ShellApplication.Namespace($ExtractFolder).Items()
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
