<#
    .SYNOPSIS
        Windows PowerShell Core bootstrapper
#>

[CmdletBinding()]
param(
    # Oh My Posh theme to configure for use
    [string] $Theme = "paradox",

    # Nerd Font to install
    [string] $NerdFont = "Hack",

    # Skip installation of fonts
    [switch] $NoFonts = $false
)

$ErrorActionPreference = "Stop"

$ROOT_INVOCATION = $MyInvocation
$NON_TERMINATING_ERROR_COUNT = 0
$TERMINAL_SETTINGS_FILE_PATH = (Join-Path $env:LOCALAPPDATA "Packages/Microsoft.WindowsTerminal_8wekyb3d8bbwe/LocalState/settings.json")
$NERD_FONT_FAMILY_NAME = $null

function Start-Main {
    if ($PSEdition -eq "Core" -and -not $IsWindows) {
        Write-Error "This script does not support the current operating system."
    }

    if (-not (Get-Command winget -ErrorAction Ignore)) {
        Write-Error "This script requires the Windows Package Manager (winget). Please refer to https://learn.microsoft.com/en-us/windows/package-manager/winget/#install-winget for installation options."
    }

    if ($PSEdition -ne "Core") {
        Write-Host "> Switching to PowerShell Core ..."

        if (-not (Get-Command pwsh -ErrorAction Ignore)) {
            winget install -e --id Microsoft.PowerShell --accept-source-agreements
            Update-Path
        }

        if ($ROOT_INVOCATION.InvocationName) {
            pwsh -NoExit -nop -wd $PWD -c $ROOT_INVOCATION.InvocationName @PSBoundParameters
        }
        else {
            pwsh -NoExit -nop -wd $PWD -ec (Get-Base64String $ROOT_INVOCATION.MyCommand.ToString())
        }

        return
    }

    $Apps = @(
        @{
            Name   = "Windows Terminal"
            Id     = "Microsoft.WindowsTerminal"
            Update = $false
        }
        @{
            Name = "Git"
            Id   = "Git.Git"
        }
        @{
            Name = "gsudo"
            Id   = "gerardog.gsudo"
        }
        @{
            Name = "Oh My Posh"
            Id   = "XP8K0HKJFRXGCK"
        }
    )

    $Modules = @(
        "Terminal-Icons"
        "posh-git"
        "z"
    )

    Write-Host "> Installing applications via winget"
    $Apps | ForEach-Object {
        Write-Host "> Installing $($_.Name) ..."

        $AdditionalArgs = @(
            "--accept-source-agreements"
            "--accept-package-agreements"
        )

        if ($_.Update -eq $false) {
            $AdditionalArgs += "--no-upgrade"
        }

        winget install -e --id $_.Id @AdditionalArgs

        Update-Path
    }

    Write-Host "> Installing PowerShell Core modules"
    $Modules | ForEach-Object {
        Write-Host "> Installing $_ ..."
        Install-OrUpdateModule $_
    }

    Update-PowerShellProfile
    Install-NerdFont

    if ($NON_TERMINATING_ERROR_COUNT -gt 0) {
        Write-Warning ("Bootstrapping completed with $NON_TERMINATING_ERROR_COUNT non-terminating errors. " `
                + "It is recommended you resolve these errors and re-run the script.")
    }
    else {
        Update-TerminalSettings
        Write-Host "> Bootstrapping Complete!"
    }
}

function Update-PowerShellProfile {
    Write-Host "> Configuring PowerShell Core profile ..."

    $PoshThemesPath = [System.Environment]::GetEnvironmentVariable("POSH_THEMES_PATH", "User")
    $ThemeFile = [System.IO.FileInfo](Join-Path $PoshThemesPath "$Theme.omp.json")
    if (-not $ThemeFile.Exists) {
        $AvailableThemes = Get-ChildItem $PoshThemesPath -Filter "*.omp.json" `
        | ForEach-Object { $_.Name -replace ".omp.json", "" } | Join-String -Separator ", "
        Write-Error -ErrorAction 'Continue' "Oh My Posh theme '$Theme' not found. Expected one of: $AvailableThemes"
        $Script:NON_TERMINATING_ERROR_COUNT++
        return
    }

    $ProfileScript = {
        Import-Module posh-git
        Import-Module Terminal-Icons

        oh-my-posh init pwsh --config "$env:POSH_THEMES_PATH\{{THEME_FILE_NAME}}" | Invoke-Expression

        Set-PSReadLineOption -PredictionSource History
        Set-PSReadLineOption -PredictionViewStyle ListView
    }.ToString() -replace "{{THEME_FILE_NAME}}", $ThemeFile.Name

    $PwshProfile = pwsh -nop -c { $PROFILE }

    if ((Test-Path $PwshProfile) -and (Get-Content $PwshProfile).Trim().Length -gt 0) {
        Write-Warning "Your PowerShell profile is not empty.`n$PwshProfile"
        if ($Host.UI.PromptForChoice($null, "Overwrite?", @("&Yes", "&No"), 1) -ne 0) {
            return
        }
    }

    $ProfileScript.Trim() -replace "        ", "" | Out-File $PwshProfile -Encoding utf8
    Write-Host "> Updated: $PwshProfile"
}

function Install-NerdFont {
    Write-Host "> Installing font: $NerdFont Nerd Font ..."

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
    if (-not $TargetAsset) {
        $AvailableFonts = $Release.assets | Where-Object { $_.name -like "*.zip" } `
        | ForEach-Object { $_.name -replace ".zip", "" } | Join-String -Separator ", "
        Write-Error -ErrorAction 'Continue' "Nerd Font '$NerdFont' not found. Expected one of: $AvailableFonts"
        $Script:NON_TERMINATING_ERROR_COUNT++
        return
    }

    # download fonts archive
    $TempDirectory = [System.IO.Path]::GetTempPath()
    $AssetArchive = (Join-Path $TempDirectory "nerd-fonts-asset-$($TargetAsset.id).zip")

    if (-not (Test-Path $AssetArchive) -or (Get-Item $AssetArchive).Length -ne $TargetAsset.size) {
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
    $Script:NERD_FONT_FAMILY_NAME = Get-FontFamilyName ($FontFiles | Select-Object -First 1).Path
    $FontsFolder = $ShellApplication.Namespace(0x14)
    $FontsFolder.CopyHere($FontFiles)
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
    if (Test-Path $TERMINAL_SETTINGS_FILE_PATH) {
        return
    }

    Write-Host "> Initializing Windows Terminal settings ..."

    wt --version # Starts a new Terminal process

    # Wait for the Terminal process to initialise its settings file, then terminate
    $LatestWindowsTerminalProcess = @(Get-Process -Name WindowsTerminal -ErrorAction Ignore) | Sort-Object -Property StartTime | Select-Object -Last 1
    while (-not (Test-Path $TERMINAL_SETTINGS_FILE_PATH)) {
        Start-Sleep -Milliseconds 50
    }
    $LatestWindowsTerminalProcess | Stop-Process
}

function Update-TerminalSettings {
    Initialize-TerminalSettings

    Write-Host "> Checking Windows Terminal settings ..."
    $TerminalSettings = Get-Content $TERMINAL_SETTINGS_FILE_PATH | ConvertFrom-Json

    $Manifest = @(
        @{
            SettingPath  = "profiles.defaults.colorScheme"
            CurrentValue = $TerminalSettings.profiles.defaults.colorScheme
            DesiredValue = "One Half Dark"
        }
        @{
            SettingPath  = "profiles.defaults.font.face"
            CurrentValue = $TerminalSettings.profiles.defaults.font.face
            DesiredValue = "$NERD_FONT_FAMILY_NAME"
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
            SettingPath  = "profiles.defaults.useAtlasEngine"
            CurrentValue = $TerminalSettings.profiles.defaults.useAtlasEngine
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
    if (-not $DiscrepantSettings) {
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
            if (-not $Ptr."$Segment") {
                $Ptr | Add-Member -notePropertyName $Segment -notePropertyValue ([PSCustomObject]@{})
            }
            $Ptr = $Ptr."$Segment"
        }

        # set desired value
        $Ptr | Add-Member -notePropertyName ($Segments | Select-Object -Last 1) -notePropertyValue $Entry.DesiredValue -Force
    }

    $TerminalSettings | ConvertTo-Json -Depth 100 | Out-File $TERMINAL_SETTINGS_FILE_PATH -Encoding utf8
    Write-Host "> Updated: $TERMINAL_SETTINGS_FILE_PATH"
}

function Update-Path {
    $env:Path = @(
        [System.Environment]::GetEnvironmentVariable("Path", "Machine")
        [System.Environment]::GetEnvironmentVariable("Path", "User")
    ) -match '.' -join ';'
}

function Get-Base64String {
    param (
        [Parameter(Mandatory)]
        [string] $Value
    )

    $Bytes = [System.Text.Encoding]::Unicode.GetBytes($Value)
    [Convert]::ToBase64String($Bytes)
}

function Install-OrUpdateModule {
    param (
        [Parameter(Mandatory)]
        [string] $Name
    )

    if (-not (Get-InstalledModule $Name -ErrorAction Ignore)) {
        Install-Module $Name -Force
    }
    else {
        Update-Module $Name
    }
}

Start-Main
