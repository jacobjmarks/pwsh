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

$rootInvocation = $MyInvocation
$nonTerminatingErrorCount = 0
$terminalSettingsFilePath = (Join-Path $env:LOCALAPPDATA "Packages/Microsoft.WindowsTerminal_8wekyb3d8bbwe/LocalState/settings.json")
$nerdFontFamilyName = $null

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

        if ($rootInvocation.InvocationName) {
            pwsh -NoExit -nop -wd $PWD -c $rootInvocation.InvocationName @PSBoundParameters
        }
        else {
            pwsh -NoExit -nop -wd $PWD -ec (Get-Base64String $rootInvocation.MyCommand.ToString())
        }

        return
    }

    $apps = @(
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

    $modules = @(
        "Terminal-Icons"
        "posh-git"
        "z"
    )

    Write-Host "> Installing applications via winget"
    $apps | ForEach-Object {
        Write-Host "> Installing $($_.Name) ..."

        $additionalArgs = @(
            "--accept-source-agreements"
            "--accept-package-agreements"
        )

        if ($_.Update -eq $false) {
            $additionalArgs += "--no-upgrade"
        }

        winget install -e --id $_.Id @additionalArgs

        Update-Path
    }

    Write-Host "> Installing PowerShell Core modules"
    $modules | ForEach-Object {
        Write-Host "> Installing $_ ..."
        Install-OrUpdateModule $_
    }

    Update-PowerShellProfile
    Install-NerdFont

    if ($nonTerminatingErrorCount -gt 0) {
        Write-Warning ("Bootstrapping completed with $nonTerminatingErrorCount non-terminating errors. " `
                + "It is recommended you resolve these errors and re-run the script.")
    }
    else {
        Update-TerminalSettings
        Write-Host "> Bootstrapping Complete!"
    }
}

function Update-PowerShellProfile {
    Write-Host "> Configuring PowerShell Core profile ..."

    $poshThemesPath = [System.Environment]::GetEnvironmentVariable("POSH_THEMES_PATH", "User")
    $themeFile = [System.IO.FileInfo](Join-Path $poshThemesPath "$Theme.omp.json")
    if (-not $themeFile.Exists) {
        $availableThemes = Get-ChildItem $poshThemesPath -Filter "*.omp.json" `
        | ForEach-Object { $_.Name -replace ".omp.json", "" } | Join-String -Separator ", "
        Write-Error -ErrorAction 'Continue' "Oh My Posh theme '$Theme' not found. Expected one of: $availableThemes"
        $Script:nonTerminatingErrorCount++
        return
    }

    $profileScript = {
        Import-Module posh-git
        Import-Module Terminal-Icons

        oh-my-posh init pwsh --config "$env:POSH_THEMES_PATH\{{THEME_FILE_NAME}}" | Invoke-Expression
        $env:POSH_GIT_ENABLED = $true

        Set-PSReadLineOption -PredictionSource History
        Set-PSReadLineOption -PredictionViewStyle ListView
    }.ToString() -replace "{{THEME_FILE_NAME}}", $themeFile.Name

    $pwshProfile = pwsh -nop -c { $PROFILE }

    if ((Test-Path $pwshProfile) -and (Get-Content $pwshProfile).Trim().Length -gt 0) {
        Write-Warning "Your PowerShell profile is not empty.`n$pwshProfile"
        if ($Host.UI.PromptForChoice($null, "Overwrite?", @("&Yes", "&No"), 1) -ne 0) {
            return
        }
    }

    $profileScript.Trim() -replace "        ", "" | Out-File $pwshProfile -Encoding utf8
    Write-Host "> Updated: $pwshProfile"
}

function Install-NerdFont {
    Write-Host "> Installing font: $NerdFont Nerd Font ..."

    if ($NoFonts) {
        Write-Warning "Skipping step."
        return
    }

    $targetRelease = "latest"
    $targetAssetName = "$NerdFont.zip"
    $fontFileFilter = "*NerdFont-*.ttf" # exclude mono/propo fonts

    # get GitHub asset
    $release = Invoke-RestMethod "https://api.github.com/repos/ryanoasis/nerd-fonts/releases/$targetRelease"
    $targetAsset = $release.assets | Where-Object { $_.name -eq $targetAssetName }
    if (-not $targetAsset) {
        $availableFonts = $release.assets | Where-Object { $_.name -like "*.zip" } `
        | ForEach-Object { $_.name -replace ".zip", "" } | Join-String -Separator ", "
        Write-Error -ErrorAction 'Continue' "Nerd Font '$NerdFont' not found. Expected one of: $availableFonts"
        $Script:nonTerminatingErrorCount++
        return
    }

    # download fonts archive
    $tempDirectory = [System.IO.Path]::GetTempPath()
    $assetArchive = (Join-Path $tempDirectory "nerd-fonts-asset-$($targetAsset.id).zip")

    if (-not (Test-Path $assetArchive) -or (Get-Item $assetArchive).Length -ne $targetAsset.size) {
        New-Item -ItemType File -Path $assetArchive | Out-Null
        Invoke-RestMethod $targetAsset.browser_download_url -OutFile $assetArchive
    }

    # extract fonts archive
    $extractFolder = (Join-Path $tempDirectory ([System.IO.Path]::GetFileNameWithoutExtension($assetArchive)))
    if (Test-Path $extractFolder) { Remove-Item -Force -Recurse $extractFolder }
    New-Item -ItemType Directory -Path $extractFolder | Out-Null
    Expand-Archive $assetArchive $extractFolder

    # install fonts
    $shellApplication = New-Object -ComObject Shell.Application
    $fontFiles = $shellApplication.Namespace($extractFolder).Items()
    $fontFiles.Filter(0x40, $fontFileFilter)
    if ($fontFiles.Count -eq 0) {
        throw "No files found within asset '$targetAssetName' matching the filter '$fontFileFilter'"
    }
    $fontFiles | ForEach-Object { Write-Host ">   $($_.Name)" }
    $Script:nerdFontFamilyName = Get-FontFamilyName ($fontFiles | Select-Object -First 1).Path
    $fontsFolder = $shellApplication.Namespace(0x14)
    $fontsFolder.CopyHere($fontFiles)
}

function Get-FontFamilyName {
    param (
        [Parameter(Mandatory)]
        [string] $File
    )

    Add-Type -AssemblyName System.Drawing
    $fontCollection = [System.Drawing.Text.PrivateFontCollection]::new()
    try {
        $fontCollection.AddFontFile($File)
        $fontCollection.Families.Name
    }
    finally {
        $fontCollection.Dispose()
    }
}

function Initialize-TerminalSettings {
    if (Test-Path $terminalSettingsFilePath) {
        return
    }

    Write-Host "> Initializing Windows Terminal settings ..."

    wt --version # Starts a new Terminal process

    # Wait for the Terminal process to initialise its settings file, then terminate
    $latestWindowsTerminalProcess = @(Get-Process -Name WindowsTerminal -ErrorAction Ignore) | Sort-Object -Property StartTime | Select-Object -Last 1
    while (-not (Test-Path $terminalSettingsFilePath)) {
        Start-Sleep -Milliseconds 50
    }
    $latestWindowsTerminalProcess | Stop-Process
}

function Update-TerminalSettings {
    Initialize-TerminalSettings

    Write-Host "> Checking Windows Terminal settings ..."
    $terminalSettings = Get-Content $terminalSettingsFilePath | ConvertFrom-Json

    $manifest = @(
        @{
            SettingPath  = "profiles.defaults.colorScheme"
            CurrentValue = $terminalSettings.profiles.defaults.colorScheme
            DesiredValue = "One Half Dark"
        }
        @{
            SettingPath  = "profiles.defaults.font.face"
            CurrentValue = $terminalSettings.profiles.defaults.font.face
            DesiredValue = "$nerdFontFamilyName"
        }
        @{
            SettingPath  = "profiles.defaults.opacity"
            CurrentValue = $terminalSettings.profiles.defaults.opacity
            DesiredValue = 75
        }
        @{
            SettingPath  = "profiles.defaults.useAcrylic"
            CurrentValue = $terminalSettings.profiles.defaults.useAcrylic
            DesiredValue = $true
        }
        @{
            SettingPath  = "profiles.defaults.useAtlasEngine"
            CurrentValue = $terminalSettings.profiles.defaults.useAtlasEngine
            DesiredValue = $true
        }
        @{
            SettingPath  = "useAcrylicInTabRow"
            CurrentValue = $terminalSettings.useAcrylicInTabRow
            DesiredValue = $true
        }
    )

    $coreTerminalProfile = $terminalSettings.profiles.list | Where-Object { $_.source -eq "Windows.Terminal.PowershellCore" } | Select-Object -First 1
    if ($coreTerminalProfile) {
        $manifest += @{
            SettingPath  = "defaultProfile"
            CurrentValue = $terminalSettings.defaultProfile
            DesiredValue = $coreTerminalProfile.guid
        }
    }
    else {
        Write-Warning "Skipping Terminal setting 'defaultProfile'; Could not find PowerShell Core profile"
    }

    $manifest = ($manifest | Sort-Object -Property SettingPath)

    $discrepantSettings = $manifest | Where-Object { $_.CurrentValue -ne $_.DesiredValue }
    if (-not $discrepantSettings) {
        return
    }

    foreach ($entry in $discrepantSettings) {
        Write-Host "    $($entry.SettingPath) = $(if($entry.CurrentValue){$entry.CurrentValue}else{"NULL"}) -> $($entry.DesiredValue)"
    }

    if ($Host.UI.PromptForChoice($null, "Update Windows Terminal settings as above? (recommended)", @("&Yes", "&No"), 1) -ne 0) {
        return
    }

    foreach ($entry in $discrepantSettings) {
        $segments = $entry.SettingPath -split "\."

        # hydrate
        $ptr = $terminalSettings
        foreach ($segment in ($segments | Select-Object -SkipLast 1)) {
            if (-not $ptr."$segment") {
                $ptr | Add-Member -NotePropertyName $segment -NotePropertyValue ([PSCustomObject]@{})
            }
            $ptr = $ptr."$segment"
        }

        # set desired value
        $ptr | Add-Member -NotePropertyName ($segments | Select-Object -Last 1) -NotePropertyValue $entry.DesiredValue -Force
    }

    $terminalSettings | ConvertTo-Json -Depth 100 | Out-File $terminalSettingsFilePath -Encoding utf8
    Write-Host "> Updated: $terminalSettingsFilePath"
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

    $bytes = [System.Text.Encoding]::Unicode.GetBytes($Value)
    [Convert]::ToBase64String($bytes)
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
        Update-Module $Name -Force
    }
}

Start-Main
