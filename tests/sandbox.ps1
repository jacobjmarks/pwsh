$ErrorActionPreference = "Stop"

function Start-Main {
    $sandbox = Get-Process -Name "WindowsSandbox" -ErrorAction Ignore
    if ($sandbox) {
        Write-Error "Please close your existing Windows Sandbox instance."
    }

    $tempFolder = Join-Path $PSScriptRoot "__sandbox__"
    if (Test-Path $tempFolder) { Remove-Item -Force -Recurse $tempFolder }
    New-Item -ItemType Directory $tempFolder | Out-Null

    $mappedFolderName = "content"

    $hostMappedFolder = Join-Path $tempFolder $mappedFolderName
    New-Item -ItemType Directory $hostMappedFolder | Out-Null

    Copy-Item (Resolve-Path "$PSScriptRoot/../bootstrap.ps1") $hostMappedFolder

    $sandboxScriptName = "main.ps1"
    $sandboxScriptContent = {
        function Wait-ForInternet {
            while (-not (Test-Connection "google.com" -Count 1 -ErrorAction Ignore)) {
                Start-Sleep -Seconds 1
            }
        }

        function Install-WinGet {
            Push-Location ([System.IO.Path]::GetTempPath())
            $ProgressPreference = "SilentlyContinue"
            try {
                $latestWingetMsixBundleUri = $(Invoke-RestMethod "https://api.github.com/repos/microsoft/winget-cli/releases/latest").assets.browser_download_url | Where-Object { $_.EndsWith(".msixbundle") }
                $latestWingetMsixBundle = $latestWingetMsixBundleUri.Split("/")[-1]
                Write-Information "Downloading winget to artifacts directory..."
                Invoke-WebRequest -Uri $latestWingetMsixBundleUri -OutFile "./$latestWingetMsixBundle"
                Invoke-WebRequest -Uri https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx -OutFile Microsoft.VCLibs.x64.14.00.Desktop.appx
                Add-AppxPackage Microsoft.VCLibs.x64.14.00.Desktop.appx
                Add-AppxPackage $latestWingetMsixBundle

            }
            finally {
                $ProgressPreference = "Continue"
                Pop-Location
            }
        }

        Write-Host "[test] Initialising"
        Write-Host "[test] Waiting for internet ..."
        Wait-ForInternet
        Write-Host "[test] Installing winget ..."
        Install-WinGet
        Write-Host "[test] Updating winget settings ..."
        $wingetSettings = @{ network = @{ downloader = "wininet" } } | ConvertTo-Json
        $wingetSettings | Out-File -Encoding utf8 "$HOME\AppData\Local\Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\settings.json"
        Write-Host "[test] Initialisation complete"
        Write-Host "[test] Starting test run ..."

        ./bootstrap.ps1
    }.ToString().Trim() -replace "(?m)^        "
    $sandboxScriptContent | Out-File (Join-Path $hostMappedFolder $sandboxScriptName)

    $sandboxDesktop = "C:\Users\WDAGUtilityAccount\Desktop"
    $sandboxMappedFolder = Join-Path $sandboxDesktop $mappedFolderName

    $wsb = Join-Path $tempFolder "Configuration.wsb"
    Out-File $wsb -InputObject @"
<Configuration>
  <MappedFolders>
    <MappedFolder>
      <HostFolder>$hostMappedFolder</HostFolder>
      <SandboxFolder>$sandboxMappedFolder</SandboxFolder>
      <ReadOnly>true</ReadOnly>
    </MappedFolder>
  </MappedFolders>
  <LogonCommand>
    <Command>PowerShell Start-Process PowerShell -WindowStyle Maximized -WorkingDirectory '$sandboxMappedFolder' -ArgumentList '-ExecutionPolicy Bypass -NoExit -NoLogo -File $sandboxScriptName'</Command>
  </LogonCommand>
</Configuration>
"@

    Invoke-Item $wsb
}

Push-Location $PSScriptRoot
try {
    Start-Main
}
finally {
    Pop-Location
}
