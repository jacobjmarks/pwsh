# pwsh

Windows PowerShell Bootstrapper

![Preview](preview.png)

## Usage

``` pwsh
irm "https://raw.githubusercontent.com/jacobjmarks/pwsh/main/bootstrap.ps1" | iex
```

## Windows Terminal Configuration

| Setting                                                                                      | Value            |
| -------------------------------------------------------------------------------------------- | ---------------- |
| Startup > Default profile                                                                    | PowerShell       |
| Startup > Default terminal application                                                       | Windows Terminal |
| Appearance > Application Theme                                                               | Dark             |
| Appearance > Use acrylic material in the tab row                                             | On               |
| Profiles: Defaults > Additional settings: Appearance > Text: Color scheme                    | One Half Dark    |
| Profiles: Defaults > Additional settings: Appearance > Text: Font face                       | Hack NF          |
| Profiles: Defaults > Additional settings: Appearance > Transparency: Background opacity      | 75%              |
| Profiles: Defaults > Additional settings: Appearance > Transparency: Enable acrylic material | On               |

## Contents

The bootstrapping script installs the latest version of the following components:

### Applications

- [Windows Terminal](https://github.com/microsoft/terminal)
- [PowerShell Core](https://github.com/PowerShell/PowerShell)
- [Git](https://git-scm.com/download/win)

### Modules / Tooling

- [gsudo](https://github.com/gerardog/gsudo)\
  A `sudo` equivalent for Windows.

- [Oh My Posh](https://github.com/jandedobbeleer/oh-my-posh) (via Microsoft Store)\
  *A prompt theme engine for any shell.*

- [Terminal-Icons](https://github.com/devblackops/Terminal-Icons)\
  Shows file and folder icons in the terminal.

- [posh-git](https://github.com/dahlbyk/posh-git)\
  Provides Git status information to the prompt.

- [z](https://github.com/badmotorfinger/z)\
  Directory jumper.

### Fonts

- [Hack Nerd Font (v2.3.3)](https://github.com/ryanoasis/nerd-fonts)

## Optional Components

While not installed via the bootstrapping script, you may find some additional useful components below.

### [7-Zip](https://www.7-zip.org/)

_7-Zip is a file archiver with a high compression ratio._

``` pwsh
winget install -e --id 7zip.7zip
```

### [Azure CLI](https://github.com/Azure/azure-cli)

_The Azure command-line interface (Azure CLI) is a set of commands used to create and manage Azure resources._

``` pwsh
winget install -e --id Microsoft.AzureCLI
```

### [Azure Functions Core Tools](https://github.com/Azure/azure-functions-core-tools)

_The Azure Functions Core Tools provide a local development experience for creating, developing, testing, running, and debugging Azure Functions._

``` pwsh
winget install -e --id Microsoft.AzureFunctionsCoreTools
```

### [NVM for Windows](https://github.com/coreybutler/nvm-windows)

_The Microsoft/npm/Google recommended Node.js version manager for Windows._

``` pwsh
winget install -e --id CoreyButler.NVMforWindows
```
