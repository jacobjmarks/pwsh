# pwsh

Windows PowerShell Bootstrapper

``` pwsh
irm "https://raw.githubusercontent.com/jacobjmarks/pwsh/main/bootstrap.ps1" | iex
```

![Preview](preview.png)

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
