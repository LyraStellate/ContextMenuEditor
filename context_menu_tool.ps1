param(
    [Parameter(Mandatory = $false)]
    [string]$AppPath,
    [Parameter(Mandatory = $false)]
    [ValidateSet("A", "R", "a", "r")]
    [string]$Action
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-KeyNameFromPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText
    )

    $normalized = [System.IO.Path]::GetFullPath($PathText).ToLowerInvariant()
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalized)
        $hashBytes = $sha1.ComputeHash($bytes)
        $hash = [System.BitConverter]::ToString($hashBytes).Replace("-", "")
        return "OpenWithCustomApp_$hash"
    }
    finally {
        $sha1.Dispose()
    }
}

function Ensure-RegistryKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
}

function Set-MenuEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,
        [Parameter(Mandatory = $true)]
        [string]$MenuName,
        [Parameter(Mandatory = $true)]
        [string]$ExePath,
        [Parameter(Mandatory = $true)]
        [string]$CommandArg
    )

    Ensure-RegistryKey -Path $BasePath
    Set-ItemProperty -Path $BasePath -Name "(default)" -Value $MenuName
    Set-ItemProperty -Path $BasePath -Name "Icon" -Value $ExePath

    $commandPath = Join-Path $BasePath "command"
    Ensure-RegistryKey -Path $commandPath
    $commandText = "`"$ExePath`" `"$CommandArg`""
    Set-ItemProperty -Path $commandPath -Name "(default)" -Value $commandText
}

function Add-ContextMenu {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath,
        [Parameter(Mandatory = $true)]
        [string]$KeyName
    )

    $appName = [System.IO.Path]::GetFileNameWithoutExtension($ExePath)
    $menuName = "$appName $([char]0x3067)$([char]0x958B)$([char]0x304F)"

    $targets = @(
        @{ Path = "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\shell\$KeyName"; Arg = "%V" },
        @{ Path = "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\Background\shell\$KeyName"; Arg = "%V" },
        @{ Path = "Registry::HKEY_CURRENT_USER\Software\Classes\*\shell\$KeyName"; Arg = "%1" }
    )

    foreach ($target in $targets) {
        Set-MenuEntry -BasePath $target.Path -MenuName $menuName -ExePath $ExePath -CommandArg $target.Arg
    }

    Write-Host ""
    Write-Host "Added to context menu."
    Write-Host "Menu label: $menuName"
    Write-Host "Use the same exe path with action [R] to remove it."
}

function Remove-ContextMenu {
    param(
        [Parameter(Mandatory = $true)]
        [string]$KeyName
    )

    $targets = @(
        "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\shell\$KeyName",
        "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\Background\shell\$KeyName",
        "Registry::HKEY_CURRENT_USER\Software\Classes\*\shell\$KeyName"
    )

    $removed = $false
    foreach ($target in $targets) {
        if (Test-Path -LiteralPath $target) {
            Remove-Item -LiteralPath $target -Recurse -Force
            $removed = $true
        }
    }

    Write-Host ""
    if ($removed) {
        Write-Host "Removed from context menu."
    }
    else {
        Write-Host "No matching entry was found."
    }
}

if ([string]::IsNullOrWhiteSpace($AppPath)) {
    $AppPath = Read-Host "Enter full path to target EXE"
}

if ([string]::IsNullOrWhiteSpace($AppPath)) {
    throw "EXE path is empty."
}

if (-not (Test-Path -LiteralPath $AppPath -PathType Leaf)) {
    throw "EXE not found: $AppPath"
}

$resolvedPath = (Resolve-Path -LiteralPath $AppPath).Path
$keyName = Get-KeyNameFromPath -PathText $resolvedPath

if ([string]::IsNullOrWhiteSpace($Action)) {
    Write-Host ""
    Write-Host "[A] Add"
    Write-Host "[R] Remove"
    $action = (Read-Host "Choose action").Trim().ToUpperInvariant()
}
else {
    $action = $Action.Trim().ToUpperInvariant()
}

switch ($action) {
    "A" { Add-ContextMenu -ExePath $resolvedPath -KeyName $keyName }
    "R" { Remove-ContextMenu -KeyName $keyName }
    default { throw "Enter A or R." }
}
