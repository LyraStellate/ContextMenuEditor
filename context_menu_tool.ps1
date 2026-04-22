param(
    [Parameter(Mandatory = $false)]
    [string]$AppPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("A", "R", "L", "a", "r", "l")]
    [string]$Action,

    [Parameter(Mandatory = $false)]
    [string[]]$Scopes,

    [Parameter(Mandatory = $false)]
    [switch]$Gui
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:ScopeCatalog = @(
    [pscustomobject]@{
        Id            = "AllFiles"
        Label         = "Files"
        UserPath      = "Registry::HKEY_CURRENT_USER\Software\Classes\*\shell"
        MachinePath   = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\*\shell"
        CommandArg    = "%1"
        DefaultForNew = $true
    },
    [pscustomobject]@{
        Id            = "Directories"
        Label         = "Folders"
        UserPath      = "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\shell"
        MachinePath   = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\Directory\shell"
        CommandArg    = "%V"
        DefaultForNew = $true
    },
    [pscustomobject]@{
        Id            = "DirectoryBackground"
        Label         = "Folder background"
        UserPath      = "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\Background\shell"
        MachinePath   = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\Directory\Background\shell"
        CommandArg    = "%V"
        DefaultForNew = $true
    },
    [pscustomobject]@{
        Id            = "Folders"
        Label         = "Folders (broad)"
        UserPath      = "Registry::HKEY_CURRENT_USER\Software\Classes\Folder\shell"
        MachinePath   = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\Folder\shell"
        CommandArg    = "%1"
        DefaultForNew = $false
    },
    [pscustomobject]@{
        Id            = "Drives"
        Label         = "Drives"
        UserPath      = "Registry::HKEY_CURRENT_USER\Software\Classes\Drive\shell"
        MachinePath   = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\Drive\shell"
        CommandArg    = "%1"
        DefaultForNew = $false
    }
)

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

function Resolve-ExePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText
    )

    if ([string]::IsNullOrWhiteSpace($PathText)) {
        throw "EXE path is empty."
    }

    if (-not (Test-Path -LiteralPath $PathText -PathType Leaf)) {
        throw "EXE not found: $PathText"
    }

    return (Resolve-Path -LiteralPath $PathText).Path
}

function Get-RegistryDefaultValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    return (Get-Item -LiteralPath $Path).GetValue("")
}

function Get-RegistryValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    return (Get-Item -LiteralPath $Path).GetValue($Name)
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

function Set-RegistryDefaultValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Value
    )

    Ensure-RegistryKey -Path $Path
    Set-Item -Path $Path -Value $Value
}

function Set-RegistryStringValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Value
    )

    Ensure-RegistryKey -Path $Path

    if ([string]::IsNullOrWhiteSpace($Value)) {
        if ((Get-Item -LiteralPath $Path).Property -contains $Name) {
            Remove-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
        }

        return
    }

    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType String -Force | Out-Null
}

function Get-DefaultMenuLabel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath
    )

    $appName = [System.IO.Path]::GetFileNameWithoutExtension($ExePath)
    return "Open with $appName"
}

function Resolve-ScopeSelection {
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$ScopeIds
    )

    if ($null -eq $ScopeIds -or $ScopeIds.Count -eq 0) {
        return @($script:ScopeCatalog | Where-Object { $_.DefaultForNew })
    }

    $selected = @()
    foreach ($scopeId in $ScopeIds) {
        $match = $script:ScopeCatalog | Where-Object {
            $_.Id.Equals($scopeId, [System.StringComparison]::OrdinalIgnoreCase)
        } | Select-Object -First 1

        if ($null -eq $match) {
            $valid = ($script:ScopeCatalog.Id -join ", ")
            throw "Invalid scope: $scopeId`nAvailable: $valid"
        }

        $selected += $match
    }

    return $selected
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

    Set-RegistryDefaultValue -Path $BasePath -Value $MenuName
    Set-RegistryStringValue -Path $BasePath -Name "Icon" -Value $ExePath

    $commandPath = Join-Path $BasePath "command"
    $commandText = "`"$ExePath`" `"$CommandArg`""
    Set-RegistryDefaultValue -Path $commandPath -Value $commandText
}

function Add-CustomAppContextMenu {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath,

        [Parameter(Mandatory = $false)]
        [string[]]$ScopeIds,

        [Parameter(Mandatory = $false)]
        [string]$MenuLabel
    )

    $resolvedPath = Resolve-ExePath -PathText $ExePath
    $selectedScopes = Resolve-ScopeSelection -ScopeIds $ScopeIds

    if ($selectedScopes.Count -eq 0) {
        throw "Select at least one scope."
    }

    $keyName = Get-KeyNameFromPath -PathText $resolvedPath
    $finalMenuLabel = if ([string]::IsNullOrWhiteSpace($MenuLabel)) {
        Get-DefaultMenuLabel -ExePath $resolvedPath
    }
    else {
        $MenuLabel.Trim()
    }

    $createdPaths = @()

    foreach ($scope in $selectedScopes) {
        $basePath = Join-Path $scope.UserPath $keyName
        Set-MenuEntry -BasePath $basePath -MenuName $finalMenuLabel -ExePath $resolvedPath -CommandArg $scope.CommandArg
        $createdPaths += $basePath
    }

    return [pscustomobject]@{
        KeyName       = $keyName
        MenuLabel     = $finalMenuLabel
        ExePath       = $resolvedPath
        RegistryPaths = $createdPaths
    }
}

function Remove-CustomAppContextMenu {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath
    )

    $resolvedPath = Resolve-ExePath -PathText $ExePath
    $keyName = Get-KeyNameFromPath -PathText $resolvedPath
    $removedPaths = @()

    foreach ($scope in $script:ScopeCatalog) {
        $target = Join-Path $scope.UserPath $keyName
        if (Test-Path -LiteralPath $target) {
            Remove-Item -LiteralPath $target -Recurse -Force
            $removedPaths += $target
        }
    }

    return [pscustomobject]@{
        KeyName       = $keyName
        ExePath       = $resolvedPath
        RemovedPaths  = $removedPaths
        WasRemoved    = ($removedPaths.Count -gt 0)
    }
}

function Remove-ContextMenuEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath
    )

    if (-not (Test-Path -LiteralPath $RegistryPath)) {
        return $false
    }

    Remove-Item -LiteralPath $RegistryPath -Recurse -Force
    return $true
}

function Set-ContextMenuEntryDisabled {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath,

        [Parameter(Mandatory = $true)]
        [bool]$Disabled
    )

    if (-not (Test-Path -LiteralPath $RegistryPath)) {
        throw "The selected registry key was not found."
    }

    if ($Disabled) {
        Set-RegistryStringValue -Path $RegistryPath -Name "LegacyDisable" -Value ""
        return
    }

    if ((Get-Item -LiteralPath $RegistryPath).Property -contains "LegacyDisable") {
        Remove-ItemProperty -Path $RegistryPath -Name "LegacyDisable" -ErrorAction SilentlyContinue
    }
}

function Get-EnumerationRoots {
    $roots = @()

    foreach ($scope in $script:ScopeCatalog) {
        $roots += [pscustomobject]@{
                ScopeId    = $scope.Id
                ScopeLabel = $scope.Label
                Source     = "User"
                Registry   = $scope.UserPath
            }

        $roots += [pscustomobject]@{
                ScopeId    = $scope.Id
                ScopeLabel = $scope.Label
                Source     = "System"
                Registry   = $scope.MachinePath
            }
    }

    return $roots
}

function Get-InvocationSummary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EntryPath
    )

    $commandPath = Join-Path $EntryPath "command"
    $command = Get-RegistryDefaultValue -Path $commandPath
    if (-not [string]::IsNullOrWhiteSpace($command)) {
        return $command
    }

    $delegateExecute = Get-RegistryValue -Path $EntryPath -Name "DelegateExecute"
    if (-not [string]::IsNullOrWhiteSpace($delegateExecute)) {
        return "DelegateExecute: $delegateExecute"
    }

    $explorerCommandHandler = Get-RegistryValue -Path $EntryPath -Name "ExplorerCommandHandler"
    if (-not [string]::IsNullOrWhiteSpace($explorerCommandHandler)) {
        return "ExplorerCommandHandler: $explorerCommandHandler"
    }

    $subCommands = Get-RegistryValue -Path $EntryPath -Name "SubCommands"
    if (-not [string]::IsNullOrWhiteSpace($subCommands)) {
        return "SubCommands: $subCommands"
    }

    return ""
}

function Get-ContextMenuEntries {
    $entries = @()

    foreach ($root in Get-EnumerationRoots) {
        if (-not (Test-Path -LiteralPath $root.Registry)) {
            continue
        }

        try {
            $children = Get-ChildItem -LiteralPath $root.Registry -ErrorAction Stop
        }
        catch {
            continue
        }

        foreach ($child in $children) {
            $entryPath = "Registry::$($child.Name)"

            try {
                $item = Get-Item -LiteralPath $entryPath -ErrorAction Stop
            }
            catch {
                continue
            }

            $displayName = $item.GetValue("")
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                $displayName = $item.GetValue("MUIVerb")
            }
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                $displayName = $child.PSChildName
            }

            $entries += [pscustomobject]@{
                    DisplayName  = $displayName
                    Scope        = $root.ScopeLabel
                    Source       = $root.Source
                    Status       = if ($null -ne $item.GetValue("LegacyDisable")) { "Disabled" } else { "Enabled" }
                    KeyName      = $child.PSChildName
                    RegistryPath = $entryPath
                    Command      = Get-InvocationSummary -EntryPath $entryPath
                    Icon         = $item.GetValue("Icon")
                    IsDisabled   = ($null -ne $item.GetValue("LegacyDisable"))
                    IsCustomApp  = ($child.PSChildName -like "OpenWithCustomApp_*")
                }
        }
    }

    return @($entries | Sort-Object DisplayName, Scope, Source, KeyName)
}

function Test-IsAdministrator {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-EntryList {
    $entries = Get-ContextMenuEntries
    if ($entries.Count -eq 0) {
        Write-Host "No context menu entries were found."
        return
    }

    $entries |
        Select-Object DisplayName, Scope, Source, Status, KeyName |
        Format-Table -AutoSize
}

function Show-ContextMenuManagerGui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Windows 11 Context Menu Editor"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(1320, 790)
    $form.MinimumSize = New-Object System.Drawing.Size(1180, 700)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

    $topInfoLabel = New-Object System.Windows.Forms.Label
    $topInfoLabel.Location = New-Object System.Drawing.Point(12, 12)
    $topInfoLabel.Size = New-Object System.Drawing.Size(1275, 42)
    $topInfoLabel.Text = if (Test-IsAdministrator) {
        "Manage existing shell-based context menu entries from a GUI. Both user and system entries are shown."
    }
    else {
        "Manage existing shell-based context menu entries from a GUI. Editing system entries (HKLM) requires administrator rights."
    }

    $splitContainer = New-Object System.Windows.Forms.SplitContainer
    $splitContainer.Location = New-Object System.Drawing.Point(12, 62)
    $splitContainer.Size = New-Object System.Drawing.Size(1275, 680)
    $splitContainer.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right
    $splitContainer.SplitterDistance = 690

    $filterLabel = New-Object System.Windows.Forms.Label
    $filterLabel.Location = New-Object System.Drawing.Point(8, 12)
    $filterLabel.Size = New-Object System.Drawing.Size(52, 20)
    $filterLabel.Text = "Search"

    $filterTextBox = New-Object System.Windows.Forms.TextBox
    $filterTextBox.Location = New-Object System.Drawing.Point(60, 9)
    $filterTextBox.Size = New-Object System.Drawing.Size(460, 27)
    $filterTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right

    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Location = New-Object System.Drawing.Point(536, 7)
    $refreshButton.Size = New-Object System.Drawing.Size(130, 30)
    $refreshButton.Text = "Refresh"
    $refreshButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Location = New-Object System.Drawing.Point(8, 44)
    $countLabel.Size = New-Object System.Drawing.Size(430, 20)
    $countLabel.Text = "0 items"

    $entryListView = New-Object System.Windows.Forms.ListView
    $entryListView.Location = New-Object System.Drawing.Point(8, 72)
    $entryListView.Size = New-Object System.Drawing.Size(660, 596)
    $entryListView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right
    $entryListView.FullRowSelect = $true
    $entryListView.HideSelection = $false
    $entryListView.MultiSelect = $false
    $entryListView.View = [System.Windows.Forms.View]::Details
    $null = $entryListView.Columns.Add("Display Name", 270)
    $null = $entryListView.Columns.Add("Scope", 120)
    $null = $entryListView.Columns.Add("Source", 90)
    $null = $entryListView.Columns.Add("Status", 80)
    $null = $entryListView.Columns.Add("Key Name", 200)

    $detailsGroup = New-Object System.Windows.Forms.GroupBox
    $detailsGroup.Location = New-Object System.Drawing.Point(10, 8)
    $detailsGroup.Size = New-Object System.Drawing.Size(548, 418)
    $detailsGroup.Text = "Selected entry"
    $detailsGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right

    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Location = New-Object System.Drawing.Point(14, 32)
    $nameLabel.Size = New-Object System.Drawing.Size(96, 20)
    $nameLabel.Text = "Display name"

    $nameTextBox = New-Object System.Windows.Forms.TextBox
    $nameTextBox.Location = New-Object System.Drawing.Point(118, 28)
    $nameTextBox.Size = New-Object System.Drawing.Size(410, 27)
    $nameTextBox.ReadOnly = $true

    $scopeLabel = New-Object System.Windows.Forms.Label
    $scopeLabel.Location = New-Object System.Drawing.Point(14, 68)
    $scopeLabel.Size = New-Object System.Drawing.Size(96, 20)
    $scopeLabel.Text = "Scope"

    $scopeTextBox = New-Object System.Windows.Forms.TextBox
    $scopeTextBox.Location = New-Object System.Drawing.Point(118, 64)
    $scopeTextBox.Size = New-Object System.Drawing.Size(180, 27)
    $scopeTextBox.ReadOnly = $true

    $sourceLabel = New-Object System.Windows.Forms.Label
    $sourceLabel.Location = New-Object System.Drawing.Point(314, 68)
    $sourceLabel.Size = New-Object System.Drawing.Size(54, 20)
    $sourceLabel.Text = "Source"

    $sourceTextBox = New-Object System.Windows.Forms.TextBox
    $sourceTextBox.Location = New-Object System.Drawing.Point(372, 64)
    $sourceTextBox.Size = New-Object System.Drawing.Size(156, 27)
    $sourceTextBox.ReadOnly = $true

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(14, 104)
    $statusLabel.Size = New-Object System.Drawing.Size(96, 20)
    $statusLabel.Text = "Status"

    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Location = New-Object System.Drawing.Point(118, 100)
    $statusTextBox.Size = New-Object System.Drawing.Size(180, 27)
    $statusTextBox.ReadOnly = $true

    $customLabel = New-Object System.Windows.Forms.Label
    $customLabel.Location = New-Object System.Drawing.Point(314, 104)
    $customLabel.Size = New-Object System.Drawing.Size(54, 20)
    $customLabel.Text = "Type"

    $customTextBox = New-Object System.Windows.Forms.TextBox
    $customTextBox.Location = New-Object System.Drawing.Point(372, 100)
    $customTextBox.Size = New-Object System.Drawing.Size(156, 27)
    $customTextBox.ReadOnly = $true

    $keyLabel = New-Object System.Windows.Forms.Label
    $keyLabel.Location = New-Object System.Drawing.Point(14, 140)
    $keyLabel.Size = New-Object System.Drawing.Size(96, 20)
    $keyLabel.Text = "Key name"

    $keyTextBox = New-Object System.Windows.Forms.TextBox
    $keyTextBox.Location = New-Object System.Drawing.Point(118, 136)
    $keyTextBox.Size = New-Object System.Drawing.Size(410, 27)
    $keyTextBox.ReadOnly = $true

    $pathLabel = New-Object System.Windows.Forms.Label
    $pathLabel.Location = New-Object System.Drawing.Point(14, 176)
    $pathLabel.Size = New-Object System.Drawing.Size(96, 20)
    $pathLabel.Text = "Registry path"

    $pathTextBox = New-Object System.Windows.Forms.TextBox
    $pathTextBox.Location = New-Object System.Drawing.Point(118, 172)
    $pathTextBox.Size = New-Object System.Drawing.Size(410, 62)
    $pathTextBox.Multiline = $true
    $pathTextBox.ReadOnly = $true
    $pathTextBox.ScrollBars = "Vertical"

    $iconLabel = New-Object System.Windows.Forms.Label
    $iconLabel.Location = New-Object System.Drawing.Point(14, 246)
    $iconLabel.Size = New-Object System.Drawing.Size(96, 20)
    $iconLabel.Text = "Icon"

    $iconTextBox = New-Object System.Windows.Forms.TextBox
    $iconTextBox.Location = New-Object System.Drawing.Point(118, 242)
    $iconTextBox.Size = New-Object System.Drawing.Size(410, 62)
    $iconTextBox.Multiline = $true
    $iconTextBox.ReadOnly = $true
    $iconTextBox.ScrollBars = "Vertical"

    $commandLabel = New-Object System.Windows.Forms.Label
    $commandLabel.Location = New-Object System.Drawing.Point(14, 316)
    $commandLabel.Size = New-Object System.Drawing.Size(96, 20)
    $commandLabel.Text = "Command"

    $commandTextBox = New-Object System.Windows.Forms.TextBox
    $commandTextBox.Location = New-Object System.Drawing.Point(118, 312)
    $commandTextBox.Size = New-Object System.Drawing.Size(410, 62)
    $commandTextBox.Multiline = $true
    $commandTextBox.ReadOnly = $true
    $commandTextBox.ScrollBars = "Vertical"

    $enableButton = New-Object System.Windows.Forms.Button
    $enableButton.Location = New-Object System.Drawing.Point(118, 382)
    $enableButton.Size = New-Object System.Drawing.Size(120, 28)
    $enableButton.Text = "Enable"
    $enableButton.Enabled = $false

    $disableButton = New-Object System.Windows.Forms.Button
    $disableButton.Location = New-Object System.Drawing.Point(248, 382)
    $disableButton.Size = New-Object System.Drawing.Size(120, 28)
    $disableButton.Text = "Disable"
    $disableButton.Enabled = $false

    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Location = New-Object System.Drawing.Point(378, 382)
    $removeButton.Size = New-Object System.Drawing.Size(150, 28)
    $removeButton.Text = "Delete"
    $removeButton.Enabled = $false

    $detailsGroup.Controls.AddRange(@(
            $nameLabel,
            $nameTextBox,
            $scopeLabel,
            $scopeTextBox,
            $sourceLabel,
            $sourceTextBox,
            $statusLabel,
            $statusTextBox,
            $customLabel,
            $customTextBox,
            $keyLabel,
            $keyTextBox,
            $pathLabel,
            $pathTextBox,
            $iconLabel,
            $iconTextBox,
            $commandLabel,
            $commandTextBox,
            $enableButton,
            $disableButton,
            $removeButton
        ))

    $newEntryGroup = New-Object System.Windows.Forms.GroupBox
    $newEntryGroup.Location = New-Object System.Drawing.Point(10, 438)
    $newEntryGroup.Size = New-Object System.Drawing.Size(548, 230)
    $newEntryGroup.Text = "Add new entry"
    $newEntryGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right

    $exeLabel = New-Object System.Windows.Forms.Label
    $exeLabel.Location = New-Object System.Drawing.Point(14, 32)
    $exeLabel.Size = New-Object System.Drawing.Size(96, 20)
    $exeLabel.Text = "Executable"

    $exePathTextBox = New-Object System.Windows.Forms.TextBox
    $exePathTextBox.Location = New-Object System.Drawing.Point(118, 28)
    $exePathTextBox.Size = New-Object System.Drawing.Size(312, 27)
    $exePathTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Location = New-Object System.Drawing.Point(438, 27)
    $browseButton.Size = New-Object System.Drawing.Size(90, 30)
    $browseButton.Text = "Browse..."
    $browseButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

    $menuLabelLabel = New-Object System.Windows.Forms.Label
    $menuLabelLabel.Location = New-Object System.Drawing.Point(14, 68)
    $menuLabelLabel.Size = New-Object System.Drawing.Size(96, 20)
    $menuLabelLabel.Text = "Menu label"

    $menuLabelTextBox = New-Object System.Windows.Forms.TextBox
    $menuLabelTextBox.Location = New-Object System.Drawing.Point(118, 64)
    $menuLabelTextBox.Size = New-Object System.Drawing.Size(410, 27)
    $menuLabelTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right

    $scopeSelectLabel = New-Object System.Windows.Forms.Label
    $scopeSelectLabel.Location = New-Object System.Drawing.Point(14, 104)
    $scopeSelectLabel.Size = New-Object System.Drawing.Size(96, 20)
    $scopeSelectLabel.Text = "Targets"

    $scopeListBox = New-Object System.Windows.Forms.CheckedListBox
    $scopeListBox.Location = New-Object System.Drawing.Point(118, 100)
    $scopeListBox.Size = New-Object System.Drawing.Size(410, 76)
    $scopeListBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right
    $scopeListBox.CheckOnClick = $true
    $scopeListBox.DisplayMember = "Label"

    foreach ($scope in $script:ScopeCatalog) {
        $index = $scopeListBox.Items.Add($scope)
        if ($scope.DefaultForNew) {
            $scopeListBox.SetItemChecked($index, $true)
        }
    }

    $noteLabel = New-Object System.Windows.Forms.Label
    $noteLabel.Location = New-Object System.Drawing.Point(118, 182)
    $noteLabel.Size = New-Object System.Drawing.Size(410, 24)
    $noteLabel.Text = "Use the list above to manage existing entries. Disabling is safer than deleting."

    $createButton = New-Object System.Windows.Forms.Button
    $createButton.Location = New-Object System.Drawing.Point(378, 198)
    $createButton.Size = New-Object System.Drawing.Size(150, 28)
    $createButton.Text = "Add entry"
    $createButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right

    $newEntryGroup.Controls.AddRange(@(
            $exeLabel,
            $exePathTextBox,
            $browseButton,
            $menuLabelLabel,
            $menuLabelTextBox,
            $scopeSelectLabel,
            $scopeListBox,
            $noteLabel,
            $createButton
        ))

    $splitContainer.Panel1.Controls.AddRange(@(
            $filterLabel,
            $filterTextBox,
            $refreshButton,
            $countLabel,
            $entryListView
        ))
    $splitContainer.Panel2.Controls.AddRange(@(
            $detailsGroup,
            $newEntryGroup
        ))

    $form.Controls.AddRange(@(
            $topInfoLabel,
            $splitContainer
        ))

    $state = @{
        AllEntries = @()
    }

    function Show-Message {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter(Mandatory = $false)]
            [string]$Title = "Windows 11 Context Menu Editor",

            [Parameter(Mandatory = $false)]
            [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
        )

        [System.Windows.Forms.MessageBox]::Show(
            $form,
            $Message,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $Icon
        ) | Out-Null
    }

    function Clear-EntryDetails {
        $nameTextBox.Text = ""
        $scopeTextBox.Text = ""
        $sourceTextBox.Text = ""
        $statusTextBox.Text = ""
        $customTextBox.Text = ""
        $keyTextBox.Text = ""
        $pathTextBox.Text = ""
        $iconTextBox.Text = ""
        $commandTextBox.Text = ""
        $enableButton.Enabled = $false
        $disableButton.Enabled = $false
        $removeButton.Enabled = $false
    }

    function Get-SelectedEntry {
        if ($entryListView.SelectedItems.Count -eq 0) {
            return $null
        }

        return $entryListView.SelectedItems[0].Tag
    }

    function Update-EntryDetails {
        $selectedEntry = Get-SelectedEntry
        if ($null -eq $selectedEntry) {
            Clear-EntryDetails
            return
        }

        $nameTextBox.Text = [string]$selectedEntry.DisplayName
        $scopeTextBox.Text = [string]$selectedEntry.Scope
        $sourceTextBox.Text = [string]$selectedEntry.Source
        $statusTextBox.Text = [string]$selectedEntry.Status
        $customTextBox.Text = if ($selectedEntry.IsCustomApp) { "Custom entry" } else { "Existing entry" }
        $keyTextBox.Text = [string]$selectedEntry.KeyName
        $pathTextBox.Text = [string]$selectedEntry.RegistryPath
        $iconTextBox.Text = [string]$selectedEntry.Icon
        $commandTextBox.Text = [string]$selectedEntry.Command
        $enableButton.Enabled = [bool]$selectedEntry.IsDisabled
        $disableButton.Enabled = -not [bool]$selectedEntry.IsDisabled
        $removeButton.Enabled = $true
    }

    function Update-EntryList {
        param(
            [Parameter(Mandatory = $false)]
            [string]$SelectedPath
        )

        $filter = $filterTextBox.Text.Trim()
        $comparison = [System.StringComparison]::OrdinalIgnoreCase

        $visibleEntries =
            if ([string]::IsNullOrWhiteSpace($filter)) {
                $state.AllEntries
            }
            else {
                @($state.AllEntries | Where-Object {
                        $_.DisplayName.IndexOf($filter, $comparison) -ge 0 -or
                        $_.KeyName.IndexOf($filter, $comparison) -ge 0 -or
                        $_.Scope.IndexOf($filter, $comparison) -ge 0 -or
                        $_.Source.IndexOf($filter, $comparison) -ge 0 -or
                        $_.Command.IndexOf($filter, $comparison) -ge 0
                    })
            }

        $entryListView.BeginUpdate()
        $entryListView.Items.Clear()

        foreach ($entry in $visibleEntries) {
            $item = New-Object System.Windows.Forms.ListViewItem($entry.DisplayName)
            $null = $item.SubItems.Add([string]$entry.Scope)
            $null = $item.SubItems.Add([string]$entry.Source)
            $null = $item.SubItems.Add([string]$entry.Status)
            $null = $item.SubItems.Add([string]$entry.KeyName)
            $item.Tag = $entry
            $null = $entryListView.Items.Add($item)
        }

        $entryListView.EndUpdate()
        $countLabel.Text = "{0} items" -f $visibleEntries.Count

        if (-not [string]::IsNullOrWhiteSpace($SelectedPath)) {
            foreach ($item in $entryListView.Items) {
                if ($item.Tag.RegistryPath -eq $SelectedPath) {
                    $item.Selected = $true
                    $item.Focused = $true
                    $item.EnsureVisible()
                    break
                }
            }
        }

        if ($entryListView.SelectedItems.Count -eq 0 -and $entryListView.Items.Count -gt 0) {
            $entryListView.Items[0].Selected = $true
        }

        Update-EntryDetails
    }

    function Refresh-Entries {
        param(
            [Parameter(Mandatory = $false)]
            [string]$SelectedPath
        )

        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $state.AllEntries = @(Get-ContextMenuEntries)
            Update-EntryList -SelectedPath $SelectedPath
        }
        catch {
            Show-Message -Message $_.Exception.Message -Title "Load error" -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    function Get-CheckedScopeIds {
        $selectedScopeIds = @()
        for ($index = 0; $index -lt $scopeListBox.Items.Count; $index++) {
            if ($scopeListBox.GetItemChecked($index)) {
                $selectedScopeIds += $scopeListBox.Items[$index].Id
            }
        }

        return $selectedScopeIds
    }

    $entryListView.Add_SelectedIndexChanged({
            Update-EntryDetails
        })

    $filterTextBox.Add_TextChanged({
            Update-EntryList -SelectedPath $null
        })

    $refreshButton.Add_Click({
            $selectedEntry = Get-SelectedEntry
            $selectedPath = if ($null -ne $selectedEntry) { $selectedEntry.RegistryPath } else { $null }
            Refresh-Entries -SelectedPath $selectedPath
        })

    $browseButton.Add_Click({
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Filter = "Executable (*.exe)|*.exe|All files (*.*)|*.*"
            $dialog.Title = "Choose an executable to add to the context menu"
            $dialog.CheckFileExists = $true
            $dialog.Multiselect = $false

            if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
                $exePathTextBox.Text = $dialog.FileName
                $menuLabelTextBox.Text = Get-DefaultMenuLabel -ExePath $dialog.FileName
            }
        })

    $createButton.Add_Click({
            try {
                $selectedScopeIds = Get-CheckedScopeIds
                if ($selectedScopeIds.Count -eq 0) {
                    Show-Message -Message "Select at least one target scope." -Title "Missing input" -Icon ([System.Windows.Forms.MessageBoxIcon]::Warning)
                    return
                }

                $result = Add-CustomAppContextMenu -ExePath $exePathTextBox.Text -ScopeIds $selectedScopeIds -MenuLabel $menuLabelTextBox.Text
                Refresh-Entries -SelectedPath $result.RegistryPaths[0]
                Show-Message -Message "Menu entry added.`n$($result.MenuLabel)" -Title "Entry added"
            }
            catch {
                Show-Message -Message $_.Exception.Message -Title "Add error" -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })

    $enableButton.Add_Click({
            $selectedEntry = Get-SelectedEntry
            if ($null -eq $selectedEntry) {
                return
            }

            try {
                Set-ContextMenuEntryDisabled -RegistryPath $selectedEntry.RegistryPath -Disabled $false
                Refresh-Entries -SelectedPath $selectedEntry.RegistryPath
            }
            catch {
                Show-Message -Message $_.Exception.Message -Title "Enable error" -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })

    $disableButton.Add_Click({
            $selectedEntry = Get-SelectedEntry
            if ($null -eq $selectedEntry) {
                return
            }

            try {
                Set-ContextMenuEntryDisabled -RegistryPath $selectedEntry.RegistryPath -Disabled $true
                Refresh-Entries -SelectedPath $selectedEntry.RegistryPath
            }
            catch {
                Show-Message -Message $_.Exception.Message -Title "Disable error" -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })

    $removeButton.Add_Click({
            $selectedEntry = Get-SelectedEntry
            if ($null -eq $selectedEntry) {
                return
            }

            $confirmation = [System.Windows.Forms.MessageBox]::Show(
                $form,
                "Delete the selected menu entry permanently?`n`n$($selectedEntry.DisplayName)`n$($selectedEntry.RegistryPath)`n`nIf you are unsure, disable it first.",
                "Confirm delete",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }

            try {
                $removed = Remove-ContextMenuEntry -RegistryPath $selectedEntry.RegistryPath
                if ($removed) {
                    Refresh-Entries -SelectedPath $null
                }
            }
            catch {
                Show-Message -Message $_.Exception.Message -Title "Delete error" -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })

    Clear-EntryDetails
    Refresh-Entries -SelectedPath $null
    [void]$form.ShowDialog()
}

$normalizedAction = if ([string]::IsNullOrWhiteSpace($Action)) { "" } else { $Action.Trim().ToUpperInvariant() }

if ($Gui -or ([string]::IsNullOrWhiteSpace($normalizedAction) -and [string]::IsNullOrWhiteSpace($AppPath))) {
    Show-ContextMenuManagerGui
    return
}

switch ($normalizedAction) {
    "A" {
        if ([string]::IsNullOrWhiteSpace($AppPath)) {
            $AppPath = Read-Host "Enter the full path to the EXE you want to add"
        }

        $result = Add-CustomAppContextMenu -ExePath $AppPath -ScopeIds $Scopes
        Write-Host ""
        Write-Host "Added to the context menu."
        Write-Host "Menu label: $($result.MenuLabel)"
        Write-Host "Key name: $($result.KeyName)"
    }

    "R" {
        if ([string]::IsNullOrWhiteSpace($AppPath)) {
            $AppPath = Read-Host "Enter the full path to the EXE you want to remove"
        }

        $result = Remove-CustomAppContextMenu -ExePath $AppPath
        Write-Host ""
        if ($result.WasRemoved) {
            Write-Host "Removed from the context menu."
            Write-Host "Key name: $($result.KeyName)"
        }
        else {
            Write-Host "No matching custom menu entry was found."
        }
    }

    "L" {
        Write-EntryList
    }

    default {
        throw "Action must be A (add), R (remove), or L (list)."
    }
}
