param(
    [Parameter(Mandatory = $false)]
    [string]$AppPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("A", "R", "L", "a", "r", "l")]
    [string]$Action,

    [Parameter(Mandatory = $false)]
    [string[]]$Scopes,

    [Parameter(Mandatory = $false)]
    [switch]$Gui,

    [Parameter(Mandatory = $false)]
    [switch]$HiddenGuiHost
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:Text = @{
    WindowTitle                     = "コンテキストメニュー エディター"
    HeaderEyebrow                   = "Windows 11"
    HeaderTitle                     = "コンテキストメニューを見やすく整理"
    HeaderDescription               = "右クリックメニューの項目を確認し、追加・無効化・削除をまとめて行えます。"
    AdminBadge                      = "管理者権限"
    AdminNote                       = "システム項目の確認と編集を行うため、管理者権限で実行します。"
    SearchLabel                     = "検索"
    SearchPlaceholder               = "表示名、キー名、コマンドを検索"
    RefreshButton                   = "更新"
    AddButton                       = "項目を追加"
    ScopeSectionTitle               = "対象"
    ScopeAll                        = "すべて"
    CountFormat                     = "{0} 件 / {1}"
    ListTitle                       = "登録済みの項目"
    ListDescription                 = "現在のコンテキストメニュー項目を一覧表示します。"
    ColumnDisplayName               = "表示名"
    ColumnScope                     = "対象"
    ColumnSource                    = "ソース"
    ColumnStatus                    = "状態"
    ColumnKeyName                   = "キー名"
    DetailTitle                     = "詳細"
    DetailDescription               = "選択した項目の状態とレジストリ情報を表示します。"
    DetailEmptyTitle                = "項目を選択してください"
    DetailEmptyDescription          = "一覧から項目を選ぶと、状態やレジストリ情報をここに表示します。"
    DetailDisplayName               = "表示名"
    DetailScope                     = "対象"
    DetailSource                    = "ソース"
    DetailStatus                    = "状態"
    DetailType                      = "種類"
    DetailKeyName                   = "キー名"
    DetailPath                      = "レジストリ パス"
    DetailIcon                      = "アイコン"
    DetailCommand                   = "コマンド"
    EnableButton                    = "有効化"
    DisableButton                   = "無効化"
    DeleteButton                    = "削除"
    ModalTitle                      = "新しい項目を追加"
    ModalDescription                = "アプリを右クリックメニューに登録します。"
    ExePathLabel                    = "実行ファイル (.exe)"
    BrowseButton                    = "参照..."
    MenuLabelLabel                  = "メニュー名"
    ScopeLabel                      = "追加先"
    ModalTip                        = "既存のシステム項目を扱うときは、削除より「無効化」のほうが安全です。"
    ModalCancelButton               = "キャンセル"
    ModalAddButton                  = "追加"
    ValidationMissingExe            = "存在する .exe ファイルを指定してください。"
    ValidationMissingScope          = "追加先を 1 つ以上選択してください。"
    ValidationReady                 = "この設定で右クリックメニューに追加できます。"
    SourceUser                      = "ユーザー"
    SourceSystem                    = "システム"
    StatusEnabled                   = "有効"
    StatusDisabled                  = "無効"
    TypeCustomApp                   = "追加した項目"
    TypeExisting                    = "既存の項目"
    ScopeAllFiles                   = "ファイル"
    ScopeDirectories                = "フォルダー"
    ScopeDirectoryBackground        = "フォルダー背景"
    ScopeFolders                    = "フォルダー全般"
    ScopeDrives                     = "ドライブ"
    ListNoneFound                   = "コンテキストメニュー項目は見つかりませんでした。"
    PromptAddAppPath                = "追加する EXE のフルパスを入力してください"
    PromptRemoveAppPath             = "削除する EXE のフルパスを入力してください"
    OutputAdded                     = "コンテキストメニューに追加しました。"
    OutputRemoved                   = "コンテキストメニューから削除しました。"
    OutputMenuLabel                 = "メニュー名: {0}"
    OutputKeyName                   = "キー名: {0}"
    OutputNoMatchingEntry           = "一致するカスタム項目は見つかりませんでした。"
    OutputElevationCancelled        = "管理者権限への昇格がキャンセルされたため、処理を中止しました。"
    ErrorExePathEmpty               = "EXE のパスが空です。"
    ErrorExeNotFound                = "EXE が見つかりません: {0}"
    ErrorInvalidScope               = "無効なスコープです: {0}`n指定可能: {1}"
    ErrorNoScopeSelected            = "追加先を 1 つ以上選択してください。"
    ErrorSelectedKeyNotFound        = "選択したレジストリ キーが見つかりません。"
    ErrorCannotDetermineScriptPath  = "昇格実行用のスクリプト パスを特定できませんでした。"
    ErrorInvalidAction              = "Action には A (追加)、R (削除)、L (一覧) のいずれかを指定してください。"
    ErrorElevationCancelled         = "管理者権限への昇格がキャンセルされたため、処理を中止しました。"
    MessageLoadErrorTitle           = "読み込みエラー"
    MessageAddErrorTitle            = "追加エラー"
    MessageAddSuccessTitle          = "項目を追加しました"
    MessageAddSuccess               = "項目を追加しました。`n{0}"
    MessageEnableErrorTitle         = "有効化エラー"
    MessageEnableSuccessTitle       = "有効化しました"
    MessageEnableSuccess            = "項目を有効化しました。"
    MessageDisableErrorTitle        = "無効化エラー"
    MessageDisableSuccessTitle      = "無効化しました"
    MessageDisableSuccess           = "項目を無効化しました。"
    MessageDeleteErrorTitle         = "削除エラー"
    MessageDeleteSuccessTitle       = "削除しました"
    MessageDeleteSuccess            = "項目を削除しました。"
    MessageDeleteConfirmTitle       = "削除の確認"
    MessageDeleteConfirm            = "選択中の項目を完全に削除しますか？`n`n{0}`n{1}`n`n迷う場合は、先に「無効化」を選ぶほうが安全です。"
    MessageInputErrorTitle          = "入力内容を確認してください"
    OpenFileDialogFilter            = "実行ファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*"
    OpenFileDialogTitle             = "追加する実行ファイルを選択してください"
    InvocationDelegateExecute       = "DelegateExecute: {0}"
    InvocationExplorerCommandHandler = "ExplorerCommandHandler: {0}"
    InvocationSubCommands           = "SubCommands: {0}"
}

$script:IconSourceCache = @{}
$script:FallbackIconSource = $null
$script:IndirectStringCache = @{}
$script:CustomAppKeyPrefix = "CustomApp_"
$script:LegacyCustomAppKeyPrefix = "OpenWithCustomApp_"

function Get-Text {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,

        [Parameter(Mandatory = $false)]
        [object[]]$Arguments
    )

    if (-not $script:Text.ContainsKey($Key)) {
        throw "Localized text not found: $Key"
    }

    $template = [string]$script:Text[$Key]
    if ($null -ne $Arguments -and $Arguments.Count -gt 0) {
        return [string]::Format([System.Globalization.CultureInfo]::CurrentCulture, $template, $Arguments)
    }

    return $template
}

function Get-SourceDisplayText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceKey
    )

    switch ($SourceKey) {
        "User" { return Get-Text -Key "SourceUser" }
        "System" { return Get-Text -Key "SourceSystem" }
        default { return $SourceKey }
    }
}

function Get-StatusDisplayText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StatusKey
    )

    switch ($StatusKey) {
        "Enabled" { return Get-Text -Key "StatusEnabled" }
        "Disabled" { return Get-Text -Key "StatusDisabled" }
        default { return $StatusKey }
    }
}

function Get-TypeDisplayText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TypeKey
    )

    switch ($TypeKey) {
        "CustomApp" { return Get-Text -Key "TypeCustomApp" }
        "Existing" { return Get-Text -Key "TypeExisting" }
        default { return $TypeKey }
    }
}

function Test-ExeFilePath {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$PathText
    )

    if ([string]::IsNullOrWhiteSpace($PathText)) {
        return $false
    }

    try {
        $extension = [System.IO.Path]::GetExtension($PathText)
    }
    catch {
        return $false
    }

    if (-not ".exe".Equals($extension, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    return (Test-Path -LiteralPath $PathText -PathType Leaf)
}

function Get-CompactSingleLineText {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text,

        [Parameter(Mandatory = $false)]
        [int]$MaxLength = 96
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $singleLine = [System.Text.RegularExpressions.Regex]::Replace($Text, "\s+", " ").Trim()
    if ($singleLine.Length -le $MaxLength) {
        return $singleLine
    }

    return $singleLine.Substring(0, [Math]::Max(0, $MaxLength - 3)) + "..."
}

function Initialize-ContextMenuShellStringInterop {
    if ("ContextMenuEditor.NativeShellStrings" -as [type]) {
        return
    }

    Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ContextMenuEditor {
    public static class NativeShellStrings {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        public static extern int SHLoadIndirectString(
            string pszSource,
            StringBuilder pszOutBuf,
            uint cchOutBuf,
            IntPtr ppvReserved);
    }
}
"@
}

function Resolve-IndirectShellString {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    if ($script:IndirectStringCache.ContainsKey($Text)) {
        return [string]$script:IndirectStringCache[$Text]
    }

    if (-not $Text.StartsWith('@', [System.StringComparison]::Ordinal)) {
        $script:IndirectStringCache[$Text] = $Text
        return $Text
    }

    Initialize-ContextMenuShellStringInterop

    try {
        $buffer = [System.Text.StringBuilder]::new(1024)
        $hr = [ContextMenuEditor.NativeShellStrings]::SHLoadIndirectString($Text, $buffer, [uint32]$buffer.Capacity, [System.IntPtr]::Zero)
        if ($hr -ge 0) {
            $resolvedText = $buffer.ToString().Trim()
            if (-not [string]::IsNullOrWhiteSpace($resolvedText)) {
                $script:IndirectStringCache[$Text] = $resolvedText
                return $resolvedText
            }
        }
    }
    catch {
    }

    $script:IndirectStringCache[$Text] = $Text
    return $Text
}

function Initialize-ContextMenuIconInterop {
    if ("ContextMenuEditor.NativeIcons" -as [type]) {
        return
    }

    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Drawing

    Add-Type -ReferencedAssemblies @("System.Drawing") @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace ContextMenuEditor {
    public static class NativeIcons {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint ExtractIconEx(string szFileName, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool DestroyIcon(IntPtr hIcon);

        public static IntPtr ExtractIndexedIconHandle(string fileName, int iconIndex) {
            if (string.IsNullOrWhiteSpace(fileName)) {
                return IntPtr.Zero;
            }

            IntPtr[] largeIcons = new IntPtr[1];
            IntPtr[] smallIcons = new IntPtr[1];

            try {
                uint extracted = ExtractIconEx(fileName, iconIndex, largeIcons, smallIcons, 1);
                if (extracted == 0) {
                    return IntPtr.Zero;
                }

                IntPtr result = smallIcons[0] != IntPtr.Zero ? smallIcons[0] : largeIcons[0];
                if (result == smallIcons[0]) {
                    smallIcons[0] = IntPtr.Zero;
                }
                if (result == largeIcons[0]) {
                    largeIcons[0] = IntPtr.Zero;
                }

                return result;
            }
            catch {
                return IntPtr.Zero;
            }
            finally {
                if (largeIcons[0] != IntPtr.Zero) {
                    DestroyIcon(largeIcons[0]);
                }
                if (smallIcons[0] != IntPtr.Zero) {
                    DestroyIcon(smallIcons[0]);
                }
            }
        }

        public static IntPtr ExtractAssociatedIconHandle(string fileName) {
            if (string.IsNullOrWhiteSpace(fileName)) {
                return IntPtr.Zero;
            }

            try {
                using (Icon icon = Icon.ExtractAssociatedIcon(fileName)) {
                    if (icon == null) {
                        return IntPtr.Zero;
                    }

                    using (Bitmap bitmap = icon.ToBitmap()) {
                        return bitmap.GetHicon();
                    }
                }
            }
            catch {
                return IntPtr.Zero;
            }
        }
    }
}
"@
}

function New-FallbackIconSource {
    if ($null -ne $script:FallbackIconSource) {
        return $script:FallbackIconSource
    }

    Initialize-ContextMenuIconInterop

    $bodyColor = [System.Windows.Media.Color]::FromRgb(234, 239, 247)
    $headerColor = [System.Windows.Media.Color]::FromRgb(208, 220, 236)
    $accentColor = [System.Windows.Media.Color]::FromRgb(116, 138, 166)
    $outlineColor = [System.Windows.Media.Color]::FromRgb(174, 188, 204)

    $bodyBrush = [System.Windows.Media.SolidColorBrush]::new($bodyColor)
    $headerBrush = [System.Windows.Media.SolidColorBrush]::new($headerColor)
    $accentBrush = [System.Windows.Media.SolidColorBrush]::new($accentColor)
    $outlineBrush = [System.Windows.Media.SolidColorBrush]::new($outlineColor)
    $outlinePen = [System.Windows.Media.Pen]::new($outlineBrush, 1)

    foreach ($freezable in @($bodyBrush, $headerBrush, $accentBrush, $outlineBrush, $outlinePen)) {
        if ($freezable.CanFreeze) {
            $freezable.Freeze()
        }
    }

    $drawingGroup = [System.Windows.Media.DrawingGroup]::new()
    $drawingGroup.Children.Add([System.Windows.Media.GeometryDrawing]::new(
            $bodyBrush,
            $outlinePen,
            [System.Windows.Media.Geometry]::Parse("M2.5,3.5 H13.5 V12.5 H2.5 Z")
        )) | Out-Null
    $drawingGroup.Children.Add([System.Windows.Media.GeometryDrawing]::new(
            $headerBrush,
            $null,
            [System.Windows.Media.Geometry]::Parse("M2.5,3.5 H13.5 V6 H2.5 Z")
        )) | Out-Null
    $drawingGroup.Children.Add([System.Windows.Media.GeometryDrawing]::new(
            $accentBrush,
            $null,
            [System.Windows.Media.Geometry]::Parse("M4.5,8 H7.25 V10.75 H4.5 Z M8.75,8 H11.5 V10.75 H8.75 Z")
        )) | Out-Null

    if ($drawingGroup.CanFreeze) {
        $drawingGroup.Freeze()
    }

    $image = [System.Windows.Media.DrawingImage]::new($drawingGroup)
    if ($image.CanFreeze) {
        $image.Freeze()
    }

    $script:FallbackIconSource = $image
    return $script:FallbackIconSource
}

function Convert-HIconToImageSource {
    param(
        [Parameter(Mandatory = $true)]
        [System.IntPtr]$Handle,

        [Parameter(Mandatory = $false)]
        [int]$Size = 18
    )

    if ($Handle -eq [System.IntPtr]::Zero) {
        return $null
    }

    Initialize-ContextMenuIconInterop

    try {
        $bitmapSource = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon(
            $Handle,
            [System.Windows.Int32Rect]::Empty,
            [System.Windows.Media.Imaging.BitmapSizeOptions]::FromWidthAndHeight($Size, $Size)
        )

        if ($bitmapSource.CanFreeze) {
            $bitmapSource.Freeze()
        }

        return $bitmapSource
    }
    finally {
        [ContextMenuEditor.NativeIcons]::DestroyIcon($Handle) | Out-Null
    }
}

function Get-NormalizedIconLocation {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$IconSpec
    )

    if ([string]::IsNullOrWhiteSpace($IconSpec)) {
        return $null
    }

    $trimmed = $IconSpec.Trim()
    $path = $null
    $index = 0

    if ($trimmed -match '^\s*"(?<path>.+)"\s*(?:,\s*(?<index>-?\d+))?\s*$') {
        $path = $matches['path']
        if ($matches['index']) {
            $index = [int]$matches['index']
        }
    }
    elseif ($trimmed -match '^\s*(?<path>.*?)(?:\s*,\s*(?<index>-?\d+))?\s*$') {
        $path = $matches['path']
        if ($matches['index']) {
            $index = [int]$matches['index']
        }
    }

    if ([string]::IsNullOrWhiteSpace($path)) {
        return $null
    }

    $expandedPath = [Environment]::ExpandEnvironmentVariables($path.Trim().Trim('"'))
    if ([string]::IsNullOrWhiteSpace($expandedPath)) {
        return $null
    }

    return [pscustomobject]@{
        Path     = $expandedPath
        Index    = $index
        CacheKey = "{0}|{1}" -f $expandedPath.ToLowerInvariant(), $index
    }
}

function Resolve-ExecutablePathFromCommandText {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$CommandText
    )

    if ([string]::IsNullOrWhiteSpace($CommandText)) {
        return $null
    }

    $candidate = $null
    if ($CommandText -match '^\s*"(?<path>[^"]+)"') {
        $candidate = $matches['path']
    }
    elseif ($CommandText -match '^\s*(?<path>\S+)') {
        $candidate = $matches['path']
    }

    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return $null
    }

    $candidate = [Environment]::ExpandEnvironmentVariables($candidate.Trim().Trim('"'))

    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
        return (Resolve-Path -LiteralPath $candidate).Path
    }

    try {
        $resolvedCommand = Get-Command -Name $candidate -CommandType Application -ErrorAction Stop | Select-Object -First 1
        if ($null -ne $resolvedCommand -and -not [string]::IsNullOrWhiteSpace($resolvedCommand.Path)) {
            return [string]$resolvedCommand.Path
        }
    }
    catch {
    }

    return $null
}

function Get-ImageSourceFromIconLocation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $false)]
        [int]$IconIndex = 0
    )

    if ([string]::IsNullOrWhiteSpace($FilePath) -or -not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
        return $null
    }

    Initialize-ContextMenuIconInterop

    $handle = [ContextMenuEditor.NativeIcons]::ExtractIndexedIconHandle($FilePath, $IconIndex)
    if ($handle -eq [System.IntPtr]::Zero) {
        $handle = [ContextMenuEditor.NativeIcons]::ExtractAssociatedIconHandle($FilePath)
    }

    if ($handle -eq [System.IntPtr]::Zero) {
        return $null
    }

    return Convert-HIconToImageSource -Handle $handle
}

function Resolve-ContextMenuEntryIconSource {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$IconSpec,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$CommandText
    )

    try {
        $iconLocation = Get-NormalizedIconLocation -IconSpec $IconSpec
        if ($null -ne $iconLocation) {
            $registryCacheKey = "registry::{0}" -f $iconLocation.CacheKey
            if ($script:IconSourceCache.ContainsKey($registryCacheKey)) {
                return $script:IconSourceCache[$registryCacheKey]
            }

            $registryIconSource = Get-ImageSourceFromIconLocation -FilePath $iconLocation.Path -IconIndex $iconLocation.Index
            if ($null -ne $registryIconSource) {
                $script:IconSourceCache[$registryCacheKey] = $registryIconSource
                return $registryIconSource
            }
        }

        $commandExecutable = Resolve-ExecutablePathFromCommandText -CommandText $CommandText
        if (-not [string]::IsNullOrWhiteSpace($commandExecutable)) {
            $commandCacheKey = "command::{0}" -f $commandExecutable.ToLowerInvariant()
            if ($script:IconSourceCache.ContainsKey($commandCacheKey)) {
                return $script:IconSourceCache[$commandCacheKey]
            }

            $commandIconSource = Get-ImageSourceFromIconLocation -FilePath $commandExecutable
            if ($null -ne $commandIconSource) {
                $script:IconSourceCache[$commandCacheKey] = $commandIconSource
                return $commandIconSource
            }
        }
    }
    catch {
    }

    if (-not $script:IconSourceCache.ContainsKey("__fallback__")) {
        $script:IconSourceCache["__fallback__"] = New-FallbackIconSource
    }

    return $script:IconSourceCache["__fallback__"]
}

function Get-NormalizedPositionKey {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$PositionValue
    )

    if ([string]::IsNullOrWhiteSpace($PositionValue)) {
        return "Normal"
    }

    switch -Regex ($PositionValue.Trim()) {
        '^(?i)top$' { return "Top" }
        '^(?i)bottom$' { return "Bottom" }
        default { return "Normal" }
    }
}

function Initialize-ContextMenuRegistryInterop {
    if ("ContextMenuEditor.NativeRegistry" -as [type]) {
        return
    }

    Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace ContextMenuEditor {
    public static class NativeRegistry {
        public static readonly UIntPtr HKEY_CLASSES_ROOT = (UIntPtr)0x80000000u;
        public static readonly UIntPtr HKEY_CURRENT_USER = (UIntPtr)0x80000001u;
        public static readonly UIntPtr HKEY_LOCAL_MACHINE = (UIntPtr)0x80000002u;

        private const int KEY_READ = 0x20019;
        private const int ERROR_SUCCESS = 0;
        private const int ERROR_MORE_DATA = 234;
        private const int ERROR_NO_MORE_ITEMS = 259;

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int RegOpenKeyEx(UIntPtr hKey, string lpSubKey, int ulOptions, int samDesired, out IntPtr phkResult);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern int RegCloseKey(IntPtr hKey);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int RegEnumKeyEx(
            IntPtr hKey,
            uint dwIndex,
            StringBuilder lpName,
            ref uint lpcchName,
            IntPtr lpReserved,
            IntPtr lpClass,
            IntPtr lpcchClass,
            IntPtr lpftLastWriteTime);

        public static string[] EnumerateSubKeyNames(UIntPtr hiveHandle, string subKey) {
            List<string> names = new List<string>();
            IntPtr openedKey = IntPtr.Zero;
            int openResult = RegOpenKeyEx(hiveHandle, subKey ?? string.Empty, 0, KEY_READ, out openedKey);
            if (openResult != ERROR_SUCCESS || openedKey == IntPtr.Zero) {
                return names.ToArray();
            }

            try {
                uint index = 0;
                while (true) {
                    uint capacity = 256;
                    while (true) {
                        StringBuilder nameBuilder = new StringBuilder((int)capacity + 1);
                        uint nameLength = capacity;
                        int enumResult = RegEnumKeyEx(
                            openedKey,
                            index,
                            nameBuilder,
                            ref nameLength,
                            IntPtr.Zero,
                            IntPtr.Zero,
                            IntPtr.Zero,
                            IntPtr.Zero);

                        if (enumResult == ERROR_SUCCESS) {
                            names.Add(nameBuilder.ToString());
                            break;
                        }

                        if (enumResult == ERROR_MORE_DATA) {
                            capacity *= 2;
                            continue;
                        }

                        if (enumResult == ERROR_NO_MORE_ITEMS) {
                            return names.ToArray();
                        }

                        break;
                    }

                    index++;
                }
            }
            finally {
                RegCloseKey(openedKey);
            }
        }
    }
}
"@
}

function Convert-RegistryPathToNativeLocation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath
    )

    $normalizedPath = $RegistryPath.Trim()
    if ($normalizedPath.StartsWith("Registry::", [System.StringComparison]::OrdinalIgnoreCase)) {
        $normalizedPath = $normalizedPath.Substring(10)
    }

    switch -Regex ($normalizedPath) {
        '^HKEY_CLASSES_ROOT(?:\\(?<sub>.*))?$' {
            return [pscustomobject]@{
                Hive   = 'HKEY_CLASSES_ROOT'
                SubKey = [string]$matches['sub']
            }
        }
        '^HKEY_CURRENT_USER(?:\\(?<sub>.*))?$' {
            return [pscustomobject]@{
                Hive   = 'HKEY_CURRENT_USER'
                SubKey = [string]$matches['sub']
            }
        }
        '^HKEY_LOCAL_MACHINE(?:\\(?<sub>.*))?$' {
            return [pscustomobject]@{
                Hive   = 'HKEY_LOCAL_MACHINE'
                SubKey = [string]$matches['sub']
            }
        }
        default {
            throw "Unsupported registry path: $RegistryPath"
        }
    }
}

function Get-NativeRegistryHiveHandle {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HiveName
    )

    Initialize-ContextMenuRegistryInterop

    switch ($HiveName) {
        'HKEY_CLASSES_ROOT' { return [ContextMenuEditor.NativeRegistry]::HKEY_CLASSES_ROOT }
        'HKEY_CURRENT_USER' { return [ContextMenuEditor.NativeRegistry]::HKEY_CURRENT_USER }
        'HKEY_LOCAL_MACHINE' { return [ContextMenuEditor.NativeRegistry]::HKEY_LOCAL_MACHINE }
        default { throw "Unsupported registry hive: $HiveName" }
    }
}

function Get-NativeRegistrySubKeyNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath
    )

    $location = Convert-RegistryPathToNativeLocation -RegistryPath $RegistryPath
    $hiveHandle = Get-NativeRegistryHiveHandle -HiveName $location.Hive
    return @([ContextMenuEditor.NativeRegistry]::EnumerateSubKeyNames($hiveHandle, $location.SubKey))
}

function Get-ShellVerbOrderSpecification {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ShellRegistryPath
    )

    $defaultValue = [string](Get-RegistryDefaultValue -Path $ShellRegistryPath)
    if ([string]::IsNullOrWhiteSpace($defaultValue)) {
        return @()
    }

    $tokens = [System.Text.RegularExpressions.Regex]::Split(($defaultValue -replace ',', ' '), '\s+') |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    $seen = @{}
    $orderedTokens = @()
    foreach ($token in $tokens) {
        $normalizedToken = $token.Trim()
        if ([string]::IsNullOrWhiteSpace($normalizedToken)) {
            continue
        }

        $lookupKey = $normalizedToken.ToLowerInvariant()
        if ($seen.ContainsKey($lookupKey)) {
            continue
        }

        $seen[$lookupKey] = $true
        $orderedTokens += $normalizedToken
    }

    return @($orderedTokens)
}

$script:ScopeCatalog = @(
    [pscustomobject]@{
        Id            = "AllFiles"
        Label         = Get-Text -Key "ScopeAllFiles"
        UserPath      = "Registry::HKEY_CURRENT_USER\Software\Classes\*\shell"
        MachinePath   = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\*\shell"
        CommandArg    = "%1"
        DefaultForNew = $true
    },
    [pscustomobject]@{
        Id            = "Directories"
        Label         = Get-Text -Key "ScopeDirectories"
        UserPath      = "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\shell"
        MachinePath   = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\Directory\shell"
        CommandArg    = "%V"
        DefaultForNew = $true
    },
    [pscustomobject]@{
        Id            = "DirectoryBackground"
        Label         = Get-Text -Key "ScopeDirectoryBackground"
        UserPath      = "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\Background\shell"
        MachinePath   = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\Directory\Background\shell"
        CommandArg    = "%V"
        DefaultForNew = $true
    },
    [pscustomobject]@{
        Id            = "Folders"
        Label         = Get-Text -Key "ScopeFolders"
        UserPath      = "Registry::HKEY_CURRENT_USER\Software\Classes\Folder\shell"
        MachinePath   = "Registry::HKEY_LOCAL_MACHINE\Software\Classes\Folder\shell"
        CommandArg    = "%1"
        DefaultForNew = $false
    },
    [pscustomobject]@{
        Id            = "Drives"
        Label         = Get-Text -Key "ScopeDrives"
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

    $hash = Get-CustomAppKeyHash -PathText $PathText
    return (Get-CustomAppKeyNameFromHash -Hash $hash)
}

function Resolve-ExePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText
    )

    if ([string]::IsNullOrWhiteSpace($PathText)) {
        throw (Get-Text -Key "ErrorExePathEmpty")
    }

    if (-not (Test-Path -LiteralPath $PathText -PathType Leaf)) {
        throw (Get-Text -Key "ErrorExeNotFound" -Arguments $PathText)
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
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    Ensure-RegistryKey -Path $Path

    if ($null -eq $Value) {
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
    return "$appName で開く"
}

function Get-CustomAppKeyHash {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText
    )

    $normalized = [System.IO.Path]::GetFullPath($PathText).ToLowerInvariant()
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalized)
        $hashBytes = $sha1.ComputeHash($bytes)
        return [System.BitConverter]::ToString($hashBytes).Replace("-", "")
    }
    finally {
        $sha1.Dispose()
    }
}

function Get-CustomAppKeyNameFromHash {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Hash,

        [Parameter(Mandatory = $false)]
        [switch]$Legacy
    )

    $prefix = if ($Legacy) { $script:LegacyCustomAppKeyPrefix } else { $script:CustomAppKeyPrefix }
    return "$prefix$Hash"
}

function Get-CustomAppKeyNamesForPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathText
    )

    $hash = Get-CustomAppKeyHash -PathText $PathText
    return @(
        (Get-CustomAppKeyNameFromHash -Hash $hash),
        (Get-CustomAppKeyNameFromHash -Hash $hash -Legacy)
    )
}

function Test-IsCustomAppKeyName {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$KeyName
    )

    if ([string]::IsNullOrWhiteSpace($KeyName)) {
        return $false
    }

    return (
        $KeyName.StartsWith($script:CustomAppKeyPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or
        $KeyName.StartsWith($script:LegacyCustomAppKeyPrefix, [System.StringComparison]::OrdinalIgnoreCase)
    )
}

function Get-MigratedCustomAppKeyName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$KeyName
    )

    if ($KeyName.StartsWith($script:LegacyCustomAppKeyPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        return ($script:CustomAppKeyPrefix + $KeyName.Substring($script:LegacyCustomAppKeyPrefix.Length))
    }

    return $KeyName
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
            throw (Get-Text -Key "ErrorInvalidScope" -Arguments $scopeId, $valid)
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
        throw (Get-Text -Key "ErrorNoScopeSelected")
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
    $keyNames = @(Get-CustomAppKeyNamesForPath -PathText $resolvedPath)
    $removedPaths = @()

    foreach ($scope in $script:ScopeCatalog) {
        foreach ($keyName in $keyNames) {
            $target = Join-Path $scope.UserPath $keyName
            if (Test-Path -LiteralPath $target) {
                Remove-Item -LiteralPath $target -Recurse -Force
                $removedPaths += $target
            }
        }
    }

    return [pscustomobject]@{
        KeyName       = $keyNames[0]
        ExePath       = $resolvedPath
        RemovedPaths  = $removedPaths
        WasRemoved    = ($removedPaths.Count -gt 0)
    }
}

function Migrate-LegacyCustomAppKeys {
    foreach ($scope in $script:ScopeCatalog) {
        if (-not (Test-Path -LiteralPath $scope.UserPath)) {
            continue
        }

        try {
            $children = @(Get-ChildItem -LiteralPath $scope.UserPath -ErrorAction Stop)
        }
        catch {
            continue
        }

        foreach ($child in $children) {
            $legacyKeyName = [string]$child.PSChildName
            if (-not $legacyKeyName.StartsWith($script:LegacyCustomAppKeyPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            $sourcePath = $child.PSPath
            $targetKeyName = Get-MigratedCustomAppKeyName -KeyName $legacyKeyName
            $targetPath = Join-Path $scope.UserPath $targetKeyName

            if ($sourcePath -eq $targetPath) {
                continue
            }

            if (-not (Test-Path -LiteralPath $targetPath)) {
                Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Recurse -Force
            }

            Remove-Item -LiteralPath $sourcePath -Recurse -Force
        }
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
        throw (Get-Text -Key "ErrorSelectedKeyNotFound")
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
    $rootIndex = 0

    foreach ($scope in $script:ScopeCatalog) {
        $roots += [pscustomobject]@{
                EnumerationRootIndex = $rootIndex
                ScopeId              = $scope.Id
                ScopeLabel           = $scope.Label
                SourceKey            = "System"
                Registry             = $scope.MachinePath
            }
        $rootIndex++

        $roots += [pscustomobject]@{
                EnumerationRootIndex = $rootIndex
                ScopeId              = $scope.Id
                ScopeLabel           = $scope.Label
                SourceKey            = "User"
                Registry             = $scope.UserPath
            }
        $rootIndex++
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
        return (Get-Text -Key "InvocationDelegateExecute" -Arguments $delegateExecute)
    }

    $explorerCommandHandler = Get-RegistryValue -Path $EntryPath -Name "ExplorerCommandHandler"
    if (-not [string]::IsNullOrWhiteSpace($explorerCommandHandler)) {
        return (Get-Text -Key "InvocationExplorerCommandHandler" -Arguments $explorerCommandHandler)
    }

    $subCommands = Get-RegistryValue -Path $EntryPath -Name "SubCommands"
    if (-not [string]::IsNullOrWhiteSpace($subCommands)) {
        return (Get-Text -Key "InvocationSubCommands" -Arguments $subCommands)
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
            $childNames = @(Get-NativeRegistrySubKeyNames -RegistryPath $root.Registry)
        }
        catch {
            continue
        }

        if ($childNames.Count -eq 0) {
            continue
        }

        $verbOrderTokens = @(Get-ShellVerbOrderSpecification -ShellRegistryPath $root.Registry)
        $verbOrderLookup = @{}
        for ($verbIndex = 0; $verbIndex -lt $verbOrderTokens.Count; $verbIndex++) {
            $verbOrderLookup[$verbOrderTokens[$verbIndex].ToLowerInvariant()] = $verbIndex
        }

        $rootEntries = @()
        $nativeOrderIndex = 0

        foreach ($childName in $childNames) {
            $entryPath = Join-Path $root.Registry $childName

            try {
                $item = Get-Item -LiteralPath $entryPath -ErrorAction Stop
            }
            catch {
                continue
            }

            $displayNameRaw = [string]$item.GetValue("")
            $displayName = $displayNameRaw
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                $displayNameRaw = [string]$item.GetValue("MUIVerb")
                $displayName = $displayNameRaw
            }
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                $displayNameRaw = $childName
                $displayName = $childName
            }
            else {
                $displayName = Resolve-IndirectShellString -Text ([string]$displayName)
            }

            $isDisabled = ($item.Property -contains "LegacyDisable")
            $commandSummary = Get-InvocationSummary -EntryPath $entryPath
            $iconValue = [string]$item.GetValue("Icon")
            $positionKey = Get-NormalizedPositionKey -PositionValue ([string]$item.GetValue("Position"))
            $isExtended = ($item.Property -contains "Extended")
            $isProgrammaticOnly = ($item.Property -contains "ProgrammaticAccessOnly")

            if ($isProgrammaticOnly -or $isExtended) {
                $nativeOrderIndex++
                continue
            }

            $sourceKey = [string]$root.SourceKey
            $statusKey = if ($isDisabled) { "Disabled" } else { "Enabled" }
            $typeKey = if (Test-IsCustomAppKeyName -KeyName $childName) { "CustomApp" } else { "Existing" }
            $verbLookupKey = $childName.ToLowerInvariant()
            $hasExplicitOrder = $verbOrderLookup.ContainsKey($verbLookupKey)
            $explicitOrderIndex = if ($hasExplicitOrder) { [int]$verbOrderLookup[$verbLookupKey] } else { [int]::MaxValue }
            $groupRank =
                if ($positionKey -eq "Top") {
                    0
                }
                elseif ($positionKey -eq "Bottom") {
                    3
                }
                elseif ($hasExplicitOrder) {
                    1
                }
                else {
                    2
                }

            $entry = [pscustomobject]@{
                DisplayName           = [string]$displayName
                DisplayNameRaw        = [string]$displayNameRaw
                ScopeId               = $root.ScopeId
                ScopeDisplay          = [string]$root.ScopeLabel
                SourceKey             = $sourceKey
                SourceDisplay         = Get-SourceDisplayText -SourceKey $sourceKey
                StatusKey             = $statusKey
                StatusDisplay         = Get-StatusDisplayText -StatusKey $statusKey
                TypeKey               = $typeKey
                TypeDisplay           = Get-TypeDisplayText -TypeKey $typeKey
                KeyName               = $childName
                RegistryPath          = $entryPath
                Command               = $commandSummary
                CommandPreview        = Get-CompactSingleLineText -Text $commandSummary -MaxLength 90
                Icon                  = $iconValue
                IconSource            = Resolve-ContextMenuEntryIconSource -IconSpec $iconValue -CommandText $commandSummary
                PositionKey           = $positionKey
                ExplicitOrderIndex    = $explicitOrderIndex
                HasExplicitOrder      = $hasExplicitOrder
                GroupRank             = $groupRank
                EnumerationRootIndex  = [int]$root.EnumerationRootIndex
                NativeOrderIndex      = [int]$nativeOrderIndex
                IsDisabled            = $isDisabled
                IsCustomApp           = ($typeKey -eq "CustomApp")
            }

            $rootEntries += $entry
            $nativeOrderIndex++
        }

        $entries += @($rootEntries | Sort-Object GroupRank, ExplicitOrderIndex, NativeOrderIndex)
    }

    return @($entries)
}

function Test-IsAdministrator {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Restart-ScriptAsAdministrator {
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$AdditionalArguments,

        [Parameter(Mandatory = $false)]
        [switch]$HideWindow
    )

    $scriptPath = if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        $PSCommandPath
    }
    else {
        $MyInvocation.MyCommand.Path
    }

    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        throw (Get-Text -Key "ErrorCannotDetermineScriptPath")
    }

    $argumentList = @(
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        (Format-ProcessArgument -Value $scriptPath)
    )

    if ($null -ne $AdditionalArguments -and $AdditionalArguments.Count -gt 0) {
        $argumentList += $AdditionalArguments
    }

    $startProcessParameters = @{
        FilePath     = "powershell.exe"
        ArgumentList = $argumentList
        Verb         = "RunAs"
        PassThru     = $true
    }

    if ($HideWindow) {
        $startProcessParameters.WindowStyle = "Hidden"
    }

    $process = Start-Process @startProcessParameters
    return ($null -ne $process)
}

function Restart-ScriptInHiddenWindow {
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$AdditionalArguments
    )

    $scriptPath = if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        $PSCommandPath
    }
    else {
        $MyInvocation.MyCommand.Path
    }

    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        throw (Get-Text -Key "ErrorCannotDetermineScriptPath")
    }

    $argumentList = @(
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-WindowStyle",
        "Hidden",
        "-File",
        (Format-ProcessArgument -Value $scriptPath)
    )

    if ($null -ne $AdditionalArguments -and $AdditionalArguments.Count -gt 0) {
        $argumentList += $AdditionalArguments
    }

    $process = Start-Process -FilePath "powershell.exe" -ArgumentList $argumentList -WindowStyle Hidden -PassThru
    return ($null -ne $process)
}

function Format-ProcessArgument {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Value
    )

    if ($Value.IndexOfAny([char[]]@(' ', "`t", '"')) -lt 0) {
        return $Value
    }

    return '"' + ($Value -replace '"', '\"') + '"'
}

function Convert-BoundParametersToArgumentList {
    param(
        [Parameter(Mandatory = $false)]
        [System.Collections.IDictionary]$BoundParameters
    )

    if ($null -eq $BoundParameters -or $BoundParameters.Count -eq 0) {
        return @()
    }

    $parameterOrder = @("AppPath", "Action", "Scopes", "Gui", "HiddenGuiHost")
    $argumentList = @()

    foreach ($parameterName in $parameterOrder) {
        if (-not $BoundParameters.ContainsKey($parameterName)) {
            continue
        }

        $value = $BoundParameters[$parameterName]

        if ($value -is [System.Management.Automation.SwitchParameter]) {
            if ($value.IsPresent) {
                $argumentList += "-$parameterName"
            }

            continue
        }

        if ($null -eq $value) {
            continue
        }

        if ($value -is [System.Array]) {
            if ($value.Count -eq 0) {
                continue
            }

            $argumentList += "-$parameterName"
            foreach ($item in $value) {
                $argumentList += (Format-ProcessArgument -Value ([string]$item))
            }

            continue
        }

        $argumentList += "-$parameterName"
        $argumentList += (Format-ProcessArgument -Value ([string]$value))
    }

    return $argumentList
}

function Ensure-AdministratorSession {
    param(
        [Parameter(Mandatory = $false)]
        [System.Collections.IDictionary]$BoundParameters,

        [Parameter(Mandatory = $false)]
        [switch]$HideWindow
    )

    if (Test-IsAdministrator) {
        return [pscustomobject]@{
            CanContinue = $true
            Relaunched  = $false
            Cancelled   = $false
        }
    }

    try {
        $argumentList = Convert-BoundParametersToArgumentList -BoundParameters $BoundParameters
        $started = Restart-ScriptAsAdministrator -AdditionalArguments $argumentList -HideWindow:$HideWindow

        if ($started) {
            return [pscustomobject]@{
                CanContinue = $false
                Relaunched  = $true
                Cancelled   = $false
            }
        }

        return [pscustomobject]@{
            CanContinue = $false
            Relaunched  = $false
            Cancelled   = $false
        }
    }
    catch [System.ComponentModel.Win32Exception] {
        if ($_.Exception.NativeErrorCode -eq 1223) {
            return [pscustomobject]@{
                CanContinue = $false
                Relaunched  = $false
                Cancelled   = $true
            }
        }

        throw
    }
}

function Set-ConsoleWindowVisibility {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Visible
    )

    if (-not ("ContextMenuEditor.ConsoleWindow" -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;

namespace ContextMenuEditor {
    public static class ConsoleWindow {
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
}
"@
    }

    $consoleHandle = [ContextMenuEditor.ConsoleWindow]::GetConsoleWindow()
    if ($consoleHandle -eq [IntPtr]::Zero) {
        return
    }

    $showMode = if ($Visible) { 5 } else { 0 }
    [ContextMenuEditor.ConsoleWindow]::ShowWindow($consoleHandle, $showMode) | Out-Null
}

function Write-EntryList {
    $entries = Get-ContextMenuEntries
    if ($entries.Count -eq 0) {
        Write-Host (Get-Text -Key "ListNoneFound")
        return
    }

    $entries |
        Select-Object `
            @{ Name = (Get-Text -Key "ColumnDisplayName"); Expression = { $_.DisplayName } }, `
            @{ Name = (Get-Text -Key "ColumnScope"); Expression = { $_.ScopeDisplay } }, `
            @{ Name = (Get-Text -Key "ColumnSource"); Expression = { $_.SourceDisplay } }, `
            @{ Name = (Get-Text -Key "ColumnStatus"); Expression = { $_.StatusDisplay } }, `
            @{ Name = (Get-Text -Key "ColumnKeyName"); Expression = { $_.KeyName } } |
        Format-Table -AutoSize
}

function Show-NotificationMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [string]$Title = (Get-Text -Key "WindowTitle"),

        [Parameter(Mandatory = $false)]
        [ValidateSet("Information", "Warning", "Error")]
        [string]$Icon = "Information"
    )

    Add-Type -AssemblyName PresentationFramework
    $messageBoxIcon = [System.Enum]::Parse([System.Windows.MessageBoxImage], $Icon)

    [System.Windows.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.MessageBoxButton]::OK,
        $messageBoxIcon
    ) | Out-Null
}

function Show-ContextMenuManagerGui {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="$($script:Text.WindowTitle)"
    Width="1280"
    Height="780"
    MinWidth="1100"
    MinHeight="680"
    WindowStartupLocation="CenterScreen"
    Background="#F4F6F8"
    Foreground="#1F2937"
    FontFamily="Segoe UI"
    FontSize="13"
    UseLayoutRounding="True">

  <Window.Resources>
    <SolidColorBrush x:Key="PageBrush" Color="#F4F6F8"/>
    <SolidColorBrush x:Key="SurfaceBrush" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="SurfaceAltBrush" Color="#FBFCFE"/>
    <SolidColorBrush x:Key="SurfaceMutedBrush" Color="#F1F5F9"/>
    <SolidColorBrush x:Key="BorderBrush" Color="#D8E1EB"/>
    <SolidColorBrush x:Key="TextBrush" Color="#1F2937"/>
    <SolidColorBrush x:Key="MutedBrush" Color="#5E6C84"/>
    <SolidColorBrush x:Key="SubtleTextBrush" Color="#667085"/>
    <SolidColorBrush x:Key="AccentBrush" Color="#0F6CBD"/>
    <SolidColorBrush x:Key="AccentHoverBrush" Color="#115EA3"/>
    <SolidColorBrush x:Key="AccentSoftBrush" Color="#EAF3FF"/>
    <SolidColorBrush x:Key="SuccessTextBrush" Color="#0F6A4A"/>
    <SolidColorBrush x:Key="SuccessSoftBrush" Color="#E7F6EE"/>
    <SolidColorBrush x:Key="WarningTextBrush" Color="#8A6514"/>
    <SolidColorBrush x:Key="WarningSoftBrush" Color="#FFF4D6"/>
    <SolidColorBrush x:Key="DangerTextBrush" Color="#B42318"/>
    <SolidColorBrush x:Key="DangerSoftBrush" Color="#FEECEC"/>
    <SolidColorBrush x:Key="InfoTextBrush" Color="#0F6CBD"/>
    <SolidColorBrush x:Key="InfoSoftBrush" Color="#EAF3FF"/>
    <SolidColorBrush x:Key="ListHoverBrush" Color="#F5F9FF"/>
    <SolidColorBrush x:Key="ListSelectedBrush" Color="#E4F0FF"/>

    <Style x:Key="SurfaceCard" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="16"/>
      <Setter Property="SnapsToDevicePixels" Value="True"/>
    </Style>

    <Style x:Key="SectionTitle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>

    <Style x:Key="LabelText" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>

    <Style x:Key="ValueText" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>

    <Style x:Key="MetaText" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource SubtleTextBrush}"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>

    <Style x:Key="SecondaryButton" TargetType="Button">
      <Setter Property="Height" Value="34"/>
      <Setter Property="Padding" Value="14,0"/>
      <Setter Property="Background" Value="{StaticResource SurfaceAltBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="ButtonBorder"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="10"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="ButtonBorder" Property="Background" Value="{StaticResource AccentSoftBrush}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="ButtonBorder" Property="Background" Value="{StaticResource ListSelectedBrush}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.55"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="PrimaryButton" BasedOn="{StaticResource SecondaryButton}" TargetType="Button">
      <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="ButtonBorder"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="10"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="ButtonBorder" Property="Background" Value="{StaticResource AccentHoverBrush}"/>
                <Setter TargetName="ButtonBorder" Property="BorderBrush" Value="{StaticResource AccentHoverBrush}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="ButtonBorder" Property="Background" Value="{StaticResource AccentHoverBrush}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="DangerButton" BasedOn="{StaticResource SecondaryButton}" TargetType="Button">
      <Setter Property="Background" Value="{StaticResource DangerSoftBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource DangerTextBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource DangerSoftBrush}"/>
    </Style>

    <Style x:Key="SegmentButton" TargetType="ToggleButton">
      <Setter Property="Height" Value="32"/>
      <Setter Property="Margin" Value="0,0,8,6"/>
      <Setter Property="Padding" Value="12,0"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ToggleButton">
            <Border x:Name="ButtonBorder"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="10"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="ButtonBorder" Property="Background" Value="{StaticResource ListHoverBrush}"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="ButtonBorder" Property="Background" Value="{StaticResource AccentSoftBrush}"/>
                <Setter TargetName="ButtonBorder" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="InputTextBox" TargetType="TextBox">
      <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="12,8"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="CaretBrush" Value="{StaticResource AccentBrush}"/>
      <Setter Property="SelectionBrush" Value="{StaticResource AccentSoftBrush}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="TextBoxBorder"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="10">
              <ScrollViewer x:Name="PART_ContentHost"
                            Margin="{TemplateBinding Padding}"
                            Background="Transparent"
                            Focusable="False"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="TextBoxBorder" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="ReadOnlyTextBox" BasedOn="{StaticResource InputTextBox}" TargetType="TextBox">
      <Setter Property="Background" Value="{StaticResource SurfaceAltBrush}"/>
    </Style>

    <Style x:Key="ScopeCheckBox" TargetType="CheckBox">
      <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>

    <Style x:Key="CompactBadgeBorder" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource SurfaceMutedBrush}"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="7,3"/>
      <Setter Property="Margin" Value="0,0,6,2"/>
    </Style>

    <Style x:Key="CompactBadgeText" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>

    <Style x:Key="ListStatusBadgeBorder" BasedOn="{StaticResource CompactBadgeBorder}" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource SuccessSoftBrush}"/>
      <Style.Triggers>
        <DataTrigger Binding="{Binding StatusKey}" Value="Disabled">
          <Setter Property="Background" Value="{StaticResource WarningSoftBrush}"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="ListStatusBadgeText" BasedOn="{StaticResource CompactBadgeText}" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource SuccessTextBrush}"/>
      <Style.Triggers>
        <DataTrigger Binding="{Binding StatusKey}" Value="Disabled">
          <Setter Property="Foreground" Value="{StaticResource WarningTextBrush}"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="ListSourceBadgeBorder" BasedOn="{StaticResource CompactBadgeBorder}" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource InfoSoftBrush}"/>
      <Style.Triggers>
        <DataTrigger Binding="{Binding SourceKey}" Value="System">
          <Setter Property="Background" Value="{StaticResource WarningSoftBrush}"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="ListSourceBadgeText" BasedOn="{StaticResource CompactBadgeText}" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource InfoTextBrush}"/>
      <Style.Triggers>
        <DataTrigger Binding="{Binding SourceKey}" Value="System">
          <Setter Property="Foreground" Value="{StaticResource WarningTextBrush}"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="ListViewItem">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="8,7"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListViewItem">
            <Border x:Name="ItemBorder"
                    Background="{TemplateBinding Background}"
                    CornerRadius="10"
                    Margin="0,0,0,4"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="ItemBorder" Property="Background" Value="{StaticResource ListHoverBrush}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="ItemBorder" Property="Background" Value="{StaticResource ListSelectedBrush}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="320"/>
    </Grid.ColumnDefinitions>

    <Border Grid.Row="0" Grid.ColumnSpan="2" Style="{StaticResource SurfaceCard}" Padding="14" Margin="0,0,0,12">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="$($script:Text.WindowTitle)"
                     Foreground="{StaticResource TextBrush}"
                     FontSize="18"
                     FontWeight="SemiBold"
                     VerticalAlignment="Center"/>
          <Border Background="{StaticResource AccentSoftBrush}" CornerRadius="10" Padding="10,4" Margin="10,0,0,0" VerticalAlignment="Center">
            <TextBlock Text="$($script:Text.AdminBadge)"
                       Foreground="{StaticResource AccentBrush}"
                       FontSize="11"
                       FontWeight="SemiBold"/>
          </Border>
        </StackPanel>

        <TextBlock x:Name="CountLabel"
                   Grid.Row="0"
                   Grid.Column="2"
                   VerticalAlignment="Center"
                   Margin="0,0,10,0"
                   Foreground="{StaticResource MutedBrush}"
                   FontSize="11"
                   FontWeight="SemiBold"/>

        <Button x:Name="RefreshButton"
                Grid.Row="0"
                Grid.Column="3"
                Content="$($script:Text.RefreshButton)"
                Style="{StaticResource SecondaryButton}"
                Margin="0,0,8,0"/>

        <Button x:Name="AddButton"
                Grid.Row="0"
                Grid.Column="4"
                Content="$($script:Text.AddButton)"
                Style="{StaticResource PrimaryButton}"/>

        <Grid Grid.Row="1" Grid.ColumnSpan="5" Margin="0,12,0,0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="270"/>
            <ColumnDefinition Width="12"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <Grid Grid.Column="0">
            <TextBox x:Name="SearchBox" Style="{StaticResource InputTextBox}"/>
            <TextBlock x:Name="SearchPlaceholder"
                       Text="$($script:Text.SearchPlaceholder)"
                       Foreground="{StaticResource MutedBrush}"
                       Margin="12,0,12,0"
                       VerticalAlignment="Center"
                       FontSize="12"
                       IsHitTestVisible="False"/>
          </Grid>

          <WrapPanel x:Name="ScopeButtonPanel" Grid.Column="2" HorizontalAlignment="Right" VerticalAlignment="Center">
            <ToggleButton x:Name="ScopeBtn_All"
                          Content="$($script:Text.ScopeAll)"
                          Style="{StaticResource SegmentButton}"
                          IsChecked="True"
                          Tag="All"/>
          </WrapPanel>
        </Grid>
      </Grid>
    </Border>

    <Border Grid.Row="1" Grid.Column="0" Style="{StaticResource SurfaceCard}" Margin="0,0,12,0">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <DockPanel Margin="14,14,14,8">
          <TextBlock Text="$($script:Text.ListTitle)" Style="{StaticResource SectionTitle}"/>
        </DockPanel>

        <ListView x:Name="EntryList"
                  Grid.Row="1"
                  Margin="10,0,10,10"
                  Background="Transparent"
                  BorderThickness="0"
                  SelectionMode="Single"
                  ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                  ScrollViewer.CanContentScroll="True"
                  VirtualizingPanel.IsVirtualizing="True"
                  VirtualizingPanel.VirtualizationMode="Recycling">
        <ListView.ItemTemplate>
            <DataTemplate>
              <Grid ToolTip="{Binding Command}">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="34"/>
                  <ColumnDefinition Width="10"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <Border Grid.Column="0"
                        Width="34"
                        Height="34"
                        Background="{StaticResource SurfaceAltBrush}"
                        BorderBrush="{StaticResource BorderBrush}"
                        BorderThickness="1"
                        CornerRadius="8"
                        VerticalAlignment="Top">
                  <Image Source="{Binding IconSource}"
                         Width="18"
                         Height="18"
                         Stretch="Uniform"
                         HorizontalAlignment="Center"
                         VerticalAlignment="Center"
                         SnapsToDevicePixels="True"/>
                </Border>

                <Grid Grid.Column="2">
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>

                  <TextBlock Grid.Row="0"
                             Grid.Column="0"
                             Text="{Binding DisplayName}"
                             Foreground="{StaticResource TextBrush}"
                             FontSize="13"
                             FontWeight="SemiBold"
                             TextTrimming="CharacterEllipsis"
                             Margin="0,1,10,0"/>

                  <WrapPanel Grid.Row="0" Grid.Column="1" HorizontalAlignment="Right">
                    <Border Style="{StaticResource ListStatusBadgeBorder}">
                      <TextBlock Text="{Binding StatusDisplay}" Style="{StaticResource ListStatusBadgeText}"/>
                    </Border>
                    <Border Style="{StaticResource ListSourceBadgeBorder}">
                      <TextBlock Text="{Binding SourceDisplay}" Style="{StaticResource ListSourceBadgeText}"/>
                    </Border>
                    <Border>
                      <Border.Style>
                        <Style TargetType="Border" BasedOn="{StaticResource CompactBadgeBorder}">
                          <Setter Property="Visibility" Value="Visible"/>
                          <Setter Property="Background" Value="{StaticResource AccentSoftBrush}"/>
                          <Style.Triggers>
                            <DataTrigger Binding="{Binding HasTypeBadge}" Value="False">
                              <Setter Property="Visibility" Value="Collapsed"/>
                            </DataTrigger>
                          </Style.Triggers>
                        </Style>
                      </Border.Style>
                      <TextBlock Text="{Binding ListTypeDisplay}"
                                 Style="{StaticResource CompactBadgeText}"
                                 Foreground="{StaticResource AccentBrush}"/>
                    </Border>
                    <Border Margin="0,0,0,2">
                      <Border.Style>
                        <Style TargetType="Border" BasedOn="{StaticResource CompactBadgeBorder}">
                          <Setter Property="Visibility" Value="Visible"/>
                          <Style.Triggers>
                            <DataTrigger Binding="{Binding HasScopeBadge}" Value="False">
                              <Setter Property="Visibility" Value="Collapsed"/>
                            </DataTrigger>
                          </Style.Triggers>
                        </Style>
                      </Border.Style>
                      <TextBlock Text="{Binding ListScopeDisplay}" Style="{StaticResource CompactBadgeText}"/>
                    </Border>
                  </WrapPanel>

                  <Grid Grid.Row="1" Grid.ColumnSpan="2" Margin="0,5,0,0">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="160"/>
                      <ColumnDefinition Width="10"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <TextBlock Grid.Column="0"
                               Text="{Binding KeyName}"
                               Style="{StaticResource MetaText}"
                               FontWeight="SemiBold"
                               ToolTip="{Binding KeyName}"/>

                    <TextBlock Grid.Column="2"
                               Text="{Binding CommandPreview}"
                               Style="{StaticResource MetaText}"
                               ToolTip="{Binding Command}"/>
                  </Grid>
                </Grid>
              </Grid>
            </DataTemplate>
          </ListView.ItemTemplate>
        </ListView>
      </Grid>
    </Border>

    <Border Grid.Row="1" Grid.Column="1" Style="{StaticResource SurfaceCard}" Padding="16">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Text="$($script:Text.DetailTitle)" Style="{StaticResource SectionTitle}"/>

        <Grid Grid.Row="1" Margin="0,14,0,14">
          <StackPanel x:Name="DetailEmptyState"
                      VerticalAlignment="Center"
                      HorizontalAlignment="Center"
                      Width="220">
            <TextBlock Text="$($script:Text.DetailEmptyTitle)"
                       TextAlignment="Center"
                       FontSize="15"
                       FontWeight="SemiBold"
                       Foreground="{StaticResource TextBrush}"/>
            <TextBlock Text="$($script:Text.DetailEmptyDescription)"
                       TextAlignment="Center"
                       Foreground="{StaticResource MutedBrush}"
                       FontSize="12"
                       Margin="0,8,0,0"
                       TextWrapping="Wrap"/>
          </StackPanel>

          <ScrollViewer x:Name="DetailContentHost"
                        Visibility="Collapsed"
                        VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock Text="$($script:Text.DetailDisplayName)" Style="{StaticResource LabelText}"/>
              <TextBlock x:Name="DetailName" Style="{StaticResource ValueText}" Margin="0,4,0,14"/>

              <Grid Margin="0,0,0,14">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Margin="0,0,8,0">
                  <TextBlock Text="$($script:Text.DetailScope)" Style="{StaticResource LabelText}"/>
                  <TextBlock x:Name="DetailScope" Style="{StaticResource ValueText}" Margin="0,4,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="1">
                  <TextBlock Text="$($script:Text.DetailSource)" Style="{StaticResource LabelText}"/>
                  <Border x:Name="SourceBadge"
                          Background="{StaticResource InfoSoftBrush}"
                          CornerRadius="8"
                          Padding="8,4"
                          Margin="0,4,0,0"
                          HorizontalAlignment="Left">
                    <TextBlock x:Name="DetailSource"
                               Foreground="{StaticResource InfoTextBrush}"
                               FontSize="11"
                               FontWeight="SemiBold"/>
                  </Border>
                </StackPanel>
              </Grid>

              <Grid Margin="0,0,0,14">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Margin="0,0,8,0">
                  <TextBlock Text="$($script:Text.DetailStatus)" Style="{StaticResource LabelText}"/>
                  <Border x:Name="StatusBadge"
                          Background="{StaticResource SuccessSoftBrush}"
                          CornerRadius="8"
                          Padding="8,4"
                          Margin="0,4,0,0"
                          HorizontalAlignment="Left">
                    <TextBlock x:Name="DetailStatus"
                               Foreground="{StaticResource SuccessTextBrush}"
                               FontSize="11"
                               FontWeight="SemiBold"/>
                  </Border>
                </StackPanel>

                <StackPanel Grid.Column="1">
                  <TextBlock Text="$($script:Text.DetailType)" Style="{StaticResource LabelText}"/>
                  <TextBlock x:Name="DetailType" Style="{StaticResource ValueText}" Margin="0,4,0,0"/>
                </StackPanel>
              </Grid>

              <TextBlock Text="$($script:Text.DetailKeyName)" Style="{StaticResource LabelText}"/>
              <TextBlock x:Name="DetailKeyName" Style="{StaticResource ValueText}" Margin="0,4,0,14"/>

              <TextBlock Text="$($script:Text.DetailPath)" Style="{StaticResource LabelText}"/>
              <TextBox x:Name="DetailPath"
                       Style="{StaticResource ReadOnlyTextBox}"
                       IsReadOnly="True"
                       TextWrapping="Wrap"
                       MaxHeight="74"
                       Margin="0,4,0,14"/>

              <TextBlock Text="$($script:Text.DetailIcon)" Style="{StaticResource LabelText}"/>
              <TextBox x:Name="DetailIcon"
                       Style="{StaticResource ReadOnlyTextBox}"
                       IsReadOnly="True"
                       TextWrapping="Wrap"
                       MaxHeight="56"
                       Margin="0,4,0,14"/>

              <TextBlock Text="$($script:Text.DetailCommand)" Style="{StaticResource LabelText}"/>
              <TextBox x:Name="DetailCommand"
                       Style="{StaticResource ReadOnlyTextBox}"
                       IsReadOnly="True"
                       TextWrapping="Wrap"
                       MaxHeight="88"
                       Margin="0,4,0,0"/>
            </StackPanel>
          </ScrollViewer>
        </Grid>

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="EnableButton"
                  Content="$($script:Text.EnableButton)"
                  Style="{StaticResource PrimaryButton}"
                  Margin="0,0,8,0"
                  IsEnabled="False"/>
          <Button x:Name="DisableButton"
                  Content="$($script:Text.DisableButton)"
                  Style="{StaticResource SecondaryButton}"
                  Margin="0,0,8,0"
                  IsEnabled="False"/>
          <Button x:Name="DeleteButton"
                  Content="$($script:Text.DeleteButton)"
                  Style="{StaticResource DangerButton}"
                  IsEnabled="False"/>
        </StackPanel>
      </Grid>
    </Border>

    <Grid x:Name="ModalOverlay" Grid.RowSpan="2" Grid.ColumnSpan="2" Visibility="Collapsed" Panel.ZIndex="100">
      <Border Background="#6A0F172A"/>

      <Border Width="520"
              Style="{StaticResource SurfaceCard}"
              Padding="22"
              HorizontalAlignment="Center"
              VerticalAlignment="Center">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <DockPanel Grid.Row="0" Margin="0,0,0,16">
            <Button x:Name="ModalCloseButton"
                    DockPanel.Dock="Right"
                    Content="X"
                    Width="34"
                    Height="34"
                    Padding="0"
                    Style="{StaticResource SecondaryButton}"/>

            <StackPanel>
              <TextBlock Text="$($script:Text.ModalTitle)"
                         FontSize="18"
                         FontWeight="SemiBold"
                         Foreground="{StaticResource TextBrush}"/>
              <TextBlock Text="$($script:Text.ModalDescription)"
                         Foreground="{StaticResource MutedBrush}"
                         FontSize="11"
                         Margin="0,4,0,0"/>
            </StackPanel>
          </DockPanel>

          <StackPanel Grid.Row="1">
            <TextBlock Text="$($script:Text.ExePathLabel)" Style="{StaticResource LabelText}"/>
            <Grid Margin="0,4,0,14">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>

              <TextBox x:Name="NewExePath"
                       Grid.Column="0"
                       Style="{StaticResource InputTextBox}"
                       Margin="0,0,10,0"/>

              <Button x:Name="BrowseButton"
                      Grid.Column="1"
                      Content="$($script:Text.BrowseButton)"
                      Style="{StaticResource SecondaryButton}"/>
            </Grid>
          </StackPanel>

          <StackPanel Grid.Row="2">
            <TextBlock Text="$($script:Text.MenuLabelLabel)" Style="{StaticResource LabelText}"/>
            <TextBox x:Name="NewMenuLabel" Style="{StaticResource InputTextBox}" Margin="0,4,0,14"/>
          </StackPanel>

          <StackPanel Grid.Row="3">
            <TextBlock Text="$($script:Text.ScopeLabel)" Style="{StaticResource LabelText}"/>
            <ItemsControl x:Name="ScopeCheckList" Margin="0,6,0,0"/>

            <Border x:Name="AddValidationBadge"
                    Background="{StaticResource WarningSoftBrush}"
                    CornerRadius="10"
                    Padding="12,8"
                    Margin="0,12,0,0">
              <TextBlock x:Name="AddValidationText"
                         Text="$($script:Text.ValidationMissingExe)"
                         Foreground="{StaticResource WarningTextBrush}"
                         FontSize="11"
                         TextWrapping="Wrap"/>
            </Border>

            <TextBlock Text="$($script:Text.ModalTip)"
                       Foreground="{StaticResource MutedBrush}"
                       FontSize="11"
                       Margin="0,10,0,0"
                       TextWrapping="Wrap"/>
          </StackPanel>

          <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,18,0,0">
            <Button x:Name="ModalCancelButton"
                    Content="$($script:Text.ModalCancelButton)"
                    Style="{StaticResource SecondaryButton}"
                    Margin="0,0,10,0"/>
            <Button x:Name="ModalAddButton"
                    Content="$($script:Text.ModalAddButton)"
                    Style="{StaticResource PrimaryButton}"
                    IsEnabled="False"
                    IsDefault="True"/>
          </StackPanel>
        </Grid>
      </Border>
    </Grid>
  </Grid>
</Window>
"@

    $xaml.Window.RemoveAttribute('x:Class')
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $ScopeButtonPanel  = $window.FindName('ScopeButtonPanel')
    $ScopeBtn_All      = $window.FindName('ScopeBtn_All')
    $SearchBox         = $window.FindName('SearchBox')
    $SearchPlaceholder = $window.FindName('SearchPlaceholder')
    $CountLabel        = $window.FindName('CountLabel')
    $RefreshButton     = $window.FindName('RefreshButton')
    $EntryList         = $window.FindName('EntryList')
    $AddButton         = $window.FindName('AddButton')
    $DetailEmptyState  = $window.FindName('DetailEmptyState')
    $DetailContentHost = $window.FindName('DetailContentHost')
    $DetailName        = $window.FindName('DetailName')
    $DetailScope       = $window.FindName('DetailScope')
    $DetailSource      = $window.FindName('DetailSource')
    $SourceBadge       = $window.FindName('SourceBadge')
    $DetailStatus      = $window.FindName('DetailStatus')
    $StatusBadge       = $window.FindName('StatusBadge')
    $DetailType        = $window.FindName('DetailType')
    $DetailKeyName     = $window.FindName('DetailKeyName')
    $DetailPath        = $window.FindName('DetailPath')
    $DetailIcon        = $window.FindName('DetailIcon')
    $DetailCommand     = $window.FindName('DetailCommand')
    $EnableButton      = $window.FindName('EnableButton')
    $DisableButton     = $window.FindName('DisableButton')
    $DeleteButton      = $window.FindName('DeleteButton')
    $ModalOverlay      = $window.FindName('ModalOverlay')
    $ModalCloseButton  = $window.FindName('ModalCloseButton')
    $BrowseButton      = $window.FindName('BrowseButton')
    $NewExePath        = $window.FindName('NewExePath')
    $NewMenuLabel      = $window.FindName('NewMenuLabel')
    $ScopeCheckList    = $window.FindName('ScopeCheckList')
    $AddValidationBadge = $window.FindName('AddValidationBadge')
    $AddValidationText = $window.FindName('AddValidationText')
    $ModalCancelButton = $window.FindName('ModalCancelButton')
    $ModalAddButton    = $window.FindName('ModalAddButton')

    $state = @{
        AllEntries                = @()
        ActiveScopeId             = 'All'
        LastSuggestedLabel        = ''
        MenuLabelManuallyEdited   = $false
        SuppressMenuLabelChanged  = $false
    }

    $scopeLabels = @{ 'All' = Get-Text -Key 'ScopeAll' }
    foreach ($scope in $script:ScopeCatalog) {
        $scopeLabels[$scope.Id] = [string]$scope.Label
    }

    foreach ($scope in $script:ScopeCatalog) {
        $btn = [System.Windows.Controls.Primitives.ToggleButton]::new()
        $btn.Content = $scope.Label
        $btn.Tag = $scope.Id
        $btn.Style = $window.Resources['SegmentButton']
        $btn.Add_Click({
            param($sender, $e)
            Set-ActiveScope -ScopeId ([string]$sender.Tag)
        })
        [void]$ScopeButtonPanel.Children.Add($btn)
    }

    foreach ($scope in $script:ScopeCatalog) {
        $checkBox = [System.Windows.Controls.CheckBox]::new()
        $checkBox.Content = $scope.Label
        $checkBox.Tag = $scope.Id
        $checkBox.IsChecked = $scope.DefaultForNew
        $checkBox.Style = $window.Resources['ScopeCheckBox']
        $checkBox.Add_Checked({ Update-AddDialogState })
        $checkBox.Add_Unchecked({ Update-AddDialogState })
        [void]$ScopeCheckList.Items.Add($checkBox)
    }

    function Show-Message {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter(Mandatory = $false)]
            [string]$Title = (Get-Text -Key 'WindowTitle'),

            [Parameter(Mandatory = $false)]
            [System.Windows.MessageBoxImage]$Icon = [System.Windows.MessageBoxImage]::Information
        )

        [System.Windows.MessageBox]::Show(
            $window,
            $Message,
            $Title,
            [System.Windows.MessageBoxButton]::OK,
            $Icon
        ) | Out-Null
    }

    function Update-SearchPlaceholder {
        $SearchPlaceholder.Visibility =
            if ([string]::IsNullOrWhiteSpace($SearchBox.Text)) {
                [System.Windows.Visibility]::Visible
            }
            else {
                [System.Windows.Visibility]::Collapsed
            }
    }

    function Clear-EntryDetails {
        $DetailName.Text = ""
        $DetailScope.Text = ""
        $DetailSource.Text = ""
        $DetailStatus.Text = ""
        $DetailType.Text = ""
        $DetailKeyName.Text = ""
        $DetailPath.Text = ""
        $DetailIcon.Text = ""
        $DetailCommand.Text = ""
        $DetailEmptyState.Visibility = [System.Windows.Visibility]::Visible
        $DetailContentHost.Visibility = [System.Windows.Visibility]::Collapsed
        $EnableButton.IsEnabled = $false
        $DisableButton.IsEnabled = $false
        $DeleteButton.IsEnabled = $false
    }

    function Get-SelectedEntry {
        return $EntryList.SelectedItem
    }

    function Update-CountLabel {
        $count = if ($null -ne $EntryList.ItemsSource) { @($EntryList.ItemsSource).Count } else { 0 }
        $CountLabel.Text = Get-Text -Key 'CountFormat' -Arguments $count, $scopeLabels[$state.ActiveScopeId]
    }

    function Get-VisibleEntriesForScope {
        param(
            [Parameter(Mandatory = $true)]
            [string]$ScopeId,

            [Parameter(Mandatory = $false)]
            [string]$Filter
        )

        $entries =
            if ($ScopeId -eq 'All') {
                $state.AllEntries
            }
            else {
                @($state.AllEntries | Where-Object { $_.ScopeId -eq $ScopeId })
            }

        if ([string]::IsNullOrWhiteSpace($Filter)) {
            return @($entries)
        }

        $comparison = [System.StringComparison]::OrdinalIgnoreCase
        return @($entries | Where-Object {
                $_.DisplayName.IndexOf($Filter, $comparison)  -ge 0 -or
                $_.ScopeDisplay.IndexOf($Filter, $comparison) -ge 0 -or
                $_.SourceDisplay.IndexOf($Filter, $comparison) -ge 0 -or
                $_.StatusDisplay.IndexOf($Filter, $comparison) -ge 0 -or
                $_.TypeDisplay.IndexOf($Filter, $comparison)  -ge 0 -or
                $_.KeyName.IndexOf($Filter, $comparison)      -ge 0 -or
                $_.Command.IndexOf($Filter, $comparison)      -ge 0
            })
    }

    function Set-EntryListPresentationData {
        param(
            [Parameter(Mandatory = $true)]
            [object[]]$Entries,

            [Parameter(Mandatory = $true)]
            [string]$ScopeId
        )

        $showScopeBadge = ($ScopeId -eq 'All')
        foreach ($entry in $Entries) {
            $listScopeDisplay = if ($showScopeBadge) { [string]$entry.ScopeDisplay } else { '' }
            $listTypeDisplay = if ([bool]$entry.IsCustomApp) { '+' } else { '' }

            Add-Member -InputObject $entry -NotePropertyName 'ListScopeDisplay' -NotePropertyValue $listScopeDisplay -Force
            Add-Member -InputObject $entry -NotePropertyName 'HasScopeBadge' -NotePropertyValue $showScopeBadge -Force
            Add-Member -InputObject $entry -NotePropertyName 'ListTypeDisplay' -NotePropertyValue $listTypeDisplay -Force
            Add-Member -InputObject $entry -NotePropertyName 'HasTypeBadge' -NotePropertyValue (-not [string]::IsNullOrWhiteSpace($listTypeDisplay)) -Force
        }

        return @($Entries)
    }

    function Set-SourceBadgeStyle {
        param(
            [Parameter(Mandatory = $true)]
            [string]$SourceKey
        )

        if ($SourceKey -eq 'System') {
            $SourceBadge.Background = $window.Resources['WarningSoftBrush']
            $DetailSource.Foreground = $window.Resources['WarningTextBrush']
            return
        }

        $SourceBadge.Background = $window.Resources['InfoSoftBrush']
        $DetailSource.Foreground = $window.Resources['InfoTextBrush']
    }

    function Set-StatusBadgeStyle {
        param(
            [Parameter(Mandatory = $true)]
            [string]$StatusKey
        )

        if ($StatusKey -eq 'Disabled') {
            $StatusBadge.Background = $window.Resources['WarningSoftBrush']
            $DetailStatus.Foreground = $window.Resources['WarningTextBrush']
            return
        }

        $StatusBadge.Background = $window.Resources['SuccessSoftBrush']
        $DetailStatus.Foreground = $window.Resources['SuccessTextBrush']
    }

    function Update-EntryDetails {
        $selectedEntry = Get-SelectedEntry
        if ($null -eq $selectedEntry) {
            Clear-EntryDetails
            return
        }

        $DetailName.Text = [string]$selectedEntry.DisplayName
        $DetailScope.Text = [string]$selectedEntry.ScopeDisplay
        $DetailSource.Text = [string]$selectedEntry.SourceDisplay
        $DetailStatus.Text = [string]$selectedEntry.StatusDisplay
        $DetailType.Text = [string]$selectedEntry.TypeDisplay
        $DetailKeyName.Text = [string]$selectedEntry.KeyName
        $DetailPath.Text = [string]$selectedEntry.RegistryPath
        $DetailIcon.Text = [string]$selectedEntry.Icon
        $DetailCommand.Text = [string]$selectedEntry.Command

        Set-SourceBadgeStyle -SourceKey ([string]$selectedEntry.SourceKey)
        Set-StatusBadgeStyle -StatusKey ([string]$selectedEntry.StatusKey)

        $DetailEmptyState.Visibility = [System.Windows.Visibility]::Collapsed
        $DetailContentHost.Visibility = [System.Windows.Visibility]::Visible
        $EnableButton.IsEnabled = [bool]$selectedEntry.IsDisabled
        $DisableButton.IsEnabled = -not [bool]$selectedEntry.IsDisabled
        $DeleteButton.IsEnabled = $true
    }

    function Update-EntryList {
        param(
            [Parameter(Mandatory = $false)]
            [string]$SelectedPath
        )

        if ([string]::IsNullOrWhiteSpace($SelectedPath)) {
            $currentSelection = Get-SelectedEntry
            if ($null -ne $currentSelection) {
                $SelectedPath = [string]$currentSelection.RegistryPath
            }
        }

        $filter = [string]$SearchBox.Text
        $entries = Get-VisibleEntriesForScope -ScopeId $state.ActiveScopeId -Filter $filter.Trim()
        $entries = Set-EntryListPresentationData -Entries $entries -ScopeId $state.ActiveScopeId

        $EntryList.ItemsSource = $null
        $EntryList.ItemsSource = $entries
        Update-CountLabel

        if (-not [string]::IsNullOrWhiteSpace($SelectedPath)) {
            $target = $entries | Where-Object { $_.RegistryPath -eq $SelectedPath } | Select-Object -First 1
            if ($null -ne $target) {
                $EntryList.SelectedItem = $target
                $EntryList.ScrollIntoView($target)
                return
            }
        }

        if ($entries.Count -gt 0) {
            $EntryList.SelectedItem = $entries[0]
            $EntryList.ScrollIntoView($entries[0])
            return
        }

        $EntryList.SelectedItem = $null
        Clear-EntryDetails
    }

    function Refresh-Entries {
        param(
            [Parameter(Mandatory = $false)]
            [string]$SelectedPath
        )

        try {
            $window.Cursor = [System.Windows.Input.Cursors]::Wait
            $state.AllEntries = @(Get-ContextMenuEntries)
            Update-EntryList -SelectedPath $SelectedPath
        }
        catch {
            Show-Message -Message $_.Exception.Message -Title (Get-Text -Key 'MessageLoadErrorTitle') -Icon ([System.Windows.MessageBoxImage]::Error)
        }
        finally {
            $window.Cursor = $null
        }
    }

    function Get-CheckedScopeIds {
        $selectedScopeIds = @()
        foreach ($item in $ScopeCheckList.Items) {
            if ($item.IsChecked -eq $true) {
                $selectedScopeIds += [string]$item.Tag
            }
        }

        return $selectedScopeIds
    }

    function Set-MenuLabelText {
        param(
            [Parameter(Mandatory = $false)]
            [AllowEmptyString()]
            [string]$Value
        )

        $state.SuppressMenuLabelChanged = $true
        $NewMenuLabel.Text = $Value
        $state.SuppressMenuLabelChanged = $false
    }

    function Update-SuggestedMenuLabel {
        $pathText = $NewExePath.Text.Trim()
        if (-not (Test-ExeFilePath -PathText $pathText)) {
            return
        }

        $suggestedLabel = Get-DefaultMenuLabel -ExePath $pathText
        $shouldApply =
            (-not $state.MenuLabelManuallyEdited) -or
            [string]::IsNullOrWhiteSpace($NewMenuLabel.Text) -or
            ($NewMenuLabel.Text -eq $state.LastSuggestedLabel)

        $state.LastSuggestedLabel = $suggestedLabel

        if ($shouldApply) {
            Set-MenuLabelText -Value $suggestedLabel
            $state.MenuLabelManuallyEdited = $false
        }
    }

    function Update-AddDialogState {
        $hasValidExe = Test-ExeFilePath -PathText $NewExePath.Text.Trim()
        $hasScope = (Get-CheckedScopeIds).Count -gt 0

        if (-not $hasValidExe) {
            $AddValidationBadge.Background = $window.Resources['WarningSoftBrush']
            $AddValidationText.Foreground = $window.Resources['WarningTextBrush']
            $AddValidationText.Text = Get-Text -Key 'ValidationMissingExe'
            $ModalAddButton.IsEnabled = $false
            return
        }

        if (-not $hasScope) {
            $AddValidationBadge.Background = $window.Resources['WarningSoftBrush']
            $AddValidationText.Foreground = $window.Resources['WarningTextBrush']
            $AddValidationText.Text = Get-Text -Key 'ValidationMissingScope'
            $ModalAddButton.IsEnabled = $false
            return
        }

        $AddValidationBadge.Background = $window.Resources['SuccessSoftBrush']
        $AddValidationText.Foreground = $window.Resources['SuccessTextBrush']
        $AddValidationText.Text = Get-Text -Key 'ValidationReady'
        $ModalAddButton.IsEnabled = $true
    }

    function Open-AddDialog {
        $state.LastSuggestedLabel = ''
        $state.MenuLabelManuallyEdited = $false
        $NewExePath.Text = ''
        Set-MenuLabelText -Value ''

        foreach ($item in $ScopeCheckList.Items) {
            $scope = $script:ScopeCatalog | Where-Object { $_.Id -eq $item.Tag } | Select-Object -First 1
            $item.IsChecked = [bool]$scope.DefaultForNew
        }

        Update-AddDialogState
        $ModalOverlay.Visibility = [System.Windows.Visibility]::Visible
        $NewExePath.Focus() | Out-Null
    }

    function Close-AddDialog {
        $ModalOverlay.Visibility = [System.Windows.Visibility]::Collapsed
    }

    function Set-ActiveScope {
        param(
            [Parameter(Mandatory = $true)]
            [string]$ScopeId
        )

        foreach ($button in $ScopeButtonPanel.Children) {
            $button.IsChecked = ($button.Tag -eq $ScopeId)
        }

        $state.ActiveScopeId = $ScopeId
        $selectedEntry = Get-SelectedEntry
        $selectedPath = if ($null -ne $selectedEntry) { [string]$selectedEntry.RegistryPath } else { $null }
        Update-EntryList -SelectedPath $selectedPath
    }

    $EntryList.Add_SelectionChanged({ Update-EntryDetails })

    $SearchBox.Add_TextChanged({
        Update-SearchPlaceholder
        $selectedEntry = Get-SelectedEntry
        $selectedPath = if ($null -ne $selectedEntry) { [string]$selectedEntry.RegistryPath } else { $null }
        Update-EntryList -SelectedPath $selectedPath
    })

    $RefreshButton.Add_Click({
        $selectedEntry = Get-SelectedEntry
        $selectedPath = if ($null -ne $selectedEntry) { [string]$selectedEntry.RegistryPath } else { $null }
        Refresh-Entries -SelectedPath $selectedPath
    })

    $ScopeBtn_All.Add_Click({ Set-ActiveScope -ScopeId 'All' })

    $AddButton.Add_Click({ Open-AddDialog })
    $ModalCloseButton.Add_Click({ Close-AddDialog })
    $ModalCancelButton.Add_Click({ Close-AddDialog })

    $BrowseButton.Add_Click({
        $dialog = [Microsoft.Win32.OpenFileDialog]::new()
        $dialog.Filter = Get-Text -Key 'OpenFileDialogFilter'
        $dialog.Title = Get-Text -Key 'OpenFileDialogTitle'
        $dialog.CheckFileExists = $true
        $dialog.Multiselect = $false

        if ($dialog.ShowDialog($window) -eq $true) {
            $NewExePath.Text = $dialog.FileName
        }
    })

    $NewExePath.Add_TextChanged({
        Update-SuggestedMenuLabel
        Update-AddDialogState
    })

    $NewMenuLabel.Add_TextChanged({
        if ($state.SuppressMenuLabelChanged) {
            return
        }

        $state.MenuLabelManuallyEdited =
            (-not [string]::IsNullOrWhiteSpace($NewMenuLabel.Text)) -and
            ($NewMenuLabel.Text -ne $state.LastSuggestedLabel)
    })

    $ModalAddButton.Add_Click({
        try {
            $selectedScopeIds = Get-CheckedScopeIds
            if ($selectedScopeIds.Count -eq 0) {
                Show-Message -Message (Get-Text -Key 'ValidationMissingScope') -Title (Get-Text -Key 'MessageInputErrorTitle') -Icon ([System.Windows.MessageBoxImage]::Warning)
                return
            }

            $result = Add-CustomAppContextMenu -ExePath $NewExePath.Text -ScopeIds $selectedScopeIds -MenuLabel $NewMenuLabel.Text
            Close-AddDialog
            Refresh-Entries -SelectedPath $result.RegistryPaths[0]
            Show-Message -Message (Get-Text -Key 'MessageAddSuccess' -Arguments $result.MenuLabel) -Title (Get-Text -Key 'MessageAddSuccessTitle')
        }
        catch {
            Show-Message -Message $_.Exception.Message -Title (Get-Text -Key 'MessageAddErrorTitle') -Icon ([System.Windows.MessageBoxImage]::Error)
        }
    })

    $EnableButton.Add_Click({
        $selectedEntry = Get-SelectedEntry
        if ($null -eq $selectedEntry) {
            return
        }

        try {
            Set-ContextMenuEntryDisabled -RegistryPath $selectedEntry.RegistryPath -Disabled $false
            Refresh-Entries -SelectedPath $selectedEntry.RegistryPath
            Show-Message -Message (Get-Text -Key 'MessageEnableSuccess') -Title (Get-Text -Key 'MessageEnableSuccessTitle')
        }
        catch {
            Show-Message -Message $_.Exception.Message -Title (Get-Text -Key 'MessageEnableErrorTitle') -Icon ([System.Windows.MessageBoxImage]::Error)
        }
    })

    $DisableButton.Add_Click({
        $selectedEntry = Get-SelectedEntry
        if ($null -eq $selectedEntry) {
            return
        }

        try {
            Set-ContextMenuEntryDisabled -RegistryPath $selectedEntry.RegistryPath -Disabled $true
            Refresh-Entries -SelectedPath $selectedEntry.RegistryPath
            Show-Message -Message (Get-Text -Key 'MessageDisableSuccess') -Title (Get-Text -Key 'MessageDisableSuccessTitle')
        }
        catch {
            Show-Message -Message $_.Exception.Message -Title (Get-Text -Key 'MessageDisableErrorTitle') -Icon ([System.Windows.MessageBoxImage]::Error)
        }
    })

    $DeleteButton.Add_Click({
        $selectedEntry = Get-SelectedEntry
        if ($null -eq $selectedEntry) {
            return
        }

        $confirmation = [System.Windows.MessageBox]::Show(
            $window,
            (Get-Text -Key 'MessageDeleteConfirm' -Arguments $selectedEntry.DisplayName, $selectedEntry.RegistryPath),
            (Get-Text -Key 'MessageDeleteConfirmTitle'),
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )

        if ($confirmation -ne [System.Windows.MessageBoxResult]::Yes) {
            return
        }

        try {
            $removed = Remove-ContextMenuEntry -RegistryPath $selectedEntry.RegistryPath
            if ($removed) {
                Refresh-Entries -SelectedPath $null
                Show-Message -Message (Get-Text -Key 'MessageDeleteSuccess') -Title (Get-Text -Key 'MessageDeleteSuccessTitle')
            }
        }
        catch {
            Show-Message -Message $_.Exception.Message -Title (Get-Text -Key 'MessageDeleteErrorTitle') -Icon ([System.Windows.MessageBoxImage]::Error)
        }
    })

    $window.Add_KeyDown({
        param($sender, $eventArgs)

        if ($eventArgs.Key -eq [System.Windows.Input.Key]::Escape -and
            $ModalOverlay.Visibility -eq [System.Windows.Visibility]::Visible) {
            Close-AddDialog
        }
    })

    $window.Add_Loaded({
        Clear-EntryDetails
        Update-SearchPlaceholder
        Update-AddDialogState
        Refresh-Entries -SelectedPath $null
    })

    [void]$window.ShowDialog()
}

$normalizedAction = if ([string]::IsNullOrWhiteSpace($Action)) { "" } else { $Action.Trim().ToUpperInvariant() }
$launchesGui = $Gui -or ([string]::IsNullOrWhiteSpace($normalizedAction) -and [string]::IsNullOrWhiteSpace($AppPath))

if ([string]::IsNullOrWhiteSpace($normalizedAction) -and -not $launchesGui) {
    throw (Get-Text -Key "ErrorInvalidAction")
}

if ($launchesGui -and -not $HiddenGuiHost) {
    Set-ConsoleWindowVisibility -Visible $false

    $hiddenHostArguments = Convert-BoundParametersToArgumentList -BoundParameters $PSBoundParameters
    $hiddenHostArguments += "-HiddenGuiHost"

    if (Restart-ScriptInHiddenWindow -AdditionalArguments $hiddenHostArguments) {
        return
    }
}

$elevationResult = Ensure-AdministratorSession -BoundParameters $PSBoundParameters -HideWindow:$launchesGui
if (-not $elevationResult.CanContinue) {
    if ($elevationResult.Cancelled) {
        if ($launchesGui) {
            Show-NotificationMessage -Message (Get-Text -Key "ErrorElevationCancelled") -Title (Get-Text -Key "WindowTitle") -Icon "Warning"
        }
        else {
            Write-Host (Get-Text -Key "OutputElevationCancelled")
        }
    }

    return
}

Migrate-LegacyCustomAppKeys

if ($launchesGui) {
    Set-ConsoleWindowVisibility -Visible $false
    Show-ContextMenuManagerGui
    return
}

switch ($normalizedAction) {
    "A" {
        if ([string]::IsNullOrWhiteSpace($AppPath)) {
            $AppPath = Read-Host (Get-Text -Key "PromptAddAppPath")
        }

        $result = Add-CustomAppContextMenu -ExePath $AppPath -ScopeIds $Scopes
        Write-Host ""
        Write-Host (Get-Text -Key "OutputAdded")
        Write-Host (Get-Text -Key "OutputMenuLabel" -Arguments $result.MenuLabel)
        Write-Host (Get-Text -Key "OutputKeyName" -Arguments $result.KeyName)
    }

    "R" {
        if ([string]::IsNullOrWhiteSpace($AppPath)) {
            $AppPath = Read-Host (Get-Text -Key "PromptRemoveAppPath")
        }

        $result = Remove-CustomAppContextMenu -ExePath $AppPath
        Write-Host ""
        if ($result.WasRemoved) {
            Write-Host (Get-Text -Key "OutputRemoved")
            Write-Host (Get-Text -Key "OutputKeyName" -Arguments $result.KeyName)
        }
        else {
            Write-Host (Get-Text -Key "OutputNoMatchingEntry")
        }
    }

    "L" {
        Write-EntryList
    }

    default {
        throw (Get-Text -Key "ErrorInvalidAction")
    }
}
