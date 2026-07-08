<#
.SYNOPSIS
    JSON ファイルで受け取ったアイテムを左右 2 つのリストボックスに表示し、
    それぞれの選択結果を JSON ファイルとして保存する汎用選択ダイアログ。

.DESCRIPTION
    入力 JSON (UTF-8 BOM 付き):
        {
            "title": "出力する列と形式を選択",              // 省略可 (既定: "項目を選択してください")
            "left": {
                "caption": "出力する列",                    // 省略可 (既定: 見出しなし)
                "multiSelect": true,                        // 省略可 (既定: true)
                "items": ["氏名", "住所", "電話番号"]
            },
            "right": {
                "caption": "出力形式",
                "multiSelect": false,
                "items": ["新規シート", "CSV ファイル"]
            }
        }

    left / right とも items は必須 (空はエラー)。
    左右のリストボックスは常に同じ幅・同じ高さで表示される。

    出力 JSON (UTF-8 BOM 付き):
        { "status": "ok",     "left": ["氏名", "住所"], "right": ["CSV ファイル"] }
        { "status": "cancel", "left": [], "right": [] }
        { "status": "error",  "left": [], "right": [], "message": "..." }

    終了コード: 0 = ok / 1 = cancel / 2 = error

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File Select-DualItems.ps1 `
        -InputPath C:\temp\in.json -OutputPath C:\temp\out.json

.NOTES
    引数なしで起動するとテストモードになる:
      - 組み込みのテストデータを左右のリストボックスに表示
      - 結果はスクリプトと同じフォルダの test_output_dual.json に出力
      - 終了時に選択結果をダイアログで表示し、出力ファイルを
        既定のアプリ (関連付けがなければメモ帳) で開く
#>
[CmdletBinding()]
param(
    [string]$InputPath,
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# UTF-8 BOM 付きでオブジェクトを JSON ファイルとして保存する
function Write-ResultFile {
    param(
        [string]$Path,
        [string]$Status,
        [string[]]$Left,
        [string[]]$Right,
        [string]$Message
    )
    $result = [ordered]@{
        status = $Status
        left   = @($Left)
        right  = @($Right)
    }
    if ($Message) { $result.message = $Message }

    $json = ConvertTo-Json -InputObject $result -Depth 5
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $json, $utf8Bom)
}

# 入力 JSON の left / right ノード 1 つ分を検証して取り出す
function Read-ListSpec {
    param(
        [psobject]$Parent,
        [string]$Key,
        [string]$Path
    )
    if (-not $Parent.PSObject.Properties[$Key] -or $null -eq $Parent.$Key) {
        throw "入力 JSON に $Key がありません: $Path"
    }
    $node = $Parent.$Key

    $items = @()
    if ($node.PSObject.Properties['items']) {
        $items = @($node.items | ForEach-Object { [string]$_ })
    }
    if ($items.Count -eq 0) {
        throw "入力 JSON の $Key.items が空、または存在しません: $Path"
    }

    $caption = ''
    if ($node.PSObject.Properties['caption'] -and $node.caption) {
        $caption = [string]$node.caption
    }

    $multiSelect = $true
    if ($node.PSObject.Properties['multiSelect'] -and $null -ne $node.multiSelect) {
        $multiSelect = [bool]$node.multiSelect
    }

    [pscustomobject]@{
        Caption     = $caption
        MultiSelect = $multiSelect
        Items       = $items
    }
}

function Read-InputSpec {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "入力ファイルが見つかりません: $Path"
    }
    # -Encoding UTF8 は BOM の有無どちらも UTF-8 として読み込む
    $raw  = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $spec = $raw | ConvertFrom-Json

    $title = '項目を選択してください'
    if ($spec.PSObject.Properties['title'] -and $spec.title) {
        $title = [string]$spec.title
    }

    [pscustomobject]@{
        Title = $title
        Left  = Read-ListSpec -Parent $spec -Key 'left'  -Path $Path
        Right = Read-ListSpec -Parent $spec -Key 'right' -Path $Path
    }
}

function Show-SelectDialog {
    param([pscustomobject]$Spec)

    Add-Type -AssemblyName PresentationFramework

    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="620" Height="480" MinWidth="420" MinHeight="240"
        WindowStartupLocation="CenterScreen" Topmost="True"
        ShowInTaskbar="True">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="12"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="lblLeft" Grid.Row="0" Grid.Column="0" Margin="0,0,0,4"/>
        <TextBlock x:Name="lblRight" Grid.Row="0" Grid.Column="2" Margin="0,0,0,4"/>
        <ListBox x:Name="lstLeft" Grid.Row="1" Grid.Column="0"
                 ScrollViewer.VerticalScrollBarVisibility="Auto"/>
        <ListBox x:Name="lstRight" Grid.Row="1" Grid.Column="2"
                 ScrollViewer.VerticalScrollBarVisibility="Auto"/>
        <StackPanel Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="3"
                    Orientation="Horizontal"
                    HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnOK" Content="OK" Width="90" Height="28"
                    IsDefault="True" Margin="0,0,10,0"/>
            <Button x:Name="btnCancel" Content="キャンセル" Width="90" Height="28"
                    IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
'@

    $reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($xaml)))
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    $btnOK     = $window.FindName('btnOK')
    $btnCancel = $window.FindName('btnCancel')

    $window.Title = $Spec.Title

    # 左右のリストボックスに見出し・選択モード・アイテムを設定する
    foreach ($side in @(
        @{ List = $window.FindName('lstLeft');  Label = $window.FindName('lblLeft');  Spec = $Spec.Left  },
        @{ List = $window.FindName('lstRight'); Label = $window.FindName('lblRight'); Spec = $Spec.Right }
    )) {
        $side.List.SelectionMode = if ($side.Spec.MultiSelect) { 'Extended' } else { 'Single' }
        foreach ($item in $side.Spec.Items) { [void]$side.List.Items.Add($item) }
        if ($side.Spec.Caption) {
            $side.Label.Text = $side.Spec.Caption
        }
        else {
            $side.Label.Visibility = 'Collapsed'
        }
    }

    $btnOK.Add_Click({ $window.DialogResult = $true })
    $btnCancel.Add_Click({ $window.DialogResult = $false })

    $dialogResult = $window.ShowDialog()

    [pscustomobject]@{
        Accepted = ($dialogResult -eq $true)
        Left     = @($window.FindName('lstLeft').SelectedItems  | ForEach-Object { [string]$_ })
        Right    = @($window.FindName('lstRight').SelectedItems | ForEach-Object { [string]$_ })
    }
}

# ---- main -------------------------------------------------------------------
try {
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw 'WPF の表示には STA モードが必要です。powershell.exe に -STA を付けて起動してください。'
    }

    # 引数なし = テストモード / 片方だけの指定はエラー
    $testMode = (-not $InputPath) -and (-not $OutputPath)
    if (-not $testMode -and ((-not $InputPath) -or (-not $OutputPath))) {
        throw '-InputPath と -OutputPath は両方指定してください (引数なしで起動するとテストモードになります)。'
    }

    if ($testMode) {
        $OutputPath = Join-Path $PSScriptRoot 'test_output_dual.json'
        $spec = [pscustomobject]@{
            Title = '[テスト] 出力する列と形式を選択してください'
            Left  = [pscustomobject]@{
                Caption     = '出力する列 (複数選択可)'
                MultiSelect = $true
                Items       = @('氏名', 'フリガナ', '郵便番号', '住所',
                                '電話番号', 'メールアドレス', '生年月日', '備考')
            }
            Right = [pscustomobject]@{
                Caption     = '出力形式 (1 つ選択)'
                MultiSelect = $false
                Items       = @('新規シート', 'CSV ファイル', 'クリップボード')
            }
        }
    }
    else {
        $spec = Read-InputSpec -Path $InputPath
    }

    $result = Show-SelectDialog -Spec $spec

    if ($result.Accepted) {
        $status = 'ok'
        Write-ResultFile -Path $OutputPath -Status $status -Left $result.Left -Right $result.Right
        $exitCode = 0
    }
    else {
        $status = 'cancel'
        Write-ResultFile -Path $OutputPath -Status $status -Left @() -Right @()
        $exitCode = 1
    }

    if ($testMode) {
        $leftLines  = if ($result.Left.Count  -gt 0) { $result.Left  -join "`n" } else { '(なし)' }
        $rightLines = if ($result.Right.Count -gt 0) { $result.Right -join "`n" } else { '(なし)' }
        [void][System.Windows.MessageBox]::Show(
            ("status : {0}`n`n左で選択されたアイテム:`n{1}`n`n右で選択されたアイテム:`n{2}`n`n出力先:`n{3}" -f `
                $status, $leftLines, $rightLines, $OutputPath),
            'テスト結果',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information)

        # 既定のアプリで開く。.json に関連付けがない場合はメモ帳で開く
        try {
            Start-Process -FilePath $OutputPath -ErrorAction Stop
        }
        catch {
            Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$OutputPath`""
        }
    }

    exit $exitCode
}
catch {
    $message = $_.Exception.Message
    if ($OutputPath) {
        try {
            Write-ResultFile -Path $OutputPath -Status 'error' -Left @() -Right @() -Message $message
        }
        catch {
            # 出力先にすら書けない場合は標準エラーのみ
        }
    }
    [Console]::Error.WriteLine($message)
    exit 2
}
