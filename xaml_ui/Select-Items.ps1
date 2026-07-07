<#
.SYNOPSIS
    JSON ファイルで受け取ったアイテムをリストボックスに表示し、
    ユーザーが選択した結果を JSON ファイルとして保存する汎用選択ダイアログ。

.DESCRIPTION
    入力 JSON (UTF-8 BOM 付き):
        {
            "title": "列を選択してください",   // 省略可 (既定: "項目を選択してください")
            "multiSelect": true,               // 省略可 (既定: true)
            "items": ["氏名", "住所", "電話番号"]
        }

    出力 JSON (UTF-8 BOM 付き):
        { "status": "ok",     "selected": ["氏名", "住所"] }
        { "status": "cancel", "selected": [] }
        { "status": "error",  "selected": [], "message": "..." }

    終了コード: 0 = ok / 1 = cancel / 2 = error

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File Select-Items.ps1 `
        -InputPath C:\temp\in.json -OutputPath C:\temp\out.json

.NOTES
    引数なしで起動するとテストモードになる:
      - 組み込みのテストデータをリストボックスに表示
      - 結果はスクリプトと同じフォルダの test_output.json に出力
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
        [string[]]$Selected,
        [string]$Message
    )
    $result = [ordered]@{
        status   = $Status
        selected = @($Selected)
    }
    if ($Message) { $result.message = $Message }

    $json = ConvertTo-Json -InputObject $result -Depth 5
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $json, $utf8Bom)
}

function Read-InputSpec {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "入力ファイルが見つかりません: $Path"
    }
    # -Encoding UTF8 は BOM の有無どちらも UTF-8 として読み込む
    $raw  = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $spec = $raw | ConvertFrom-Json

    $items = @()
    if ($spec.PSObject.Properties['items']) {
        $items = @($spec.items | ForEach-Object { [string]$_ })
    }
    if ($items.Count -eq 0) {
        throw "入力 JSON の items が空、または存在しません: $Path"
    }

    $title = '項目を選択してください'
    if ($spec.PSObject.Properties['title'] -and $spec.title) {
        $title = [string]$spec.title
    }

    $multiSelect = $true
    if ($spec.PSObject.Properties['multiSelect'] -and $null -ne $spec.multiSelect) {
        $multiSelect = [bool]$spec.multiSelect
    }

    [pscustomobject]@{
        Title       = $title
        MultiSelect = $multiSelect
        Items       = $items
    }
}

function Show-SelectDialog {
    param([pscustomobject]$Spec)

    Add-Type -AssemblyName PresentationFramework

    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="380" Height="480" MinWidth="280" MinHeight="240"
        WindowStartupLocation="CenterScreen" Topmost="True"
        ShowInTaskbar="True">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <ListBox x:Name="lstItems" Grid.Row="0"
                 ScrollViewer.VerticalScrollBarVisibility="Auto"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal"
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

    $listBox   = $window.FindName('lstItems')
    $btnOK     = $window.FindName('btnOK')
    $btnCancel = $window.FindName('btnCancel')

    $window.Title = $Spec.Title
    $listBox.SelectionMode = if ($Spec.MultiSelect) { 'Extended' } else { 'Single' }
    foreach ($item in $Spec.Items) { [void]$listBox.Items.Add($item) }

    $btnOK.Add_Click({ $window.DialogResult = $true })
    $btnCancel.Add_Click({ $window.DialogResult = $false })
    # ダブルクリックで即決定 (選択中のアイテムがある場合のみ)
    $listBox.Add_MouseDoubleClick({
        if ($listBox.SelectedItems.Count -gt 0) { $window.DialogResult = $true }
    })

    $dialogResult = $window.ShowDialog()

    [pscustomobject]@{
        Accepted = ($dialogResult -eq $true)
        Selected = @($listBox.SelectedItems | ForEach-Object { [string]$_ })
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
        $OutputPath = Join-Path $PSScriptRoot 'test_output.json'
        $spec = [pscustomobject]@{
            Title       = '[テスト] 列を選択してください'
            MultiSelect = $true
            Items       = @('氏名', 'フリガナ', '郵便番号', '住所',
                            '電話番号', 'メールアドレス', '生年月日', '備考')
        }
    }
    else {
        $spec = Read-InputSpec -Path $InputPath
    }

    $result = Show-SelectDialog -Spec $spec

    if ($result.Accepted) {
        $status = 'ok'
        Write-ResultFile -Path $OutputPath -Status $status -Selected $result.Selected
        $exitCode = 0
    }
    else {
        $status = 'cancel'
        Write-ResultFile -Path $OutputPath -Status $status -Selected @()
        $exitCode = 1
    }

    if ($testMode) {
        $lines = if ($result.Selected.Count -gt 0) { $result.Selected -join "`n" }
                 else                              { '(なし)' }
        [void][System.Windows.MessageBox]::Show(
            ("status : {0}`n`n選択されたアイテム:`n{1}`n`n出力先:`n{2}" -f $status, $lines, $OutputPath),
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
            Write-ResultFile -Path $OutputPath -Status 'error' -Selected @() -Message $message
        }
        catch {
            # 出力先にすら書けない場合は標準エラーのみ
        }
    }
    [Console]::Error.WriteLine($message)
    exit 2
}
