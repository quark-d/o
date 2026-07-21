# Show-LedgerDialog.ps1
#
# 在庫管理台帳の入力ダイアログ (PowerShell 5.1 + WPF/XAML 版)。
# VBA UserForm 版の DPI 問題時のフォールバック実装。
#
# 起動:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File Show-LedgerDialog.ps1 `
#       -InputPath <入力JSON> -OutputPath <出力JSON>
#
# 引数なしで起動するとテストモード (組み込みデータで単体テスト。-TestDialog で画面指定)。
# 入出力 JSON は UTF-8 (BOM 付き)。スキーマは README.md 参照 (schemaVersion 1.2)。
# 出力は OK/キャンセル/エラーのいずれでも必ず作成し、既存ファイルは無条件上書き。
# 終了コード: 0=OK / 1=キャンセル / 2=エラー
#
# 注意: XAML にイベント属性 (Click= 等) や x:Class を書かないこと (README 参照)。

param(
    [string]$InputPath,
    [string]$OutputPath,
    [ValidateSet('carryIn', 'progress', 'defect', 'exception')]
    [string]$TestDialog = 'defect'
)

$ErrorActionPreference = 'Stop'
$script:SchemaVersion = '1.2'

Add-Type -AssemblyName PresentationFramework | Out-Null

#==============================================================================
# 出力 JSON
#==============================================================================

# 全キーを常に出力する (VBA 側の簡易パーサが分岐なしで読めるように)
function New-EmptyOperation {
    [ordered]@{
        kind            = ''
        eventKind       = ''
        localId         = ''
        itemNo          = ''
        lotId           = ''
        fromState       = ''
        qty             = 0
        actionDate      = ''
        recorder        = ''
        note            = ''
        refill          = $false
        refillFromState = ''
        refillQty       = 0
        targetLotId     = ''
    }
}

function Write-ResultJson {
    param([string]$Status, [string]$Message, $Operation)
    if (-not $Operation) { $Operation = New-EmptyOperation }
    $doc = [ordered]@{
        schemaVersion = $script:SchemaVersion
        status        = $Status
        message       = $Message
        operation     = $Operation
    }
    $json = $doc | ConvertTo-Json -Depth 5
    $enc = New-Object System.Text.UTF8Encoding($true)   # BOM 付き
    [System.IO.File]::WriteAllText($script:OutPath, $json, $enc)
}

#==============================================================================
# 工程パターン (LedgerCore.bas と同一のロジック)
#==============================================================================

function Get-PatternChain([int]$Pattern) {
    $tail = @('梱包中', '梱包済み', '発送準備中', '発送準備完了')
    switch ($Pattern) {
        1 { @('未処理') + $tail }
        2 { @('未処理', 'アセンブリ中', 'アセンブリ済み') + $tail }
        3 { @('未処理', 'アニール中', 'アニール済み') + $tail }
        4 { @('未処理', 'アニール中', 'アニール済み', 'アセンブリ中', 'アセンブリ済み') + $tail }
        default { @() }
    }
}

function Get-EventToState([string]$EventKind) {
    switch ($EventKind) {
        'アニール開始'   { 'アニール中' }
        'アニール完了'   { 'アニール済み' }
        'アセンブリ開始' { 'アセンブリ中' }
        'アセンブリ完了' { 'アセンブリ済み' }
        '梱包開始'       { '梱包中' }
        '梱包完了'       { '梱包済み' }
        '発送準備開始'   { '発送準備中' }
        '発送準備完了'   { '発送準備完了' }
        default          { '' }
    }
}

function Get-AllowedEvents([int]$Pattern) {
    $chain = Get-PatternChain $Pattern
    $events = @()
    $all = @('アニール開始', 'アニール完了', 'アセンブリ開始', 'アセンブリ完了',
             '梱包開始', '梱包完了', '発送準備開始', '発送準備完了')
    for ($i = 1; $i -lt $chain.Count; $i++) {
        foreach ($ev in $all) {
            if ((Get-EventToState $ev) -eq $chain[$i]) { $events += $ev; break }
        }
    }
    $events + '発送'
}

function Get-FromState([int]$Pattern, [string]$EventKind) {
    if ($EventKind -eq '発送') { return '発送準備完了' }
    $to = Get-EventToState $EventKind
    if (-not $to) { return '' }
    $chain = Get-PatternChain $Pattern
    for ($i = 1; $i -lt $chain.Count; $i++) {
        if ($chain[$i] -eq $to) { return $chain[$i - 1] }
    }
    ''
}

#==============================================================================
# 在庫スナップショットの参照ヘルパー
#==============================================================================

# 入力 JSON の stock をハッシュテーブル化する ("L|ロットID"/"U|localId" → 状態→数量)
function Build-StockMap($InObj) {
    $map = @{}
    foreach ($entry in @($InObj.stock.lots)) {
        if ($null -eq $entry) { continue }
        $h = @{}
        foreach ($p in $entry.byState.PSObject.Properties) { $h[$p.Name] = [int]$p.Value }
        $map['L|' + $entry.lotId] = $h
    }
    foreach ($entry in @($InObj.stock.unallocated)) {
        if ($null -eq $entry) { continue }
        $h = @{}
        foreach ($p in $entry.byState.PSObject.Properties) { $h[$p.Name] = [int]$p.Value }
        $map['U|' + $entry.localId] = $h
    }
    $map
}

function Get-Avail([string]$Bucket, [string]$Id, [string]$State) {
    $h = $script:StockMap[$Bucket + '|' + $Id]
    if ($h -and $h.ContainsKey($State)) { [Math]::Max(0, $h[$State]) } else { 0 }
}

function Get-StatesWithStock([string]$Bucket, [string]$Id) {
    @($script:States | Where-Object { (Get-Avail $Bucket $Id $_) -gt 0 })
}

function Describe-States([string]$Bucket, [string]$Id) {
    $parts = @()
    foreach ($st in $script:States) {
        $q = Get-Avail $Bucket $Id $st
        if ($q -gt 0) { $parts += "$st $q" }
    }
    if ($parts.Count -eq 0) { 'なし' } else { $parts -join ' / ' }
}

#==============================================================================
# WPF 共通
#==============================================================================

function New-WindowFromXaml([string]$XamlText) {
    $xml = [xml]$XamlText
    $reader = New-Object System.Xml.XmlNodeReader $xml
    [Windows.Markup.XamlReader]::Load($reader)
}

function Show-Error([string]$Message) {
    [void][System.Windows.MessageBox]::Show($Message, '台帳入力',
        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
}

# 数量テキストの検証 (NG ならメッセージを出して $null)
function Parse-Qty([string]$Text) {
    $q = 0
    if (-not [int]::TryParse($Text, [ref]$q) -or $q -lt 1) {
        Show-Error '数量は 1 以上の整数で入力してください。'
        return $null
    }
    $q
}

# 実施日・記録者の共通検証 → @{ actionDate=...; recorder=... } または $null
function Read-CommonInputs($DatePicker, $RecorderCombo) {
    if (-not $DatePicker.SelectedDate) {
        Show-Error '実施日を選択してください。'
        return $null
    }
    if (-not $RecorderCombo.SelectedItem) {
        Show-Error '記録者を選択してください。'
        return $null
    }
    @{
        actionDate = $DatePicker.SelectedDate.ToString('yyyy-MM-dd')
        recorder   = [string]$RecorderCombo.SelectedItem
    }
}

# 共通コントロール (実施日・記録者) に既定値を入れる
function Init-CommonInputs($DatePicker, $RecorderCombo) {
    foreach ($r in @($script:In.recorders)) { [void]$RecorderCombo.Items.Add($r) }
    if ($script:In.defaults.recorder) { $RecorderCombo.SelectedItem = $script:In.defaults.recorder }
    $DatePicker.SelectedDate = [datetime]::ParseExact(
        $script:In.defaults.actionDate, 'yyyy-MM-dd', $null)
}

# 共通フッター XAML (実施日/記録者/備考/OK/キャンセル)
$script:CommonFooterXaml = @'
      <TextBlock Text="実施日" Margin="0,10,0,2"/>
      <DatePicker x:Name="dpDate"/>
      <TextBlock Text="記録者" Margin="0,10,0,2"/>
      <ComboBox x:Name="cmbRecorder"/>
      <TextBlock Text="備考" Margin="0,10,0,2"/>
      <TextBox x:Name="txtNote"/>
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
        <Button x:Name="btnOK" Content="OK" Width="90" Height="26" IsDefault="True" Margin="0,0,8,0"/>
        <Button x:Name="btnCancel" Content="キャンセル" Width="90" Height="26" IsCancel="True"/>
      </StackPanel>
'@

function New-DialogXaml([string]$Title, [string]$BodyXaml) {
    @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" SizeToContent="WidthAndHeight" ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen" Topmost="True" FontSize="13">
  <StackPanel Margin="16" Width="440">
$BodyXaml
$($script:CommonFooterXaml)
  </StackPanel>
</Window>
"@
}

#==============================================================================
# 1. 搬入ダイアログ (localId / 品番 / 品名 の 3 欄連動入力)
#==============================================================================

# 品名重複時の候補選択ウィンドウ。選択された item または $null を返す
function Show-ItemCandidatePicker($Candidates) {
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="候補の選択" SizeToContent="WidthAndHeight" ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen" Topmost="True" FontSize="13">
  <StackPanel Margin="16" Width="420">
    <TextBlock Text="品名が複数の品番に一致しました。選択してください:" Margin="0,0,0,6"/>
    <ListBox x:Name="lstCands" Height="120"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
      <Button x:Name="btnPickOK" Content="OK" Width="90" Height="26" IsDefault="True" Margin="0,0,8,0"/>
      <Button x:Name="btnPickCancel" Content="キャンセル" Width="90" Height="26" IsCancel="True"/>
    </StackPanel>
  </StackPanel>
</Window>
'@
    $win = New-WindowFromXaml $xaml
    $lst = $win.FindName('lstCands')
    foreach ($it in $Candidates) {
        [void]$lst.Items.Add("$($it.localId)  $($it.itemNo)  $($it.itemName)")
    }
    $lst.SelectedIndex = 0
    $script:pickWin = $win
    $win.FindName('btnPickOK').Add_Click({ $script:pickWin.DialogResult = $true })
    $result = $win.ShowDialog()
    if ($result -and $lst.SelectedIndex -ge 0) { $Candidates[$lst.SelectedIndex] } else { $null }
}

function Show-CarryInDialog {
    $body = @'
      <TextBlock Text="localId (内部ID)" Margin="0,0,0,2"/>
      <ComboBox x:Name="cmbLocalId" IsEditable="True"/>
      <TextBlock Text="品番" Margin="0,8,0,2"/>
      <ComboBox x:Name="cmbItemNo" IsEditable="True"/>
      <TextBlock Text="品名" Margin="0,8,0,2"/>
      <ComboBox x:Name="cmbItemName" IsEditable="True"/>
      <TextBlock x:Name="lblPool" Margin="0,4,0,0" Foreground="Gray" TextWrapping="Wrap"/>
      <TextBlock Text="数量" Margin="0,10,0,2"/>
      <TextBox x:Name="txtQty"/>
'@
    $wa = if ($script:In.labels -and $script:In.labels.warehouseA) { [string]$script:In.labels.warehouseA } else { '倉庫A' }
    $wb = if ($script:In.labels -and $script:In.labels.warehouseB) { [string]$script:In.labels.warehouseB } else { '倉庫B' }
    $win = New-WindowFromXaml (New-DialogXaml "搬入 ($wa → $wb)" $body)
    $ui = @{}
    foreach ($n in 'cmbLocalId', 'cmbItemNo', 'cmbItemName', 'lblPool', 'txtQty',
                   'dpDate', 'cmbRecorder', 'txtNote', 'btnOK') {
        $ui[$n] = $win.FindName($n)
    }
    $script:ui = $ui
    $script:uiItems = @($script:In.items)
    $script:carrySel = $null      # 確定済み item (未確定は $null)
    $script:carryUpd = $false     # 3 欄へのプログラム反映中

    foreach ($it in $script:uiItems) {
        [void]$ui.cmbLocalId.Items.Add([string]$it.localId)
        [void]$ui.cmbItemNo.Items.Add([string]$it.itemNo)
        [void]$ui.cmbItemName.Items.Add([string]$it.itemName)
    }
    Init-CommonInputs $ui.dpDate $ui.cmbRecorder

    # 確定した item を 3 欄へ反映
    $script:applyItem = {
        param($it)
        $ui = $script:ui
        $script:carrySel = $it
        $script:carryUpd = $true
        $ui.cmbLocalId.Text = [string]$it.localId
        $ui.cmbItemNo.Text = [string]$it.itemNo
        $ui.cmbItemName.Text = [string]$it.itemName
        $script:carryUpd = $false
        $ui.lblPool.Text = '現在の未割当在庫: ' + (Describe-States 'U' $it.localId)
    }

    $script:clearItemFields = {
        $ui = $script:ui
        $script:carrySel = $null
        $script:carryUpd = $true
        $ui.cmbLocalId.Text = ''
        $ui.cmbItemNo.Text = ''
        $ui.cmbItemName.Text = ''
        $script:carryUpd = $false
        $ui.lblPool.Text = ''
    }

    # ドロップダウンから選択されたときの即時反映 (3 欄とも items と同順)
    $script:carrySelectionChanged = {
        param($cmb)
        if ($script:carryUpd) { return }
        if ($cmb.SelectedIndex -ge 0) {
            & $script:applyItem $script:uiItems[$cmb.SelectedIndex]
        }
    }
    $ui.cmbLocalId.Add_SelectionChanged({ & $script:carrySelectionChanged $script:ui.cmbLocalId })
    $ui.cmbItemNo.Add_SelectionChanged({ & $script:carrySelectionChanged $script:ui.cmbItemNo })
    $ui.cmbItemName.Add_SelectionChanged({ & $script:carrySelectionChanged $script:ui.cmbItemName })

    # 欄を離れたときの確定チェック。未登録なら他欄に「!」+ 再入力/キャンセル
    $script:carryFieldExit = {
        param([string]$src, $cmb, [string]$caption)
        if ($script:carryUpd) { return }
        $ui = $script:ui
        $txt = $cmb.Text.Trim()
        if (-not $txt -or $txt -eq '!') { return }

        if ($src -eq 'name') {
            $cands = @($script:uiItems | Where-Object { [string]$_.itemName -eq $txt })
            if ($cands.Count -gt 1) {
                # 品名の複数一致 → 候補選択 (キャンセルは入力クリア)
                $pick = Show-ItemCandidatePicker $cands
                if ($pick) { & $script:applyItem $pick } else { & $script:clearItemFields }
                return
            }
            $it = $cands | Select-Object -First 1
        } elseif ($src -eq 'local') {
            $it = @($script:uiItems | Where-Object { [string]$_.localId -eq $txt }) | Select-Object -First 1
        } else {
            $it = @($script:uiItems | Where-Object { [string]$_.itemNo -eq $txt }) | Select-Object -First 1
        }
        if ($it) {
            & $script:applyItem $it
            return
        }

        # 未登録 → 他欄に「!」を表示して再入力かキャンセルを促す
        $script:carrySel = $null
        $script:carryUpd = $true
        if ($src -ne 'local') { $ui.cmbLocalId.Text = '!' }
        if ($src -ne 'itemNo') { $ui.cmbItemNo.Text = '!' }
        if ($src -ne 'name') { $ui.cmbItemName.Text = '!' }
        $script:carryUpd = $false
        $ui.lblPool.Text = ''

        $ans = [System.Windows.MessageBox]::Show(
            "入力された$caption「$txt」は M_品番 にありません。`n再入力しますか? (いいえ = 入力をクリア)",
            '未登録の' + $caption,
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
        if ($ans -eq [System.Windows.MessageBoxResult]::Yes) {
            $script:carryUpd = $true
            if ($src -ne 'local') { $ui.cmbLocalId.Text = '' }
            if ($src -ne 'itemNo') { $ui.cmbItemNo.Text = '' }
            if ($src -ne 'name') { $ui.cmbItemName.Text = '' }
            $script:carryUpd = $false
            [void]$cmb.Focus()
        } else {
            & $script:clearItemFields
        }
    }
    $ui.cmbLocalId.Add_LostFocus({ & $script:carryFieldExit 'local' $script:ui.cmbLocalId 'localId' })
    $ui.cmbItemNo.Add_LostFocus({ & $script:carryFieldExit 'itemNo' $script:ui.cmbItemNo '品番' })
    $ui.cmbItemName.Add_LostFocus({ & $script:carryFieldExit 'name' $script:ui.cmbItemName '品名' })

    $ui.btnOK.Add_Click({
        $ui = $script:ui
        if ($null -eq $script:carrySel) {
            Show-Error 'localId / 品番 / 品名 のいずれかで品番を確定してください。'; return
        }
        $qty = Parse-Qty $ui.txtQty.Text
        if ($null -eq $qty) { return }
        $common = Read-CommonInputs $ui.dpDate $ui.cmbRecorder
        if ($null -eq $common) { return }

        $op = New-EmptyOperation
        $op.kind = 'carryIn'
        $op.eventKind = '搬入'
        $op.localId = [string]$script:carrySel.localId
        $op.itemNo = [string]$script:carrySel.itemNo
        $op.qty = $qty
        $op.actionDate = $common.actionDate
        $op.recorder = $common.recorder
        $op.note = $ui.txtNote.Text.Trim()
        $script:dlgOp = $op
        $script:dlgWin.DialogResult = $true
    })

    $script:dlgWin = $win
    $script:dlgOp = $null
    [void]$win.ShowDialog()
    $script:dlgOp
}

#==============================================================================
# 2. 工程進行ダイアログ (発送含む)
#==============================================================================
function Show-ProgressDialog {
    $body = @'
      <TextBlock Text="ロット" Margin="0,0,0,2"/>
      <ComboBox x:Name="cmbLot"/>
      <TextBlock x:Name="lblItem" Margin="0,4,0,0" Foreground="Gray" TextWrapping="Wrap"/>
      <TextBlock Text="イベント" Margin="0,10,0,2"/>
      <ComboBox x:Name="cmbEvent"/>
      <TextBlock x:Name="lblFrom" Margin="0,4,0,0" Foreground="Gray"/>
      <TextBlock Text="数量" Margin="0,10,0,2"/>
      <TextBox x:Name="txtQty"/>
'@
    $win = New-WindowFromXaml (New-DialogXaml '工程進行 / 発送' $body)
    $ui = @{}
    foreach ($n in 'cmbLot', 'lblItem', 'cmbEvent', 'lblFrom', 'txtQty', 'dpDate',
                   'cmbRecorder', 'txtNote', 'btnOK') {
        $ui[$n] = $win.FindName($n)
    }
    $script:ui = $ui
    $script:uiLots = @(@($script:In.lots) | Where-Object { $_.planStatus -eq '有効' })

    foreach ($lt in $script:uiLots) {
        $it = @($script:In.items) | Where-Object { $_.localId -eq $lt.localId } | Select-Object -First 1
        $nm = if ($it) { "  $($it.itemName)" } else { '' }
        [void]$ui.cmbLot.Items.Add("$($lt.lotId)$nm  (必要 $($lt.required))")
    }
    Init-CommonInputs $ui.dpDate $ui.cmbRecorder

    $ui.cmbLot.Add_SelectionChanged({
        $ui = $script:ui
        $ui.cmbEvent.Items.Clear()
        $ui.lblFrom.Text = ''
        if ($ui.cmbLot.SelectedIndex -lt 0) { return }
        $lt = $script:uiLots[$ui.cmbLot.SelectedIndex]
        $it = @($script:In.items) | Where-Object { $_.localId -eq $lt.localId } | Select-Object -First 1
        if ($null -eq $it) { $ui.lblItem.Text = "localId $($lt.localId) が M_品番 にありません"; return }
        $ui.lblItem.Text = "$($it.localId)  $($it.itemNo)  $($it.itemName)  (パターン $($it.pattern))`n" +
                           '在庫: ' + (Describe-States 'L' $lt.lotId)
        foreach ($ev in (Get-AllowedEvents ([int]$it.pattern))) {
            [void]$ui.cmbEvent.Items.Add($ev)
        }
    })

    $ui.cmbEvent.Add_SelectionChanged({
        $ui = $script:ui
        $ui.lblFrom.Text = ''
        if ($ui.cmbLot.SelectedIndex -lt 0 -or $ui.cmbEvent.SelectedIndex -lt 0) { return }
        $lt = $script:uiLots[$ui.cmbLot.SelectedIndex]
        $it = @($script:In.items) | Where-Object { $_.localId -eq $lt.localId } | Select-Object -First 1
        if ($null -eq $it) { return }
        $from = Get-FromState ([int]$it.pattern) ([string]$ui.cmbEvent.SelectedItem)
        $avail = Get-Avail 'L' $lt.lotId $from
        $ui.lblFrom.Text = "元状態: $from  (可用 $avail 個)"
        if ($avail -gt 0 -and -not $ui.txtQty.Text.Trim()) { $ui.txtQty.Text = "$avail" }
    })

    $ui.btnOK.Add_Click({
        $ui = $script:ui
        if ($ui.cmbLot.SelectedIndex -lt 0) { Show-Error 'ロットを選択してください。'; return }
        if ($ui.cmbEvent.SelectedIndex -lt 0) { Show-Error 'イベントを選択してください。'; return }
        $qty = Parse-Qty $ui.txtQty.Text
        if ($null -eq $qty) { return }
        $common = Read-CommonInputs $ui.dpDate $ui.cmbRecorder
        if ($null -eq $common) { return }

        $op = New-EmptyOperation
        $op.kind = 'progress'
        $op.eventKind = [string]$ui.cmbEvent.SelectedItem
        $op.lotId = $script:uiLots[$ui.cmbLot.SelectedIndex].lotId
        $op.qty = $qty
        $op.actionDate = $common.actionDate
        $op.recorder = $common.recorder
        $op.note = $ui.txtNote.Text.Trim()
        $script:dlgOp = $op
        $script:dlgWin.DialogResult = $true
    })

    $script:dlgWin = $win
    $script:dlgOp = $null
    [void]$win.ShowDialog()
    $script:dlgOp
}

#==============================================================================
# 3. 不良+補填ダイアログ
#==============================================================================
function Show-DefectDialog {
    $body = @'
      <TextBlock Text="ロット" Margin="0,0,0,2"/>
      <ComboBox x:Name="cmbLot"/>
      <TextBlock x:Name="lblItem" Margin="0,4,0,0" Foreground="Gray"/>
      <TextBlock Text="不良になった在庫の元状態" Margin="0,10,0,2"/>
      <ComboBox x:Name="cmbState"/>
      <TextBlock x:Name="lblStock" Margin="0,4,0,0" Foreground="Gray"/>
      <TextBlock Text="不良数量" Margin="0,10,0,2"/>
      <TextBox x:Name="txtQty"/>
      <CheckBox x:Name="chkRefill" Content="未割当在庫から補填する (充当行を自動生成)" Margin="0,12,0,0"/>
      <TextBlock Text="補填元の状態" Margin="0,8,0,2"/>
      <ComboBox x:Name="cmbRefillState" IsEnabled="False"/>
      <TextBlock Text="補填数量 (上限 = 不良数量)" Margin="0,8,0,2"/>
      <TextBox x:Name="txtRefillQty" IsEnabled="False"/>
      <TextBlock x:Name="lblPool" Margin="0,4,0,0" Foreground="Gray" TextWrapping="Wrap"/>
'@
    $win = New-WindowFromXaml (New-DialogXaml '不良発生 (+補填)' $body)
    $ui = @{}
    foreach ($n in 'cmbLot', 'lblItem', 'cmbState', 'lblStock', 'txtQty', 'chkRefill',
                   'cmbRefillState', 'txtRefillQty', 'lblPool', 'dpDate', 'cmbRecorder',
                   'txtNote', 'btnOK') {
        $ui[$n] = $win.FindName($n)
    }
    $script:ui = $ui
    $script:uiLots = @(@($script:In.lots) | Where-Object { $_.planStatus -eq '有効' })

    foreach ($lt in $script:uiLots) {
        $it = @($script:In.items) | Where-Object { $_.localId -eq $lt.localId } | Select-Object -First 1
        $nm = if ($it) { "  $($it.itemName)" } else { '' }
        [void]$ui.cmbLot.Items.Add("$($lt.lotId)$nm  (必要 $($lt.required))")
    }
    Init-CommonInputs $ui.dpDate $ui.cmbRecorder

    # 補填数量の既定 = MIN(不良数量, 補填元状態の未割当在庫)
    $script:syncRefillQty = {
        $ui = $script:ui
        if (-not $ui.chkRefill.IsChecked) { return }
        if ($ui.cmbLot.SelectedIndex -lt 0 -or $ui.cmbRefillState.SelectedIndex -lt 0) { return }
        $q = 0
        if (-not [int]::TryParse($ui.txtQty.Text, [ref]$q)) { return }
        $lt = $script:uiLots[$ui.cmbLot.SelectedIndex]
        $pool = Get-Avail 'U' $lt.localId ([string]$ui.cmbRefillState.SelectedItem)
        $ui.txtRefillQty.Text = "$([Math]::Min($q, $pool))"
    }

    $ui.cmbLot.Add_SelectionChanged({
        $ui = $script:ui
        $ui.cmbState.Items.Clear()
        $ui.cmbRefillState.Items.Clear()
        $ui.lblStock.Text = ''
        if ($ui.cmbLot.SelectedIndex -lt 0) { return }
        $lt = $script:uiLots[$ui.cmbLot.SelectedIndex]
        $it = @($script:In.items) | Where-Object { $_.localId -eq $lt.localId } | Select-Object -First 1
        $ui.lblItem.Text = if ($it) { "$($it.localId)  $($it.itemNo)  $($it.itemName)" } else { $lt.localId }
        foreach ($st in (Get-StatesWithStock 'L' $lt.lotId)) { [void]$ui.cmbState.Items.Add($st) }
        foreach ($st in (Get-StatesWithStock 'U' $lt.localId)) { [void]$ui.cmbRefillState.Items.Add($st) }
        $ui.lblPool.Text = '未割当在庫: ' + (Describe-States 'U' $lt.localId)
    })

    $ui.cmbState.Add_SelectionChanged({
        $ui = $script:ui
        if ($ui.cmbLot.SelectedIndex -lt 0 -or $ui.cmbState.SelectedIndex -lt 0) { return }
        $lt = $script:uiLots[$ui.cmbLot.SelectedIndex]
        $avail = Get-Avail 'L' $lt.lotId ([string]$ui.cmbState.SelectedItem)
        $ui.lblStock.Text = "可用 $avail 個"
    })

    $ui.chkRefill.Add_Click({
        $ui = $script:ui
        $on = [bool]$ui.chkRefill.IsChecked
        $ui.cmbRefillState.IsEnabled = $on
        $ui.txtRefillQty.IsEnabled = $on
        if ($on) { & $script:syncRefillQty }
    })
    $ui.txtQty.Add_TextChanged({ & $script:syncRefillQty })
    $ui.cmbRefillState.Add_SelectionChanged({ & $script:syncRefillQty })

    $ui.btnOK.Add_Click({
        $ui = $script:ui
        if ($ui.cmbLot.SelectedIndex -lt 0) { Show-Error 'ロットを選択してください。'; return }
        if ($ui.cmbState.SelectedIndex -lt 0) {
            Show-Error '不良になった在庫の元状態を選択してください。'; return
        }
        $qty = Parse-Qty $ui.txtQty.Text
        if ($null -eq $qty) { return }
        $common = Read-CommonInputs $ui.dpDate $ui.cmbRecorder
        if ($null -eq $common) { return }

        $op = New-EmptyOperation
        $op.kind = 'defect'
        $op.eventKind = '不良発生'
        $op.lotId = $script:uiLots[$ui.cmbLot.SelectedIndex].lotId
        $op.fromState = [string]$ui.cmbState.SelectedItem
        $op.qty = $qty
        $op.actionDate = $common.actionDate
        $op.recorder = $common.recorder
        $op.note = $ui.txtNote.Text.Trim()

        if ($ui.chkRefill.IsChecked) {
            if ($ui.cmbRefillState.SelectedIndex -lt 0) {
                Show-Error '補填元の状態を選択してください。'; return
            }
            $rq = 0
            if (-not [int]::TryParse($ui.txtRefillQty.Text, [ref]$rq) -or $rq -lt 1) {
                Show-Error '補填数量を 1 以上の整数で入力してください。'; return
            }
            if ($rq -gt $qty) {
                Show-Error ("補填数量 ($rq) は不良数量 ($qty) を超えられません。" +
                            '追加の割当は例外操作の「充当」で行ってください。')
                return
            }
            $op.refill = $true
            $op.refillFromState = [string]$ui.cmbRefillState.SelectedItem
            $op.refillQty = $rq
        }

        $script:dlgOp = $op
        $script:dlgWin.DialogResult = $true
    })

    $script:dlgWin = $win
    $script:dlgOp = $null
    [void]$win.ShowDialog()
    $script:dlgOp
}

#==============================================================================
# 4. 例外操作ダイアログ (廃棄 / 倉庫戻し / 余剰化 / 充当)
#==============================================================================
function Show-ExceptionDialog {
    $body = @'
      <TextBlock Text="操作" Margin="0,0,0,2"/>
      <StackPanel Orientation="Horizontal">
        <RadioButton x:Name="rbDiscard" GroupName="op" Content="廃棄" IsChecked="True" Margin="0,0,12,0"/>
        <RadioButton x:Name="rbReturn" GroupName="op" Content="倉庫戻し" Margin="0,0,12,0"/>
        <RadioButton x:Name="rbSurplus" GroupName="op" Content="余剰化" Margin="0,0,12,0"/>
        <RadioButton x:Name="rbAllocate" GroupName="op" Content="充当"/>
      </StackPanel>
      <TextBlock Text="対象 (ロット / 未割当)" Margin="0,10,0,2"/>
      <ComboBox x:Name="cmbTarget"/>
      <TextBlock Text="元状態" Margin="0,10,0,2"/>
      <ComboBox x:Name="cmbState"/>
      <TextBlock x:Name="lblStock" Margin="0,4,0,0" Foreground="Gray"/>
      <TextBlock Text="数量" Margin="0,10,0,2"/>
      <TextBox x:Name="txtQty"/>
      <TextBlock Text="充当先ロット (充当のみ)" Margin="0,10,0,2"/>
      <ComboBox x:Name="cmbTargetLot" IsEnabled="False"/>
'@
    $win = New-WindowFromXaml (New-DialogXaml '例外操作 (廃棄 / 倉庫戻し / 余剰化 / 充当)' $body)
    $ui = @{}
    foreach ($n in 'rbDiscard', 'rbReturn', 'rbSurplus', 'rbAllocate', 'cmbTarget',
                   'cmbState', 'lblStock', 'txtQty', 'cmbTargetLot', 'dpDate',
                   'cmbRecorder', 'txtNote', 'btnOK') {
        $ui[$n] = $win.FindName($n)
    }
    $script:ui = $ui
    Init-CommonInputs $ui.dpDate $ui.cmbRecorder

    $script:currentEvent = {
        $ui = $script:ui
        if ($ui.rbDiscard.IsChecked) { '廃棄' }
        elseif ($ui.rbReturn.IsChecked) { '倉庫戻し' }
        elseif ($ui.rbSurplus.IsChecked) { '余剰化' }
        else { '充当' }
    }

    # 対象リストを操作種別に合わせて作り直す (VBA 版 RebuildTargets と同一ロジック)
    $script:rebuildTargets = {
        $ui = $script:ui
        $ev = & $script:currentEvent
        $ui.cmbTarget.Items.Clear()
        $ui.cmbState.Items.Clear()
        $ui.cmbTargetLot.Items.Clear()
        $ui.lblStock.Text = ''
        $script:uiTargets = @()

        foreach ($lt in @($script:In.lots)) {
            $has = if ($ev -eq '廃棄') { (Get-Avail 'L' $lt.lotId '不良') -gt 0 }
                   else { (Get-StatesWithStock 'L' $lt.lotId).Count -gt 0 }
            if ($has) {
                $script:uiTargets += , @($lt.lotId, $lt.localId)
                $suffix = if ($lt.planStatus -ne '有効') { "  [$($lt.planStatus)]" } else { '' }
                [void]$ui.cmbTarget.Items.Add("ロット: $($lt.lotId)$suffix")
            }
        }
        if ($ev -ne '余剰化') {
            foreach ($it in @($script:In.items)) {
                $has = if ($ev -eq '廃棄') { (Get-Avail 'U' $it.localId '不良') -gt 0 }
                       else { (Get-StatesWithStock 'U' $it.localId).Count -gt 0 }
                if ($has) {
                    $script:uiTargets += , @('', $it.localId)
                    [void]$ui.cmbTarget.Items.Add("未割当: $($it.localId)  $($it.itemNo)  $($it.itemName)")
                }
            }
        }
        $ui.cmbTargetLot.IsEnabled = ($ev -eq '充当')
        $ui.cmbState.IsEnabled = ($ev -ne '廃棄')
        if ($ev -eq '廃棄') { [void]$ui.cmbState.Items.Add('不良') }
    }

    foreach ($rb in $ui.rbDiscard, $ui.rbReturn, $ui.rbSurplus, $ui.rbAllocate) {
        $rb.Add_Checked({ & $script:rebuildTargets })
    }

    $ui.cmbTarget.Add_SelectionChanged({
        $ui = $script:ui
        $ui.lblStock.Text = ''
        if ($ui.cmbTarget.SelectedIndex -lt 0) { return }
        $t = $script:uiTargets[$ui.cmbTarget.SelectedIndex]
        $ev = & $script:currentEvent

        if ($ev -eq '廃棄') {
            $ui.cmbState.Items.Clear()
            [void]$ui.cmbState.Items.Add('不良')
            $ui.cmbState.SelectedIndex = 0
        } else {
            $ui.cmbState.Items.Clear()
            $bucket = if ($t[0]) { 'L' } else { 'U' }
            $id = if ($t[0]) { $t[0] } else { $t[1] }
            foreach ($st in (Get-StatesWithStock $bucket $id)) { [void]$ui.cmbState.Items.Add($st) }
        }

        if ($ev -eq '充当') {
            $ui.cmbTargetLot.Items.Clear()
            $script:uiTargetLots = @(@($script:In.lots) | Where-Object {
                $_.planStatus -eq '有効' -and $_.localId -eq $t[1] -and $_.lotId -ne $t[0] })
            foreach ($lt in $script:uiTargetLots) {
                [void]$ui.cmbTargetLot.Items.Add("$($lt.lotId)  (必要 $($lt.required))")
            }
        }
    })

    $ui.cmbState.Add_SelectionChanged({
        $ui = $script:ui
        if ($ui.cmbTarget.SelectedIndex -lt 0 -or $ui.cmbState.SelectedIndex -lt 0) { return }
        $t = $script:uiTargets[$ui.cmbTarget.SelectedIndex]
        $bucket = if ($t[0]) { 'L' } else { 'U' }
        $id = if ($t[0]) { $t[0] } else { $t[1] }
        $avail = Get-Avail $bucket $id ([string]$ui.cmbState.SelectedItem)
        $ui.lblStock.Text = "可用 $avail 個"
    })

    $ui.btnOK.Add_Click({
        $ui = $script:ui
        $ev = & $script:currentEvent
        if ($ui.cmbTarget.SelectedIndex -lt 0) { Show-Error '対象を選択してください。'; return }
        if ($ui.cmbState.SelectedIndex -lt 0) { Show-Error '元状態を選択してください。'; return }
        if ($ev -eq '充当' -and $ui.cmbTargetLot.SelectedIndex -lt 0) {
            Show-Error '充当先ロットを選択してください。'; return
        }
        $qty = Parse-Qty $ui.txtQty.Text
        if ($null -eq $qty) { return }
        $common = Read-CommonInputs $ui.dpDate $ui.cmbRecorder
        if ($null -eq $common) { return }

        $t = $script:uiTargets[$ui.cmbTarget.SelectedIndex]
        $op = New-EmptyOperation
        $op.kind = 'exception'
        $op.eventKind = $ev
        $op.lotId = $t[0]
        $op.localId = $t[1]
        $op.fromState = [string]$ui.cmbState.SelectedItem
        $op.qty = $qty
        $op.actionDate = $common.actionDate
        $op.recorder = $common.recorder
        $op.note = $ui.txtNote.Text.Trim()
        if ($ev -eq '充当') {
            $op.targetLotId = $script:uiTargetLots[$ui.cmbTargetLot.SelectedIndex].lotId
        }
        $script:dlgOp = $op
        $script:dlgWin.DialogResult = $true
    })

    & $script:rebuildTargets

    $script:dlgWin = $win
    $script:dlgOp = $null
    [void]$win.ShowDialog()
    $script:dlgOp
}

#==============================================================================
# テストモード用の組み込みデータ (現在のサンプルブック相当)
#==============================================================================
$script:TestInputJson = @'
{
  "schemaVersion": "1.2",
  "dialog": "defect",
  "defaults": { "actionDate": "2026-07-20", "recorder": "山田" },
  "labels": { "warehouseA": "倉庫A", "warehouseB": "倉庫B" },
  "recorders": ["山田", "佐藤"],
  "states": ["未処理", "アニール中", "アニール済み", "アセンブリ中", "アセンブリ済み",
             "梱包中", "梱包済み", "発送準備中", "発送準備完了", "不良"],
  "items": [
    { "localId": "a_001", "itemNo": "10000001", "itemName": "サンプル部品A(梱包のみ)", "pattern": 1 },
    { "localId": "a_002", "itemNo": "10000002", "itemName": "サンプル部品B(アセンブリ)", "pattern": 2 },
    { "localId": "b_001", "itemNo": "10000003", "itemName": "サンプル部品C(アニール)", "pattern": 3 },
    { "localId": "b_002", "itemNo": "10000004", "itemName": "サンプル部品D(アニール+アセンブリ)", "pattern": 4 }
  ],
  "lots": [
    { "lotId": "a_002_2026-07-13_納期先A", "localId": "a_002", "itemNo": "10000002",
      "shipWeek": "2026-07-13", "dest": "納期先A", "required": 5, "planStatus": "有効" },
    { "lotId": "a_002_2026-07-13_納期先B", "localId": "a_002", "itemNo": "10000002",
      "shipWeek": "2026-07-13", "dest": "納期先B", "required": 6, "planStatus": "有効" },
    { "lotId": "a_001_2026-07-20", "localId": "a_001", "itemNo": "10000001",
      "shipWeek": "2026-07-20", "dest": "", "required": 10, "planStatus": "有効" },
    { "lotId": "b_001_2026-07-20", "localId": "b_001", "itemNo": "10000003",
      "shipWeek": "2026-07-20", "dest": "", "required": 8, "planStatus": "有効" }
  ],
  "stock": {
    "lots": [
      { "lotId": "a_002_2026-07-13_納期先A", "byState": { "梱包済み": 8, "未処理": -3 } },
      { "lotId": "a_002_2026-07-13_納期先B", "byState": { "アセンブリ済み": 6 } },
      { "lotId": "a_001_2026-07-20", "byState": { "未処理": 10 } },
      { "lotId": "b_001_2026-07-20", "byState": { "アニール中": 6 } }
    ],
    "unallocated": [
      { "localId": "a_001", "byState": { "未処理": 2 } },
      { "localId": "a_002", "byState": { "未処理": 9 } }
    ]
  }
}
'@

#==============================================================================
# メイン
#==============================================================================

$testMode = (-not $InputPath -and -not $OutputPath)

if (-not $testMode -and (-not $InputPath -or -not $OutputPath)) {
    Write-Error 'InputPath と OutputPath は両方指定してください。'
    exit 2
}

if ($testMode) {
    $script:OutPath = Join-Path $PSScriptRoot 'ledger_dialog_test_out.json'
} else {
    $script:OutPath = $OutputPath
}

try {
    if ($testMode) {
        $script:In = $script:TestInputJson | ConvertFrom-Json
        $dialog = $TestDialog
    } else {
        if (-not (Test-Path -LiteralPath $InputPath)) {
            throw "入力ファイルが見つかりません: $InputPath"
        }
        $script:In = Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $dialog = [string]$script:In.dialog
    }

    # schemaVersion のメジャー不一致はエラー
    $major = ([string]$script:In.schemaVersion) -split '\.' | Select-Object -First 1
    $myMajor = $script:SchemaVersion -split '\.' | Select-Object -First 1
    if ($major -ne $myMajor) {
        throw "schemaVersion が一致しません (入力: $($script:In.schemaVersion) / 対応: $script:SchemaVersion)"
    }

    $script:States = @($script:In.states)
    $script:StockMap = Build-StockMap $script:In

    $op = switch ($dialog) {
        'carryIn'   { Show-CarryInDialog }
        'progress'  { Show-ProgressDialog }
        'defect'    { Show-DefectDialog }
        'exception' { Show-ExceptionDialog }
        default     { throw "不明な dialog です: $dialog" }
    }

    if ($null -eq $op) {
        Write-ResultJson -Status 'cancel' -Message '' -Operation $null
        if ($testMode) {
            [void][System.Windows.MessageBox]::Show('キャンセルされました。', 'テストモード')
        }
        exit 1
    }

    Write-ResultJson -Status 'ok' -Message '' -Operation $op
    if ($testMode) {
        $summary = ($op.GetEnumerator() | ForEach-Object { "$($_.Key) = $($_.Value)" }) -join "`n"
        [void][System.Windows.MessageBox]::Show($summary, 'テストモード: 入力結果')
        Invoke-Item $script:OutPath
    }
    exit 0
}
catch {
    Write-ResultJson -Status 'error' -Message $_.Exception.Message -Operation $null
    if ($testMode) {
        [void][System.Windows.MessageBox]::Show($_.Exception.Message, 'テストモード: エラー')
    }
    exit 2
}
