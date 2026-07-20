# build_xlsm.ps1
#
# 台帳入力ツール.xlsm を Excel COM + VBE 拡張で自動生成する。
#   - 「操作」シート (ボタン群) と「設定」シート (名前付きセル) を作成
#   - UserForm 4 画面をデザイン時コントロール込みで生成 (.frx 不要)
#   - .bas / .cls をインポートし、フォームコード (.formcode.txt) を注入
#   - 保存後に SmokeTest マクロで疎通確認
#
# 前提: Excel のオプションで「VBA プロジェクト オブジェクト モデルへの
# アクセスを信頼する」が有効であること (ファイル > オプション >
# トラスト センター > トラスト センターの設定 > マクロの設定)。
# このスクリプトは設定の確認のみ行い、変更はしない。
#
# 使い方:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File build_xlsm.ps1
#   (省略可) -OutPath <出力xlsm> -LedgerPath <台帳xlsx>

param(
    [string]$ToolDir = $PSScriptRoot,
    [string]$OutPath = (Join-Path $PSScriptRoot '台帳入力ツール.xlsm'),
    [string]$LedgerPath = 'E:\LocalAppsWorks\InventoryManager\在庫管理台帳.xlsx'
)

$ErrorActionPreference = 'Stop'
$sjis = [System.Text.Encoding]::GetEncoding(932)

#==============================================================================
# コントロール生成ヘルパー
#==============================================================================

function Add-Ctl {
    param($Designer, [string]$Type, [string]$Name,
          [double]$L, [double]$T, [double]$W, [double]$H, [string]$Caption = '')
    $progId = switch ($Type) {
        'label'  { 'Forms.Label.1' }
        'text'   { 'Forms.TextBox.1' }
        'combo'  { 'Forms.ComboBox.1' }
        'check'  { 'Forms.CheckBox.1' }
        'option' { 'Forms.OptionButton.1' }
        'button' { 'Forms.CommandButton.1' }
    }
    $c = $Designer.Controls.Add($progId, $Name, $true)
    $c.Left = $L; $c.Top = $T; $c.Width = $W; $c.Height = $H
    switch ($Type) {
        'label' {
            $c.Caption = $Caption
            $c.WordWrap = $true
            $c.AutoSize = $false
        }
        'combo'  { $c.Style = 2 }   # fmStyleDropDownList (選択のみ・自由入力不可)
        'check'  { $c.Caption = $Caption }
        'option' { $c.Caption = $Caption }
        'button' {
            $c.Caption = $Caption
            if ($Name -eq 'btnOK') { $c.Default = $true }
            if ($Name -eq 'btnCancel') { $c.Cancel = $true }
        }
    }
    $c
}

# 共通フッター (実施日/記録者/備考/OK/キャンセル)。フォームの必要高さを返す
function Add-Footer {
    param($Designer, [double]$Y)
    [void](Add-Ctl $Designer 'label' 'lblDateCap' 12 $Y 100 12 '実施日')
    [void](Add-Ctl $Designer 'text' 'txtDate' 12 ($Y + 14) 100 18)
    [void](Add-Ctl $Designer 'label' 'lblRecCap' 12 ($Y + 40) 100 12 '記録者')
    [void](Add-Ctl $Designer 'combo' 'cmbRecorder' 12 ($Y + 54) 140 18)
    [void](Add-Ctl $Designer 'label' 'lblNoteCap' 12 ($Y + 80) 100 12 '備考')
    [void](Add-Ctl $Designer 'text' 'txtNote' 12 ($Y + 94) 316 18)
    [void](Add-Ctl $Designer 'button' 'btnOK' 172 ($Y + 124) 76 24 'OK')
    [void](Add-Ctl $Designer 'button' 'btnCancel' 252 ($Y + 124) 76 24 'キャンセル')
    $Y + 124 + 24 + 12   # コントロール下端 + 余白
}

function New-FormComponent {
    param($VbProject, [string]$Name, [string]$Caption, [scriptblock]$Layout)
    $comp = $VbProject.VBComponents.Add(3)   # vbext_ct_MSForm
    $comp.Name = $Name
    $comp.Properties.Item('Caption').Value = $Caption
    $bottom = & $Layout $comp.Designer
    $comp.Properties.Item('Width').Value = 348
    $comp.Properties.Item('Height').Value = $bottom + 28   # タイトルバー分
    $comp
}

#==============================================================================
# 事前チェック
#==============================================================================

foreach ($f in @('m_JsonLite.bas', 'm_LedgerCore.bas', 'm_LedgerDialogs.bas', 'm_LedgerFormUtil.bas',
                 'c_ItemInfo.cls', 'c_LotInfo.cls', 'c_StockSnapshot.cls', 'c_DialogContext.cls',
                 'c_LedgerOperation.cls', 'c_LedgerDialogResult.cls', 'c_ILedgerEventDialog.cls',
                 'c_FormLedgerEventDialog.cls', 'c_PsLedgerEventDialog.cls',
                 'CarryInForm.formcode.txt', 'ProgressForm.formcode.txt',
                 'DefectForm.formcode.txt', 'ExceptionForm.formcode.txt')) {
    if (-not (Test-Path -LiteralPath (Join-Path $ToolDir $f))) {
        throw "必要なファイルがありません: $f"
    }
}

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # VBE 信頼設定の確認 (変更はしない)
    $ver = $excel.Version
    $accessVBOM = 0
    try {
        $accessVBOM = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ver\Excel\Security" `
            -Name AccessVBOM -ErrorAction Stop).AccessVBOM
    } catch { $accessVBOM = 0 }
    if ($accessVBOM -ne 1) {
        throw ('VBA プロジェクトへのアクセスが信頼されていません。Excel の ' +
               'ファイル > オプション > トラスト センター > トラスト センターの設定 > ' +
               'マクロの設定 で「VBA プロジェクト オブジェクト モデルへのアクセスを' +
               '信頼する」を有効にしてから再実行してください (ビルド時のみ必要)。')
    }

    if (Test-Path -LiteralPath $OutPath) { Remove-Item -LiteralPath $OutPath -Force }

    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(2).Delete() }

    #--------------------------------------------------------------------------
    # 「操作」シート
    #--------------------------------------------------------------------------
    $wsOp = $wb.Worksheets.Item(1)
    $wsOp.Name = '操作'
    $wsOp.Range('B2').Value2 = '在庫管理台帳 入力ツール'
    $wsOp.Range('B2').Font.Size = 14
    $wsOp.Range('B2').Font.Bold = $true
    $wsOp.Range('B3').Value2 = '通常の入力は上の 4 つのボタンから。台帳を直接修正するときだけ保護解除を使う。'

    $buttons = @(
        @('搬入 (倉庫A→倉庫B)',        'ShowCarryInDialog',   30, 60),
        @('工程進行 / 発送',           'ShowProgressDialog',  30, 100),
        @('不良発生 (+補填)',          'ShowDefectDialog',    30, 140),
        @('例外操作 (廃棄・戻し 等)',  'ShowExceptionDialog', 30, 180),
        @('保護解除 (直接修正用)',     'UnprotectLedgerNow',  240, 60),
        @('再保護 + 保存',             'ProtectLedgerNow',    240, 100)
    )
    foreach ($b in $buttons) {
        $shape = $wsOp.Shapes.AddFormControl(0, $b[2], $b[3], 180, 30)  # 0 = xlButtonControl
        $shape.TextFrame.Characters().Text = $b[0]
        $shape.OnAction = $b[1]
    }

    #--------------------------------------------------------------------------
    # 「設定」シート (名前付きセル)
    #--------------------------------------------------------------------------
    $wsCfg = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wsOp)
    $wsCfg.Name = '設定'
    $wsCfg.Range('A1').Value2 = '設定 (B 列を編集する)'
    $wsCfg.Range('A1').Font.Bold = $true
    $wsCfg.Range('A3').Value2 = '台帳パス'
    $wsCfg.Range('B3').Value2 = $LedgerPath
    $wsCfg.Range('A4').Value2 = 'UI 実装 (Forms / PowerShell)'
    $wsCfg.Range('B4').Value2 = 'Forms'
    $wsCfg.Range('A5').Value2 = 'PS スクリプトパス (空 = ツールと同じフォルダ)'
    $wsCfg.Range('B5').Value2 = ''
    $wsCfg.Range('A6').Value2 = 'JSON フォルダ (空 = %TEMP%)'
    $wsCfg.Range('B6').Value2 = ''
    $wsCfg.Columns.Item(1).ColumnWidth = 42
    $wsCfg.Columns.Item(2).ColumnWidth = 60
    [void]$wsCfg.Range('B4').Validation.Add(3, 1, 1, 'Forms,PowerShell')   # 3 = xlValidateList

    [void]$wb.Names.Add('cfgLedgerPath', '=設定!$B$3')
    [void]$wb.Names.Add('cfgUiKind', '=設定!$B$4')
    [void]$wb.Names.Add('cfgPsScriptPath', '=設定!$B$5')
    [void]$wb.Names.Add('cfgJsonDir', '=設定!$B$6')

    #--------------------------------------------------------------------------
    # UserForm 4 画面 (デザイン時コントロール生成)
    #--------------------------------------------------------------------------
    $vbp = $wb.VBProject

    $forms = @{}
    $forms['CarryInForm'] = New-FormComponent $vbp 'CarryInForm' '搬入' {
        param($d)
        [void](Add-Ctl $d 'label' 'lblItemCap' 12 10 100 12 '品番')
        [void](Add-Ctl $d 'combo' 'cmbItem' 12 24 316 18)
        [void](Add-Ctl $d 'label' 'lblPool' 12 46 316 24 '')
        [void](Add-Ctl $d 'label' 'lblQtyCap' 12 76 100 12 '数量')
        [void](Add-Ctl $d 'text' 'txtQty' 12 90 100 18)
        Add-Footer $d 116
    }

    $forms['ProgressForm'] = New-FormComponent $vbp 'ProgressForm' '工程進行 / 発送' {
        param($d)
        [void](Add-Ctl $d 'label' 'lblLotCap' 12 10 100 12 'ロット')
        [void](Add-Ctl $d 'combo' 'cmbLot' 12 24 316 18)
        [void](Add-Ctl $d 'label' 'lblItem' 12 46 316 24 '')
        [void](Add-Ctl $d 'label' 'lblEvCap' 12 76 100 12 'イベント')
        [void](Add-Ctl $d 'combo' 'cmbEvent' 12 90 160 18)
        [void](Add-Ctl $d 'label' 'lblFrom' 12 112 316 12 '')
        [void](Add-Ctl $d 'label' 'lblQtyCap' 12 130 100 12 '数量')
        [void](Add-Ctl $d 'text' 'txtQty' 12 144 100 18)
        Add-Footer $d 170
    }

    $forms['DefectForm'] = New-FormComponent $vbp 'DefectForm' '不良発生 (+補填)' {
        param($d)
        [void](Add-Ctl $d 'label' 'lblLotCap' 12 10 100 12 'ロット')
        [void](Add-Ctl $d 'combo' 'cmbLot' 12 24 316 18)
        [void](Add-Ctl $d 'label' 'lblItem' 12 46 316 12 '')
        [void](Add-Ctl $d 'label' 'lblStCap' 12 64 200 12 '不良になった在庫の元状態')
        [void](Add-Ctl $d 'combo' 'cmbState' 12 78 160 18)
        [void](Add-Ctl $d 'label' 'lblStock' 180 80 130 12 '')
        [void](Add-Ctl $d 'label' 'lblQtyCap' 12 102 100 12 '不良数量')
        [void](Add-Ctl $d 'text' 'txtQty' 12 116 100 18)
        [void](Add-Ctl $d 'check' 'chkRefill' 12 142 316 16 '未割当在庫から補填する (充当行を自動生成)')
        [void](Add-Ctl $d 'label' 'lblRsCap' 12 162 200 12 '補填元の状態')
        [void](Add-Ctl $d 'combo' 'cmbRefillState' 12 176 160 18)
        [void](Add-Ctl $d 'label' 'lblRqCap' 12 200 200 12 '補填数量 (上限 = 不良数量)')
        [void](Add-Ctl $d 'text' 'txtRefillQty' 12 214 100 18)
        [void](Add-Ctl $d 'label' 'lblPool' 12 238 316 24 '')
        Add-Footer $d 268
    }

    $forms['ExceptionForm'] = New-FormComponent $vbp 'ExceptionForm' '例外操作' {
        param($d)
        [void](Add-Ctl $d 'label' 'lblOpCap' 12 10 100 12 '操作')
        [void](Add-Ctl $d 'option' 'optDiscard' 12 24 56 16 '廃棄')
        [void](Add-Ctl $d 'option' 'optReturn' 76 24 84 16 '倉庫A戻し')
        [void](Add-Ctl $d 'option' 'optSurplus' 168 24 64 16 '余剰化')
        [void](Add-Ctl $d 'option' 'optAllocate' 240 24 56 16 '充当')
        [void](Add-Ctl $d 'label' 'lblTgtCap' 12 46 200 12 '対象 (ロット / 未割当)')
        [void](Add-Ctl $d 'combo' 'cmbTarget' 12 60 316 18)
        [void](Add-Ctl $d 'label' 'lblStCap' 12 84 100 12 '元状態')
        [void](Add-Ctl $d 'combo' 'cmbState' 12 98 160 18)
        [void](Add-Ctl $d 'label' 'lblStock' 180 100 130 12 '')
        [void](Add-Ctl $d 'label' 'lblQtyCap' 12 122 100 12 '数量')
        [void](Add-Ctl $d 'text' 'txtQty' 12 136 100 18)
        [void](Add-Ctl $d 'label' 'lblAlCap' 12 160 200 12 '充当先ロット (充当のみ)')
        [void](Add-Ctl $d 'combo' 'cmbTargetLot' 12 174 316 18)
        Add-Footer $d 200
    }

    #--------------------------------------------------------------------------
    # モジュールのインポートとフォームコードの注入
    #--------------------------------------------------------------------------
    foreach ($f in @('m_JsonLite.bas', 'm_LedgerCore.bas', 'm_LedgerFormUtil.bas', 'm_LedgerDialogs.bas',
                     'c_ItemInfo.cls', 'c_LotInfo.cls', 'c_StockSnapshot.cls', 'c_DialogContext.cls',
                     'c_LedgerOperation.cls', 'c_LedgerDialogResult.cls', 'c_ILedgerEventDialog.cls',
                     'c_FormLedgerEventDialog.cls', 'c_PsLedgerEventDialog.cls')) {
        [void]$vbp.VBComponents.Import((Join-Path $ToolDir $f))
    }

    foreach ($name in $forms.Keys) {
        $code = [System.IO.File]::ReadAllText((Join-Path $ToolDir "$name.formcode.txt"), $sjis)
        $cm = $forms[$name].CodeModule
        # VBE の「変数の宣言を強制する」設定で自動挿入される Option Explicit を除去
        # (注入コード側の Option Explicit と重複するとコンパイルエラーになる)
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString($code)
    }

    #--------------------------------------------------------------------------
    # 疎通確認と保存 (保存前の新規ブックならマクロ設定に関係なく Run できる)
    #--------------------------------------------------------------------------
    $result = $excel.Run('SmokeTest')
    if ($result -ne 'ok') { throw "SmokeTest が失敗しました: $result" }

    $wsOp.Activate()
    $wb.SaveAs($OutPath, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled

    Write-Host "ビルド完了: $OutPath (SmokeTest: $result)"
}
finally {
    if ($wb) { $wb.Close($true) }
    if ($excel) { $excel.Quit() }
    foreach ($o in $wb, $excel) {
        if ($o) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
