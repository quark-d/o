VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_LedgerOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' ダイアログの出力 (操作レベルの入力内容)
'
' UI はこれを返すだけ。台帳行への展開・検証は m_LedgerCore が行う。
' フラットな 1 クラス + Kind で、種別ごとに使うフィールドが異なる
' (設計合意 2026-07-20。未使用フィールドは既定値のまま)。
'==============================================================================

Public Kind As LedgerOpKind        ' 操作種別 (m_LedgerCore.bas の Enum)
Public EventKind As String         ' イベント種類 (M_リスト A列の値)
Public ItemNo As String            ' 品番 (搬入・未割当対象のとき。ロット対象なら導出)
Public LotId As String             ' 対象ロットID (空 = 未割当/品番プール)
Public FromState As String         ' 元状態 (工程進行はコアが導出して設定)
Public Qty As Long                 ' 数量
Public ActionDate As Date          ' 実施日
Public Recorder As String          ' 記録者
Public Note As String              ' 備考

' 不良+補填のみ
Public Refill As Boolean           ' 補填する
Public RefillFromState As String   ' 補填元状態 (未割当プールの状態)
Public RefillQty As Long           ' 補填数量 (上限 = 不良数量)

' 例外操作の充当のみ (不良+補填ではコアが LotId を設定する)
Public TargetLotId As String       ' 充当先ロットID
