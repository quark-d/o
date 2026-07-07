VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IItemSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' インターフェイス: アイテム選択ダイアログ
'
' 実装クラス:
'   PsItemSelector   ... PowerShell (WPF) 版。JSON ファイルで受け渡す
'   FormItemSelector ... VBA UserForm 版。Collection を直接受け渡す
'
' 利用側は SelectItemsDialog.CreateItemSelector() で実装を取得する。
'==============================================================================

' アイテム選択ダイアログを表示し、結果を返す
'   items       : 表示するアイテム (文字列の Collection)
'   title       : ウィンドウタイトル
'   multiSelect : True = 複数選択可
Public Function SelectItems(ByVal items As Collection, ByVal title As String, _
                            ByVal multiSelect As Boolean) As SelectItemsResult
End Function
