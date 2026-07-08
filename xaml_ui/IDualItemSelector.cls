VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDualItemSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' インターフェイス: 2 リストボックス (左右) アイテム選択ダイアログ
'
' 実装クラス:
'   PsDualItemSelector   ... PowerShell (WPF) 版。JSON ファイルで受け渡し
'   FormDualItemSelector ... VBA UserForm 版。オブジェクトを直接受け渡し
'
' 利用側は SelectItemsDialog.CreateDualItemSelector() で実装を取得する。
'==============================================================================

' 左右 2 つのリストボックスを持つ選択ダイアログを表示し、結果を返す
'   leftList / rightList : 各リストの仕様 (アイテム・見出し・選択モード)
'   title                : ウィンドウタイトル
Public Function SelectItems(ByVal leftList As ItemListSpec, _
                            ByVal rightList As ItemListSpec, _
                            ByVal title As String) As SelectDualItemsResult
End Function
