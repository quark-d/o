VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_StockSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' V_在庫内訳のスナップショット (ロット×状態 / 未割当×localId×状態 → 数量)
'
' キーは "L|ロットID|状態" または "U|localId|状態"。
' ブックの既知の不整合により数量が負のバケットが存在し得る。
' 消費可能量としては負は 0 として扱う (AvailableQty)。
'==============================================================================

Private mDict As Object        ' Scripting.Dictionary: key → Long
Private mStates As Collection  ' 状態の表示順 (M_リスト B列の順)

Private Sub Class_Initialize()
    Set mDict = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Init(ByVal states As Collection)
    Set mStates = states
End Sub

Public Property Get States() As Collection
    Set States = mStates
End Property

'------------------------------------------------------------------------------
' 登録・取得
'------------------------------------------------------------------------------
Public Sub SetLotQty(ByVal lotId As String, ByVal state As String, ByVal qty As Long)
    mDict(Me.BucketKey("L", lotId, state)) = qty
End Sub

Public Sub SetUnallocQty(ByVal localId As String, ByVal state As String, ByVal qty As Long)
    mDict(Me.BucketKey("U", localId, state)) = qty
End Sub

' 生の数量 (負もあり得る)。未登録は 0
Public Function RawQty(ByVal bucket As String, ByVal id As String, _
                       ByVal state As String) As Long
    Dim key As String: key = Me.BucketKey(bucket, id, state)
    If mDict.Exists(key) Then RawQty = mDict(key)
End Function

' 消費可能量 (負は 0 に丸める)
Public Function AvailableQty(ByVal bucket As String, ByVal id As String, _
                             ByVal state As String) As Long
    Dim q As Long: q = Me.RawQty(bucket, id, state)
    If q > 0 Then AvailableQty = q
End Function

' 在庫が正の状態の一覧 (mStates の順で返す)
Public Function StatesWithStock(ByVal bucket As String, ByVal id As String) As Collection
    Dim result As New Collection
    Dim st As Variant
    For Each st In mStates
        If Me.AvailableQty(bucket, id, CStr(st)) > 0 Then result.Add CStr(st)
    Next
    Set StatesWithStock = result
End Function

'------------------------------------------------------------------------------
' シミュレーション用 (m_LedgerCore の検証でのみ使用)
'------------------------------------------------------------------------------
' 検証用の複製辞書を返す。キー体系はこのクラスと同一
Public Function CloneDict() As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    For Each key In mDict.Keys
        result(key) = mDict(key)
    Next
    Set CloneDict = result
End Function

' 現時点で負のバケットのキー一覧 (書込後の負在庫チェックの除外リストに使う)
Public Function NegativeBucketKeys() As Collection
    Dim result As New Collection
    Dim key As Variant
    For Each key In mDict.Keys
        If mDict(key) < 0 Then result.Add CStr(key)
    Next
    Set NegativeBucketKeys = result
End Function

Public Function BucketKey(ByVal bucket As String, ByVal id As String, _
                          ByVal state As String) As String
    BucketKey = bucket & "|" & id & "|" & state
End Function
