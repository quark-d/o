VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PsDualItemSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' IDualItemSelector 実装: PowerShell (WPF) 版
'
' Select-DualItems.ps1 を同期実行し、UTF-8 (BOM 付き) の JSON ファイルで
' データを受け渡す。JSON の生成・解析はこのクラスと JsonLite
' (内部実装用モジュール) に閉じている。
'==============================================================================

Implements IDualItemSelector

' Select-DualItems.ps1 のパス (既定値は Class_Initialize で設定。差し替え可)
Public ScriptPath As String
' 受け渡し用 JSON ファイルのパス (既定は %TEMP%。差し替え可)
Public InputPath As String
Public OutputPath As String

Private Sub Class_Initialize()
    ScriptPath = "D:\Apps\PowerShell\xaml_ui\Select-DualItems.ps1"
    InputPath = Environ$("TEMP") & "\select_dual_items_in.json"
    OutputPath = Environ$("TEMP") & "\select_dual_items_out.json"
End Sub

'------------------------------------------------------------------------------
' IDualItemSelector
'------------------------------------------------------------------------------
Private Function IDualItemSelector_SelectItems(ByVal leftList As ItemListSpec, _
                                               ByVal rightList As ItemListSpec, _
                                               ByVal title As String) As SelectDualItemsResult
    Dim result As New SelectDualItemsResult
    On Error GoTo Fail

    ' 古い出力ファイルが残っていれば削除 (誤読防止)
    On Error Resume Next
    Kill OutputPath
    On Error GoTo Fail

    JsonLite.WriteTextUtf8Bom InputPath, BuildInputJson(leftList, rightList, title)

    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA" & _
          " -File """ & ScriptPath & """" & _
          " -InputPath """ & InputPath & """" & _
          " -OutputPath """ & OutputPath & """"

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True   ' 0=コンソール非表示, True=終了まで待機

    If Dir(OutputPath) = "" Then
        result.Init "error", Nothing, Nothing, _
                    "出力ファイルが作成されませんでした: " & OutputPath
    Else
        ParseOutputJson JsonLite.ReadTextUtf8(OutputPath), result
    End If

    Set IDualItemSelector_SelectItems = result
    Exit Function

Fail:
    result.Init "error", Nothing, Nothing, Err.Description
    Set IDualItemSelector_SelectItems = result
End Function

'------------------------------------------------------------------------------
' 入力 JSON の生成
'------------------------------------------------------------------------------
Private Function BuildInputJson(ByVal leftList As ItemListSpec, _
                                ByVal rightList As ItemListSpec, _
                                ByVal title As String) As String
    Dim json As String
    json = "{" & vbCrLf
    json = json & "  ""title"": """ & JsonLite.JsonEscape(title) & """," & vbCrLf
    json = json & "  ""left"": " & BuildListJson(leftList) & "," & vbCrLf
    json = json & "  ""right"": " & BuildListJson(rightList) & vbCrLf
    json = json & "}"
    BuildInputJson = json
End Function

' リスト 1 つ分 { "caption": ..., "multiSelect": ..., "items": [...] } を作る
Private Function BuildListJson(ByVal spec As ItemListSpec) As String
    Dim json As String
    Dim first As Boolean
    Dim v As Variant

    json = "{ ""caption"": """ & JsonLite.JsonEscape(spec.Caption) & """, " & _
           """multiSelect"": " & LCase$(CStr(spec.MultiSelect)) & ", " & _
           """items"": ["
    first = True
    For Each v In spec.Items
        If Not first Then json = json & ", "
        json = json & """" & JsonLite.JsonEscape(CStr(v)) & """"
        first = False
    Next
    json = json & "] }"

    BuildListJson = json
End Function

'------------------------------------------------------------------------------
' 出力 JSON の解析
'------------------------------------------------------------------------------
Private Sub ParseOutputJson(ByVal json As String, ByRef result As SelectDualItemsResult)
    Dim status As String
    status = JsonLite.JsonExtractString(json, "status")

    Select Case status
        Case "ok"
            Dim leftSel As New Collection
            Dim rightSel As New Collection
            JsonLite.JsonExtractStringArray json, "left", leftSel
            JsonLite.JsonExtractStringArray json, "right", rightSel
            result.Init "ok", leftSel, rightSel, ""
        Case "cancel"
            result.Init "cancel", Nothing, Nothing, ""
        Case Else
            result.Init "error", Nothing, Nothing, _
                        JsonLite.JsonExtractString(json, "message")
    End Select
End Sub
