VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PsItemSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' IItemSelector 実装: PowerShell (WPF) 版
'
' Select-Items.ps1 を同期実行し、UTF-8 (BOM 付き) の JSON ファイルで
' データを受け渡す。JSON の生成・解析はこのクラス内に閉じている。
'==============================================================================

Implements IItemSelector

' Select-Items.ps1 のパス (既定値は Class_Initialize で設定。差し替え可)
Public ScriptPath As String
' 受け渡し用 JSON ファイルのパス (既定は %TEMP%。差し替え可)
Public InputPath As String
Public OutputPath As String

Private Sub Class_Initialize()
    ScriptPath = "D:\Apps\PowerShell\xaml_ui\Select-Items.ps1"
    InputPath = Environ$("TEMP") & "\select_items_in.json"
    OutputPath = Environ$("TEMP") & "\select_items_out.json"
End Sub

'------------------------------------------------------------------------------
' IItemSelector
'------------------------------------------------------------------------------
Private Function IItemSelector_SelectItems(ByVal items As Collection, _
                                           ByVal title As String, _
                                           ByVal multiSelect As Boolean) As SelectItemsResult
    Dim result As New SelectItemsResult
    On Error GoTo Fail

    ' 古い出力ファイルが残っていれば削除 (誤読防止)
    On Error Resume Next
    Kill OutputPath
    On Error GoTo Fail

    WriteTextUtf8Bom InputPath, BuildInputJson(items, title, multiSelect)

    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA" & _
          " -File """ & ScriptPath & """" & _
          " -InputPath """ & InputPath & """" & _
          " -OutputPath """ & OutputPath & """"

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True   ' 0=コンソール非表示, True=終了まで待機

    If Dir(OutputPath) = "" Then
        result.Init "error", Nothing, "出力ファイルが作成されませんでした: " & OutputPath
    Else
        ParseOutputJson ReadTextUtf8(OutputPath), result
    End If

    Set IItemSelector_SelectItems = result
    Exit Function

Fail:
    result.Init "error", Nothing, Err.Description
    Set IItemSelector_SelectItems = result
End Function

'------------------------------------------------------------------------------
' 入力 JSON の生成
'------------------------------------------------------------------------------
Private Function BuildInputJson(ByVal items As Collection, ByVal title As String, _
                                ByVal multiSelect As Boolean) As String
    Dim json As String
    Dim first As Boolean
    Dim v As Variant

    json = "{" & vbCrLf
    json = json & "  ""title"": """ & JsonEscape(title) & """," & vbCrLf
    json = json & "  ""multiSelect"": " & LCase$(CStr(multiSelect)) & "," & vbCrLf
    json = json & "  ""items"": ["
    first = True
    For Each v In items
        If Not first Then json = json & ", "
        json = json & """" & JsonEscape(CStr(v)) & """"
        first = False
    Next
    json = json & "]" & vbCrLf & "}"

    BuildInputJson = json
End Function

'------------------------------------------------------------------------------
' 出力 JSON の解析
'------------------------------------------------------------------------------
Private Sub ParseOutputJson(ByVal json As String, ByRef result As SelectItemsResult)
    Dim status As String
    status = JsonExtractString(json, "status")

    Select Case status
        Case "ok"
            Dim selected As New Collection
            JsonExtractStringArray json, "selected", selected
            result.Init "ok", selected, ""
        Case "cancel"
            result.Init "cancel", Nothing, ""
        Case Else
            result.Init "error", Nothing, JsonExtractString(json, "message")
    End Select
End Sub

'------------------------------------------------------------------------------
' ファイル I/O (UTF-8)
'------------------------------------------------------------------------------
Private Sub WriteTextUtf8Bom(ByVal path As String, ByVal text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2               ' adTypeText
        .Charset = "UTF-8"      ' BOM 付きで出力される
        .Open
        .WriteText text
        .SaveToFile path, 2     ' adSaveCreateOverWrite
        .Close
    End With
End Sub

Private Function ReadTextUtf8(ByVal path As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2               ' adTypeText
        .Charset = "UTF-8"      ' BOM の有無どちらも可
        .Open
        .LoadFromFile path
        ReadTextUtf8 = .ReadText(-1)   ' adReadAll
        .Close
    End With
End Function

'------------------------------------------------------------------------------
' JSON ヘルパー (Select-Items.ps1 の固定フォーマット専用の簡易実装)
'------------------------------------------------------------------------------
Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

Private Function JsonUnescape(ByVal s As String) As String
    Dim result As String
    Dim i As Long
    i = 1
    Do While i <= Len(s)
        Dim c As String
        c = Mid$(s, i, 1)
        If c = "\" And i < Len(s) Then
            Dim nxt As String
            nxt = Mid$(s, i + 1, 1)
            Select Case nxt
                Case """", "\", "/"
                    result = result & nxt
                    i = i + 2
                Case "n"
                    result = result & vbLf
                    i = i + 2
                Case "r"
                    result = result & vbCr
                    i = i + 2
                Case "t"
                    result = result & vbTab
                    i = i + 2
                Case "u"
                    result = result & ChrW$(CLng("&H" & Mid$(s, i + 2, 4)))
                    i = i + 6
                Case Else
                    result = result & c
                    i = i + 1
            End Select
        Else
            result = result & c
            i = i + 1
        End If
    Loop
    JsonUnescape = result
End Function

' "key": "value" の value を取り出す
Private Function JsonExtractString(ByVal json As String, ByVal key As String) As String
    Dim p As Long, q As Long
    p = InStr(json, """" & key & """")
    If p = 0 Then Exit Function
    p = InStr(p, json, ":")
    If p = 0 Then Exit Function
    p = InStr(p, json, """")
    If p = 0 Then Exit Function
    ' 閉じ引用符を探す (エスケープされた \" は読み飛ばす)
    q = p + 1
    Do While q <= Len(json)
        If Mid$(json, q, 1) = """" And Mid$(json, q - 1, 1) <> "\" Then Exit Do
        q = q + 1
    Loop
    JsonExtractString = JsonUnescape(Mid$(json, p + 1, q - p - 1))
End Function

' "key": [ "a", "b", ... ] の配列要素を Collection に取り出す
Private Sub JsonExtractStringArray(ByVal json As String, ByVal key As String, _
                                   ByRef result As Collection)
    Dim p As Long
    p = InStr(json, """" & key & """")
    If p = 0 Then Exit Sub
    p = InStr(p, json, "[")
    If p = 0 Then Exit Sub

    Dim arrEnd As Long
    arrEnd = InStr(p, json, "]")
    If arrEnd = 0 Then Exit Sub

    Dim i As Long
    i = p + 1
    Do While i < arrEnd
        If Mid$(json, i, 1) = """" Then
            Dim q As Long
            q = i + 1
            Do While q <= Len(json)
                If Mid$(json, q, 1) = """" And Mid$(json, q - 1, 1) <> "\" Then Exit Do
                q = q + 1
            Loop
            result.Add JsonUnescape(Mid$(json, i + 1, q - i - 1))
            i = q + 1
            ' 文字列内に ] があった場合に備えて配列の終端を取り直す
            arrEnd = InStr(i, json, "]")
            If arrEnd = 0 Then Exit Sub
        Else
            i = i + 1
        End If
    Loop
End Sub
