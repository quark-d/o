Attribute VB_Name = "JsonLite"
Option Explicit

'==============================================================================
' 簡易 JSON ヘルパー + UTF-8 ファイル I/O (内部実装用モジュール)
'
' Select-Items.ps1 / Select-DualItems.ps1 との受け渡し JSON
' (固定フォーマット) 専用の簡易実装。汎用 JSON パーサーではない。
' PsItemSelector / PsDualItemSelector の内部でのみ使うこと。
' JSON を利用側コード・インターフェイスに露出させない (README 参照)。
'==============================================================================

'------------------------------------------------------------------------------
' ファイル I/O (UTF-8)
'------------------------------------------------------------------------------
Public Sub WriteTextUtf8Bom(ByVal path As String, ByVal text As String)
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

Public Function ReadTextUtf8(ByVal path As String) As String
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
' JSON ヘルパー
'------------------------------------------------------------------------------
Public Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

Public Function JsonUnescape(ByVal s As String) As String
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
Public Function JsonExtractString(ByVal json As String, ByVal key As String) As String
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
Public Sub JsonExtractStringArray(ByVal json As String, ByVal key As String, _
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
