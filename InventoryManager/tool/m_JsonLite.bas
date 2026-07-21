Attribute VB_Name = "m_JsonLite"
Option Explicit

'==============================================================================
' 簡易 JSON ヘルパー + UTF-8 ファイル I/O (内部実装用モジュール)
'
' Show-LedgerDialog.ps1 との受け渡し JSON (固定フォーマット) 専用の簡易実装。
' 汎用 JSON パーサーではない。c_PsLedgerEventDialog の内部でのみ使うこと。
' JSON を利用側コード・インターフェイスに露出させないこと (README 参照)。
'
' D:\Apps\PowerShell\xaml_ui\JsonLite.bas を元に、数値/真偽値の抽出を追加。
'==============================================================================

'------------------------------------------------------------------------------
' ファイル I/O (UTF-8)
'------------------------------------------------------------------------------
Public Sub WriteTextUtf8Bom(ByVal path As String, ByVal text As String)
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
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
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
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
    Dim i As Long: i = 1
    Do While i <= Len(s)
        Dim c As String: c = Mid$(s, i, 1)
        If c = "\" And i < Len(s) Then
            Dim nxt As String: nxt = Mid$(s, i + 1, 1)
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

' "key": "value" の value を取り出す (見つからなければ空文字)
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
    JsonExtractString = m_JsonLite.JsonUnescape(Mid$(json, p + 1, q - p - 1))
End Function

' "key": 123 の数値を取り出す (見つからなければ既定値)
Public Function JsonExtractLong(ByVal json As String, ByVal key As String, _
                                Optional ByVal defaultValue As Long = 0) As Long
    Dim raw As String: raw = ExtractRawValue(json, key)
    If raw = "" Or Not IsNumeric(raw) Then
        JsonExtractLong = defaultValue
    Else
        JsonExtractLong = CLng(Val(raw))
    End If
End Function

' "key": true/false を取り出す (見つからなければ既定値)
Public Function JsonExtractBool(ByVal json As String, ByVal key As String, _
                                Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim raw As String: raw = LCase$(ExtractRawValue(json, key))
    Select Case raw
        Case "true"
            JsonExtractBool = True
        Case "false"
            JsonExtractBool = False
        Case Else
            JsonExtractBool = defaultValue
    End Select
End Function

' key の直後の「引用符なし生値」(数値・true/false・null) を取り出す内部関数
Private Function ExtractRawValue(ByVal json As String, ByVal key As String) As String
    Dim p As Long, q As Long
    p = InStr(json, """" & key & """")
    If p = 0 Then Exit Function
    p = InStr(p, json, ":")
    If p = 0 Then Exit Function
    p = p + 1
    ' 空白を読み飛ばす
    Do While p <= Len(json) And (Mid$(json, p, 1) = " " Or Mid$(json, p, 1) = vbTab _
                                 Or Mid$(json, p, 1) = vbCr Or Mid$(json, p, 1) = vbLf)
        p = p + 1
    Loop
    ' 区切り文字 (, } ] または空白) まで読む
    q = p
    Do While q <= Len(json)
        Select Case Mid$(json, q, 1)
            Case ",", "}", "]", " ", vbCr, vbLf, vbTab
                Exit Do
        End Select
        q = q + 1
    Loop
    ExtractRawValue = Trim$(Mid$(json, p, q - p))
End Function
