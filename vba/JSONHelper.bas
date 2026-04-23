Attribute VB_Name = "JSONHelper"
'==============================================================================
' JSONHelper.bas
' JSON ファイル読込ユーティリティ
' 外部ライブラリ不要 — Scripting.Runtime (Dictionary) のみ使用
'
' 対応フォーマット:
'   LoadCasesJSON      : cases.json (Python が出力する配列形式)
'   LoadImageMapping   : image_mapping.json (ユーザー作成のオブジェクト形式)
'==============================================================================
Option Explicit

' --- module-level parser state ---
Private mText As String
Private mPos  As Long

'------------------------------------------------------------------------------
' Public: cases.json を読込む
' 戻り値: Collection of Scripting.Dictionary
'   各 Dict のキー: case_id, config_file, conditions,
'                   max_height, dispersion_height, gas_holdup
'------------------------------------------------------------------------------
Public Function LoadCasesJSON(filePath As String) As Collection
    If Not FileExists(filePath) Then
        Err.Raise vbObjectError + 1, "JSONHelper", "cases.json が見つかりません: " & filePath
    End If

    mText = ReadFileUTF8(filePath)
    mPos  = 1
    SkipWS

    Dim result As New Collection
    ' expect top-level array [
    AssertChar "["
    mPos = mPos + 1
    SkipWS

    Do While Peek() <> "]"
        SkipWS
        Dim caseDict As Scripting.Dictionary
        Set caseDict = ParseObject()
        result.Add caseDict
        SkipWS
        If Peek() = "," Then mPos = mPos + 1
        SkipWS
    Loop

    Set LoadCasesJSON = result
End Function

'------------------------------------------------------------------------------
' Public: image_mapping.json を読込む
' 戻り値: Scripting.Dictionary
'   Key   : case key (例 "case1a")
'   Value : Scripting.Dictionary  (IMG_Main, IMG_Graph1, IMG_Graph2)
'------------------------------------------------------------------------------
Public Function LoadImageMapping(filePath As String) As Scripting.Dictionary
    If Not FileExists(filePath) Then
        Err.Raise vbObjectError + 1, "JSONHelper", _
            "image_mapping.json が見つかりません: " & filePath
    End If

    mText = ReadFileUTF8(filePath)
    mPos  = 1
    SkipWS

    ' outer object → Scripting.Dictionary of Scripting.Dictionary
    Dim result As Scripting.Dictionary
    Set result = ParseObjectOfObjects()
    Set LoadImageMapping = result
End Function

'==============================================================================
' Private helpers
'==============================================================================

' Parse { "k": "v", "k2": 123, ... } → Scripting.Dictionary (scalar values)
Private Function ParseObject() As Scripting.Dictionary
    Dim d As New Scripting.Dictionary
    AssertChar "{"
    mPos = mPos + 1
    SkipWS

    Do While Peek() <> "}"
        SkipWS
        Dim key As String
        key = ParseString()
        SkipWS
        AssertChar ":"
        mPos = mPos + 1
        SkipWS
        Dim val As Variant
        val = ParseScalar()
        d(key) = val
        SkipWS
        If Peek() = "," Then mPos = mPos + 1
        SkipWS
    Loop
    mPos = mPos + 1  ' skip }

    Set ParseObject = d
End Function

' Parse { "case1a": { "IMG_Main": "...", ... }, ... }
' → outer Dictionary whose values are inner Dictionaries
Private Function ParseObjectOfObjects() As Scripting.Dictionary
    Dim outer As New Scripting.Dictionary
    AssertChar "{"
    mPos = mPos + 1
    SkipWS

    Do While Peek() <> "}"
        SkipWS
        Dim outerKey As String
        outerKey = ParseString()
        SkipWS
        AssertChar ":"
        mPos = mPos + 1
        SkipWS
        Dim inner As Scripting.Dictionary
        Set inner = ParseObject()
        Set outer(outerKey) = inner
        SkipWS
        If Peek() = "," Then mPos = mPos + 1
        SkipWS
    Loop
    mPos = mPos + 1  ' skip }

    Set ParseObjectOfObjects = outer
End Function

' Parse a scalar value (string or number; null → "")
Private Function ParseScalar() As Variant
    Dim c As String
    c = Peek()
    If c = Chr(34) Then
        ParseScalar = ParseString()
    ElseIf c = "t" Then           ' true
        ParseScalar = True
        mPos = mPos + 4
    ElseIf c = "f" Then           ' false
        ParseScalar = False
        mPos = mPos + 5
    ElseIf c = "n" Then           ' null
        ParseScalar = ""
        mPos = mPos + 4
    Else
        ParseScalar = ParseNumber()
    End If
End Function

' Parse a JSON string (handles \n \r \t \" \\ \/)
Private Function ParseString() As String
    AssertChar Chr(34)
    mPos = mPos + 1               ' skip opening "
    Dim result As String
    result = ""

    Do While mPos <= Len(mText)
        Dim c As String
        c = Mid(mText, mPos, 1)

        If c = "\" Then
            mPos = mPos + 1
            Dim esc As String
            esc = Mid(mText, mPos, 1)
            Select Case esc
                Case "n":     result = result & Chr(10)
                Case "r":     result = result & Chr(13)
                Case "t":     result = result & Chr(9)
                Case Chr(34): result = result & Chr(34)
                Case "\":     result = result & "\"
                Case "/":     result = result & "/"
                Case "u"
                    ' \uXXXX
                    Dim hex4 As String
                    hex4 = Mid(mText, mPos + 1, 4)
                    result = result & ChrW(CLng("&H" & hex4))
                    mPos = mPos + 4
                Case Else:    result = result & esc
            End Select
            mPos = mPos + 1

        ElseIf c = Chr(34) Then   ' closing "
            mPos = mPos + 1
            ParseString = result
            Exit Function

        Else
            result = result & c
            mPos = mPos + 1
        End If
    Loop

    Err.Raise vbObjectError + 2, "JSONHelper", "JSON文字列が閉じられていません"
End Function

' Parse a JSON number → Double
Private Function ParseNumber() As Double
    Dim start As Long
    start = mPos
    Do While mPos <= Len(mText)
        Dim c As String
        c = Mid(mText, mPos, 1)
        If InStr("0123456789.+-eE", c) > 0 Then
            mPos = mPos + 1
        Else
            Exit Do
        End If
    Loop
    ParseNumber = CDbl(Mid(mText, start, mPos - start))
End Function

' Skip whitespace (space, tab, CR, LF)
Private Sub SkipWS()
    Do While mPos <= Len(mText)
        Dim c As String
        c = Mid(mText, mPos, 1)
        If c = " " Or c = Chr(9) Or c = Chr(10) Or c = Chr(13) Then
            mPos = mPos + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function Peek() As String
    If mPos <= Len(mText) Then
        Peek = Mid(mText, mPos, 1)
    Else
        Peek = ""
    End If
End Function

Private Sub AssertChar(expected As String)
    If Peek() <> expected Then
        Err.Raise vbObjectError + 3, "JSONHelper", _
            "JSONパースエラー: 位置 " & mPos & " で '" & expected & "' を期待, 実際: '" & Peek() & "'"
    End If
End Sub

' Read entire file as UTF-8 string using FileSystemObject
Private Function ReadFileUTF8(filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    ' -2 = system default encoding; for UTF-8 without BOM: use ADODB.Stream if needed
    Set ts = fso.OpenTextFile(filePath, 1, False, -2)
    ReadFileUTF8 = ts.ReadAll()
    ts.Close
End Function

' Check file existence without raising an error
Private Function FileExists(path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
End Function
