Attribute VB_Name = "Phase2_3_Main"
'==============================================================================
' Phase2_3_Main.bas
' Phase 2 & 3 メイン処理
'
' ■ 前提
'   - Phase 1 (Python) 完了済み: data/work/cases.json, data/work/database.db
'   - image_mapping.json をユーザーが配置済み: data/work/image_mapping.json
'   - テンプレート PPTX が存在する: data/input/template.pptx
'
' ■ 実行方法
'   VBA エディタ(Alt+F11) → Run_Phase2_3 を実行
'   または Excel/PowerPoint のマクロダイアログから実行
'
' ■ 必要な参照設定 (Tools → References)
'   - Microsoft Scripting Runtime
'   - Microsoft PowerPoint XX.X Object Library
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' 設定値 — 環境に合わせて変更すること
'------------------------------------------------------------------------------
Private Const DB_PATH            As String = "data\work\database.db"
Private Const CASES_JSON_PATH    As String = "data\work\cases.json"
Private Const IMAGE_MAPPING_PATH As String = "data\work\image_mapping.json"
Private Const TEMPLATE_PPT_PATH  As String = "data\input\template.pptx"
Private Const OUTPUT_DIR         As String = "data\output"

' 処理の基準ディレクトリ（このスクリプトと同じ場所が起点）
' Excel から実行する場合は ThisWorkbook.Path、
' 独立した .bas 実行の場合は適宜変更してください。
Private Function BaseDir() As String
    On Error Resume Next
    BaseDir = ThisWorkbook.Path
    If BaseDir = "" Then BaseDir = CurDir()
    On Error GoTo 0
End Function

Private Function AbsPath(rel As String) As String
    AbsPath = BaseDir() & "\" & rel
End Function

'==============================================================================
' エントリポイント
'==============================================================================
Public Sub Run_Phase2_3()
    On Error GoTo ErrHandler

    Dim casesPath    As String
    Dim mappingPath  As String
    Dim templatePath As String
    Dim outputDir    As String

    casesPath    = AbsPath(CASES_JSON_PATH)
    mappingPath  = AbsPath(IMAGE_MAPPING_PATH)
    templatePath = AbsPath(TEMPLATE_PPT_PATH)
    outputDir    = AbsPath(OUTPUT_DIR)

    ' ---- Phase 2: データセット構築 ----
    Dim caseList As Collection
    Set caseList = Phase2_BuildDataset(casesPath, mappingPath, templatePath)

    If caseList.Count = 0 Then
        MsgBox "処理するケースが見つかりませんでした。" & vbLf & _
               "cases.json を確認してください: " & casesPath, vbExclamation
        Exit Sub
    End If

    ' ---- Phase 3: PPT 複数生成 ----
    Phase3_GeneratePPTs caseList, templatePath, outputDir

    MsgBox "✓ 全 " & caseList.Count & " 件のPPTを生成しました。" & vbLf & _
           "出力フォルダ: " & outputDir, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました:" & vbLf & Err.Description, vbCritical
End Sub

'==============================================================================
' Phase 2: cases.json + image_mapping.json → ケースデータの Collection 構築
'==============================================================================
Private Function Phase2_BuildDataset( _
    casesJsonPath    As String, _
    imageMappingPath As String, _
    templatePath     As String _
) As Collection

    ' ---- 1. cases.json 読込 ----
    Dim cases As Collection
    Set cases = JSONHelper.LoadCasesJSON(casesJsonPath)

    ' ---- 2. image_mapping.json 読込 ----
    Dim imgMap As Scripting.Dictionary
    Set imgMap = JSONHelper.LoadImageMapping(imageMappingPath)

    ' ---- 3. テンプレート検証 ----
    ValidateTemplate templatePath

    ' ---- 4. 各ケースにテキスト＆画像パスを付加 ----
    Dim result As New Collection
    Dim i As Long
    For i = 1 To cases.Count
        Dim caseData As Scripting.Dictionary
        Set caseData = cases(i)

        Dim caseId As String
        caseId = CStr(caseData("case_id"))

        ' テキスト置換データを設定
        Dim textData As New Scripting.Dictionary
        textData("TXT_ConfigFile") = "設定ファイル名：" & caseData("config_file")
        textData("TXT_Conditions") = CStr(caseData("conditions"))
        textData("TXT_Dispersion") = "時刻20 sec における90%分散高さ：" & _
                                     Format(CDbl(caseData("dispersion_height")), "0.000") & " m"
        textData("TXT_GasHU")      = "時刻20 sec におけるガスホールドアップ：" & _
                                     Format(CDbl(caseData("gas_holdup")), "0.000") & " %"

        ' 画像置換データを設定
        Dim imageData As New Scripting.Dictionary
        Dim mappingKey As String
        mappingKey = "case" & caseId   ' 例: "case1a"

        If imgMap.Exists(mappingKey) Then
            Dim imgs As Scripting.Dictionary
            Set imgs = imgMap(mappingKey)
            Dim imgKey As Variant
            For Each imgKey In imgs.Keys
                imageData(CStr(imgKey)) = CStr(imgs(imgKey))
            Next imgKey
        Else
            Debug.Print "警告: image_mapping.json にキーがありません → " & mappingKey
        End If

        ' ケースエントリをまとめる
        Dim entry As New Scripting.Dictionary
        entry("case_id")        = caseId
        entry("output_filename") = "output_case" & caseId & ".pptx"
        Set entry("text_data")  = textData
        Set entry("image_data") = imageData

        result.Add entry
    Next i

    Set Phase2_BuildDataset = result
End Function

'==============================================================================
' Phase 3: PPT 複数生成
'==============================================================================
Private Sub Phase3_GeneratePPTs( _
    caseList     As Collection, _
    templatePath As String, _
    outputDir    As String _
)
    ' PowerPoint アプリを起動（非表示で実行）
    Dim pptApp As Object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = msoFalse

    Dim errorList As String
    errorList = ""

    Dim i As Long
    For i = 1 To caseList.Count
        Dim entry As Scripting.Dictionary
        Set entry = caseList(i)

        Dim caseId       As String
        Dim outputFile   As String
        caseId     = CStr(entry("case_id"))
        outputFile = outputDir & "\" & entry("output_filename")

        On Error GoTo LoopError
        PPTGenerator.GeneratePPT _
            templatePath:=templatePath, _
            outputPath:=outputFile, _
            textData:=entry("text_data"), _
            imageData:=entry("image_data"), _
            pptApp:=pptApp
        On Error GoTo 0

        Debug.Print "生成完了 [" & i & "/" & caseList.Count & "]: " & entry("output_filename")
        GoTo LoopNext

LoopError:
        errorList = errorList & "  - case" & caseId & ": " & Err.Description & vbLf
        On Error GoTo 0

LoopNext:
    Next i

    pptApp.Quit
    Set pptApp = Nothing

    If errorList <> "" Then
        MsgBox "以下のケースでエラーが発生しました:" & vbLf & errorList, vbExclamation
    End If
End Sub

'==============================================================================
' テンプレート PPTX の必須シェイプ確認
'==============================================================================
Private Sub ValidateTemplate(templatePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(templatePath) Then
        Err.Raise vbObjectError + 10, "Phase2_3_Main", _
            "テンプレートPPTXが見つかりません: " & templatePath
    End If

    Dim requiredShapes(6) As String
    requiredShapes(0) = "TXT_ConfigFile"
    requiredShapes(1) = "TXT_Conditions"
    requiredShapes(2) = "TXT_Dispersion"
    requiredShapes(3) = "TXT_GasHU"
    requiredShapes(4) = "IMG_Main"
    requiredShapes(5) = "IMG_Graph1"
    requiredShapes(6) = "IMG_Graph2"

    ' PowerPoint を一時起動して確認
    Dim pptApp As Object
    Set pptApp = CreateObject("PowerPoint.Application")
    Dim prs As Object
    Set prs = pptApp.Presentations.Open(templatePath, msoTrue, msoFalse, msoFalse)

    Dim slide As Object
    Set slide = prs.Slides(1)

    Dim foundNames As New Scripting.Dictionary
    Dim shp As Object
    For Each shp In slide.Shapes
        foundNames(shp.Name) = True
    Next shp

    Dim missing As String
    missing = ""
    Dim j As Integer
    For j = 0 To UBound(requiredShapes)
        If Not foundNames.Exists(requiredShapes(j)) Then
            missing = missing & "  - " & requiredShapes(j) & vbLf
        End If
    Next j

    prs.Close False
    pptApp.Quit

    If missing <> "" Then
        Err.Raise vbObjectError + 11, "Phase2_3_Main", _
            "テンプレートに必須シェイプが存在しません:" & vbLf & missing
    End If
End Sub
