Attribute VB_Name = "PPTGenerator"
'==============================================================================
' PPTGenerator.bas
' テンプレート PPTX のコピー → テキスト・画像置換 → 保存
'
' 呼び出し元: Phase2_3_Main.bas の Phase3_GeneratePPTs()
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' 1ケース分の PPT を生成する
'   templatePath  : テンプレート PPTX のフルパス
'   outputPath    : 出力 PPTX のフルパス（上書き OK）
'   textData      : Scripting.Dictionary  (シェイプ名 → テキスト)
'   imageData     : Scripting.Dictionary  (シェイプ名 → 画像フルパス)
'   pptApp        : 起動済み PowerPoint.Application オブジェクト
'------------------------------------------------------------------------------
Public Sub GeneratePPT( _
    templatePath  As String, _
    outputPath    As String, _
    textData      As Scripting.Dictionary, _
    imageData     As Scripting.Dictionary, _
    pptApp        As Object _
)
    ' ---- 出力フォルダを作成 ----
    Dim outDir As String
    outDir = Left(outputPath, InStrRev(outputPath, "\") - 1)
    EnsureDir outDir

    ' ---- テンプレートをコピー ----
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile templatePath, outputPath, True

    ' ---- PPTX を開く ----
    Dim prs As Object
    Set prs = pptApp.Presentations.Open(outputPath, msoFalse, msoFalse, msoFalse)

    Dim slide As Object
    Set slide = prs.Slides(1)

    ' ---- テキスト置換 ----
    Dim shpName As Variant
    For Each shpName In textData.Keys
        ReplaceText slide, CStr(shpName), textData(shpName)
    Next shpName

    ' ---- 画像置換 ----
    For Each shpName In imageData.Keys
        ReplacePictureHard slide, CStr(shpName), imageData(shpName)
    Next shpName

    ' ---- 保存 & 閉じる ----
    prs.Save
    prs.Close
End Sub

'------------------------------------------------------------------------------
' テキストボックスの文字列を置換する
'------------------------------------------------------------------------------
Public Sub ReplaceText(slide As Object, shapeName As String, newText As String)
    Dim shp As Object
    For Each shp In slide.Shapes
        If shp.Name = shapeName Then
            If shp.HasTextFrame Then
                ' フォント情報を保持したまま文字列だけ変更
                With shp.TextFrame.TextRange
                    .Text = newText
                End With
            End If
            Exit Sub
        End If
    Next shp
    ' シェイプが見つからなくてもエラーにしない（警告ログのみ）
    Debug.Print "警告: シェイプが見つかりません → " & shapeName
End Sub

'------------------------------------------------------------------------------
' 画像シェイプを新しい画像ファイルで置換する
' 元シェイプの位置・サイズを維持し、同名で再配置
'------------------------------------------------------------------------------
Public Sub ReplacePictureHard(slide As Object, shapeName As String, imgPath As String)
    ' ---- 画像ファイルの存在確認 ----
    If imgPath = "" Then
        Debug.Print "警告: 画像パスが空です → " & shapeName
        Exit Sub
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(imgPath) Then
        Debug.Print "警告: 画像ファイルが見つかりません → " & imgPath
        Exit Sub
    End If

    ' ---- 対象シェイプを検索 ----
    Dim shp As Object
    Dim targetShp As Object
    For Each shp In slide.Shapes
        If shp.Name = shapeName Then
            Set targetShp = shp
            Exit For
        End If
    Next shp

    If targetShp Is Nothing Then
        Debug.Print "警告: シェイプが見つかりません → " & shapeName
        Exit Sub
    End If

    ' ---- 位置・サイズを記録 ----
    Dim L As Single, T As Single, W As Single, H As Single
    L = targetShp.Left
    T = targetShp.Top
    W = targetShp.Width
    H = targetShp.Height

    ' ---- 元シェイプを削除して画像を挿入 ----
    targetShp.Delete

    Dim newShp As Object
    Set newShp = slide.Shapes.AddPicture( _
        FileName:=imgPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=L, Top:=T, Width:=W, Height:=H _
    )
    newShp.Name = shapeName
End Sub

'------------------------------------------------------------------------------
' フォルダが存在しない場合は作成する（再帰的）
'------------------------------------------------------------------------------
Private Sub EnsureDir(dirPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dirPath) Then
        fso.CreateFolder dirPath
    End If
End Sub
