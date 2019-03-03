' 定数定義
'' 対象のアドインファイル
Const FILLE_NAME="OhinaWork.ppam"

' メイン処理
Call installAddin

' アドインをインストール
Sub installAddin()
    Dim objPowerPoint
    Dim objExcel
    Dim addinFolderPath
    Dim currentFolderPath
    Dim addinFilePath
    Dim currentFilePath
    Dim objFileSys
    Dim objAddin

    ' Office Object定義
    Set objExcel = CreateObject("Excel.Application")
    Set objPowerPoint = CreateObject("PowerPoint.Application")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")

    ' アドインフォルダー取得
    addinFolderPath = objExcel.Application.UserLibraryPath

    ' カレントフォルダー取得
    currentFolderPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

    ' アドインファイルパスを取得
    addinFilePath   = objFileSys.BuildPath(addinFolderPath, FILLE_NAME)

    ' カレントファイルパスを取得
    currentFilePath   = objFileSys.BuildPath(currentFolderPath, FILLE_NAME)

    ' 対象のアドインファイルをアドインフォルダに移動
    objFileSys.CopyFile currentFilePath, addinFilePath

    ' パワーポイントを非表示で起動
    objPowerPoint.Presentations.Add(False)

    ' アドインを追加
    Set objAddin = objPowerPoint.AddIns.Add(addinFilePath)

    ' 起動時に自動的に読む
    objAddin.AutoLoad = True

    ' パワーポイントを停止
    objPowerPoint.Quit
    Set objPowerPoint = Nothing
    Set objExcel = Nothing
    Set objFileSys = Nothing

    MsgBox "アドインを" & addinFilePath & " にインストールしました。"
End Sub
