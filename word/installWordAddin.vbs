Option Explicit
Const FILLE_NAME="OhinaWorki.dotm"

Call CopyToAddinFolder

Sub CopyToAddinFolder()
    Dim objWord
    Dim strAddPath
    Dim strCurrentPath
    Dim strAddCopy
    Dim strCurrentCp
    Dim objFileSys
    Dim oAdd

    Set objWord   = CreateObject("Word.Application")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")

    strAddPath = objWord.StartupPath
    strCurrentPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    strAddCopy   = objFileSys.BuildPath(strAddPath, FILLE_NAME)
    strCurrentCp   = objFileSys.BuildPath(strCurrentPath, FILLE_NAME)

    objFileSys.CopyFile strCurrentCp, strAddCopy

    Set objWord   = Nothing
    Set objFileSys = Nothing

    MsgBox strAddPath
End Sub
