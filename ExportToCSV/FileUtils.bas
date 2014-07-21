Attribute VB_Name = "FileUtils"
Function BaseName(ByVal FileName As String, Optional ByVal WithExtension As Boolean = False)
    FileWithoutPath = Mid(FileName, InStrRev(FileName, "\", -1, vbTextCompare) + 1)
    If WithExtension Then
        BaseName = FileWithoutPath
    Else
        BaseName = WithoutExtension(FileWithoutPath)
    End If
End Function

Function WithoutExtension(ByVal FileName As String)
    WithoutExtension = Left(FileName, (InStrRev(FileName, ".", -1, vbTextCompare) - 1))
End Function

Sub DeleteFile(ByVal FileName As String)
   If ExistsFile(FileName) Then
      SetAttr FileName, vbNormal
      Kill FileName
   End If
End Sub

Function ExistsFile(ByVal FileName As String) As Boolean
   On Error Resume Next
   ExistsFile = ((GetAttr(FileName) And vbArchive) = vbArchive)
End Function

Function ExistsDir(ByVal Path As String) As Boolean
   On Error Resume Next
   ExistsDir = ((GetAttr(Path) And vbDirectory) = vbDirectory)
End Function

Sub MkDir(ByVal Path As String)
    If Not ExistsDir(Path) Then
        FileSystem.MkDir Path
    End If
End Sub

Sub MkDirHidden(ByVal Path As String)
    MkDir Path
    SetAttr Path, vbHidden
End Sub
