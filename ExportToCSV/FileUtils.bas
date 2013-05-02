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
   ExistsFile = Exists(FileName, vbArchive)
End Function

Function ExistsDir(ByVal Path As String) As Boolean
   ExistsDir = Exists(Path, vbDirectory)
End Function

Private Function Exists(ByVal Resource As String, ByVal Attributes As VbFileAttribute)
    Exists = (Dir(Resource, Attributes) <> "")
End Function

Sub MkDir(ByVal Path As String)
    If Not ExistsDir(Path) Then
        FileSystem.MkDir Path
    End If
End Sub
