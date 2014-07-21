Attribute VB_Name = "Tracking"
Private ChangedSheets As Object
Private OriginalFileName As String

Sub SetFileName(ByVal file_name As String)
    OriginalFileName = file_name
End Sub

Function GetFileName()
    GetFileName = OriginalFileName
End Function

Sub Change(ByVal Sheet As Worksheet)
    If ChangedSheets Is Nothing Then
        Set ChangedSheets = CreateObject("Scripting.Dictionary")
    End If
    Set ChangedSheets(Sheet) = Sheet
End Sub

Sub Clear()
    Set ChangedSheets = Nothing
End Sub

Function AllChangedSheets()
    Set AllChangedSheets = New Collection
    If Not ChangedSheets Is Nothing Then
        For Each Sheet In ChangedSheets
            On Error Resume Next
            AllChangedSheets.Add Sheet
        Next
    End If
End Function

Sub ShowChangedSheets()
    Dim SheetNames As String
    For Each Sheet In Sheets
        If SheetNames = "" Then
            SheetNames = Sheet.Name
        Else
            SheetNames = SheetNames & ", " & Sheet.Name
        End If
    Next
    If SheetNames = "" Then
        MsgBox "There are no changes"
    Else
        MsgBox "There are changes on: " & SheetNames
    End If
End Sub

