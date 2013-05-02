Attribute VB_Name = "Synchronization"
Sub Synchronize()
    Application.DisplayStatusBar = True
    
    Set CurrentWorkbook = ActiveWorkbook
    Set CurrentSheet = ActiveWorkbook.ActiveSheet
    
    RemoveInvalidFiles CurrentWorkbook
    SaveMissingFiles CurrentWorkbook
    SaveChangedSheets
    Tracking.Clear
    
    CurrentWorkbook.Activate
    CurrentSheet.Select
    
    Application.DisplayStatusBar = False
End Sub

Private Sub RemoveInvalidFiles(ByVal Workbook As Workbook, Optional ByVal Force As Boolean = False)
    Set Files = TargetFiles(Workbook)
    I = 1
    For Each File In Files
        ShowProgress "Removing invalid files", I / Files.Count, FileUtils.BaseName(File, True)
        If Not ExistsSheet(FileUtils.BaseName(File), Workbook) Or Force Then
            FileUtils.DeleteFile File
        End If
        I = I + 1
    Next
    If TargetFiles(Workbook).Count = 0 Then
        FileSystem.RmDir TargetPath(Workbook)
    End If
End Sub

Private Sub SaveMissingFiles(ByVal Workbook As Workbook)
    Set AllSheets = Workbook.Sheets
    For I = 1 To AllSheets.Count
        FileName = SheetFileName(AllSheets(I))
        ShowProgress "Saving missing files", I / AllSheets.Count, FileUtils.BaseName(FileName, True)
        If Not FileUtils.ExistsFile(FileName) Then
            SaveAsCSV AllSheets(I)
        End If
    Next
End Sub

Private Sub SaveChangedSheets()
    Set AllSheets = Tracking.Sheets
    For I = 1 To AllSheets.Count
        ShowProgress "Saving changes", I / AllSheets.Count, AllSheets(I).Name
        SaveAsCSV AllSheets(I)
    Next
End Sub

Sub ExportAllSheets()
    Application.DisplayStatusBar = True
    
    Set CurrentWorkbook = ActiveWorkbook
    Set CurrentSheet = ActiveWorkbook.ActiveSheet
    
    SheetsCount = CurrentWorkbook.Worksheets.Count
    For I = 1 To SheetsCount
        ShowProgress
        Application.StatusBar = "Progress: " & Format(I / SheetsCount, "0%") & " | Exporting sheet: " & CurrentWorkbook.Sheets(I).Name
        SaveAsCSV CurrentWorkbook.Sheets(I)
    Next
    
    CurrentWorkbook.Activate
    CurrentSheet.Select
    
    Application.DisplayStatusBar = False
End Sub

Private Sub SaveAsCSV(ByVal Sheet As Worksheet)
    Application.DisplayAlerts = False
    Sheet.Select
    Sheet.Copy
    ActiveWorkbook.SaveAs FileName:=SheetFileName(Sheet), _
                            FileFormat:=xlCSV, _
                            CreateBackup:=False, _
                            ConflictResolution:=xlLocalSessionChanges
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
End Sub

Private Function TargetPath(ByVal Workbook As Workbook)
    TargetPath = Workbook.Path & "\" & "." & FileUtils.WithoutExtension(Workbook.Name)
    FileUtils.MkDir (TargetPath)
End Function

Private Function SheetFileName(ByVal Sheet As Worksheet)
    SheetFileName = TargetPath(Sheet.Parent) & "\" & Sheet.Name & ".csv"
End Function

Private Function TargetFiles(ByVal Workbook As Workbook)
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set Folder = FS.GetFolder(TargetPath(Workbook))
    Set TargetFiles = Folder.Files
End Function

Private Function FindSheet(ByVal Name As String, ByVal Workbook As Workbook)
    Set FindSheet = Nothing
    For Each Sheet In Workbook.Sheets
        If Sheet.Name = Name Then
            Set FindSheet = Sheet
            Exit Function
        End If
    Next
End Function

Private Function ExistsSheet(ByVal Name As String, ByVal Workbook As Workbook)
    ExistsSheet = Not FindSheet(Name, Workbook) Is Nothing
End Function

Private Sub ShowProgress(ByVal Title As String, ByVal Percentage As Variant, ByVal Description As String)
    If Percentage > 100 Then
        Percentage = 100
    End If
    Application.StatusBar = Title & " | Progress: " & Format(Percentage, "0%") & " | " & Description
End Sub
