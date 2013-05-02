VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If Success Then
        Synchronization.Synchronize
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sheet As Object, ByVal Target As Range)
    Tracking.Change Sheet
End Sub
