Attribute VB_Name = "Pages"
Sub GoPage()
Attribute GoPage.VB_ProcData.VB_Invoke_Func = "W\n14"

' Macro para seleccionar página.

    On Error Resume Next

    Sheets(ActiveCell.Value).Visible = True
    Sheets(ActiveCell.Value).Select

End Sub



Sub ShowPages()
Attribute ShowPages.VB_ProcData.VB_Invoke_Func = "Q\n14"

' Macro para mostrar listado de páginas

    On Error Resume Next

    'Me paro en la pagina del listado

    Sheets("Paginas").Select

    Dim j As Integer

    Dim NumSheets As Integer

    NumSheets = Sheets.Count

    For j = 1 To NumSheets

        Cells(j, 1) = Sheets(j).Name

    Next j

End Sub


Sub ShowAll()

' Macro para mostrar listado de páginas

    On Error Resume Next

    Dim j As Integer

    Dim NumSheets As Integer

    NumSheets = Sheets.Count

    For j = 1 To NumSheets
        Sheets(j).Visible = True
    Next j
End Sub

Sub HideAllNonOntologies()

' Macro para mostrar listado de páginas

    On Error Resume Next

    Dim j As Integer

    Dim NumSheets As Integer

    Sheets("Paginas").Select

    NumSheets = Sheets.Count

    For j = 1 To NumSheets
        If Sheets(j).Name <> "Paginas" And Left(Sheets(j).Name, 3) <> "ont" Then
            Sheets(j).Visible = False
        End If
    Next j
End Sub
