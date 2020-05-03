Attribute VB_Name = "Groupstage"
Sub main_groupstage()
    Application.ScreenUpdating = False
    Call set_globals
    Dim i As Integer
    Dim n As Integer
    Dim parts_r As Range
    Dim dynamic_screen As Range
    n = 0
    
    For i = tables_vStart To max_participants
        If IsEmpty(Cells(i, 2).Value) = False Then
            n = n + 1
        End If
    Next i
    
    n_par = n
    If n > max_participants Then
        MsgBox "More participants entered than the maximum of the max_participants value. Things will go to shit from here on out.."
    End If
    
    Set parts_r = Range(Cells(tables_vStart, 2), Cells(tables_vStart + n - 1, 2))
    ThisWorkbook.Names.Add Name:="Parts", RefersTo:=parts_r
    
    Set dynamic_screen = Sheets("Groupstage").Range(Cells(tables_vStart - 1, 6), Cells(tables_vStart + max_participants * 3, max_participants * 3))
    
    
    Call clear_area(dynamic_screen)
    
    
    If n <> 0 Then
        Call create_points_table(parts_r)
        Call create_matchups(parts_r)
        Call create_standings(parts_r)
    End If
    
    
    Application.ScreenUpdating = True
End Sub



