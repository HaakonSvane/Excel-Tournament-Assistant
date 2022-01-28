Attribute VB_Name = "Groupstage"

Sub main_groupstage()
    Application.ScreenUpdating = False
    Call delete_all_names
    Call set_globals
    Dim parts_r As Range
    Dim dynamic_screen As Range

    Set parts_r = Range(Cells(tables_vStart, 2), Cells(tables_vStart + n_par - 1, 2))
    ThisWorkbook.names.Add name:="Parts", RefersTo:=parts_r
    
    Set dynamic_screen = Sheets("Groupstage").Range(Cells(tables_vStart - 1, 6), Cells(tables_vStart + max_participants * 5, max_participants * 5))
    Call clear_area(dynamic_screen)
   
    If n_par <> 0 Then
        Call create_points_table(parts_r)
        Call create_matchups(parts_r)
        Call create_standings(parts_r)
    End If
    
    
    Application.ScreenUpdating = True
End Sub



