Attribute VB_Name = "MainStage"

Sub main_mainstage()
    Application.ScreenUpdating() = False
    Application.EnableEvents = False
    
    Call set_globals
    Dim standings As Range
    Dim parts As Range
    
    Set standings = Sheets("Groupstage").Range("Standings")
    Set parts = Sheets("Groupstage").Range("Parts")
    
    Dim dynamic_screen As Range
    Set dynamic_screen = Worksheets("Mainstage").Range(Cells(tables_vStart - 1, 6), Cells(tables_vStart + max_participants * 3, max_participants * 3))
    
    Call clear_area(dynamic_screen)
    Call create_upperbracket(standings)
    

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
