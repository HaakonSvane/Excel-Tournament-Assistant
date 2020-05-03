Attribute VB_Name = "GlobalFunctions"
Public COLOR_BACKGROUND As Long
Public COLOR_FOREGROUND_1 As Long
Public COLOR_FOREGROUND_2 As Long
Public COLOR_HEADER As Long
Public COLOR_PASS As Long
Public COLOR_FAIL As Long
Public COLOR_ERROR As Long

Public tables_vStart As Integer
Public tables_hStart As Integer
Public max_participants As Integer
Public group_first_to As Integer
Public tiebreaker_first_to As Integer
Public finals_first_to As Integer
Public winner_advantage As Boolean

Public Sub set_globals()
    Dim active_sheet As Worksheet
    Set active_sheet = ThisWorkbook.ActiveSheet
    
    tables_vStart = 3
    tables_hStart = 7
    
    Dim color_field As Range
    Dim bool_field As Range
    Dim value_field As Range

    With Sheets("Preferences")
        .Activate
        Set color_field = .Range(Cells(3, 4), Cells(13, 9))
        Set bool_field = .Range(Cells(3, 11), Cells(13, 17))
        Set value_field = .Range(Cells(3, 18), Cells(13, 22))
    End With
    
    ThisWorkbook.Names.Add Name:="ColorOptions", RefersTo:=color_field
    ThisWorkbook.Names.Add Name:="BoolOptions", RefersTo:=bool_field
    ThisWorkbook.Names.Add Name:="ValueOptions", RefersTo:=value_field
    
    ' Fetching colors
    COLOR_FOREGROUND_1 = color_field.Cells(1, 5).Interior.Color
    COLOR_FOREGROUND_2 = color_field.Cells(2, 5).Interior.Color
    COLOR_BACKGROUND = color_field.Cells(3, 5).Interior.Color
    COLOR_HEADER = color_field.Cells(4, 5).Interior.Color
    COLOR_PASS = color_field.Cells(5, 5).Interior.Color
    COLOR_FAIL = color_field.Cells(6, 5).Interior.Color
    COLOR_ERROR = color_field.Cells(7, 5).Interior.Color
    
    ' Fetching bools
    If bool_field.Cells(1, 6).Value = 1 Then
        winner_advantage = True
    Else
        winner_advangtage = False
    End If
    
    ' Fetching values
    
    group_first_to = (value_field.Cells(1, 5).Value + 1) / 2
    tiebreaker_first_to = (value_field.Cells(2, 5).Value + 1) / 2
    final_first_to = (value_field.Cells(3, 5).Value + 1) / 2
    max_participants = value_field.Cells(4, 5).Value
    
    active_sheet.Activate
End Sub

