Attribute VB_Name = "Utilities"
Option Explicit
Public dict As New Collection
Public n_par As Integer

Sub clear_area(field As Range)
    With field
        .Cells.ClearContents
        .Borders.LineStyle = xlNone
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Interior.Color = COLOR_BACKGROUND
        .Cells.UnMerge
        .Font.ColorIndex = 1
        .FormatConditions.Delete
    End With
End Sub
Sub outer_border(field As Range)
    With field
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
    End With
End Sub
Sub outer_border_small(field As Range)
    With field
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
End Sub
Sub cross_field(field As Range)
    field.Borders(xlDiagonalDown).LineStyle = xlContinuous
    field.Borders(xlDiagonalUp).LineStyle = xlContinuous
End Sub
Sub set_color_condition_matches(field As Range, first_to)

    Dim cond_green As FormatCondition, cond_first As FormatCondition, cond_last As FormatCondition, cond_fail As FormatCondition
    
    Set cond_green = field.FormatConditions.Add(xlCellValue, xlEqual, first_to)
    cond_green.Priority = 2
    
    Set cond_first = field(1).FormatConditions.Add(Type:=xlExpression, Formula1:="=" & field(2).Address & "=" & first_to)
    cond_first.Priority = 2
    Set cond_last = field(2).FormatConditions.Add(Type:=xlExpression, Formula1:="=" & field(1).Address & "=" & first_to)
    cond_last.Priority = 2

    Set cond_fail = field.FormatConditions.Add(Type:=xlExpression, Formula1:="=OR(SUM(" & field.Address & ") > " & 2 * first_to - 1 & ";OR(" & field(1).Address & "< 0;" & field(2).Address & " < 0))")
    cond_fail.Priority = 1
    
    With cond_green
        .Interior.Color = COLOR_PASS
    End With
    
    With cond_first
        .Interior.Color = COLOR_FAIL
    End With
    
    With cond_last
        .Interior.Color = COLOR_FAIL
    End With
    
    With cond_fail
        .Interior.Color = COLOR_ERROR
        .Font.Color = vbWhite
    End With

End Sub
Sub set_color_condition_played(field As Range)
    Dim cond_green As FormatCondition
    Set cond_green = field.FormatConditions.Add(xlCellValue, xlEqual, field.Rows.Count - 1)
    With cond_green
        .Interior.Color = COLOR_PASS
    End With
End Sub
Sub color_diagonal(field As Range)
    Dim c As Long
    c = RGB(100, 100, 100)
    
    With field
        Dim i As Integer
        For i = 1 To .Rows.Count
            Range(.Cells(i, 2 * i - 1), .Cells(i, 2 * i)).Merge
            .Cells(i, 2 * i - 1).Interior.Color = c
        Next i
    End With
End Sub
Sub inside_lines(field As Range)
    With field
        Dim i As Integer
        For i = 1 To .Rows.Count
            'Horisontale linjer
            Range(.Cells(i, 1), Cells(i, .Columns.Count + 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Next i
        For i = 1 To Int(.Columns.Count / 2)
            'Vertikale linjer'
            Range(.Cells(1, 2 * (i - 1) + 2), .Cells(.Rows.Count, i)).Borders(xlEdgeRight).LineStyle = xlContinuous
        Next i
    End With
End Sub
Sub create_header(field As Range, text As String)
    field.Merge
    field.Interior.Color = COLOR_HEADER
    field.Value = text
    field.NumberFormat = "@"
    field.HorizontalAlignment = xlCenter
    field.VerticalAlignment = xlCenter
    field.Font.Size = 22
    field.Font.Bold = True
    Call outer_border(field)
End Sub
Sub create_container(x As Integer, y As Integer, w As Integer, h As Integer, Optional inner_header As String = "[NONE]")
    Dim container As Range
    Set container = Range(Cells(y, x), Cells(y + h - 1, x + w - 1))
    
    container.Interior.Color = COLOR_FOREGROUND_1
    
    If Not inner_header = "[NONE]" Then
        With container.Range(Cells(1, 1), Cells(1, 2))
            .Merge
            .Value = inner_header
            .Font.Bold = True
            .Font.Size = 20
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If
    Call outer_border(container)
    
End Sub
Function create_match(p1 As String, p2 As String, x As Integer, y As Integer, first_to) As Range
    Dim field As Range
    Set field = Range(Cells(y, x), Cells(y + 1, x + 2))
       
    With field
        .Cells.Interior.ColorIndex = 0
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
    
        Call outer_border_small(field)
        .Range(Cells(1, 1), Cells(1, 2)).Merge
        .Cells(1, 1).Value = p1
        
        .Range(Cells(2, 1), Cells(2, 2)).Merge
        .Cells(2, 1).Value = p2
        
        If p1 = "[NONE]" Or p2 = "[NONE]" Then
            Call cross_field(.Range(Cells(1, 3), Cells(2, 3)))
        Else
            Call set_color_condition_matches(.Range(Cells(1, 3), Cells(2, 3)), first_to)
        End If
        
    End With
    
    Set create_match = field.Range(Cells(1, 3), Cells(2, 3))
End Function
Sub create_matchups(parts_range As Range)
    Dim n As Integer
    Dim n_real As Integer
    Dim round_nr As Integer
    Dim match_nr As Integer
    
    Dim lookups_h As Range
    Dim lookups_v As Range
    Dim match_address As Range
    Dim table As Range
    
    Dim parts As New Collection

    n = parts_range.Rows.Count
    n_real = parts_range.Rows.Count
    
    Set lookups_h = Range(Cells(tables_vStart, 9), Cells(tables_vStart, 9 + (n_real - 1) * 2))
    Set lookups_v = Range(Cells(tables_vStart + 1, 7), Cells(tables_vStart + n_real, 7))
    Set table = Range(Cells(tables_vStart + 1, 7 + 2), Cells(tables_vStart + n_real, 7 + n_real * 2))

    Dim i As Integer
    For i = 1 To n
        parts.Add parts_range.Cells(i, 1).Value
    Next i
    If n Mod 2 <> 0 Then
        parts.Add "[NONE]"
        n = n + 1
    End If
    
    Dim extra_pad As Integer
    extra_pad = 0
    If n = 3 Then extra_pad = 2
    
    Dim fake_parts As Integer
    If n Mod 2 = 0 Then
        fake_parts = n
    Else
        fake_parts = n + 1
    End If
    
    Dim x0 As Integer
    Dim y0 As Integer
    x0 = 6 + 2 * n + 4 + extra_pad
    y0 = tables_vStart
    
    Dim w As Integer
    w = fake_parts / 2 * 4
    
    Dim header_field As Range
    Set header_field = Range(Cells(y0 - 1, x0), Cells(y0 - 1, x0 + 3))
    Call create_header(header_field, "Matchups")
    
    For round_nr = 1 To (n - 1)
        Call create_container(x0, y0 + (round_nr - 1) * 4, w, 3, "Round " & CStr(round_nr))
        For match_nr = 1 To n / 2
            Set match_address = create_match(CStr(parts(match_nr)), CStr(parts(n - (match_nr - 1))), x0 + 4 * (match_nr - 1) + 1, y0 + 4 * (round_nr - 1) + 1, group_first_to)
            If CStr(parts(match_nr)) <> "[NONE]" And CStr(parts(n - (match_nr - 1))) <> "[NONE]" Then
                Dim c As Integer
                For c = 1 To Int(lookups_h.Columns.Count / 2 + 1)
                    Dim r As Integer
                    For r = 1 To lookups_v.Rows.Count
                    'MsgBox lookups_h(2 * c - 1).Value & " = " & parts(match_nr) & vbCrLf & lookups_v(r).Value & " = " & parts(n - (match_nr - 1))'
                        If lookups_h(2 * c - 1).Value = parts(match_nr) And lookups_v(r).Value = parts(n - (match_nr - 1)) Then
                            'MsgBox r & ", " & 2 * c - 1'
                            table.Cells(r, 2 * c - 1).formula = "=" & match_address.Cells(2, 1).Address
                            table.Cells(r, 2 * c).formula = "=" & match_address.Cells(1, 1).Address
                            
                            table.Cells(c, 2 * r).formula = "=" & table.Cells(r, 2 * c - 1).Address
                            table.Cells(c, 2 * r - 1).formula = "=" & table.Cells(r, 2 * c).Address
                        End If
                    Next r
                Next c
            End If

        Next match_nr
        parts.Add parts(n), After:=1
        parts.Remove (n + 1)
        
    Next round_nr
    
End Sub

Sub create_points_table(parts As Range)
    Dim table_field As Range
    Dim table_inside As Range
    Dim header_field As Range
    Set table_field = Range(Cells(tables_vStart, 7), Cells(tables_vStart + parts.Rows.Count, 6 + (parts.Rows.Count + 1) * 2))
    Set table_inside = Range(Cells(tables_vStart + 1, 7 + 2), Cells(tables_vStart + parts.Rows.Count, 6 + (parts.Rows.Count + 1) * 2))
    Set header_field = Range(Cells(tables_vStart - 1, 7), Cells(tables_vStart - 1, 10))
    
    ThisWorkbook.Names.Add Name:="Points", RefersTo:=table_field
    
    With table_field
    .Interior.ColorIndex = 0
        Dim i As Integer
        For i = 1 To parts.Rows.Count + 1
            'For kolonne'
            .Range(Cells(i, 1), Cells(i, 2)).Merge
            .Cells(i + 1, 1).Value = parts.Cells(i, 1).Value
            
            .Cells(i, 1).Interior.Color = COLOR_FOREGROUND_1
            .Cells(i, 1).NumberFormat = "@"
            .Cells(i, 1).Font.Size = 20
            .Cells(i, 1).Font.Bold = True
            .Cells(i, 1).HorizontalAlignment = xlCenter
            
            'For rad'
            .Range(Cells(1, 2 * i - 1), Cells(1, 2 * i)).Merge
            .Cells(1, 2 * i - 1 + 2).Value = parts.Cells(i, 1).Value
            .Cells(1, 2 * i - 1).Interior.Color = COLOR_FOREGROUND_1
            .Cells(1, 2 * i - 1).NumberFormat = "@"
            .Cells(1, 2 * i - 1).Font.Size = 20
            .Cells(1, 2 * i - 1).Font.Bold = True
            .Cells(1, 2 * i - 1).HorizontalAlignment = xlCenter
            
        Next i
    End With
    
    
    With table_inside
        .NumberFormat = "General"
        .HorizontalAlignment = xlCenter
        .Font.Size = 20
        .Font.Bold = True
    End With

    Call outer_border(table_field)
    Call inside_lines(table_field)
    Call color_diagonal(table_field)
    Call create_header(header_field, "PointsTable:")

End Sub
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub
Sub create_standings(parts As Range)
    Dim field As Range
    Dim header_field As Range
    Dim table As Range
    Dim table_over As Range
    Dim sum_field As Range
    
    Set field = Range(Cells(tables_vStart + parts.Rows.Count + 3, 7), Cells(tables_vStart + parts.Rows.Count + 3 + parts.Rows.Count, 16))
    ThisWorkbook.Names.Add Name:="Standings", RefersTo:=field
    Set header_field = Range(Cells(tables_vStart + parts.Rows.Count + 2, 7), Cells(tables_vStart + parts.Rows.Count + 2, 10))
    Set table = Range(Cells(tables_vStart + 1, 7 + 2), Cells(tables_vStart + parts.Rows.Count, 7 + parts.Rows.Count * 2 + 1))
    Set table_over = Range(Cells(tables_vStart, 7 + 2), Cells(tables_vStart, 7 + parts.Rows.Count * 2 + 1))
    
    Set sum_field = Range(Cells(tables_vStart + 1, 7 + 2 + table.Columns.Count), Cells(tables_vStart + parts.Rows.Count, 7 + 2 + table.Columns.Count))

    Call create_header(header_field, "Standings:")
    field.Interior.ColorIndex = 0
    
    
    With field.Range(Cells(1, 1), Cells(1, 10))
        .Font.Size = 22
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    With field.Range(Cells(2, 1), Cells(field.Rows.Count, 1))
        .Font.Size = 22
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With field.Range(Cells(2, 3), Cells(field.Rows.Count, field.Columns.Count))
        .Range(Cells(1, 1), Cells(.Rows.Count, 1)).Font.Bold = False
        .Range(Cells(1, 1), Cells(.Rows.Count, 1)).Font.Size = 22
        .Range(Cells(1, 1), Cells(.Rows.Count, 1)).HorizontalAlignment = xlCenter
        .Range(Cells(1, 1), Cells(.Rows.Count, 1)).NumberFormat = "General"
        
        .Range(Cells(1, 3), Cells(.Rows.Count, 7)).Font.Bold = True
        .Range(Cells(1, 3), Cells(.Rows.Count, 7)).Font.Size = 22
        .Range(Cells(1, 3), Cells(.Rows.Count, 7)).HorizontalAlignment = xlCenter
        .Range(Cells(1, 3), Cells(.Rows.Count, 7)).NumberFormat = "General"

    End With
    
    
    With field
        Dim i As Integer
        For i = 1 To field.Rows.Count
            .Range(Cells(i, 1), Cells(i, 2)).Merge
            .Range(Cells(i, 3), Cells(i, 4)).Merge
            .Range(Cells(i, 5), Cells(i, 6)).Merge
            .Range(Cells(i, 7), Cells(i, 8)).Merge
            .Range(Cells(i, 9), Cells(i, 10)).Merge
            If i > 1 Then
                .Cells(i, 1).NumberFormat = "@"
                .Cells(i, 1).Interior.Color = COLOR_FOREGROUND_1
                .Cells(i, 1).Value = CStr(i - 1 & ".")
            End If
        Next i
    
        .Cells(1, 1).Value = "Plass:"
        .Cells(1, 1).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 3).Value = "Navn:"
        .Cells(1, 3).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 5).Value = "Poeng:"
        .Cells(1, 5).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 7).Value = "Kamper:"
        .Cells(1, 7).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 9).Value = "Seiere:"
        .Cells(1, 9).Interior.Color = COLOR_FOREGROUND_1
        
        
    End With
    

    Call inside_lines(field)
    Call outer_border(field)
    
    Dim adds As Range
    With table
        For i = 1 To parts.Rows.Count
            Dim n As Integer
            For n = 1 To parts.Rows.Count
                If n = 1 Then
                    Set adds = Range(.Cells(i, 2 * n - 1), .Cells(i, 2 * n - 1))
                Else
                    Set adds = Union(adds, .Cells(i, 2 * n - 1))
                End If
            Next n
            sum_field.Cells(i, 1).formula = "=SUM(" & adds.Address & ",0.0" & (max_participants - i + 1) & ")"
        Next i
        sum_field.Font.Color = COLOR_BACKGROUND

        
        Dim offset_formula As String
        
        Dim p As Integer
        For p = 1 To parts.Rows.Count
            offset_formula = "OFFSET(" & table_over.Address & ", MATCH(LARGE(" & sum_field.Address & "," & p & " )," & sum_field.Address & ",0),0,1," & .Columns.Count & ")"
            field.Cells(1 + p, 3).formula = "=INDEX(" & parts.Address & ", MATCH(LARGE(" & sum_field.Address & "," & p & ")," & sum_field.Address & ",0))"
            field.Cells(1 + p, 5).formula = "=INT(LARGE(" & sum_field.Address & "," & p & "))"
            field.Cells(1 + p, 7).formula = "=COUNTIF(" & offset_formula & ",  "">= 2"" )"
           
            field.Range(Cells(1 + p, 9), Cells(1 + p, 10)).UnMerge
            field.Cells(1 + p, 9).FormulaArray = "=COUNT(IF(IF(MOD(COLUMN(" & "XOX" & ")+1" & "," & "2)=0" & "," & "XOX" & "," & "0)=" & group_first_to & "," & "XOX" & "," & """""))"
            field.Cells(1 + p, 9).Replace What:="XOX", Replacement:=offset_formula
            field.Range(Cells(1 + p, 9), Cells(1 + p, 10)).Merge
        Next p
        Call set_color_condition_played(Range(field.Cells(2, 7), field.Cells(field.Rows.Count, 7)))
       
        End With
End Sub
Public Function get_match_winner(first As String, second As String) As Integer
'Funksjonen returnerer en integer (1 eller 2) som indikerer vinneren av kampen'
    Dim winner As Integer
    With Sheets("Groupstage").Range("Points")
        Dim down As Range
        Dim side As Range
        Set down = Range(.Cells(2, 1), .Cells(.Rows.Count, 1))
        Set side = Range(.Cells(1, 3), .Cells(1, .Columns.Count))
        
        Dim y As Integer
        For y = 1 To down.Rows.Count
            Dim x As Integer
            For x = 1 To side.Columns.Count / 2
                If down.Cells(y, 1).Value = first And side.Cells(1, 2 * x - 1).Value = second Then
                    If .Cells(1 + y, 2 * x - 1 + 2) = first_to Then
                        winner = 1
                    ElseIf .Cells(1 + y, 2 * x - 1 + 3) = first_to Then
                        winner = 2
                    Else
                        MsgBox "Could not find winner between " & first & " and " & second
                    End If
                End If
            Next x
        Next y
    End With
    get_match_winner = winner
    
End Function

Public Function get_wins_count(player As String)
    Dim wins As Integer
    With Sheets("Groupstage").Range("Standings")
        
        Dim n As Integer
        n = 1
        Do While Not (.Cells(1 + n, 3).Value = player)
            n = n + 1
        Loop
        wins = .Cells(1 + n, 9).Value
    End With
    
    get_wins_count = wins
End Function
Sub create_adjusted_standings(dict As Dictionary, points As Dictionary, extra_points As Dictionary)
    Dim parts As Range
    Dim field As Range
    Dim stand As Range
    Dim header_field As Range
    
    Set stand = Sheets("Groupstage").Range("Standings")
    Set parts = Range(stand.Cells(2, 3), stand.Cells(stand.Rows.Count, 3))
    
    Set header_field = Sheets("Mainstage").Range(Cells(tables_vStart - 1, tables_hStart), Cells(tables_vStart - 1, tables_hStart + 3))
    Set field = Sheets("Mainstage").Range(Cells(tables_vStart, tables_hStart), Cells(tables_vStart + stand.Rows.Count - 1, tables_hStart + 7))
    ThisWorkbook.Names.Add Name:="Adjusted_Standings", RefersTo:=field
    
    Call create_header(header_field, "Adjusted standings:")
    
    field.Interior.ColorIndex = 0
    
    
    With field.Range(Cells(1, 1), Cells(1, 10))
        .Font.Size = 22
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "@"
        
    End With

    With field.Range(Cells(2, 1), Cells(field.Rows.Count, 1))
        .Font.Size = 22
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With field.Range(Cells(2, 3), Cells(field.Rows.Count, field.Columns.Count))
        .Range(Cells(1, 1), Cells(.Rows.Count, 2)).Font.Bold = False
        .Range(Cells(1, 1), Cells(.Rows.Count, 2)).Font.Size = 22
        .Range(Cells(1, 1), Cells(.Rows.Count, 2)).HorizontalAlignment = xlCenter
        .Range(Cells(1, 1), Cells(.Rows.Count, 2)).VerticalAlignment = xlCenter
        .Range(Cells(1, 1), Cells(.Rows.Count, 2)).NumberFormat = "@"
        
        .Range(Cells(1, 3), Cells(.Rows.Count, 4)).Font.Bold = True
        .Range(Cells(1, 3), Cells(.Rows.Count, 4)).Font.Size = 22
        .Range(Cells(1, 3), Cells(.Rows.Count, 4)).HorizontalAlignment = xlCenter
        .Range(Cells(1, 3), Cells(.Rows.Count, 4)).VerticalAlignment = xlCenter
        .Range(Cells(1, 3), Cells(.Rows.Count, 4)).NumberFormat = "@"
 
        .Range(Cells(1, 5), Cells(.Rows.Count, 6)).Font.Bold = False
        .Range(Cells(1, 5), Cells(.Rows.Count, 6)).Font.Size = 22
        .Range(Cells(1, 5), Cells(.Rows.Count, 6)).HorizontalAlignment = xlCenter
        .Range(Cells(1, 5), Cells(.Rows.Count, 6)).VerticalAlignment = xlCenter
        .Range(Cells(1, 5), Cells(.Rows.Count, 6)).NumberFormat = "@"
        .Range(Cells(1, 5), Cells(.Rows.Count, 6)).Font.Color = COLOR_FOREGROUND_2
    End With
    
    
    With field
        Dim i As Integer
        For i = 1 To field.Rows.Count
            .Range(Cells(i, 1), Cells(i, 2)).Merge
            .Range(Cells(i, 3), Cells(i, 4)).Merge
            .Range(Cells(i, 5), Cells(i, 6)).Merge
            .Range(Cells(i, 7), Cells(i, 8)).Merge
            If i > 1 Then
                .Cells(i, 1).NumberFormat = "@"
                .Cells(i, 1).VerticalAlignment = xlCenter
                .Cells(i, 1).Interior.Color = COLOR_FOREGROUND_1
                .Cells(i, 1).Value = CStr(i - 1 & ".")
            End If
        Next i
    
        .Cells(1, 1).Value = "Plass:"
        .Cells(1, 1).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 3).Value = "Navn:"
        .Cells(1, 3).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 5).Value = "Poeng:"
        .Cells(1, 5).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 7).Value = "Ekstra:"
        .Cells(1, 7).Interior.Color = COLOR_FOREGROUND_1
        
    End With

    Call inside_lines(field)
    Call outer_border(field)
    
    Dim orig_stand As Range
    Set orig_stand = Sheets("Groupstage").Range("Standings")
    Dim orig_data As Range
    Set orig_data = Range(orig_stand.Cells(2, 3), orig_stand.Cells(orig_stand.Rows.Count, orig_stand.Columns.Count))

    With orig_data
        Dim anchor As Integer
        anchor = 1
        Dim cluster_points As Integer
        
        For i = 1 To field.Rows.Count - 1
            Dim auto_matched As Boolean
            Dim needs_play As Boolean
            Dim adjusted_points As Boolean
            auto_matched = False
            needs_play = False
            adjusted_points = False
            
            Dim new_player As Integer
            Dim player_name As String
            Dim cluster_rating As Integer
            cluster_rating = dict.Item(parts.Cells(i).Value)
            
            ' If the player has zero cluster rating, it shall remain in the same position as in the original standings
            If cluster_rating = 0 Then
                new_player = i
            
            ' If the player has more than zero in cluster rating, it must be a part of a cluster and may get a new standing
            ElseIf cluster_rating > 0 Then
                    adjusted_points = True
                    If Not .Cells(i, 3).Value = .Cells(anchor, 3) Then
                        anchor = i
                    End If
                    
                    ' Counting cluster size
                    Dim cluster_s As Integer
                    cluster_s = 1
                    cluster_points = .Cells(i, 3)
                    Dim p As Integer
                    For p = anchor To field.Rows.Count
                        If cluster_points = .Cells(p, 3).Value Then
                            cluster_s = cluster_s + 1
                        End If
                    Next p
                    
                    'Debug.Print "Anchor: " & anchor & ", Cluster size: " & cluster_s
                    For p = anchor To anchor + cluster_s - 1
                    'Debug.Print "Checking if " & parts.Cells(p).Value & " [" & dict.Item(parts.Cells(p).Value) & "] has Value " & i - anchor + 1
                        If dict.Item(parts.Cells(p).Value) = i - anchor + 1 Then
                            new_player = p
                            Exit For
                        End If
                    Next
            
            ' If the player has -1 as cluster rating, it is automatically matched to his equal due to tournament structure and does not requre a tiebreaker match
            ElseIf cluster_rating = -1 Then
                auto_matched = True
                new_player = i
                field.Cells(1 + i, 3).Interior.Color = COLOR_PASS
                
            ' If the player has -2 as cluster rating, it is to play a tiebreaker set against the other player in the cluster
            ElseIf cluster_rating = -2 Then
                needs_play = True
                new_player = i
                field.Cells(1 + i, 3).Interior.Color = COLOR_FAIL
            End If
                              
            player_name = .Cells(new_player, 1).Value
            field.Cells(1 + i, 3).Value = player_name
            field.Cells(1 + i, 5).Value = points.Item(player_name)
            
            If extra_points.Item(player_name) = 0 Then
                field.Cells(1 + i, 7).Value = "-"
            Else
                field.Cells(1 + i, 7).Value = extra_points.Item(player_name)
            End If
        Next i
    End With
    
End Sub

Sub create_tiebrakers(stand As Range, extra_matches As Integer)
    
    Dim n_p As Long
    Dim dict As Dictionary
    Dim points_dict As Dictionary
    Dim extra_points_dict As Dictionary
    Set dict = New Dictionary
    Set points_dict = New Dictionary
    Set extra_points_dict = New Dictionary
    
    Dim parts As Range
    Set parts = Range(stand.Cells(2, 3), stand.Cells(stand.Rows.Count, 3))
    
    'Populating the score dictionary
    Dim o As Integer
    For o = 2 To stand.Rows.Count
        points_dict.Add Key:=stand.Cells(o, 3).Value, Item:=stand.Cells(o, 5).Value
    Next o
    

    
    With Range(stand.Cells(2, 5), stand.Cells(stand.Rows.Count, 5))
        n_p = .Rows.Count
        
        Dim i As Integer
        i = 1
        Do While i <= .Rows.Count
            Dim n As Integer
            Dim anchor As Integer
            Dim cluster_size As Integer
            n = i + 1
            anchor = i
            cluster_size = 1
            
            For n = i + 1 To .Rows.Count
                If .Cells(anchor, 1).Value = .Cells(n, 1).Value Then
                    cluster_size = cluster_size + 1
                Else
                    Exit For
                End If
            Next n
            
            'Debug.Print "i, anchor, cluster size and n_p"
            'Debug.Print i & ", "; anchor & ", " & cluster_size & ", " & n_p
            
            i = i + cluster_size

            
            '------------------------------ Regler ----------------------------------'
            Dim valid As Boolean
            valid = True
            
            'Dersom clusteret er større enn 2 må måtiebreakeren løses med andre midler'
            If cluster_size = 2 Then
            
                extra_points_dict.Add Key:=parts(anchor).Value, Item:=0
                extra_points_dict.Add Key:=parts(anchor + 1).Value, Item:=0
                'Dersom de to midterste har lik poengsum. Disse skal møtes i første kamp uansett'
                If n_p Mod 2 = 0 And anchor = n_p / 2 And extra_matches = 0 Then
                    valid = False
                End If
                
                'For deltakertall som ikke tilfredsstiller 2^n må de lavest seeded spillerene spille introduksjonskamper uansett
                Dim k As Integer
                'For k = 1 To extra_matches
                    'If anchor = n_p - (2 * k - 1) Then valid = False
                'Next k
                
                'Paret slipper å spille tiebraker'
                If valid = False Then
                    dict.Add Key:=parts(anchor).Value, Item:=-1
                    dict.Add Key:=parts(anchor + 1).Value, Item:=-1
                    
                'Paret må avgjøre stilling i en tiebraker'
                Else:
                    dict.Add Key:=parts(anchor).Value, Item:=-2
                    dict.Add Key:=parts(anchor + 1).Value, Item:=-2
                End If
                
            'Håndtering av tilfellet hvor clusteret er større enn 2'
            ElseIf cluster_size > 2 Then
                Dim win As Integer
                Dim points() As Integer
                Dim equals_points() As Integer
                
                'Teller opp samlet seiere for hver spiller i clusteret'
                ReDim points(1 To cluster_size) As Integer
                ReDim equals_points(1 To cluster_size) As Integer
                Dim equal_clusters As Integer
                
                For k = anchor To cluster_size - 1 + anchor
                    points(k - anchor + 1) = points(k - anchor + 1) + get_wins_count(parts.Cells(k).Value)
                    'Debug.Print parts.Cells(k).Value & " has " & points(k - anchor + 1) & " wins and cluster index: " & k - anchor + 1
                    
                    'Dersom noen av spillerene har like mange seiere telles kamper de har spilt mot hverandre tidligere i gruppespillet som avgjørende'
                    Dim j As Integer
                    For j = k + 1 To cluster_size - 1 + anchor
                        win = get_match_winner(parts.Cells(k).Value, parts.Cells(j).Value)
                        If win = 1 Then
                            points(k - anchor + 1) = points(k - anchor + 1) + 1
                            'Debug.Print parts.Cells(k).Value & " won over " & parts.Cells(j) & " and now has " & points(k - anchor + 1) & " points."
                        ElseIf win = 2 Then
                            points(j - anchor + 1) = points(j - anchor + 1) + 1
                            'Debug.Print parts.Cells(j).Value & " won over " & parts.Cells(k) & " and now has " & points(j - anchor + 1) & " points."
                        End If
                    Next j
                Next k
                
                For k = anchor To cluster_size - 1 + anchor
                    'Dersom spillerene enda har like poengsummer velges vinneren av clusteret tilfeldig :('
                    For j = k + 1 To cluster_size - 1 + anchor
                        If points(k - anchor + 1) = points(j - anchor + 1) Then
                            'MsgBox parts.Cells(k).Value & ": p = " & equals_points(k - anchor + 1) & ", i = " & k - anchor + 1 & " EQUALS " & Cells(j + anchor - 1).Value & ": p = " & equals_points(j) & ", i = " & j
                            'If equals_points(k - anchor + 1) = equals_points(j - anchor + 1) Then
                            equals_points(j - anchor + 1) = equals_points(j - anchor + 1) + Int(100 / cluster_size * Rnd) + 1 'Byttes ut mot en funksjon som genererer unike verdier'
                            equals_points(k - anchor + 1) = equals_points(k - anchor + 1) + Int(100 / cluster_size * Rnd) + 1
                            'End If
                        End If
                    Next j
                Next k

            
                'Lager justerte standings etter behandling av algoritmen over'
                Dim pointsC() As Integer
                Dim equals_pointsC() As Integer
                pointsC = points
                equals_pointsC = equals_points
            
                Call QuickSort(pointsC, 1, CLng(cluster_size))
                Call QuickSort(equals_pointsC, 1, CLng(cluster_size))
                'Debug.Print "Sorted points: "
                'Dim jj As Long
                'For jj = 1 To cluster_size
                    'Debug.Print pointsC(jj) & ", " & equals_pointsC(jj)
                'Next jj
                Dim z As Integer
                z = cluster_size
                For k = 1 To cluster_size
                Dim p As Integer
                    For p = 1 To cluster_size
                        For j = 1 To cluster_size
                            'Debug.Print pointsC(k) & "=" & points(j) & " and " & equals_pointsC(p) & "=" & equals_points(j) & ", indicies=" & k & p & j
                            If pointsC(k) = points(j) And equals_pointsC(p) = equals_points(j) Then
                                If Not dict.Exists(parts(j + anchor - 1).Value) Then
                                    dict.Add Key:=parts(j + anchor - 1).Value, Item:=z
                                    Dim ite As Double
                                    ite = equals_points(j) / 100 + CDbl(points(j))
                                    extra_points_dict.Add Key:=parts(j + anchor - 1).Value, Item:=ite
                                    z = z - 1
                                    Exit For
                                End If
                            End If
                        Next j
                    Next p
                Next k
            
            'Dersom clusteret bare består av en spiller, altså at spilleren har garantert standing'
            Else:
                If Not dict.Exists(parts(anchor).Value) Then
                    dict.Add Key:=parts(anchor).Value, Item:=0
                    extra_points_dict.Add Key:=parts(anchor).Value, Item:=0
                End If
            End If
        Loop
        
        'Debug.Print "Debugging dictionary of size " & extra_points_dict.Count & " :"
        'Dim ii As Long
        'For ii = 0 To extra_points_dict.Count - 1
            'Debug.Print extra_points_dict.Keys()(ii), extra_points_dict.Items()(ii)
        'Next ii
    
    End With
    
    Call create_adjusted_standings(dict, points_dict, extra_points_dict)
    Call create_match("TEST1", "TEST2", 20, tables_vStart, tiebreaker_first_to)
    
    
    
End Sub

Sub create_upperbracket(stand As Range)
    Dim byes As Integer
    Dim extra_matches As Integer
    Dim n As Integer
    Dim i As Integer
    
    Dim field As Range
    Dim parts As Range
    Set parts = Sheets("Groupstage").Range("Parts")
    
    i = 0
    
    Do While n < parts.Rows.Count
        n = 2 ^ i
        i = i + 1
    Loop
    
    byes = n - parts.Rows.Count
    
    extra_matches = parts.Rows.Count - byes
    Call create_tiebrakers(stand, extra_matches)

    
    'Set field = Range(Cells(), Cells())
    'MsgBox extra_matches
    
    
End Sub
