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

    Set cond_fail = field.FormatConditions.Add(Type:=xlExpression, Formula1:="=OR(SUM(" & field.Address & ") > " & 2 * first_to - 1 & ";OR(" & field(1).Address & "< 0;" & field(2).Address & " < 0); OR(" & field(1).Address & ">" & first_to & "; " & field(2).Address & ">" & first_to & "))")
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
        .VerticalAlignment = xlCenter
    
        Call outer_border_small(field)
        .Range(Cells(1, 1), Cells(1, 2)).Merge
        .Cells(1, 1).Value = p1
        
        .Range(Cells(2, 1), Cells(2, 2)).Merge
        .Cells(2, 1).Value = p2
        .Cells(1, 3).NumberFormat = "0"
        .Cells(2, 3).NumberFormat = "0"
        
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
    Call create_header(header_field, "Matchups:")
    
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
            'For columns'
            .Range(Cells(i, 1), Cells(i, 2)).Merge
            .Cells(i + 1, 1).Value = parts.Cells(i, 1).Value
            
            .Cells(i, 1).Interior.Color = COLOR_FOREGROUND_1
            .Cells(i, 1).NumberFormat = "@"
            .Cells(i, 1).Font.Size = 20
            .Cells(i, 1).Font.Bold = True
            .Cells(i, 1).HorizontalAlignment = xlCenter
            
            'For rows'
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
    Call create_header(header_field, "Points table:")

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
    
        .Cells(1, 1).Value = "Place:"
        .Cells(1, 1).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 3).Value = "Name:"
        .Cells(1, 3).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 5).Value = "Points:"
        .Cells(1, 5).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 7).Value = "Matches:"
        .Cells(1, 7).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 9).Value = "Wins:"
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
            field.Cells(1 + p, 7).formula = "=COUNTIF(" & offset_formula & ",  "">= " & group_first_to & """ )"
           
            field.Range(Cells(1 + p, 9), Cells(1 + p, 10)).UnMerge
            field.Cells(1 + p, 9).FormulaArray = "=COUNT(IF(IF(MOD(COLUMN(" & "XOX" & ")+1" & "," & "2)=0" & "," & "XOX" & "," & "0)=" & group_first_to & "," & "XOX" & "," & """""))"
            field.Cells(1 + p, 9).Replace What:="XOX", Replacement:=offset_formula
            field.Range(Cells(1 + p, 9), Cells(1 + p, 10)).Merge
        Next p
        Call set_color_condition_played(Range(field.Cells(2, 7), field.Cells(field.Rows.Count, 7)))
       
        End With
End Sub
Public Function get_match_winner(first As String, second As String, first_to) As Integer
'Returns an integer (1 or 2) based on the winner of the match. Fails to 3'
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
    ThisWorkbook.Names.Add Name:="AdjustedStandings", RefersTo:=field
    
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
    
        .Cells(1, 1).Value = "Place:"
        .Cells(1, 1).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 3).Value = "Name:"
        .Cells(1, 3).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 5).Value = "Points:"
        .Cells(1, 5).Interior.Color = COLOR_FOREGROUND_1
        .Cells(1, 7).Value = "Extra:"
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
        
        Dim tiebreaker_parts As New Collection
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
                    
                    For p = anchor To anchor + cluster_s - 1
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
                new_player = i
                player_name = .Cells(new_player, 1).Value
                tiebreaker_parts.Add player_name
                needs_play = True
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
        
        ' Setting up the tiebreaker matches (if any)
        Dim k As Integer
        Dim h As Integer
        Dim w As Integer
        Dim s As Integer
        Dim match As Range
        h = IIf(tiebreaker_parts.Count / 2 < field.Rows.Count \ 3, tiebreaker_parts.Count / 2, field.Rows.Count \ 3)
        w = Application.WorksheetFunction.RoundUp((tiebreaker_parts.Count / 2) / CDbl(h), 0)
        For k = 1 To tiebreaker_parts.Count / 2
            If k = 1 Then
                Call create_container(16, tables_vStart, w * 4, h * 3)
                Call create_header(Range(Cells(tables_vStart - 1, 16), Cells(tables_vStart - 1, 19)), "Tiebreakers:")
            End If
            Set match = create_match(tiebreaker_parts(2 * k - 1), tiebreaker_parts(2 * k), 17 + ((k - 1) \ h) * 4, tables_vStart + 1 + ((k - 1) Mod h) * 3, tiebreaker_first_to)
            ThisWorkbook.Names.Add Name:="Tiebreaker" & CStr(k), RefersTo:=match
            
            
            Dim cond_pass As FormatCondition
            Dim index As Integer
            index = Application.WorksheetFunction.match(tiebreaker_parts(2 * k - 1), field.Range(Cells(2, 3), Cells(field.Rows.Count, 3)), 0)
            Debug.Print index
            Set cond_pass = field.Range(Cells(1 + index, 3), Cells(2 + index, 3)).FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(OR(" & match(1).Address & "=" & tiebreaker_first_to & ";" & match(2).Address & "=" & tiebreaker_first_to & ");SUM(" & match.Address & ")<=" & 2 * tiebreaker_first_to - 1 & ";NOT(OR(" & match(1).Address & "<0;" & match(2).Address & "<0)))")
            With cond_pass
                .Interior.ColorIndex = 0
            End With
        Next k
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
            
            i = i + cluster_size

            
            '------------------------------ Rules ----------------------------------'
            'Clusters are groups of participants that have the same score after the group stage'
            'Since there can exist several clusters of size 1, 2 or more, each of these three cases _
            must be handled differently.
            
            'For clusters of size three or more, the rules are written such that each player is given _
            extra points for: _
                * Total number of wins _
                * Number of games won against players in the cluster
            'If this cluster fails to be solved by the addition of these extra points, the standings of _
            the cluster is chosen randomly. I hate random..
            
            Dim valid As Boolean
            valid = True
            
            'For the case where a cluster is of size 2'
            If cluster_size = 2 Then
            
                extra_points_dict.Add Key:=parts(anchor).Value, Item:=0
                extra_points_dict.Add Key:=parts(anchor + 1).Value, Item:=0
                ' If the the two players in the middle of the standings are equal in points, they will meet in the first round of mainstage anyways.
                If n_p Mod 2 = 0 And anchor = n_p / 2 And extra_matches = 0 Then
                    valid = False
                End If
                
                'In the case of a number of participants that does not satisfy 2^n, the following players will meet in the first round of mainstage anyways.
                'Dim k As Integer
                'For k = 1 To extra_matches
                    'If anchor = n_p - (2 * k - 1) Then valid = False
                'Next k
                
                'If the cluster of 2 players does not need to meet for a tiebreaker match'
                If valid = False Then
                    dict.Add Key:=parts(anchor).Value, Item:=-1
                    dict.Add Key:=parts(anchor + 1).Value, Item:=-1
                    
                'If the cluster of 2 players need to meet for a tiebreaker match'
                Else:
                    dict.Add Key:=parts(anchor).Value, Item:=-2
                    dict.Add Key:=parts(anchor + 1).Value, Item:=-2
                End If
                
            'If the cluster is of size 2'
            ElseIf cluster_size > 2 Then
                Dim win As Integer
                Dim points() As Integer
                Dim equals_points() As Integer
                
                'Counting the number of victories for each player in the cluster'
                ReDim points(1 To cluster_size) As Integer
                ReDim equals_points(1 To cluster_size) As Integer
                Dim equal_clusters As Integer
                
                Dim k As Integer
                For k = anchor To cluster_size - 1 + anchor
                    points(k - anchor + 1) = points(k - anchor + 1) + get_wins_count(parts.Cells(k).Value)
                    
                    'If the addition of extra points for the victories still has not resolved the cluster, points are given for each win against players in the cluster'
                    Dim j As Integer
                    For j = k + 1 To cluster_size - 1 + anchor
                        win = get_match_winner(parts.Cells(k).Value, parts.Cells(j).Value, group_first_to)
                        If win = 1 Then
                            points(k - anchor + 1) = points(k - anchor + 1) + 1
                        ElseIf win = 2 Then
                            points(j - anchor + 1) = points(j - anchor + 1) + 1
                        End If
                    Next j
                Next k
                
                For k = anchor To cluster_size - 1 + anchor
                    'If the cluster still has not resolved, the standings are chosen randomly :(
                    For j = k + 1 To cluster_size - 1 + anchor
                        If points(k - anchor + 1) = points(j - anchor + 1) Then
                            equals_points(j - anchor + 1) = equals_points(j - anchor + 1) + Int(100 / cluster_size * Rnd) + 1 'Byttes ut mot en funksjon som genererer unike verdier'
                            equals_points(k - anchor + 1) = equals_points(k - anchor + 1) + Int(100 / cluster_size * Rnd) + 1
                        End If
                    Next j
                Next k

            
                'Algorithm for sorting and matching the players in the cluster after additional points are given'
                Dim pointsC() As Integer
                Dim equals_pointsC() As Integer
                pointsC = points
                equals_pointsC = equals_points
            
                Call QuickSort(pointsC, 1, CLng(cluster_size))
                Call QuickSort(equals_pointsC, 1, CLng(cluster_size))
                Dim z As Integer
                z = cluster_size
                For k = 1 To cluster_size
                Dim p As Integer
                    For p = 1 To cluster_size
                        For j = 1 To cluster_size
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
            
            'If the cluster is of size 1, the player has a guaranteed standing'
            Else:
                If Not dict.Exists(parts(anchor).Value) Then
                    dict.Add Key:=parts(anchor).Value, Item:=0
                    extra_points_dict.Add Key:=parts(anchor).Value, Item:=0
                End If
            End If
        Loop
    End With
    
    Call create_adjusted_standings(dict, points_dict, extra_points_dict)
    
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
    
    extra_matches = (parts.Rows.Count - byes) / 2
    Call create_tiebrakers(stand, extra_matches)
    
    Dim ad_stand As Range
    Set ad_stand = Sheets("Mainstage").Range("Adjusted_Standings")
    
    Dim k As Integer
    
    For k = byes + 1 To parts.Rows.Count
        'Call create_match("=3+2", "=3*6", 22, tables_vStart + k * 3, 2)
        If ad_stand.Cells(1 + k, 3).Interior.Color = COLOR_FAIL Then
        End If
        
    Next k
    
    
End Sub
