Sub SET_Calendario()

    Dim count_feriados As Integer, dia_do_feriado As Date
    
    Dim position_row As Integer, position_column As Integer
    
    'Limpa a sequencia de dias uteis gerada
    Range("O16:AM16").ClearContents

    'Gera a sequencia de dias uteis do mês
    Range("N16").DataSeries Rowcol:=xlRows, Type:=xlChronological, Date:= _
    xlWeekday, Step:=1, Stop:=Range("K17").Value, Trend:=False

    'Limpa as celulas do calendario
    Range("B4:F4,B6:F6,B8:F8,B10:F10,B12:F12").Clear

    'Chama funçao coloca borda
    coloca_borda
    
    'Caso nao tenha feriado, nao vai começar o loop para pintar as celulas
    If Range("L17").Value Then
        count_feriados = Range("L16").End(xlDown).Row
        
        For Index = 17 To count_feriados
            dia_do_feriado = Cells(Index, 12).Value
            
            position_row = 3
            position_column = 2
            
            For count_dias = 1 To 25
                dia = Cells(position_row, position_column).Value
                
                If dia = dia_do_feriado Then
                    Cells(position_row + 1, position_column).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.149998474074526
                        .PatternTintAndShade = 0
                    End With
                End If
                
                position_column = position_column + 1
                 
                If position_column >= 6 Then
                    position_row = position_row + 2
                    position_column = 2
                End If
                              
            Next
            
        Next
    
    End If

End Sub

Sub coloca_borda()
'
' coloca_borda Macro
'
    Range("B2:F11").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
