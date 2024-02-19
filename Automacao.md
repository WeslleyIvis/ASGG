Sub SET_Calendario()

      Sub SET_Calendario()
        Dim count_feriados As Integer, dia_do_feriado As Date
        Dim sizeColumnWIdth As Double
        Dim sizeRowHeight As Double
        Dim position_row As Integer
        Dim position_column As Integer
        Dim gap As Integer
        
        Sheets("Calendario").Activate
    
        'Limpa a sequencia de dias uteis gerada
        Range("O16:AM16").ClearContents
    
        'Gera a sequencia de dias uteis do mês
        Range("N16").DataSeries Rowcol:=xlRows, Type:=xlChronological, Date:= _
        xlWeekday, Step:=1, Stop:=Range("K17").Value, Trend:=False
    
        'Limpa as celulas do calendario
        Range("B4:F4,B6:F6,B8:F8,B10:F10,B12:F12").Clear
    
        'Define o tamanho das celulas da planilha
        
        Range("B:F").ColumnWidth = 30
        Range("B4,B6,B8,B10,B12").RowHeight = 150
        
        sizeColumnWIdth = Range("B:F").ColumnWidth
        sizeRowHeight = Range("B4,B6,B8,B10,B12").RowHeight
    
        'Chama funçao coloca borda
        SET_borda
        
        'Pinta as celulas que são feriados
        If Range("L17").Value Then
            count_feriados = Range("L16").End(xlDown).Row
            
         
            
            For Index = 17 To count_feriados
                position_row = 3
                position_column = 2
                
                dia_do_feriado = Cells(Index, 12).Value
                                    
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
                     
                    If position_column >= 7 Then
                        position_row = position_row + 2
                        position_column = 2
                    End If
                Next
            Next
    End If
    
 
    position_row = 4
    position_column = 2
    
    For Index = 1 To 5
        gap = 5
        
        For x = 23 To Range("L22").End(xlDown).Row
              If Index = Range("K" & x).Value Then
                Call Cria_Dados_Rentagulos(Cells(position_row, Range("L" & x).Value + 1).Left, _
                Cells(position_row, 2).Top, _
                Cells(4, 2).ColumnWidth, _
                sizeRowHeight, _
                Range("J" & x), gap, _
                Range("M" & x).Value)
                
                gap = gap + sizeRowHeight / 7 + 5
              End If
        Next
    position_row = position_row + 2
    Next
       
    

    End Sub


    Sub SET_borda()

        Range("B2:F12").Select
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
        
        
        Range( _
            "B4,C4,D4,E4,F4,B6,C6,D6,E6,F6,B8,C8,D8,E8,F8,B10,C10,D10,E10,F10,B12,C12,D12,E12,F12" _
            ).Select
        Range("F12").Activate
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
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    End Sub

    Sub Cria_Dados_Rentagulos(cellPositionX As Double, cellPositionY As Double, celWidht As Double, celHeight As Double, _
      descricao As String, gap As Integer, qtdSemana As Integer)
          
          
          'Argumentos (tipoGeometrico, x, y, largura, altura)
          ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
              cellPositionX + 2, _
              cellPositionY + gap, _
              celWidht * qtdSemana, _
              celHeight / 7).Select
          Selection.ShapeRange.ScaleWidth 5.25, msoScaleFromTopLeft
          Selection.ShapeRange.ShapeStyle = msoShapeStylePreset12
          Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = descricao
          
      End Sub


