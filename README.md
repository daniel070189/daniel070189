Sub duplicar()
Set B = Sheets("planilha1")
B.Columns("D:H").ClearContents
B.Columns("A:C").AutoFilter
B.Columns("D:H").Font.Size = 14
B.AutoFilter.Sort.SortFields.Add Key:=Range("C1:C3705"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
B.AutoFilter.Sort.SortFields.Add Key:=Range("A1:A3705"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
B.Columns("A:C").AutoFilter
B.Columns("A:A").ColumnWidth = 0
B.Columns("B:B").ColumnWidth = 0
B.Columns("C:C").ColumnWidth = 0

For A = 2 To B.Cells(2, 3).End(xlDown).Row

If B.Cells(A, 1) = "UTIN" Or B.Cells(A, 1) = "UCIN" Or B.Cells(A, 1) = "CENTRO OBSTETRICO" Or B.Cells(A, 1) = "PRE-PARTO" Or B.Cells(A, 1) = "CLINICA CIRURGICA" Or B.Cells(A, 1) = 

"NEUROLOGIA" Or B.Cells(A, 1) = "CPN" Or B.Cells(A, 1) = "ANEXO OBSTETRICO" Then

For D = 1 To 2
With B.Cells(B.Rows.Count, 4).End(xlUp)

.Offset(1, 0).WrapText = True
.Offset(1, 0).RowHeight = 90

.Offset(1, 0) = B.Cells(A, 1) ' SETOR
.Offset(1, 0).ColumnWidth = 20
.Offset(1, 0).HorizontalAlignment = xlLeft
.Offset(1, 0).VerticalAlignment = xlCenter
.Offset(1, 0).Borders(xlEdgeTop).LineStyle = xlContinuous
.Offset(1, 0).Borders(xlEdgeLeft).LineStyle = xlContinuous
.Offset(1, 0).Borders(xlEdgeBottom).LineStyle = xlContinuous
.Offset(1, 0).Borders(xlEdgeRight).LineStyle = xlContinuous

G = G + 1
.Offset(1, G) = B.Cells(A, 2) 'LEITO
.Offset(1, G).ColumnWidth = 20
.Offset(1, G).HorizontalAlignment = xlLeft
.Offset(1, G).VerticalAlignment = xlCenter
.Offset(1, G).Borders(xlEdgeTop).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeLeft).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeBottom).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeRight).LineStyle = xlContinuous

G = G + 1
.Offset(1, G) = B.Cells(A, 3) 'NOME
.Offset(1, G).ColumnWidth = 60
.Offset(1, G).HorizontalAlignment = xlLeft
.Offset(1, G).VerticalAlignment = xlCenter
.Offset(1, G).Borders(xlEdgeTop).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeLeft).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeBottom).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeRight).LineStyle = xlContinuous

G = G + 1
.Offset(1, G) = Date + 1 'DATA
.Offset(1, G).ColumnWidth = 15
.Offset(1, G).HorizontalAlignment = xlLeft
.Offset(1, G).VerticalAlignment = xlCenter
.Offset(1, G).Borders(xlEdgeTop).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeLeft).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeBottom).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeRight).LineStyle = xlContinuous

G = G + 1
.Offset(1, G) = IIf(D = 1, "2hs as 8hs", "9hs as 14hs") 'INTERVALO
.Offset(1, G).ColumnWidth = 15
.Offset(1, G).HorizontalAlignment = xlLeft
.Offset(1, G).VerticalAlignment = xlCenter
.Offset(1, G).Borders(xlEdgeTop).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeLeft).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeBottom).LineStyle = xlContinuous
.Offset(1, G).Borders(xlEdgeRight).LineStyle = xlContinuous

G = 0
H = .Offset(1, G).Row
End With

Next D


End If
Next A

    With B.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
    End With

B.PageSetup.PrintArea = "$A$1:$H$" & H

End Sub
