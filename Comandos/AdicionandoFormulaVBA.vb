'https://pt.stackoverflow.com/questions/209353/bloquear-c%C3%A9lula-preenchida-com-vba

ActiveSheet(1)

        strSQL = "" _
                & " SELECT   " _
                & "   1 AS COL_A  " _
                & " , 4 AS COL_B  "

        rs = cnn.Execute(strSQL)

        xlApp.DisplayAlerts = False

        'grava dados do select no excel
        With xlWorkSheet.Range("A2")
            .CopyFromRecordset(rs)
        End With

        Rows = xlWorkSheet.UsedRange.Rows.Count

        With xlWorkSheet.Range("A1:C1")
            .Font.Bold = True
        End With

        With xlWorkSheet.Range("A1:C" & Rows)
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Color = RGB(0, 0, 0)
        End With

        ' adicionando formula
        xlWorkSheet.Range("C2").Formula = "=SOMA(A2+B2)"

        ' exemplo2
        xlWorkSheet.Range("C2").Formula = "=IF(OR(G4="""",$B$19=0),"""",ROUND(G4/$B$19,2))"


        'AJUSTES FINAIS AUTOFIT
        xlWorkSheet.Cells.EntireColumn.AutoFit()
        xlWorkSheet.Cells.EntireRow.AutoFit()
        xlWorkSheet.Range("A1").Select()
