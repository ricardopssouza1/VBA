-- ===================
--- WHILE  
-- ===================

-- exemplo 1

Dim x As Integer

            x = 1

            While x <= Rows
                If Not xlWorkSheet.Range("C" & x).Value Is Nothing Then
                    If xlWorkSheet.Range("C" & x).Value.ToString() = "VIDA" Then
                        xlWorkSheet.Range("A2:AF" & x).Borders.ColorIndex = 1
                    End If
                End If


                x = x + 1

            End While


-- exemplo 2

        i = 2

        While i <= Rows
            If xlWorkSheet.Range("A" & i).Value.ToString() = "2" Then
                With xlWorkSheet.Range("A" & i & ":G" & i)
                    .Borders.ColorIndex = 1
                    .Font.Name = "Calibri"
                    .Font.Size = 10
                    .Interior.ColorIndex = 6
                End With
            End If

            i = i + 1

        End While


-- exemplo 3

Dim x As Integer = 1
        While x <= ROWS
       
            If Not xlWorkSheet.Range("A" & x).Value Is Nothing Then
                If xlWorkSheet.Range("A" & x).Value.ToString() = "Razão social" Then
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Font.Name = "Calibri"
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Font.Size = 10
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Font.Bold = True
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Borders.ColorIndex = 1
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Interior.Color = RGB(242, 242, 242)
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Borders.Weight = Excel.XlBorderWeight.xlThick
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End If
            End If


            x = x + 1
        End While
		
-- exemplo 4


        If Rows > 1 Then
            xlWorkSheet.Range("L1:L" & Rows).AutoFilter(Field:=1, Criteria1:="1")
            With xlWorkSheet.Range("A1:K" & Rows)
                .Font.Color = RGB(255, 0, 0)
                .Font.Bold = True
            End With
            If xlWorkSheet.AutoFilterMode = True Then
                xlWorkSheet.AutoFilterMode = False
            End If
        End If		


-- ===================
--- RETORNA NOME DA SHEET  
-- ===================


  Dim plan As String

        plan = xlWorkSheet.Name.ToString

        MsgBox(plan)

        With xlWorkSheet.Range("C9")
            .CopyFromRecordset(rs)
        End With


-- ===================
--- CONVERT LINHAS  
-- ===================



        Dim A As Excel.Worksheet = CType(xlWorkbook.Worksheets(4), Excel.Worksheet)
        A.UsedRange.Formula = A.UsedRange.Formula
        A.Calculate()


        A = CType(xlWorkbook.Worksheets(4), Excel.Worksheet)
        A.UsedRange.Formula = A.UsedRange.Formula
        A.Calculate()



-- ===================
--- REMOVE COLUNA  
-- ===================

        With xlWorkSheet.Range("A1")
            .EntireColumn.Delete()
        End With

-- ===================
-- COLA DADOS NO XLS 
-- =================== 


	CString_DB_ADMIX_IN


        Dim CaminhoArquivo As String
        Dim Rows As Integer
        Dim strSQL As String
        Dim cnn As New ADODB.Connection
        Dim rs As ADODB.Recordset

        cnn.Open(Dts.Variables("CString_DB_ADMIX_IN").Value.ToString)


        strSQL = "SELECT TOP 1 CONVERT(VARCHAR,VL_CUSTO) AS VL_US FROM SULAMERICA_FAT_MOV " _
                & " WHERE 1=1 " _
                & " AND DT_COMPETENCIA = '" + Dts.Variables("Comp").Value.ToString() + "'" _
                & " AND RIGHT(REPLICATE('0',16) + CD_APOLICE,16) = RIGHT(REPLICATE('0',16) + '" + Dts.Variables("Contrato").Value.ToString() + "',16) " _
                & " AND RIGHT(REPLICATE('0',16) + CD_SUB,16) = RIGHT(REPLICATE('0',16) + '" + Dts.Variables("vSub").Value.ToString() + "',16) " _
                & " AND VL_CUSTO NOT LIKE '%0.0000%'"

        rs = cnn.Execute(strSQL)


		'grava dados do select no excel
		With xlWorkSheet.Range("I9")
			.CopyFromRecordset(rs)
		End With

		Rows = xlWorkSheet.UsedRange.Rows.Count
		With xlWorkSheet.Range("A1:Q" & Rows)
			.Font.Name = "Calibri"
			.Font.Size = 11
		End With


         rs.Close()
            cnn.Close()


-- ===================
-- BORDAS 
-- ===================

.Borders.Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin
.Borders.Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin



            .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
			
            With xlWorkSheet.Range("A1:M1")
                '.Interior.ColorIndex = 15
                .Interior.Color = RGB(0, 51, 102)
                .Font.ColorIndex = 2
                .Font.Name = "Calibri"
                .Font.Size = 10
                .Font.Bold = True
                .Borders.ColorIndex = 1
            End With

            With xlWorkSheet.Range("A1:M" & Rows)
                '.Borders.ColorIndex = 1
                .Font.Name = "Calibri"
                .Font.Size = 10
            End With

			Worksheets(1).Range("A1").Borders.LineStyle = xlDouble
			
			Worksheets("Sheet1").Range("A1:G1").Borders(xlEdgeBottom).Color = RGB(255, 0, 0)
			
            Worksheets("Sheet1").Range("B2:D4").Borders(xlInsideVertical).LineStyle = xlContinuous
            Worksheets("Sheet1").Range("B2:D4").Borders(xlInsideHorizontal).LineStyle = xlContinuous
            Worksheets("Sheet1").Range("B2:D4").Borders(xlInsideVertical).Weight = xlMedium
            Worksheets("Sheet1").Range("B2:D4").Borders(xlInsideHorizontal).Weight = xlMedium
			
			
			.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlBorderWeight.xlThick
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlBorderWeight.xlThick
            .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlBorderWeight.xlThick
            .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlBorderWeight.xlThick
			
			
			Worksheets("Sheet1").Range("B2:D4").Borders.LineStyle = xlContinuous
            Worksheets("Sheet1").Range("B2:D4").Borders.Weight = xlMedium			


-- ===================
--- LINHAS PONTILHADAS  
-- ===================

.Range("A" & i & ":H" & i).Borders.Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDot
			
			
-- ===================
-- DELETE COLUNA OU LINHA
-- ===================

xlWorkSheet.Range("L1:L" & Rows).Delete()
xlWorkSheet.Range("A1:L1").Delete()






















-- ===================
--- WHILE  
-- ===================

-- exemplo 1

Dim x As Integer

            x = 1

            While x <= Rows
                If Not xlWorkSheet.Range("C" & x).Value Is Nothing Then
                    If xlWorkSheet.Range("C" & x).Value.ToString() = "VIDA" Then
                        xlWorkSheet.Range("A2:AF" & x).Borders.ColorIndex = 1
                    End If
                End If


                x = x + 1

            End While


-- exemplo 2

        i = 2

        While i <= Rows
            If xlWorkSheet.Range("A" & i).Value.ToString() = "2" Then
                With xlWorkSheet.Range("A" & i & ":G" & i)
                    .Borders.ColorIndex = 1
                    .Font.Name = "Calibri"
                    .Font.Size = 10
                    .Interior.ColorIndex = 6
                End With
            End If

            i = i + 1

        End While


-- exemplo 3

Dim x As Integer = 1
        While x <= ROWS
       
            If Not xlWorkSheet.Range("A" & x).Value Is Nothing Then
                If xlWorkSheet.Range("A" & x).Value.ToString() = "Razão social" Then
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Font.Name = "Calibri"
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Font.Size = 10
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Font.Bold = True
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Borders.ColorIndex = 1
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Interior.Color = RGB(242, 242, 242)
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).Borders.Weight = Excel.XlBorderWeight.xlThick
                    xlWorkSheet.Range("A" & x & ":E" & x + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End If
            End If


            x = x + 1
        End While


-- ===================
--- RETORNA NOME DA SHEET  
-- ===================


  Dim plan As String

        plan = xlWorkSheet.Name.ToString

        MsgBox(plan)

        With xlWorkSheet.Range("C9")
            .CopyFromRecordset(rs)
        End With


-- ===================
--- CONVERT LINHAS  
-- ===================



        Dim A As Excel.Worksheet = CType(xlWorkbook.Worksheets(4), Excel.Worksheet)
        A.UsedRange.Formula = A.UsedRange.Formula
        A.Calculate()


        A = CType(xlWorkbook.Worksheets(4), Excel.Worksheet)
        A.UsedRange.Formula = A.UsedRange.Formula
        A.Calculate()



-- ===================
--- REMOVE COLUNA  
-- ===================

        With xlWorkSheet.Range("A1")
            .EntireColumn.Delete()
        End With


-- ===================
--- LINHAS PONTILHADAS  
-- ===================

.Range("A" & i & ":H" & i).Borders.Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDot

-- ===================
--- BORDA PADRÃO  
-- ===================


            With xlWorkSheet.Range("A1:M1")
                '.Interior.ColorIndex = 15
                .Interior.Color = RGB(0, 51, 102)
                .Font.ColorIndex = 2
                .Font.Name = "Calibri"
                .Font.Size = 10
                .Font.Bold = True
                .Borders.ColorIndex = 1
            End With

            With xlWorkSheet.Range("A1:M" & Rows)
                '.Borders.ColorIndex = 1
                .Font.Name = "Calibri"
                .Font.Size = 10
            End With


-- ===================
-- COLA DADOS NO XLS 
-- =================== 


	CString_DB_ADMIX_IN


        Dim CaminhoArquivo As String
        Dim Rows As Integer
        Dim strSQL As String
        Dim cnn As New ADODB.Connection
        Dim rs As ADODB.Recordset

        cnn.Open(Dts.Variables("CString_DB_ADMIX_IN").Value.ToString)


        strSQL = "SELECT TOP 1 CONVERT(VARCHAR,VL_CUSTO) AS VL_US FROM SULAMERICA_FAT_MOV " _
                & " WHERE 1=1 " _
                & " AND DT_COMPETENCIA = '" + Dts.Variables("Comp").Value.ToString() + "'" _
                & " AND RIGHT(REPLICATE('0',16) + CD_APOLICE,16) = RIGHT(REPLICATE('0',16) + '" + Dts.Variables("Contrato").Value.ToString() + "',16) " _
                & " AND RIGHT(REPLICATE('0',16) + CD_SUB,16) = RIGHT(REPLICATE('0',16) + '" + Dts.Variables("vSub").Value.ToString() + "',16) " _
                & " AND VL_CUSTO NOT LIKE '%0.0000%'"

        rs = cnn.Execute(strSQL)


		'grava dados do select no excel
		With xlWorkSheet.Range("I9")
			.CopyFromRecordset(rs)
		End With

		Rows = xlWorkSheet.UsedRange.Rows.Count
		With xlWorkSheet.Range("A1:Q" & Rows)
			.Font.Name = "Calibri"
			.Font.Size = 11
		End With


         rs.Close()
            cnn.Close()


-- ===================
-- BORDAS 
-- ===================

.Borders.Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin
.Borders.Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin



            .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter



-- ===================
-- DELETE COLUNA OU LINHA
-- ===================

xlWorkSheet.Range("L1:L" & Rows).Delete()
xlWorkSheet.Range("A1:L1").Delete()