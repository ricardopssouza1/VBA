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
-- INSERE DADOS DO BD NO XLS 
-- =================== 

        Dim Rows As Integer
        Dim strSQL As String
        Dim cnn As New ADODB.Connection
        Dim rs As ADODB.Recordset

        cnn.Open(Dts.Variables("CString_DB_IN").Value.ToString)


        strSQL = "SELECT TOP 1 VL_CUSTO AS VL_US FROM TABELA " _
                & " WHERE 1=1 " _
                & " AND DT_COMPETENCIA = '" + Dts.Variables("Comp").Value.ToString() + "'"

        rs = cnn.Execute(strSQL)

		'grava dados do select no excel
		With xlWorkSheet.Range("I9")
			.CopyFromRecordset(rs)
		End With

		'lê quantas linhas foram inseridas
		Rows = xlWorkSheet.UsedRange.Rows.Count
		
		'formata cabeçalho
		With xlWorkSheet.Range("A1:C1")
			.Font.Name = "Calibri"
			.Font.Size = 11
		End With

		'fecha conexões
         rs.Close()
         cnn.Close()


-- ===================
-- BORDAS 
-- ===================

.Borders.Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin
.Borders.Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin


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
-- DELETE COLUNA OU LINHA
-- ===================

xlWorkSheet.Range("L1:L" & Rows).Delete()
xlWorkSheet.Range("A1:L1").Delete()

-- ===================
--- REMOVE COLUNA  
-- ===================

        With xlWorkSheet.Range("A1")
            .EntireColumn.Delete()
        End With

-- ===================
-- VETOR
-- ===================

            Dim vetor() As String
            Dim I As String

            vetor = Split("A2,B2,C2,D2,E2,F2,G2,H2,I2,J2,K2,L2,M2,N2,O2,P2,Q2,R2,S2,T2,U2,V2,W2,X2,Y2,Z2", ",")

            For Each I In vetor
                xlWorkSheet.Range(I.ToString()).Borders.Item(Excel.XlBordersIndex.xlEdgeRight).Color = RGB(255, 255, 255)
                xlWorkSheet.Range(I.ToString()).Borders.Item(Excel.XlBordersIndex.xlEdgeLeft).Color = RGB(255, 255, 255)
            Next
			
			
-- ===================
-- FORMATAÇÃO MOEDA
-- ===================			
			
            With xlWorkSheet.Range("W2:W" & Rows - 1)
                .NumberFormat = "$#,##0.00"
            End With		


-- ===================
-- ALINHAMENTO
-- ===================

            With xlWorkSheet.Range("A1")
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter				
            End With			
			
			
-- ===================
-- FORMULA
-- ===================			
			
        ' adicionando formula
        xlWorkSheet.Range("C2").Formula = "=SOMA(A2+B2)"			
		
		
-- ===================
-- ARQUIVO DE ERRO 
-- ===================			
		
    Public Sub Main()

        ' Tratamento de erro
        On Error GoTo aviso

        '
		' script
		'

        'Armazena arquivo TXT com a msg de erro
aviso:  If Err.Number <> 0 Then
            txt = "C:\ARQUIVO_NOTIFICAÇÃO.TXT"
            sw = My.Computer.FileSystem.OpenTextFileWriter(txt, True)
            sw.WriteLine("Erro número #" & Str$(Err.Number) & " na Linha " & Str$(Erl) & " - " & Err.Description & " - Gerado por " & Err.Source)
            sw.Close()
        End If

		
    End Sub		

-- ===================
-- EXECUTA OUTRO DTS
-- ===================	

        Dim pkgLocation As String
        Dim pkg As New Package
        Dim app As New Application
        Dim pkgResults As DTSExecResult

		pkgLocation = "C:\Arquivos.dtsx"		
		
        pkg = app.LoadPackage(pkgLocation, Nothing)

		'parametros da dts
        pkg.Variables("Comp").Value = Dts.Variables("Comp").Value.ToString()
        pkg.Variables("CString_DB_IN").Value = Dts.Variables("CString_DB_IN").Value.ToString()
        pkg.Variables("Contrato").Value = Dts.Variables("Contrato").Value.ToString()
        pkg.Variables("Empresa").Value = Dts.Variables("Empresa").Value.ToString()
        pkg.Variables("Destino").Value = Dts.Variables("Destino").Value.ToString()

        pkgResults = pkg.Execute()		
	
			
-- ===================
-- DIVERSOS
-- ===================

'https://bettersolutions.com/vba/macros/sendkeys.htm
'https://ferramentasexcelvba.wordpress.com/2018/05/11/send-keys-no-vba/
'https://github.com/OfficeDev/VBA-content/blob/master/VBA/Excel-VBA/articles/application-sendkeys-method-excel.md

Send Keys no VBA
'Um comando que pode ser muito útil, especialmente para automatizar processos repetitivos, é o “SendKeys”. Este comando simplesmente emula o teclado.
'A sintaxe é muito simples, algo como:

Application.SendKeys "Bom dia!", True

'Para enviar um “Enter”, utilizar o símbolo “~”:

Application.SendKeys "~", True

'Exemplo:
'Esta rotina vai abrir o Bloco de Notas, esperar um segundo, dar um Enter e escrever “Bom dia!”.

Shell "NotePad.exe", vbMaximizedFocus

Application.Wait DateTime.Now + DateTime.TimeValue(“00:00:01”)

Application.SendKeys “~”, True 'Enter

Application.SendKeys “Bom dia!”, True

xlWorkSheet.Application.SendKeys("~", True)


'outro exemplo
'https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/sendkeys-statement

Dim ReturnValue, I 
ReturnValue = Shell("CALC.EXE", 1)    ' Run Calculator. 
AppActivate ReturnValue     ' Activate the Calculator. 
For I = 1 To 100    ' Set up counting loop. 
    SendKeys I & "{+}", True    ' Send keystrokes to Calculator 
Next I    ' to add each value of I. 
SendKeys "=", True    ' Get grand total. 
SendKeys "%{F4}", True    ' Send ALT+F4 to close Calculator. 



			