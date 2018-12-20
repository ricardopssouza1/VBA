
' Microsoft SQL Server Integration Services Script Task
' Write scripts using Microsoft Visual Basic
' The ScriptMain class is the entry point of the Script Task.

Imports System
Imports System.Data
Imports System.Math
Imports Microsoft.SqlServer.Dts.Runtime
Imports System.IO
Imports System.Xml
Imports Microsoft.Office.Interop
Imports ADODB

Public Class ScriptMain

    Public xlApp As Excel.Application
    Public xlWorkbook As Excel.Workbook
    Public xlWorkSheet As Excel.Worksheet


    Public Sub Main()

        Dim CaminhoArquivo As String
        Dim Rows As Integer
        Dim strSQL As String
        Dim cnn As New ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim range As Excel.Range
        Dim Obj As Object

        'identifica caminho e abre conexão
        CaminhoArquivo = "C:\Temp\ARQUIVO.xlsx"

        OpenExcel(CaminhoArquivo)

        cnn.Open("Data Source=SERVIDOR;User ID=USUARIO;Password=SENHA;Initial Catalog=BANCODEDADOS;Provider=SQLNCLI.1;Auto Translate=False; ")
		
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '                                   INSERT NO ARQUIVO EXCEL
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ActiveSheet(1)

            'Select na tabela passando variaveis competência e contrato
            strSQL = "SELECT A.CAMPO1, A.CAMPO2, A.CAMPO3 FROM TB_TABELA A "
			
            rs = cnn.Execute(strSQL)

            'Grava dados do select no excel
            With xlWorkSheet.Range("A1")
                .CopyFromRecordset(rs)
            End With

			'Conta o numero de linhas inseridas
            Rows = xlWorkSheet.UsedRange.Rows.Count

			'Formata conteúdo inserido
             With xlWorkSheet.Range("A1:C" & Rows)
                 .Font.Name = "Calibri"
                 .Font.Size = 10
             End With
                
            'Ajustes Finais
            xlWorkSheet.Cells.EntireColumn.AutoFit()
            xlWorkSheet.Cells.EntireRow.AutoFit()
            xlWorkSheet.Range("A1").Select()

    End Sub
    Public Sub OpenExcel(ByVal strCaminho As String)

        xlApp = New Excel.Application
        xlWorkbook = xlApp.Workbooks.Open(strCaminho)

    End Sub
    Public Sub ActiveSheet(ByVal intSheet As Integer)
        xlWorkSheet = CType(xlWorkbook.Worksheets(intSheet), Excel.Worksheet)
        xlWorkSheet.Activate()
    End Sub
    Public Sub SaveExcel(ByVal strCaminho As String)
        'Salva Excel
        xlApp.DisplayAlerts = False
        If (Not xlWorkbook Is Nothing) Then xlWorkbook.SaveAs(strCaminho)
        If (Not xlApp Is Nothing) Then xlApp.Quit()
    End Sub
    Public Sub CloseExcel()

        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkbook)
        releaseObject(xlWorkSheet)

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Class
		