Imports System
Imports System.Data
Imports System.Math
Imports Microsoft.SqlServer.Dts.Runtime
Imports Microsoft.Office.Interop
Imports System.Xml
Imports ADODB
Imports System.IO

Public Class ScriptMain

    Public xlApp As Excel.Application
    Public xlWorkbook As Excel.Workbook
    Public xlWorkSheet As Excel.Worksheet

    Public Sub Main()

        Dim CaminhoArquivo As String
        Dim Rows As Integer
        Dim a As Excel.Worksheet
        Dim strSQL As String
        Dim cnn As New ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim rCnt As Integer
        Dim cCnt As Integer
        Dim txt As String
        Dim sw As StreamWriter

        CaminhoArquivo = Dts.Variables("Caminho").Value.ToString()
        OpenExcel(CaminhoArquivo)

        '--------------------------------------------------------------
        ' exemplo utilizando vetor
        '--------------------------------------------------------------

        ActiveSheet(1)
        xlApp.DisplayAlerts = False

        Dim vetor() As String
        Dim I As String

        'pintando celula de VERMELHO
        vetor = Split("A1,C1,G1,B2,D2", ",")

        For Each I In vetor
            If (I.ToString() Like "*1*") Then
                Dim t As String = I.ToString().Replace("1", "")
                'MsgBox(t + " vermelho")
                xlWorkSheet.Range(t.ToString() & "1").Interior.Color = RGB(255, 0, 0)
            ElseIf (I.ToString() Like "*2*") Then
                Dim t As String = I.ToString().Replace("2", "")
                'MsgBox(t + " amarelo")
                xlWorkSheet.Range(t.ToString() & "2").Interior.Color = RGB(255, 255, 0)
            End If
        Next

        'pintando celula de AMARELO
        vetor = Split("B1,D1,E1", ",")

        For Each I In vetor
            xlWorkSheet.Range(I.ToString()).Interior.Color = RGB(255, 255, 0)
        Next

        'pintando celula de CINZA
        vetor = Split("F1,E1,H1", ",")

        For Each I In vetor
            xlWorkSheet.Range(I.ToString()).Interior.Color = RGB(191, 191, 191)
        Next

        'AJUSTES FINAIS AUTOFIT
        xlWorkSheet.Cells.EntireColumn.AutoFit()
        xlWorkSheet.Cells.EntireRow.AutoFit()
        xlWorkSheet.Range("A1").Select()

        'SaveExcel(CaminhoArquivo2)
        xlWorkbook.Close(True)

        CloseExcel()

    End Sub
