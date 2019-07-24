Imports System
Imports System.Data
Imports System.Math
Imports Microsoft.SqlServer.Dts.Runtime
Imports Microsoft.Office.Interop
Imports ADODB
Imports System.IO
Imports System.Data.SqlClient


<Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute()> _
<System.CLSCompliantAttribute(False)> _
Partial Public Class ScriptMain
    Inherits Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase


    Public xlApp As Excel.Application
    Public xlWorkbook As Excel.Workbook
    Public xlWorkSheet As Excel.Worksheet

    Public Sub Main()

        Dim CaminhoArquivo As String
        Dim Rows As Integer
        Dim strSQL As String
        Dim cnn As New ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim StringConexao As String

        Dim txt As String
        Dim sw As StreamWriter

        StringConexao = "Data Source=LOCALHOST;User ID=sa;Password=SQL123;Initial Catalog=OLTP;Provider=SQLNCLI11.1;Auto Translate=False;"
        CaminhoArquivo = "C:\DTS\Arquivo.xls"
        OpenExcel(CaminhoArquivo)
        cnn.Open(StringConexao)

        xlApp.DisplayAlerts = False

        'grava dados do select no excel
        ActiveSheet(1)

        strSQL = "" _
        & vbNewLine & " SELECT  a.ID_ESCOLA " _
        & vbNewLine & "       , a.REGIAO " _
        & vbNewLine & "       , a.CODMUN " _
        & vbNewLine & " 	  , ISNULL(b.DS_MUNICIPIO,'NAO CADASTRADO') AS DS_MUNICIPIO " _
        & vbNewLine & "       , a.ESTRATOGEO " _
        & vbNewLine & "       , a.CAPITAL " _
        & vbNewLine & "       , a.UPA " _
        & vbNewLine & "       , a.PESO_ESCOLA " _
        & vbNewLine & "       , a.PUBPRIV " _
        & vbNewLine & "       , a.DEPENDADM " _
        & vbNewLine & "    FROM [dbo].[TMP_ESCOLAS] a " _
        & vbNewLine & "    LEFT JOIN dbo.TMP_MUNICIPIOS b on b.CD_IBGE = a.CODMUN "
        rs = cnn.Execute(strSQL)

        With xlWorkSheet.Range("A2")
            .CopyFromRecordset(rs)
        End With

        Rows = xlWorkSheet.UsedRange.Rows.Count

        With xlWorkSheet.Range("A1:J" & Rows)
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        End With

        'grava dados do select no txt
        strSQL = "" _
        & vbNewLine & " SELECT  a.ID_ESCOLA " _
        & vbNewLine & "       +';'+ a.REGIAO " _
        & vbNewLine & "       +';'+ a.CODMUN " _
        & vbNewLine & " 	  +';'+ ISNULL(b.DS_MUNICIPIO,'NAO CADASTRADO') " _
        & vbNewLine & "       +';'+ a.ESTRATOGEO " _
        & vbNewLine & "       +';'+ a.CAPITAL " _
        & vbNewLine & "       +';'+ a.UPA " _
        & vbNewLine & "       +';'+ a.PESO_ESCOLA " _
        & vbNewLine & "       +';'+ a.PUBPRIV " _
        & vbNewLine & "       +';'+ a.DEPENDADM AS CAMPO" _
        & vbNewLine & "    FROM [dbo].[TMP_ESCOLAS] a " _
        & vbNewLine & "    LEFT JOIN dbo.TMP_MUNICIPIOS b on b.CD_IBGE = a.CODMUN "
        rs = cnn.Execute(strSQL)

txt = "C:\DTS\Arquivo_Texto.txt"
        sw = My.Computer.FileSystem.OpenTextFileWriter(txt, True)
        sw.WriteLine(rs.GetString())
        sw.Close()

        'AJUSTES FINAIS AUTOFIT
        xlWorkSheet.Cells.EntireColumn.AutoFit()
        xlWorkSheet.Cells.EntireRow.AutoFit()
        xlWorkSheet.Range("A1").Select()

        'fecha conex�o
        rs.Close()
        cnn.Close()

        'SaveExcel(CaminhoArquivo2)
        xlWorkbook.Close(True)

        CloseExcel()

        Dts.TaskResult = ScriptResults.Success
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


    Enum ScriptResults
        Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success
        Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
    End Enum



End Class
