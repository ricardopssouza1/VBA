Imports System
Imports System.Data
Imports System.Math
Imports Microsoft.SqlServer.Dts.Runtime
Imports System.IO
Imports ADODB

Public Class ScriptMain

	Public Sub Main()

        Dim strSQL As String
        Dim cnn As New ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim txt As String
        Dim sw As StreamWriter

        Dim versaoSQL As String
        Dim retorno As String
 
        'Setando a versão do SQL que está sendo utilizada
        cnn.Open(Dts.Variables("ConexaoString").Value.ToString())
        strSQL = "select right(left(@@VERSION,25),4)"
        rs = cnn.Execute(strSQL)
        versaoSQL = rs.Fields(0).Value.ToString()

        Select Case versaoSQL
            Case "2016"
                retorno = "2016"
            Case Else
                retorno = "2008"
        End Select

        txt = Dts.Variables("CaminhoArquivo").Value.ToString() + "\NOTIFICAÇÃO.TXT"
        sw = My.Computer.FileSystem.OpenTextFileWriter(txt, True)
        sw.WriteLine(retorno)
        sw.Close()

        'fecha conexão
        rs.Close()
        cnn.Close()
        '
        Dts.TaskResult = Dts.Results.Success
    End Sub

End Class