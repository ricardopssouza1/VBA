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
