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