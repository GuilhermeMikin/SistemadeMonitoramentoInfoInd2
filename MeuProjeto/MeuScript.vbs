Sub sleep(Timesec)
    Set objwsh = CreateObject("WScript.Shell") 
    objwsh.Run "Timeout /T " & Timesec & " /nobreak" ,0 ,true
    Set objwsh = Nothing
End Sub

Sub ChangeValue(valor)
    PV.innerHTML = Cstr(valor)
End Sub

Sub LerOPC_QuickOPC()
    ' Create EasyOPC-DA component
	Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")

	for i = 1 to 2
        valor = Client.ReadItemValue("", "Graybox.Simulator.1", "numeric.sin.float")
		ChangeValue valor
		InserirDB valor
		sleep 1
	next   
End Sub

Sub LerOPC()
    VL.innerHTML = Cstr(" ")
    LerOPC_QuickOPC
    VL.innerHTML = Cstr("Valores lidos e inseridos no Banco de Dados com sucesso!!")
End Sub

Sub InserirDB(valor)
	Dim Connection
    Dim Command
    Dim SQL
    
    cmd = "python pyScriptDB.py"
    set objShell = CreateObject("WScript.Shell")
	objShell.Run cmd, 0, true
	Set objShell = Nothing

    Set objDateToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("datevalue.txt",1)
	strDateText = objDateToRead.ReadAll()
	objDateToRead.Close
	Set objDateToRead = Nothing

    SQL = "INSERT INTO POINTVALUES (VALOR, TIMESTAMP1) VALUES ("& valor &", '"& strDateText &"')"

    Set Connection = CreateObject("ADODB.Connection")
    Set Command = CreateObject("ADODB.Command")

    'Abrindo a conexão com ODBC SQLite3:
    Connection.Open "DSN=SQLiteODBCServer"
    Command.ActiveConnection = Connection
    Command.CommandText = SQL
    Command.Execute

    Set Command = Nothing

    Connection.Close
    Set Connection = Nothing

End Sub

