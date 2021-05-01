
Sub sleep(Timesec)
    Set objwsh = CreateObject("WScript.Shell") 
    objwsh.Run "Timeout /T " & Timesec & " /nobreak" ,0 ,true
    Set objwsh = Nothing
End Sub


Sub ChangeValue(valor)
    PV.innerHTML = Cstr(valor)
    'MsgBox valor
End Sub


Sub LerTeste()
    Dim i
	for i = 1 to 10
		ChangeValue i
		sleep 2
	next
End Sub

Sub LerOPC_OpenOPC()

	cmd = "python leituraOPC.py"
	
	for i = 1 to 30
		set objShell = CreateObject("WScript.Shell")
		objShell.Run cmd, 0, true
		Set objShell = Nothing
		
		Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("opcvalue.txt",1)
		strFileText = objFileToRead.ReadAll()
		objFileToRead.Close
		Set objFileToRead = Nothing
		ChangeValue strFileText
		'InserirDB(strFileText)
		sleep 0.5
	next

End Sub

Sub LerOPC_QuickOPC()
    ' Create EasyOPC-DA component
    Dim Client
	Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")

	for i = 1 to 10
        valor = Client.ReadItemValue ("", "Graybox.Simulator.1", "numeric.sin.uint8")
		ChangeValue valor
		'InserirDB valor
		sleep 1
	next   
    
End Sub

Sub LerOPC()
    'LerOPC_QuickOPC
    LerOPC_OpenOPC
    'Msgbox "Ler OPC"
End Sub


Sub InserirDB(valor)
	Dim Connection
	Dim Command
	Dim SQL
End Sub


Sub Window_OnClose()
	Close Me
End Sub


Sub HelloWorld()
    MsgBox "Hello World"
End Sub

Sub Inserir2(valor)
    Dim SQL
    SQL = "INSERT INTO POINTVALUES (VALOR, TIMESTAMP1)  VALUES ("  &  valor &  ", #" &  Now & "#)"
    MsgBox SQL
End Sub


'Inserir o dado no banco de dados
Sub Inserir()

    Dim Connection
    Dim Command
    Dim SQL

    SQL = "INSERT INTO POINTVALUES (VALOR, TIMESTAMP1)  VALUES (444, #2021-05-01 13:00:00#)"

    Set Connection = CreateObject("ADODB.Connection")
    Set Command = CreateObject("ADODB.Command")

    'Abrir a conexao
    '1) USANDO-SE OLEDB
    'Connection.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=.\database.mdb;Persist Security Info=False"

    '2) USANDO-SE ODBC
    Connection.Open "DNS=MeuServidor32"

    Command.ActiveConnection = Connection
    Command.CommandText = SQL
    Command.Execute

    Set Command = Nothing

    Connection.Close
    Set Connection = Nothing

    Msgbox "Valor inserido!"

End Sub



'Ler  o dado a partir do banco de dados
Sub Ler()
    Dim Connection
    Dim RecordSet
    Dim SQL
    
    SQL = "SELECT * FROM POINTVALUES"

    Set Connection = CreateObject("ADODB.Connection")
    Set RecordSet = CreateObject("ADODB.RecordSet")    

    'Abrir a conexao
    '1) USANDO-SE OLEDB
    Connection.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=.\database.mdb;Persist Security Info=False"

    RecordSet.Open SQL, Connection

    If RecordSet.EOF Then
        MsgBox "Nenhum registro encontrado."
    Else
        Do While Not RecordSet.EOF
            MsgBox RecordSet("Valor") & " - " &  RecordSet("Timestamp1")
            RecordSet.MoveNext                        
        Loop
    End If

    RecordSet.Close
    Set RecordSet = Nothing

    Connection.Close
    Set Connection = Nothing

    'Msgbox "FIM!"

End Sub

Sub MeuTeste()
    MsgBox "Meu Teste"
End Sub

MeuTeste
'Inserir
'Ler
'Inserir2 777

