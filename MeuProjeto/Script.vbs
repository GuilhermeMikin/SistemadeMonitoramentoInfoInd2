Sub Inserir()

    Dim Connection
    Dim Command
    Dim SQL

        SQL = "INSERT INTO ALUNOS(NOME, MAT) VALUES ('Guilherme Balduino Lopes', '11511EAU011')"
        Set Connection = CreateObject("ADODB.Connection")
        Set Command = CreateObject("ADODB.Command")

        Connection.Open "DSN=MeuServidor32"
        
        Command.ActiveConnection = Connection
        Command.CommandText = SQL
        Command.Execute
        Set Command = Nothing
        Connection.Close
        Set Connection = Nothing

End Sub


Sub Ler()

    Dim Connection
    Dim Recordset
    Dim SQL
    
        SQL = "SELECT * FROM ALUNOS"
        Set Connection = CreateObject("ADODB.Connection")
        Set Recordset = CreateObject("ADODB.Recordset")
        Connection.Open "DSN=MeuServidor32"
        Recordset.Open SQL, Connection
        If Recordset.EOF then
        MsgBox "Nenhum registro encontrado."
        Else
        Do While Not Recordset.EOF
        Msgbox Recordset("NOME")
        Recordset.MoveNext
        Loop
        End If
        Recordset.Close
        Set Recordset = Nothing
        Connection.Close
        Set Connection = Nothing

End Sub

' Sub Fler OPC()
'     Dim Client

'     Ser Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDACliente")

'     valor = client.ReadItemValue("", "Graybox.Simulator.1", "numeric.sin.float")

' 	MsgBox valor

Sub GravarDado(valor)
    
    SQL = "INSERT INTO POINTVALUES(VALOR) VALUES (" & valor &")"

End Sub

Sub MeuTeste()
    MsgBox "Meu teste"

End Sub

MeuTeste
'Inserir
'Ler