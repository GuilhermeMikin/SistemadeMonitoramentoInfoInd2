
Sub sleep(Timesec)
    Set objwsh = CreateObject("WScript.Shell") 
    objwsh.Run "Timeout /T " & Timesec & " /nobreak" ,0 ,true
    Set objwsh = Nothing
End Sub


Sub ChangeValue(valor)
    MsgBox valor
End Sub


Sub LerOPC_QuickOPC()
    ' Create EasyOPC-DA component
    Dim Client
	Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")

	for i = 1 to 10
        valor = Client.ReadItemValue("", "Graybox.Simulator.1", "numeric.sin.uint8")
		ChangeValue valor
		'InserirDB valor
		sleep 2
	next   
    
End Sub

Sub MeuTeste()
    MsgBox "Meu Teste"
End Sub

MeuTeste
'LerOPC_QuickOPC


