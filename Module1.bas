Attribute VB_Name = "ADO"
'faz a conex�o do BD
Public connect As ADODB.Connection
'Agente que faz a comunica��o (insert, delete, update) na tabela
Public record As ADODB.Recordset
'representa o db, guarda a string de conex�o (frase que guarda pr� configs)
Public db As String

Public Sub Conexao()
'instaciamos a conex�o
Set connect = New ADODB.Connection
'instaciamos o Recordset
Set record = New ADODB.Recordset
'
'conecta ao bd espec�fico do acess em seu devido caminho
db = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DB\DB.mdb"

'connect ao banco da vari�vel db (acess e no caminho tal)
connect.Open db
End Sub

