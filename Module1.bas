Attribute VB_Name = "ADO"
'faz a conexão do BD
Public connect As ADODB.Connection
'Agente que faz a comunicação (insert, delete, update) na tabela
Public record As ADODB.Recordset
'representa o db, guarda a string de conexão (frase que guarda pré configs)
Public db As String

Public Sub Conexao()
'instaciamos a conexão
Set connect = New ADODB.Connection
'instaciamos o Recordset
Set record = New ADODB.Recordset
'
'conecta ao bd específico do acess em seu devido caminho
db = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DB\DB.mdb"

'connect ao banco da variável db (acess e no caminho tal)
connect.Open db
End Sub

