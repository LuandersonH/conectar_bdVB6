VERSION 5.00
Begin VB.Form FLogin 
   Caption         =   "Login"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "FLogin"
   ScaleHeight     =   7140
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox checkSenha 
      Caption         =   "Senha Visível"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1950
      TabIndex        =   5
      Top             =   2910
      Width           =   2850
   End
   Begin VB.CommandButton btnEntrar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   3465
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3405
      Width           =   1935
   End
   Begin VB.TextBox TSenha 
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2325
      Width           =   4950
   End
   Begin VB.TextBox TUsuario 
      Height          =   450
      Left            =   1875
      TabIndex        =   2
      Top             =   900
      Width           =   4950
   End
   Begin VB.Label LblSenha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   1
      Top             =   1875
      Width           =   750
   End
   Begin VB.Label LblUsuario 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3960
      TabIndex        =   0
      Top             =   390
      Width           =   1020
   End
End
Attribute VB_Name = "FLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEntrar_Click()
    Dim queryVerificarAcesso As String
    Dim usuarioDigitado As String
    Dim senhaDigitada As String

   'Tira espaços em branco do login e senha
    HandleUsuarioLogin = Trim(TUsuario.Text)
    HandleSenhaLogin = Trim(TSenha.Text)
    
    If HandleUsuarioLogin = "" Or HandleSenhaLogin = "" Then
        MsgBox "Por favor, preencha os campos de usuário e senha", vbExclamation, "Aviso"
        Exit Sub
    End If
      
   ' Faço minha query e depois consulto no banco (conexão feita no arquivo ADO)
   queryVerificarAcesso = "SELECT * FROM tb_login WHERE usuario = '" & HandleUsuarioLogin & "' AND senha = '" & HandleSenhaLogin & "'"
   
   ' queryVerificarAcesso = comando SQL que irá ser executado no banco de dados
   ' connect = variável que guarda a conexão ADO, que já está configurada com o tipo de conexão e caminho do banco de dados Access
   ' adOpenStatic = tipo de cursor, que indica que os dados lidos não serão atualizados automaticamente após a consulta; é um cursor "estático"
   ' adLockReadOnly = tipo de bloqueio, indicando que os dados são somente para leitura, ou seja, não podem ser modificados
   record.Open queryVerificarAcesso, connect, adOpenStatic, adLockReadOnly

      
    'Verifica se deu certo o login
    If Not record.EOF Then
        MsgBox "Login realizado com sucesso!", vbInformation, "Sucesso"
        Form1.Show
        Me.Visible = False
    Else
        MsgBox "Usuário ou senha inválidos.", vbCritical, "Erro"
    End If

   'Fecha para não ocupar memória
    record.Close
End Sub

'Função para exibir ou não oq é digitado na senha
Private Sub checkSenha_Click()
If Me.checkSenha.Value Then
   Me.TSenha.PasswordChar = ""
Else
    Me.TSenha.PasswordChar = "*"
End If

End Sub

Private Sub Form_Load()
Call Conexao
End Sub
