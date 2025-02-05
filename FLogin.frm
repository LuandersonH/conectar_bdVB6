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
If Me.TUsuario.Text = "" And Me.TSenha.Text = "" Then
   MsgBox "Preencha os campos de Login e Senha", vbCritical, "Aviso!"
   Exit Sub
End If

MsgBox "teste"
End Sub

Private Sub checkSenha_Click()
If Me.checkSenha.Value Then
   Me.TSenha.PasswordChar = ""
Else
    Me.TSenha.PasswordChar = "*"
End If

End Sub
