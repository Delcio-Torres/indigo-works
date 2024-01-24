VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3600
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3380.205
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1050
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   540
      Width           =   2325
   End
   Begin VB.TextBox txtLogin 
      Height          =   345
      Left            =   1050
      TabIndex        =   1
      Top             =   150
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   255
      TabIndex        =   3
      Top             =   1035
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1860
      TabIndex        =   4
      Top             =   1035
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Usuário:"
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   195
      Width           =   585
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Senha:"
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   5
      Top             =   585
      Width           =   510
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()

'On Error GoTo Trata_Erro
    
    Dim flag As Boolean
    Dim rt As Recordset
    Set rt = New ADODB.Recordset
    
    abreConexao
    
    rs.Open "Select * from usuario", db, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then Exit Sub
    rs.MoveFirst
    
 While Not rs.EOF
  
    If txtLogin.Text = rs!login And txtPassword.Text = rs!senha Then
       flag = True
       varCodigoUsuario = rs!codigousuario
       varTipoUsuario = rs!tipo
       varNomeUsuario = rs!nome
       
       rs!lastlogin = Format(Date, "dd/mm/yy")
       rs.update
       
    End If
    
    rs.MoveNext
 
Wend
    
If flag Then
    If varTipoUsuario = "Administrador" Then
        frmPrincipal.Label2.Caption = "Adm: " & varNomeUsuario & " - " & varCodigoUsuario
    ElseIf varTipoUsuario = "Usuário" Or varTipoUsuario = "Usuário-Ex" Then
        frmPrincipal.Label2.Caption = "Usuário: " & varNomeUsuario & " - " & varCodigoUsuario
    Else
        frmPrincipal.Label2.Caption = "Visitante: " & varNomeUsuario & " - " & varCodigoUsuario
    End If
    rs.Close
    Unload Me
    LoginSucceeded = True
Else
    MsgBox "Usuário não cadastrado.", vbCritical
    txtLogin.SetFocus
End If

End Sub

Private Sub Form_Load()
    Me.Left = frmPrincipal.Left + frmPrincipal.Width / 2 - Me.Width / 2
    Me.Top = frmPrincipal.Top + frmPrincipal.Height / 2 - Me.Height / 2
End Sub

Private Sub txtLogin_GotFocus()
    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(txtLogin.Text)
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub


