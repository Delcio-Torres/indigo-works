VERSION 5.00
Begin VB.Form frmPrincipal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Indigo Works"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8520
   Icon            =   "frmPrincipal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   4320
      TabIndex        =   9
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Relatório &OS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   4320
      TabIndex        =   8
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Expedição"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   1800
      TabIndex        =   7
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Relatório Vendas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Administração"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Orçamento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   5640
         Picture         =   "frmPrincipal.frx":08CA
         ScaleHeight     =   1095
         ScaleWidth      =   2895
         TabIndex        =   5
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Indigo Works 3.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   540
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3600
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Usuário: Cique aqui ou  tecle F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5475
      MouseIcon       =   "frmPrincipal.frx":12C0
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5520
      Width           =   2820
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub Command1_Click(Index As Integer)

   If Index <> 5 Then
      If varCodigoUsuario = 0 Then
         frmLogin.Show 1
      End If
   
      If Not frmLogin.LoginSucceeded Then
         Exit Sub
      End If
   End If
    
   Select Case Index
   
      Case 0
         frmPesquisaOs.Show 1
      Case 1, 3
   
         If varTipoUsuario = "Administrador" Then
            If Index = 1 Then frmAdministracao.Show 1
            If Index = 3 Then frmRelatorioOs.Show 1
         Else
            MsgBox "Área resrita para administrador!", vbCritical
         End If
      Case 2
         frmRelatorioVendas.Show 1
      Case 4
      
         frmExpedicao.Show 1
      
      Case 5
         Unload Me
   
   End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyCancel Then
        Unload Me
    ElseIf KeyCode = vbKeyF3 Then
        frmLogin.Show 1
    End If
    
End Sub

Private Sub Form_Load()

update

Dim cap As String

    If App.PrevInstance Then
        cap = Me.Caption
        Me.Caption = ""
        AppActivate cap
        Unload Me
    
    End If
    
    abreConexao
    rs.Open "SELECT * FROM usuario", db, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        MsgBox "Está é a primeira vez que o programa é executado." & Chr(13) _
                & "Cadastre um administrador agora.", vbInformation
        rs.Close
        db.Close
        frmAdministracao.Show 1
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Label2.ForeColor = &H0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        End
        'Unload frm
        'Unload frmPesquisaOs
        'Unload frmPrevisaodeentrega
        'Unload frmPesquisaCliente
        'Unload frmListaOrcamento
        'Unload frmOrcamento
        'Unload frmSplash
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Label2.ForeColor = &HFF0000
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Label2.ForeColor = &H0
    frmLogin.Show 1
End Sub
