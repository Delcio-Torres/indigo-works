VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPesquisaUsuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pesquisa de usuário"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "frmPesquisaOrcamento.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   5400
      Width           =   2295
   End
   Begin MSComctlLib.StatusBar stbUsuario 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   5895
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Del - Remove usuário"
            TextSave        =   "Del - Remove usuário"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14993
            MinWidth        =   14993
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   5400
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid grdUsuario 
      Height          =   4935
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   17
      Cols            =   5
      FixedCols       =   0
      BackColorSel    =   -2147483646
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisa nome do usuário:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   2445
   End
End
Attribute VB_Name = "frmPesquisaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varOldCriterio As String


Public Sub cmdEditar_Click()

    If grdUsuario.TextMatrix(grdUsuario.RowSel, 1) = "" Then
        Exit Sub
    End If
        
    frmAdministracao.preencheControleUsuario grdUsuario.TextMatrix(grdUsuario.RowSel, 1)

    Unload Me
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    abreConexao
    rs.Open "SELECT nome, codigousuario, telefone, celular, tipo FROM usuario ORDER BY nome", db, adOpenStatic, adLockOptimistic
    rs.MoveFirst
    preencheGridUsuario
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me

    rs.Close
    db.Close
End Sub


Private Sub grdUsuario_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        stbUsuario.Panels(2).Text = grdUsuario.TextMatrix(grdUsuario.Row, 1) & "-" & grdUsuario.RowSel
    End If
End Sub

Public Sub preencheGridUsuario()
        
With grdUsuario
    rs.Update
    Screen.MousePointer = vbHourglass
    .Row = 1
    .Clear
    .Refresh

    .Visible = False
    
    .FormatString = "<Nome|^Código|<Telefone|<Celular|<Tipo"

    'define o numero de linhas e colunas e configura o grid
    
    .Rows = rs.RecordCount + 17 - rs.RecordCount
    .Cols = rs.Fields.Count
    .ColWidth(0) = 4585
    .ColWidth(1) = 1000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 2000
    .Row = 1
    .Col = 0
    .RowSel = .Rows - 1
    .ColSel = .Cols - 1
    
    rs.MoveFirst
    'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
    .Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    .Visible = True
     
    .TopRow = .Rows - 16
    'Para selecionar toda a linha...
    .Row = .TopRow
    .Col = 0
    .ColSel = .Cols - 1
End With
    
    Screen.MousePointer = vbDefault
End Sub




Private Sub Text1_Change()
    rs.Close
    rs.Open "SELECT nome, codigousuario, telefone, celular, tipo FROM usuario WHERE nome LIKE '%" & Me.Text1 & "%' ORDER BY nome", db, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        stbUsuario.Panels(2).Text = "Não foi encontrado nenhum registro!!!"
        grdUsuario.Clear
        Exit Sub
    Else
        stbUsuario.Panels(2).Text = ""
        rs.MoveFirst
    End If
    
    preencheGridUsuario
End Sub


