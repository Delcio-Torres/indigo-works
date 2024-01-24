VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPesquisaCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pesquisa de Clientes"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13485
   Icon            =   "frmPesquisaCliente.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
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
      TabIndex        =   2
      Top             =   5400
      Width           =   2295
   End
   Begin MSComctlLib.StatusBar stbUsuario 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   5895
      Width           =   13485
      _ExtentX        =   23786
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19279
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
      Caption         =   "&OK"
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
      TabIndex        =   1
      Top             =   5400
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid grdCliente 
      Height          =   4935
      Left            =   75
      TabIndex        =   3
      Top             =   120
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   17
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   -2147483646
      ScrollTrack     =   -1  'True
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
      Caption         =   "Pesquisa nome do cliente:"
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
      Width           =   2370
   End
End
Attribute VB_Name = "frmPesquisaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varOldCriterio As String

Public Sub cmdEditar_Click()
    
    If grdCliente.TextMatrix(grdCliente.RowSel, 1) = "" Then
        Exit Sub
    End If
    
    If grdCliente.TextMatrix(grdCliente.RowSel, 4) = "Bloqueado" Then
        'MsgBox "Verificar o cadastro do cliente: '" & grdUsuario.TextMatrix(grdUsuario.RowSel, 1) & "' na administração.", vbCritical
    End If
    
    preencheCampoCliente grdCliente.TextMatrix(grdCliente.RowSel, 1)
    Unload Me
    frmOrcamento.txtOrcamento(1).SetFocus
    frmOrcamento.txtCliente(0).Locked = True
    frmOrcamento.txtTelefone.Enabled = False
    frmOrcamento.txtCliente(1).Enabled = False
    frmOrcamento.txtCliente(2).Enabled = False
    frmOrcamento.txtCPF.Enabled = False
    frmOrcamento.Option1(0).Enabled = False
    frmOrcamento.Option1(1).Enabled = False
    'frmOrcamento.optPagamento(4).Value = True
    'frmOrcamento.Frame9.Enabled = False
    
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    abreConexao
    rs.Open "SELECT nome, idcliente, telefone, cnpj, condicao, contato FROM cliente ORDER BY nome", db, adOpenStatic, adLockOptimistic
    rs.MoveFirst

    If rs.RecordCount = 0 Then
        Unload Me
    End If
    rs.Close
    
    preencheGridCliente

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    db.Close
End Sub

Private Sub grdUsuario_DblClick()

    cmdEditar_Click

End Sub

Private Sub grdUsuario_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        stbUsuario.Panels(2).Text = grdCliente.TextMatrix(grdCliente.Row, 1) & "-" & grdCliente.RowSel
    End If
End Sub

Public Sub preencheGridCliente()

With grdCliente

    rs.Open "SELECT nome, idcliente, telefone, cnpj, condicao FROM cliente WHERE nome LIKE '%" & Me.Text1.Text & "%' ORDER BY nome", db, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        stbUsuario.Panels(2).Text = "Não foi encontrado nenhum registro!!!"
        grdCliente.Clear
        
        .FormatString = "<Nome|^Código|<Telefone|>CNPJ/CPF|^Condiçao"
    
        .Rows = rs.RecordCount + 1
        .Cols = rs.Fields.Count
        .ColWidth(0) = 4585
        .ColWidth(1) = 1000
        .ColWidth(2) = 1500
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .Rows = 17
        .Col = 0
        .Row = 1
        .ColSel = .Cols - 1
        rs.Close
        Exit Sub
    Else
        stbUsuario.Panels(2).Text = ""
        rs.MoveFirst
    End If

    Screen.MousePointer = vbHourglass
    '.Row = 1
    .Clear
    .Refresh

    .Visible = False

    .FormatString = "<Nome|^Código|<Telefone|>CNPJ/CPF|^Condiçao"

    .Rows = rs.RecordCount + 1
    .Cols = rs.Fields.Count
    .ColWidth(0) = 4585
    .ColWidth(1) = 1000
    .ColWidth(2) = 1500
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
    '.Row = 1
    .Col = 0
    .RowSel = .Rows - 1
    .ColSel = .Cols - 1

    rs.MoveFirst
    'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
    .Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    .Visible = True
    
    If rs.RecordCount < 16 Then .Rows = 17
    
    .TopRow = 1
    'Para selecionar toda a linha...
    .Row = .TopRow
    .Col = 0
    .ColSel = .Cols - 1
End With

    Screen.MousePointer = vbDefault
rs.Close

End Sub

Private Sub Text1_Change()
    
    preencheGridCliente
    
    
End Sub
Private Function preencheCampoCliente(codigo As Single)
    abreConexao
    
    rs.Open "SELECT idcliente, nome, endereco, bairro, telefone, cnpj, condicao FROM cliente WHERE idcliente=" & codigo, db, adOpenStatic, adLockOptimistic
    'rs.Open "SELECT * FROM cliente", db, adOpenStatic, adLockOptimistic
    
With frmOrcamento
    .txtCliente(0).Text = rs!nome
    .txtCliente(1).Text = rs!endereco
    .txtCliente(2).Text = rs!bairro
    .txtTelefone.Text = rs!telefone
    .txtCodigoCliente.Text = rs!idcliente
    If Len(rs!CNPJ) = 11 Then
        .Option1(0).Value = True
    Else
        .Option1(1).Value = True
    End If
    .txtCPF = rs!CNPJ

    If rs!condicao = "Bloqueado" Then
        .txtCliente(0).BackColor = &HC0FFFF
        .txtCliente(1).BackColor = &HC0FFFF
        .txtCliente(2).BackColor = &HC0FFFF
        .txtTelefone.BackColor = &HC0FFFF
        .txtCPF.BackColor = &HC0FFFF
    Else
        .txtCliente(0).BackColor = &HFFFFC0
        .txtCliente(1).BackColor = &HFFFFC0
        .txtCliente(2).BackColor = &HFFFFC0
        .txtTelefone.BackColor = &HFFFFC0
        .txtCPF.BackColor = &HFFFFC0
    End If

End With

    rs.Close
    
End Function
