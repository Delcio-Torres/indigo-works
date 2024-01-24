VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmListaOrcamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13065
   Icon            =   "frmListaOrcamento.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisa cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   6360
      Width           =   3855
      Begin VB.TextBox Text1 
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   3375
      End
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
      Left            =   7560
      TabIndex        =   3
      Top             =   6720
      Width           =   2295
   End
   Begin MSComctlLib.StatusBar stbServicos 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   7425
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3563
            MinWidth        =   3563
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19403
            MinWidth        =   19403
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   11113
      _Version        =   393216
      Rows            =   18
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   -2147483636
      Redraw          =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      DisabledPicture =   "frmListaOrcamento.frx":0ECA
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
      Left            =   10080
      TabIndex        =   1
      Top             =   6720
      Width           =   2295
   End
End
Attribute VB_Name = "frmListaOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub mostraRegistro(numeroOs As Integer)

'On Error GoTo Trata_Erro
    
    abreConexao
    
    rs.Open "Select * from servicos where numeroos=" & numeroOs, db, adOpenStatic, adLockOptimistic
 
If rs.RecordCount = 0 Then Exit Sub

 
 With Form1
    .Caption = "Oraçamento Nº " & rs!numeroOs
    .txtNumero = rs!numeroOs
    .Text1.Text = rs!nome
    .MaskEdBox1 = rs!telefone
    .Label8 = rs!Data
    .Label9 = rs!hora
    .Text5 = rs!quantidade
    .text7.Text = rs!descricao

    .optMidia_Click (rs!midia)
    
    .optMidia(rs!midia).Value = True
    .optGramatura(rs!gramatura).Value = True
    .Option2(rs!cores).Value = True
    .Text6 = rs!quantlaminacao
    .Option1(rs!tipolaminacao).Value = True
    .Text8 = rs!quantcapa
    .Option3(rs!tipocapa).Value = True
    .Text4 = rs!quantwireo
    .Option4(rs!tipowireo).Value = True
    .Check1.Value = rs!embutir
    .Text10 = rs!corte
    .Text9 = rs!meiocorte
    .Text11 = rs!transfer
End With
    varModoEdicao = True
    Unload Me
    
Exit Sub

'Trata_Erro:
    'MsgBox "houve um erro ai - " & Err

End Sub

Private Sub Command1_Click()
'Para rolar para a linha 35...
Grid.TopRow = 10

'Para selecionar toda a linha...
Grid.Row = Grid.TopRow
Grid.Col = 0
Grid.ColSel = Grid.Cols - 1
End Sub



Private Sub Form_Load()
    preencheGridServicos
End Sub

Private Sub Grid_DblClick()
    mostraRegistro (Grid.TextMatrix(Grid.Row, 0))
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mostraRegistro (Grid.TextMatrix(Grid.Row, 0))
    End If
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbEnter Then
    'End If
End Sub


Private Sub Text1_Change()
    rs.Close
    rs.Open "SELECT numeroos, nome, descricao, telefone FROM servicos WHERE nome LIKE '%" & Me.Text1 & "%' ORDER BY nome", db, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        Grid.TextMatrix(1, 0) = "Número"
        stbServicos.Panels(2).Text = "Não foi encontrado nenhum registro!!!"
        Grid.Clear
        Exit Sub
    Else
        stbServicos.Panels(2).Text = ""
        rs.MoveFirst
    End If
    
    preencheGridServicos
End Sub

Public Sub preencheGridServicos()

abreConexao

rs.Open "SELECT numeroos, nome, descricao, telefone FROM servicos", db, adOpenStatic, adLockReadOnly



Screen.MousePointer = vbHourglass
Grid.Row = 1
Grid.Clear
Grid.Refresh

Grid.Visible = False

Grid.FormatString = "^Número|<Cliente|<Descrição|<Vendedor"


Grid.Rows = 17
Grid.Cols = rs.Fields.Count
Grid.ColWidth(0) = 1000
Grid.ColWidth(1) = 5000
Grid.ColWidth(2) = 5000
Grid.Row = 1
Grid.RowSel = Grid.Rows - 1
Grid.ColSel = Grid.Cols - 1

rs.MoveFirst
'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
Grid.Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
Grid.Visible = True


'Para selecionar toda a linha...
Grid.TopRow = Grid.Rows - 1
Grid.Row = Grid.TopRow
Grid.Col = 0
Grid.ColSel = Grid.Cols - 1

Screen.MousePointer = vbDefault
    If Grid.TextMatrix(1, 0) = "" Then
        Grid.TextMatrix(0, 0) = "Número"
    End If

End Sub
