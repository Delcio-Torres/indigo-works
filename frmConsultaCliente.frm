VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConsultaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa Cliente"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   12405
   Icon            =   "frmConsultaCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
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
      Left            =   10200
      TabIndex        =   3
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Abrir OS"
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
      Left            =   8040
      TabIndex        =   2
      Top             =   6240
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6765
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21819
            MinWidth        =   176
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
   Begin MSFlexGridLib.MSFlexGrid grdCliente 
      Height          =   6090
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10742
      _Version        =   393216
      Rows            =   17
      Cols            =   6
      FixedCols       =   0
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
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
End
Attribute VB_Name = "frmConsultaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

   If grdCliente.TextMatrix(grdCliente.RowSel, 0) <> "" Then frmPesquisaOs.preencheCliente (grdCliente.TextMatrix(grdCliente.RowSel, 0))

End Sub

Private Sub Command2_Click()
   Unload Me
End Sub
Public Sub preencheGridCliente(Optional criterio As String)

With frmConsultaCliente.grdCliente

   Screen.MousePointer = vbHourglass
   abreConexao

   Dim nomeCliente As String
   Dim idcliente As Long
   Dim rd As Recordset
   Set rd = New ADODB.Recordset
   
   If frmPesquisaOs.Option1(0).Value = True Then
      rs.Open "SELECT os.idos, os.nomeCliente, os.data, os.hora, usuario.nome, os.baixa FROM OS INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE os.idos =" & criterio & " ORDER BY idos", db, adOpenForwardOnly, adLockReadOnly
   Else
      rs.Open "SELECT os.idos, os.nomeCliente, os.data, os.hora, usuario.nome, os.baixa FROM OS INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE nomeCliente LIKE" & "'%" & criterio & "%'" & " ORDER BY idos", db, adOpenForwardOnly, adLockReadOnly
   End If
        
   If rs.RecordCount = 0 Then
      .Clear
      .Row = 1
      .ColSel = .Cols - 1
      Screen.MousePointer = vbDefault
      frmProgresso.bar1.Value = 100
      
      Dim msgRetorno As Integer
      msgRetorno = MsgBox("Não foi encontrado nenhum registro com esse critério.", vbOKOnly + vbInformation + vbDefaultButton2)
      Unload Me
      Unload frmConsultaCliente
      Exit Sub
    End If
    
    If rs.RecordCount <> 0 Then rs.MoveFirst
    Dim X As Single

   .FormatString = "^Nº OS|<Nome|^Data|^Hora|<Usuario|^Baixa"

   'define o numero de linhas e colunas e configura o grid

   .ColWidth(0) = 800 'OS
   .ColWidth(1) = 4300 'Nome cliente
   .ColWidth(2) = 1200 'Data
   .ColWidth(3) = 1200 'Hora
   .ColWidth(4) = 3000 'Usuario
   .ColWidth(5) = 1500 'Baixa
   
   .Rows = rs.RecordCount + 1
   .Cols = 6
   .Row = 1
   .Col = 0
   .RowSel = .Rows - 1
   .ColSel = .Cols - 1
    
   .Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
   .Row = 1
   .Visible = True
    
    
   If rs.RecordCount > 10 Then
      .Rows = rs.RecordCount + 10
   Else
      .Rows = .Rows + 20 - rs.RecordCount
   End If
    
   .Row = rs.RecordCount
    
   If rs.RecordCount = 0 Then
      .Rows = 21
      .Row = 1
   End If
   
   .TopRow = .Rows - 20
   Dim linhaos As Long
    
   If linhaos <> 0 Then
      .Row = linhaos
      .TopRow = linhaos
   End If
   .ColSel = .Cols - 1
   linhaos = .Row

   If rs.RecordCount = 0 Then
      frmConsultaCliente.StatusBar1.Panels(1) = "Total de OS: Nenhum registro encontrado"
   ElseIf rs.RecordCount = 1 Then
      frmConsultaCliente.StatusBar1.Panels(1) = "Total de OS: " & rs.RecordCount & " registro encontrado"
   Else
      frmConsultaCliente.StatusBar1.Panels(1) = "Total de OS: " & rs.RecordCount & " registros encontrados"
   End If
End With
   
   rs.Close
   Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
   preencheGridCliente frmPesquisaOs.txtPesquisa.Text
End Sub
