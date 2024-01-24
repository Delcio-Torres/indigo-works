VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProgresso 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   675
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8535
   ControlBox      =   0   'False
   Icon            =   "frmProgresso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar bar1 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1
   End
End
Attribute VB_Name = "frmProgresso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub preencheGridCliente(Optional criterio As String)

With frmConsultaCliente.grdCliente

   Load frmConsultaCliente
   Dim mensagem As String
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
   
   mensagem = "Total de OS: " & rs.RecordCount
   mensagem = rs.RecordCount & " registros encontrados"
        
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
    
    
    X = 1
   Dim taxa As Double
   Dim registros As Long
   Dim contador As Double
   Dim vProgresso As Double
   
   frmProgresso.bar1.Max = 101
   frmProgresso.bar1.Min = 0
   registros = rs.RecordCount
   taxa = 100 / registros
   
'   While Not rs.EOF
'      contador = contador + taxa
'      If contador > 100 Then contador = 100
'      frmProgresso.bar1.Value = contador
'
'      .Rows = X + 1
'      .TextMatrix(X, 0) = Format$(rs!idos, "####0000")
'      .TextMatrix(X, 1) = rs!nomeCliente
'      .TextMatrix(X, 2) = Format$(rs!Data, "dd/mm/YY")
'      .TextMatrix(X, 3) = Format$(rs!hora, "hh:mm")
'      .TextMatrix(X, 4) = rs!nome
'      If Not IsNull(rs!baixa) Then .TextMatrix(X, 5) = rs!baixa
'      X = X + 1
'      rs.MoveNext
'   Wend
    
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
    Me.Hide
    frmConsultaCliente.Show 1

End Sub

Private Sub Form_Activate()
   preencheGridCliente frmPesquisaOs.txtPesquisa.Text
   Me.BackColor = &H80000004
End Sub

