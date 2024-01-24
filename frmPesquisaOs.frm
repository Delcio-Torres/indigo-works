VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPesquisaOs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indigo Works"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   Icon            =   "frmPesquisaOs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Busca"
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
      Left            =   5760
      TabIndex        =   9
      Top             =   9240
      Width           =   1215
   End
   Begin VB.TextBox txtPesquisa 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   3
      Top             =   9240
      Width           =   2775
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Incluir"
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
      Left            =   9600
      TabIndex        =   5
      Top             =   9240
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Nome Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   9240
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Nº OS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   9240
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
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
      Left            =   11400
      TabIndex        =   6
      Top             =   9240
      Width           =   1695
   End
   Begin MSComctlLib.StatusBar stbUsuario 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   9765
      Width           =   13320
      _ExtentX        =   23495
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
            Object.Width           =   18988
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
      Left            =   7800
      TabIndex        =   4
      Top             =   9240
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdUsuario 
      Height          =   8930
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   15743
      _Version        =   393216
      Rows            =   17
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   -2147483646
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisa Nº OS"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   9000
      Width           =   1455
   End
End
Attribute VB_Name = "frmPesquisaOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varOldCriterio As String
Public linhaos As Long
Public os As Integer
Public coluna As Long
Public nomeUsuario As String
Public nosFinal As Long

Private Sub cmdBuscar_Click()
    
   If txtPesquisa.Text = "" Then Exit Sub
   
   abreConexao
   
   Dim rd As Recordset
   Set rd = New ADODB.Recordset
   
   If frmPesquisaOs.Option1(0).Value = True Then
      rs.Open "SELECT os.idos, os.nomeCliente, os.data, os.hora, usuario.nome, os.baixa FROM OS INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE os.idos =" & txtPesquisa.Text & " ORDER BY idos", db, adOpenForwardOnly, adLockReadOnly
   Else
      rs.Open "SELECT os.idos, os.nomeCliente, os.data, os.hora, usuario.nome, os.baixa FROM OS INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE nomeCliente LIKE " & "'%" & txtPesquisa.Text & "%'" & "  ORDER BY idos", db, adOpenForwardOnly, adLockReadOnly
   End If
        
   If rs.RecordCount = 0 Then
      
      Dim msgRetorno As Integer
      msgRetorno = MsgBox("Não foi encontrado nenhum registro com esse critério.", vbOKOnly + vbInformation + vbDefaultButton2)
      
      rs.Close
      
      Exit Sub
    End If
    
   frmConsultaCliente.Show 1
   Unload frmConsultaCliente

End Sub

Public Sub cmdEditar_Click()

    If grdUsuario.TextMatrix(grdUsuario.RowSel, 1) = "" Then Exit Sub
    linhaos = grdUsuario.Row
    preencheCliente grdUsuario.TextMatrix(grdUsuario.RowSel, 0)

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdIncluir_Click()

On Error GoTo meuerro
   linhaos = 0
   
   abreConexao
   
   rs.Open "SELECT * FROM os ORDER BY idos", db, adOpenStatic, adLockOptimistic
   
   If rs.RecordCount <> 0 Then
    rs.MoveLast
    osInicial = rs!idos
    frmOrcamento.linhaInicial = rs.RecordCount
    rs.Close
   End If
   
   frmOrcamento.Show 1
   grdUsuario.SetFocus

Exit Sub

meuerro:

    MsgBox (Err.Description)
   
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
      If varTipoUsuario = "Administrador" Then
         Load frmBaixarOs
         frmBaixarOs.Show 1
      End If
   End If

   If KeyCode = vbKeyF12 Then
      If varCodigoUsuario = 6 Then
         
         Dim numeroOs As Long
         Dim rC As Recordset
         Set rC = New ADODB.Recordset
         
         rs.Open "SELECT * FROM os WHERE idcliente > 0", db, adOpenStatic, adLockOptimistic
         rs.MoveFirst
               
         While Not rs.EOF
            rC.Open "SELECT * FROM cliente WHERE idcliente =" & rs!idcliente, db, adOpenStatic, adLockOptimistic
            rs!nomeCliente = rC!nome
            rs.MoveNext
            rC.Close
         Wend
         rs.Close
      End If
      MsgBox "Tudo Limpo!!!!!!"
   End If

End Sub
Private Sub Form_Load()
    preencheGridCliente
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub grdUsuario_DblClick()

    If grdUsuario.TextMatrix(grdUsuario.RowSel, 1) = "" Then
        Exit Sub
    End If
    linhaos = grdUsuario.Row
    preencheCliente grdUsuario.TextMatrix(grdUsuario.RowSel, 0)

End Sub
Private Sub grdUsuario_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        stbUsuario.Panels(2).Text = grdUsuario.TextMatrix(grdUsuario.Row, 1) & "-" & grdUsuario.RowSel
    End If
End Sub
Public Sub preencheGridCliente(Optional criterio As String)
   
abreConexao

With grdUsuario
   
   .Refresh
   
   Dim mensagem As String
   Dim registro As Integer
   
   rs.Open "SELECT os.idos, os.nomeCliente, os.data, os.hora, usuario.nome, os.baixa FROM OS INNER JOIN usuario ON os.idusuario=usuario.codigousuario ORDER BY idos", db, adOpenForwardOnly, adLockReadOnly
   registro = rs.RecordCount
   
   If registro = 0 Then
      .Rows = 30
   Else
      rs.MoveFirst
      .Rows = registro + 1
   End If
   
   Screen.MousePointer = vbHourglass
   .Visible = False
   
   mensagem = "Total de OS: " & registro

   .FormatString = "^Nº OS|<Nome|^Data|^Hora|<Usuario|^Baixa"

   'define o número de linhas e colunas e configura o grid

   .ColWidth(0) = 800  'OS
   .ColWidth(1) = 4300 'Nome cliente
   .ColWidth(2) = 1200 'Data
   .ColWidth(3) = 1200 'Hora
   .ColWidth(4) = 3000 'Usuario
   .ColWidth(5) = 1500 'Baixa
   'define o numero de linhas e colunas e configura o grid
   
   .Cols = 6
   .Row = 1
   .Col = 0
   .RowSel = .Rows - 1
   .ColSel = .Cols - 1

   'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
   
   If registro > 0 Then .Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
   
   '.Row = 1
   .Visible = True
   
   
   Screen.MousePointer = vbDefault

   stbUsuario.Panels(2).Text = mensagem
   
   If rs.RecordCount < 21 Then
      .Rows = rs.RecordCount + (31 - rs.RecordCount)
      .TopRow = 1
   Else
      .Rows = .Rows + 9
      .TopRow = .Rows - 2 '29
   End If

   .Row = rs.RecordCount

   .ColSel = .Cols - 1
   linhaos = .Row

End With
   rs.Close
   frmOrcamento.btnAdd(2).Enabled = True
   Screen.MousePointer = vbDefault
End Sub
Private Sub Option1_Click(Index As Integer)
   
   txtPesquisa.Text = ""
   txtPesquisa.SetFocus
   
   Select Case Index
      Case 0
         Label1.Caption = "Pesquisa Nº OS:"
         txtPesquisa.MaxLength = 5
      Case 1
         Label1.Caption = "Pesquisa nome cliente:"
         txtPesquisa.MaxLength = 0
   End Select

End Sub
Public Function preencheCliente(codigo As Long)
   Dim codigoCliente As String
   Dim nome As String
   Dim telefone As String
   Dim rd As Recordset
   Set rd = New ADODB.Recordset
   
   nOS = codigo
   
   Load frmOrcamento
   
   
   rs.Open "SELECT * FROM os WHERE idos=" & codigo, db, adOpenStatic, adLockPessimistic
    
   If rs.RecordCount = 0 Then
      Dim a As Integer
      a = MsgBox("Esta OS foi excluída pelo administrador", vbInformation)
      rs.Close
      Exit Function
   Else
      rd.Open "SELECT * FROM cliente WHERE idcliente=" & rs!idcliente, db, adOpenStatic, adLockPessimistic
   End If

With frmOrcamento
    
    If rs!idcliente = 0 Then
    
            If Len(rs!telefonecliente) < 15 Then
            .txtTelefone.Mask = "(##)####-####"
        
        Else
            .txtTelefone.Mask = "(##)#####-####"
        End If
        
        .txtTelefone.Text = rs!telefonecliente   'telefone
        .txtCliente(0).Text = rs!nomeCliente 'nome
        
        If Not IsNull(rs!CNPJ) Then
            If Len(rs!CNPJ) = 14 Then .Option1(1).Value = True Else .Option1(0).Value = True
            .txtCPF.Text = rs!CNPJ
            If Not IsNull(rs!endereço) Then .txtCliente(1) = rs!endereço
            If Not IsNull(rs!bairro) Then .txtCliente(2) = rs!bairro
            If Not IsNull(rs!photobook) Then .Check4.Value = rs!photobook
            If Not IsNull(rs!Montagem) Then .Check5.Value = rs!Montagem
            If Not IsNull(rs!Pagamento) Then .optPagamento.Item(rs!Pagamento).Value = True
        End If
    Else
        '.txtCliente(0).Locked = False
        '.txtCliente(1).Enabled = False
        '.txtTelefone.Enabled = False
        '.txtCliente(2).Enabled = False
        '.Option1(1).Enabled = False
        '.txtCPF.Enabled = False
        .txtCliente(0).Text = rd!nome
        .txtCliente(1).Text = rd!endereco
        
        If Len(rd!telefone) < 5 Then
            .txtTelefone.Mask = "(##)####-####)"
        
        Else
            .txtTelefone.Mask = "(##)#####-####)"
        End If
        
        .txtTelefone.Text = rd!telefone
        
        .txtCliente(2).Text = rd!bairro
        .Option1(1).Value = True
        .txtCPF.Text = rd!CNPJ
        .txtCodigoCliente = rs!idcliente
        
        If Not IsNull(rs!Pagamento) Then
            .optPagamento.Item(rs!Pagamento).Value = True
        Else
            .optPagamento(4).Value = True
        End If
        
        If rd!condicao = "Bloqueado" Then
            .txtCliente(0).BackColor = &HC0FFFF
            .txtCliente(1).BackColor = &HC0FFFF
            .txtCliente(2).BackColor = &HC0FFFF
            .txtTelefone.BackColor = &HC0FFFF
            .txtCPF.BackColor = &HC0FFFF
        End If
    End If
    
    rd.Close
    
    rd.Open "SELECT * FROM usuario WHERE codigoUsuario=" & rs!idUsuario, db, adOpenStatic, adLockOptimistic
        If rd.RecordCount > 0 Then
            .txtUsuario = rd!nome
        Else
            .txtUsuario = "Usuário deletado"
        End If
    rd.Close
    
    .txtData = Format$(rs!Data, "dd/mm/yy")
    .txtHora = Format$(rs!hora, "hh:mm")
    .txtAlteradoPor = rs!alteradopor
'planos
    Dim w As Long

    rd.Open "SELECT * FROM plano WHERE idos=" & codigo, db, adOpenStatic, adLockOptimistic
     rd.MoveFirst
    w = 0
    While Not rd.EOF
        w = w + 1
        .gd.TextMatrix(w, 1) = rd!quantidade
        .gd.TextMatrix(w, 2) = rd!descricao
        
        If rd!formato <> 0 Then .gd.TextMatrix(w, 3) = rd!formato
        
        .gd.TextMatrix(w, 4) = rd!midia
        If rd!cores <> 0 Then .gd.TextMatrix(w, 5) = rd!cores
        If rd!laminacao <> 0 Then .gd.TextMatrix(w, 6) = rd!laminacao
        If rd!encadernacao <> 0 Then .gd.TextMatrix(w, 7) = rd!encadernacao
        If rd!capa <> 0 Then .gd.TextMatrix(w, 8) = rd!capa
        If rd!wireo <> 0 Then .gd.TextMatrix(w, 9) = rd!wireo
        
        If rd!corte <> 0 Then .gd.TextMatrix(w, 10) = rd!corte
        If rd!corte = 0 Then .gd.TextMatrix(w, 10) = ""
        
        .gd.TextMatrix(w, 11) = rd!meiocorte
        If rd!meiocorte = 0 Then .gd.TextMatrix(w, 11) = ""
        
        If Not IsNull(rd!picote) Then
         .gd.TextMatrix(w, 12) = rd!picote
        End If
        If rd!picote = 0 Then .gd.TextMatrix(w, 12) = ""
        
        .gd.TextMatrix(w, 13) = rd!vinco
        If rd!vinco = 0 Then .gd.TextMatrix(w, 13) = ""
    
        .gd.TextMatrix(w, 14) = rd!valor

        rd.MoveNext
    Wend
    frmOrcamento.contadorDePlano = w
    
    If Not IsNull(rs!acrescimo) Then
        .txtDescricaoAcrescimo = rs!acrescimo
    End If
        If Not IsNull(rs!descricaodesconto) Then
        .txtDescricaoDesconto = rs!descricaodesconto
    End If
    
    'verifica se OS esta baixada
    If rs!baixa = "Baixado" Then
        .Check3.Value = 1
    Else
        .Check3.Value = 0
    End If
    
    .txtExemplar = 2
    .txtAcrescimo = rs!outros
    .txtDesconto = rs!desconto
    .txtValorAcrescimo = Format$(rs!outros, "#,##0.00")
    .txtValorDesconto = Format$(rs!desconto * -1, "#,##0.00")
    .txtExemplar = rs!exemplar
    
    If Not IsNull(rs!entrega) Then
        .optEntrega.Item(rs!entrega).Value = True
    End If

    .modoEdicao = True
    .nOS = codigo
    rs.Close
    
    .Caption = "Ordem de serviço: " & Format$(codigo, "####0000")
    .btnAdd(2).Enabled = True

    .Show 1
    
End With
End Function
Private Function nomeCliente(id As Single) As String

    Dim rd As Recordset
    Set rd = New ADODB.Recordset

    rd.Open "SELECT nome FROM cliente WHERE idcliente=" & id, db, adOpenStatic, adLockOptimistic
    
    If rd.RecordCount > 0 Then nomeCliente = rd!nome
    
    rd.Close
    
End Function
Private Sub txtPesquisa_GotFocus()

    cmdBuscar.Default = True

End Sub

Private Sub txtPesquisa_KeyPress(KeyAscii As Integer)
    If Option1(0).Value = True Then
        If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Static Function linhagrid(linhatop As Long) As Long
    With grdUsuario
        .TopRow = linhatop
        .Row = linhatop
        .ColSel = 5
    End With
End Function

Private Sub txtPesquisa_LostFocus()
    cmdEditar.Default = True
End Sub
