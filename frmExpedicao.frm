VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmExpedicao 
   Caption         =   "Controle de Expedição"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19545
   Icon            =   "frmExpedicao.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   19545
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CommandButton Controle 
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
      Height          =   615
      Index           =   3
      Left            =   6240
      TabIndex        =   4
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Controle 
      Caption         =   "&Incluir Serviço"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   4200
      TabIndex        =   3
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Controle 
      Caption         =   "Visualizar &Saída"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Controle 
      Caption         =   "Visualizar &Entrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   2960
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5212
      _Version        =   393216
      Rows            =   23
      Cols            =   3
      ForeColor       =   0
      BackColorSel    =   16711680
      ForeColorSel    =   16777215
      FocusRect       =   0
      HighLight       =   2
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
Attribute VB_Name = "frmExpedicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim varModoVisao As String


Public Sub preencheGridSaida(Optional criterio As String)

Dim linhaos As Integer

    varModoVisao = "saida"
    flex.ForeColor = &HFF&
    Me.Caption = "Controle de Expedição - Saída"
    grid_de_saida
    
    abreExpedicao

With flex
    Screen.MousePointer = vbHourglass
    
    rse.Open "SELECT nomecliente, nos, tipo, loc, lançamento, datachegada, datasaida, vendedor, operadorsaida FROM expedicao WHERE entrega=true ORDER BY nomecliente", dbe, adOpenForwardOnly, adLockOptimistic
    If rse.RecordCount = 0 Then
        .Clear
        grid_de_saida
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Dim X As Integer
    
    For X = 1 To rse.RecordCount
        .TextMatrix(X, 0) = Format(X, "00")
    Next
    
    rse.MoveFirst
    'define o numero de linhas e colunas e configura o grid
    .Rows = rse.RecordCount + 1

    .Row = 1
    .Col = 1
    .RowSel = .Rows - 1
    .ColSel = .Cols - 1
    
    
    'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
    .Clip = rse.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    .Row = 1
    .Visible = True
    
    Screen.MousePointer = vbDefault

    If rse.RecordCount <= 23 Then
        Dim qLinhas As Integer
    
        qLinhas = 23 - rse.RecordCount
    
        .Rows = rse.RecordCount + qLinhas
    Else
        .Rows = rse.RecordCount
    End If
    
    If rse.RecordCount = 0 Then
        .Rows = 23
        .Row = 1
    End If


End With

    rse.Close
    Screen.MousePointer = vbDefault
    
End Sub
Public Sub preencheGridEntrada(Optional criterio As String)

    Dim linhaos As Integer
    
    varModoVisao = "entrada"
    flex.ForeColor = &H0&
    Me.Caption = "Controle de Expedição - Entrada"
    grid_de_entrada
    
    abreExpedicao

With flex
    Screen.MousePointer = vbHourglass
    
    rse.Open "SELECT nomecliente, nos, tipo, loc, datachegada, vendedor, operadorentrada FROM expedicao WHERE entrega=false ORDER BY nomecliente", dbe, adOpenForwardOnly, adLockOptimistic
    If rse.RecordCount = 0 Then
        .Clear
        grid_de_entrada
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Dim X As Integer
    
    For X = 1 To rse.RecordCount
        .TextMatrix(X, 0) = Format(X, "00")
    Next
    
    rse.MoveFirst
    'define o numero de linhas e colunas e configura o grid
    .Rows = rse.RecordCount + 1

    .Row = 1
    .Col = 1
    .RowSel = .Rows - 1
    .ColSel = .Cols - 1
    
    
    'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
    .Clip = rse.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    .Row = 1
    .Visible = True
    
    Screen.MousePointer = vbDefault
    
    If rse.RecordCount <= 23 Then
        Dim qLinhas As Integer
    
        qLinhas = 23 - rse.RecordCount
    
        .Rows = rse.RecordCount + qLinhas
    Else
        .Rows = rse.RecordCount
    End If
    
    If rse.RecordCount = 0 Then
        .Rows = 23
        .Row = 1
    End If
        
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1

End With
    
    rse.Close
    Screen.MousePointer = vbDefault

End Sub





Private Sub Controle_Click(Index As Integer)
    
    flex.Refresh
    
    Select Case Index
        
        Case 0

            varModoVisao = "entrada"
            preencheGridEntrada
            
        Case 1

            varModoVisao = "saída"
            preencheGridSaida
        
        Case 2
        
            If permissao Then
                frmEntradaExp.Show 1
            End If
        
        Case 3
        
            Unload Me
            
    End Select
End Sub

Private Sub flex_DblClick()
    
    If varTipoUsuario <> "Usuário-Ex" And varTipoUsuario <> "Administrador" Then
        MsgBox "Você não tem permissão para usar esse serviço", vbCritical
        Exit Sub
    End If
    
    If flex.TextMatrix(flex.RowSel, 1) = "" Then Exit Sub
    If varModoVisao = "entrada" Then
        With frmEntregaExp
            Load frmEntregaExp
            .Label1(0) = flex.TextMatrix(flex.RowSel, 2)
            .Label1(1) = flex.TextMatrix(flex.RowSel, 1)
            .Label1(2) = flex.TextMatrix(flex.RowSel, 3)
            .Label1(3) = flex.TextMatrix(flex.RowSel, 4)
            .Label1(4) = Format(Date, "dd/mm/yy")
            .Label1(5) = varNomeUsuario
            
            .Show 1
            
        End With
    Else
        
        Load frmDetalhesEntrega
        frmDetalhesEntrega.varNos = flex.TextMatrix(flex.RowSel, 3)
        frmDetalhesEntrega.Show 1
    
    End If

End Sub

Private Sub flex_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Del = 46
    If Not permissao Then Exit Sub
    
    If flex.TextMatrix(flex.RowSel, 3) = "" Then Exit Sub
    
    If KeyCode = 46 Then
        Dim op As Integer
        
        op = MsgBox("Deseja realmente excluir o item selecionado?", vbCritical + vbYesNo, "Indigo Works")
        ' sim = 6
        ' não = 7
        
        If op = 6 Then
            
            abreExpedicao
            
            rse.Open "SELECT * FROM expedicao WHERE nos =" & flex.TextMatrix(flex.RowSel, 2), dbe, adOpenStatic, adLockPessimistic
            
            rse.Delete adAffectCurrent
            rse.Update
            
            rse.Close
            dbe.Close
            
            If varModoVisao = "entrada" Then
                preencheGridEntrada
            Else
                preencheGridSaida
            End If
                    
        End If
        
    End If
    
    
    
    
End Sub

Private Sub Form_Load()

    varModoVisao = "entrada"
    Me.Caption = "Controle de Expedição - Entrada"

    flex.Width = Me.ScaleWidth
    flex.Height = Me.ScaleHeight - Controle.Item(0).Height
    flex.Left = 0
    flex.Top = 0
    Dim fdp As Integer
    fdp = Controle.Item(0).Height
    'Controle.Item(0). = Me.ScaleHeight - fdp
    Controle.Item(1).Top = Me.ScaleHeight - Controle.Item(1).Height - 100
    Controle.Item(2).Top = Me.ScaleHeight - Controle.Item(2).Height - 100
    preencheGridEntrada
    
End Sub

Private Sub Form_Resize()
    If Me.Width < 10000 Then
        Me.Width = 10000
        Exit Sub
    End If
    If Me.Height < 5000 Then
        Me.Height = 5000
        Exit Sub
    End If

    flex.Width = Me.ScaleWidth
    flex.Height = Me.ScaleHeight - Controle.Item(0).Height - 200
    Controle.Item(0).Top = Me.ScaleHeight - Controle.Item(0).Height - 100
    Controle.Item(1).Top = Me.ScaleHeight - Controle.Item(1).Height - 100
    Controle.Item(2).Top = Me.ScaleHeight - Controle.Item(2).Height - 100
    Controle.Item(3).Top = Me.ScaleHeight - Controle.Item(3).Height - 100
    
    
    Dim modo As Integer
    Dim menos As Integer
    
    If varModoVisao = "entrada" Then
        modo = 8
        menos = 13860
    Else
        modo = 10
        menos = 15350
    End If
    
    If Me.ScaleWidth - menos < 10 Then
        flex.ColWidth(modo) = 10
    Else
        flex.ColWidth(modo) = Me.ScaleWidth - menos
    End If
        
End Sub



Public Sub grid_de_entrada()
    With flex
        .Clear
        .Cols = 9
        .FormatString = "^|<Nome|^OS|^Tipo|^Local|^Data|<Vendedor|<Operador"
        .ColWidth(0) = 500  ' ordem
        .ColWidth(1) = 3500 ' Cliente
        .ColWidth(2) = 900  ' OS
        .ColWidth(3) = 900  ' Tipo
        .ColWidth(4) = 900  ' Local
        .ColWidth(5) = 1800 ' Data
        .ColWidth(6) = 2500 ' Atendente
        .ColWidth(7) = 2500 ' Operador

        If Me.ScaleWidth - 13860 < 10 Then
            flex.ColWidth(8) = 10
        Else
            flex.ColWidth(8) = Me.ScaleWidth - 13860
        End If
        
    End With

End Sub
Public Sub grid_de_saida()
    With flex
        '.Clear
        .Cols = 11
        .FormatString = "^|<Nome|^N° OS|^Tipo|^NS|^Lanç.|^Entrada|^Saída|<Vendedor|<Operador"
        .ColWidth(0) = 500  ' ordem
        .ColWidth(1) = 3500 ' Cliente
        .ColWidth(2) = 900  ' OS
        .ColWidth(3) = 900  ' Tipo
        .ColWidth(4) = 900  ' Local
        .ColWidth(5) = 900  ' Lançamento
        .ColWidth(6) = 1200 ' Data
        .ColWidth(7) = 1200 ' Saída
        .ColWidth(8) = 2500 ' Atendente
        .ColWidth(9) = 2500 ' Operador


        If Me.ScaleWidth - 15350 < 10 Then
            flex.ColWidth(10) = 10
        Else
            flex.ColWidth(10) = Me.ScaleWidth - 15350
        End If
        
    End With
End Sub
Public Function permissao() As Boolean

    If varTipoUsuario = "Usuário-Ex" Or varTipoUsuario = "Administrador" Then
        permissao = True
    Else
        MsgBox "Você não tem permissão para usar esse serviço", vbCritical
    End If

End Function
