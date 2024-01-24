VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBaixarOs 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixar OS"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11985
   Icon            =   "frmBaixarOs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Deletar"
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
      Index           =   2
      Left            =   6360
      TabIndex        =   3
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdLote 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Incluir &Lote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Baixar"
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
      Left            =   8760
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Fatura"
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
      Left            =   7560
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grdos 
      Height          =   4950
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8731
      _Version        =   393216
      Rows            =   17
      Cols            =   6
      FixedCols       =   0
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
   Begin VB.CommandButton cmdBaixar 
      BackColor       =   &H8000000A&
      Caption         =   "Baixar tudo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9960
      TabIndex        =   6
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      TabIndex        =   0
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdIncluir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Incluir"
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
      Height          =   420
      Index           =   0
      Left            =   2880
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "OS Nº:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   5400
      Width           =   705
   End
End
Attribute VB_Name = "frmBaixarOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tos As Long
Public linha As Long

Private Sub alinhaColunas()

With grdos
    .ColAlignment(0) = 3
    .ColAlignment(1) = 1
    .ColAlignment(2) = 6
    .ColAlignment(3) = 3
    .ColAlignment(4) = 1
End With

End Sub

Private Sub cmdIncluir_Click(Index As Integer)

Dim idUsuario As Integer
Dim numeroOs As Integer

    If Text2.Text = "" Then Text2.SetFocus: Exit Sub
        
    Dim p As Long
    For p = 1 To grdos.Rows - 1
        If Text2.Text = grdos.TextMatrix(p, 0) Then
            MsgBox "'OS' já incluída", vbInformation
            Text2.Text = ""
            Text2.SetFocus
            Exit Sub
        End If
    Next
    
    
    
    abreConexao
    
    Dim rs2 As Recordset
    Set rs2 = New ADODB.Recordset
    
    rs.Open "Select * from os WHERE idos=" & Text2.Text, db, adOpenStatic, adLockOptimistic
        
    If rs.RecordCount = 0 Then
        MsgBox "Não foi possível encontrar a OS com esse número.", vbInformation
        Text2.Text = ""
        Text2.SetFocus
        Exit Sub
    'ElseIf rs!Baixa = "Baixado" Then 'Or rs!baixa = "Faturado" Then
    '    MsgBox "Esta OS já foi baixada.", vbInformation
    '    Text2.Text = ""
    '    Text2.SetFocus
    '    Exit Sub
    End If

    idUsuario = rs!idUsuario
    linha = linha + 1
    If linha > 16 Then
        grdos.Rows = grdos.Rows + 1
    End If
    
    grdos.TextMatrix(linha, 0) = rs!idos
    If Not IsNull(rs!baixa) Then grdos.TextMatrix(linha, 5) = rs!baixa
    

   grdos.TextMatrix(linha, 1) = rs!nomeCliente

    
    grdos.TextMatrix(linha, 3) = rs!Data
    grdos.TextMatrix(linha, 2) = Format(calculaValorOs(rs!idos), "#,##0.00")

    
    rs.Open "SELECT * FROM usuario WHERE codigousuario=" & idUsuario, db, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount <> 0 Then
        grdos.TextMatrix(linha, 4) = rs!nome
    End If
    
    If linha > 16 Then
        grdos.TopRow = linha - 15
    End If
    
    grdos.Row = linha
    grdos.ColSel = 5
    
    rs.Close
    Text2.Text = ""
    Text2.SetFocus
End Sub

Private Sub cmdBaixar_Click()

Dim w As Long
Dim numeroOs As Long
Dim linha2 As Long

abreConexao

grdos.Row = 1
grdos.ColSel = 5

For w = 1 To grdos.Rows - 1
  numeroOs = Val(grdos.TextMatrix(1, 0))
    If numeroOs <> 0 Then
      rs.Open "Select * from os where idos=" & numeroOs, db, adOpenStatic, adLockOptimistic
      
      If Option1(0).Value = True Then
        rs!baixa = "Baixado"
      ElseIf Option1(1).Value = True Then
        rs!baixa = "Faturado"
      ElseIf Option1(2).Value = True Then
        Dim op As Integer
        op = MsgBox("Todas as OS selecionadas serão apagadas." & Chr(13) & "Quer continuar?", vbOKCancel + vbCritical)
        
        Debug.Print op
        Exit Sub
        If op = 1 Then
          'rs.Delete adAffectCurrent
        End If
      End If
      rs!databaixa = Date
      rs.update
      rs.Close
      
      Dim atendente As String
      atendente = grdos.TextMatrix(1, 4)
      grdos.RemoveItem (grdos.Row)
      linha = linha - 1
         
      '----------------------------------------------------------------------
      With frmPesquisaOs.grdUsuario
        For linha2 = 1 To .Rows - 1
          If atendente <> "" And .TextMatrix(linha2, 0) <> "" Then
            If numeroOs = .TextMatrix(linha2, 0) Then
              If Option1(0).Value = True Then
                .TextMatrix(linha2, 5) = "Baixado"
                ElseIf Option1(1).Value = True Then
                .TextMatrix(linha2, 5) = "Faturado"
              End If
              Exit For
            End If
          End If
        Next
      End With
      '----------------------------------------------------------------------
      
         If grdos.Rows < 17 Then
            grdos.Rows = grdos.Rows + 1
         End If
      End If
Next

Text2.SetFocus
Text2.Text = ""

End Sub

Private Sub cmdLote_Click()
    frmBaixarOsLote.Show 1
End Sub



Private Sub Form_Activate()
    Text2.SetFocus
    If varCodigoUsuario = 3 Or varCodigoUsuario = 6 Then
        Option1(0).Visible = True
    End If
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyCancel Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()

linha = 0
With grdos
    
    alinhaColunas
    'defineLargura
    .ColWidth(0) = 800 'OS
    .ColWidth(1) = 4100 'Nome cliente
    .ColWidth(2) = 1000 'Valor
    .ColWidth(3) = 1100 'Data
    .ColWidth(4) = 3000 'Usuario
    .ColWidth(5) = 1000 'Baixa
    
'------------Título
    .TextMatrix(0, 0) = "OS"
    .TextMatrix(0, 1) = "Cliente"
    .TextMatrix(0, 2) = "Valor"
    .TextMatrix(0, 3) = "Data"
    .TextMatrix(0, 4) = "Atendente"
    .TextMatrix(0, 5) = "Baixa"
End With

End Sub

Private Sub grdOS_DblClick()
    If grdos.TextMatrix(grdos.RowSel, 1) = "" Then
        Exit Sub
    End If
    frmPesquisaOs.preencheCliente grdos.TextMatrix(grdos.RowSel, 0)
End Sub

Private Sub grdOS_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 46 Then
    If grdos.TextMatrix(grdos.Row, 0) <> "" Then
        grdos.RemoveItem (grdos.Row)
        linha = linha - 1
        If grdos.Rows < 17 Then
            If grdos.Row > linha Then
                grdos.Row = linha
                grdos.ColSel = 5
            End If
            grdos.Rows = grdos.Rows + 1
        End If
    End If
End If

End Sub

Private Sub Option1_Click(Index As Integer)
    cmdBaixar.Enabled = True
    Select Case Index
        
        Case 0
            cmdBaixar.Caption = "Baixar tudo"
        Case 1
            cmdBaixar.Caption = "Faturar tudo"
        Case 2
            cmdBaixar.Caption = "Deletar tudo"
    
    End Select

End Sub

Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
    KeyAscii = 0
End If

End Sub
