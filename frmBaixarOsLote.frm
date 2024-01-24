VERSION 5.00
Begin VB.Form frmBaixarOsLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adicionar OS em Lote"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3795
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBaixarOsLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtInicial 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Text            =   "os final"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtInicial 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Text            =   "os inicial"
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "OS Final:"
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
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   360
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "OS Inicial:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   885
   End
End
Attribute VB_Name = "frmBaixarOsLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()

Dim idUsuario As Integer
Dim numeroOs As Integer

    If txtInicial(0).Text = "" Then txtInicial(0).SetFocus: Exit Sub
    If txtInicial(1).Text = "" Then txtInicial(1).SetFocus: Exit Sub
    
    If txtInicial(1).Text < txtInicial(0).Text Then
        Dim a As Integer
        a = MsgBox("'OS' inicial maior que 'OS' final", vbExclamation)
        Exit Sub
    End If

    abreConexao
    
    Dim rs2 As Recordset
    Set rs2 = New ADODB.Recordset

With frmBaixarOs
 
.linha = 0
Dim w As Long
For w = txtInicial(0).Text To txtInicial(1).Text
    rs.Open "Select * from os WHERE idos=" & w, db, adOpenStatic, adLockOptimistic
        
    If rs.RecordCount <> 0 Then
    
        idUsuario = rs!idUsuario
        .linha = .linha + 1
        If .linha > 16 Then
            .grdos.Rows = .grdos.Rows + 1
        End If
        
        .grdos.TextMatrix(.linha, 0) = Format(rs!idos, "####00000")
        
        If rs!idcliente <> 0 Then
            rs2.Open "SELECT * FROM cliente WHERE idcliente=" & rs!idcliente, db, adOpenStatic, adLockOptimistic
            .grdos.TextMatrix(.linha, 1) = rs2!nome
            rs2.Close
        Else
            .grdos.TextMatrix(.linha, 1) = rs!nomeCliente
        End If
        
        If Not IsNull(rs!baixa) Then .grdos.TextMatrix(.linha, 5) = rs!baixa
        .grdos.TextMatrix(.linha, 3) = rs!Data
        .grdos.TextMatrix(.linha, 2) = Format(calculaValorOs(rs!idos), "#,##0.00")
        
        rs.Open "SELECT * FROM usuario WHERE codigousuario=" & idUsuario, db, adOpenStatic, adLockOptimistic
        
        If rs.RecordCount <> 0 Then
            .grdos.TextMatrix(.linha, 4) = rs!nome
        End If
        
        If .linha > 16 Then
            .grdos.TopRow = .linha - 15
        End If
        
        .grdos.Row = .linha
        .grdos.ColSel = 5
    
    End If
    rs.Close

Next
End With
    Unload Me
End Sub




Private Sub Form_Load()
    txtInicial(0).Text = ""
    txtInicial(1).Text = ""
    Me.Top = frmBaixarOs.Top + frmBaixarOs.Height / 2 - Me.Height / 2
    Me.Left = frmBaixarOs.Left + frmBaixarOs.Width / 2 - Me.Width / 2

End Sub

Private Sub txtInicial_GotFocus(Index As Integer)
    txtInicial(Index).SelStart = 0
    txtInicial(Index).SelLength = Len(txtInicial(Index).Text)
End Sub

Private Sub txtInicial_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub


