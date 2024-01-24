VERSION 5.00
Begin VB.Form frmEntradaExp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada para Expedição"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6300
   Icon            =   "frmEntradaExp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2640
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1800
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "Offset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "Digital"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4560
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Vendedor:"
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
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Operador:"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Local:"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nome Cliente: "
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
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número OS:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1245
   End
End
Attribute VB_Name = "frmEntradaExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim varExpNomeCliente As String
Dim varExpOS As Long
Dim varExpVendedor As String
Dim varExpOperador As String

Function digital()
    
    abreConexao
    
    rs.Open "Select nomecliente, idos , idusuario from os WHERE idos=" & Text1(0).Text, db, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        MsgBox "OS não encontrada."
        Text1(0).SetFocus
        Exit Function
    End If
    
    varExpNomeCliente = rs!nomeCliente
    varExpOS = rs!idos
    Text1(2).Text = rs!nomeCliente
    
    Dim varIdusuario As Integer
    varIdusuario = rs!idUsuario
    
    rs.Close
    rs.Open "SELECT nome FROM usuario WHERE codigousuario=" & varIdusuario, db, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        varExpVendedor = "Vendedor excluído"
    Else
        varExpVendedor = rs!nome
    End If
    
    Text1(3).Text = varExpVendedor
    
    rs.Close
    db.Close

End Function


Function ofsset()
    
End Function


Private Sub Command1_Click(Index As Integer)

    varExpOperador = "delcio"
    Select Case Index
        Case 0
            
            abreExpedicao
            rse.Open "Select * from expedicao", dbe, adOpenStatic, adLockOptimistic
            
            rse.AddNew
            
            rse!nomeCliente = Text1(2).Text
            rse!nOS = Text1(0).Text
            rse!Loc = Text1(1).Text
            rse!vendedor = varExpVendedor
            rse!dataChegada = Format(Date, "dd/mm/yy")
            rse!horachegada = Format(Time, "hh:mm")
            rse!operadorEntrada = Label5.Caption
            
            If optTipo(0).Value = True Then
                rse!tipo = "Digital"
            Else
                rse!tipo = "Offset"
            End If
            
            rse.Update
            
            frmExpedicao.preencheGridEntrada
            
        Case 1
            Unload Me
    End Select
        
        Unload Me
    
End Sub

Private Sub Form_Activate()
    Text1(0).SetFocus
    Label5.Caption = varNomeUsuario
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Label5.Caption = ""
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub



Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Text1(1).Text = UCase(Text1(1).Text)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    
    Label5.Caption = varNomeUsuario
    
    Select Case Index
        
        Case 0
            
            If Text1(0).Text = "" Or optTipo(0).Value = False Then Exit Sub
            
            abreConexao
            
            abreExpedicao
            rse.Open "SELECT * FROM expedicao WHERE nos=" & Text1(0).Text, dbe, adOpenStatic, adLockOptimistic
            If rse.RecordCount > 1 Then
                MsgBox "Já foi dado entrada dessa OS na expedição."
                Text1(0).SetFocus
                Exit Sub
            End If
            rse.Close
            db.Close
            
            If optTipo(0).Value = True Then
                digital
            Else
                'offset
            End If
            
            
        Case 1

            
        Case 2
        Case 3
        Case 4
    End Select
End Sub


