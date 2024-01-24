VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRelatorioVendas 
   Caption         =   "Relatório de Vendas"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   Icon            =   "frmRelatorioVendas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6015
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   8055
      Begin VB.Frame Frame3 
         Height          =   3375
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   7935
         Begin VB.Frame Frame5 
            Height          =   2775
            Left            =   6000
            TabIndex        =   40
            Top             =   360
            Width           =   1815
            Begin VB.TextBox txtTotalGeral 
               Alignment       =   1  'Right Justify
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
               Left            =   240
               TabIndex        =   43
               Text            =   "0,00"
               Top             =   1560
               Width           =   1335
            End
            Begin VB.TextBox txtValorOsNaoBaixada 
               Alignment       =   1  'Right Justify
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
               Left            =   240
               TabIndex        =   41
               Text            =   "0,00"
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Total Geral:"
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
               Index           =   4
               Left            =   240
               TabIndex        =   44
               Top             =   1320
               Width           =   1050
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "OS Não Baixadas"
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
               Index           =   3
               Left            =   120
               TabIndex        =   42
               Top             =   360
               Width           =   1620
            End
         End
         Begin VB.Frame Frame4 
            Height          =   2775
            Left            =   4080
            TabIndex        =   33
            Top             =   360
            Width           =   1815
            Begin VB.TextBox txtValorOsFaturada 
               Alignment       =   1  'Right Justify
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
               Left            =   240
               TabIndex        =   36
               Text            =   "0,00"
               Top             =   1320
               Width           =   1335
            End
            Begin VB.TextBox txtValorOsBaixada 
               Alignment       =   1  'Right Justify
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
               Left            =   240
               TabIndex        =   35
               Text            =   "0,00"
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox txtTotal 
               Alignment       =   1  'Right Justify
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
               Left            =   240
               TabIndex        =   34
               Text            =   "0,00"
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "OS Faturadas"
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
               Index           =   2
               Left            =   240
               TabIndex        =   39
               Top             =   1080
               Width           =   1245
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "OS Baixadas"
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
               Left            =   240
               TabIndex        =   38
               Top             =   360
               Width           =   1185
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Total"
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
               TabIndex        =   37
               Top             =   1800
               Width           =   465
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H8000000A&
            Caption         =   "OS emitida nos últimos 6 meses"
            Height          =   2775
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   3855
            Begin VB.PictureBox pic 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Index           =   5
               Left            =   240
               ScaleHeight     =   15
               ScaleWidth      =   375
               TabIndex        =   17
               Top             =   2520
               Width           =   375
            End
            Begin VB.PictureBox pic 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Index           =   4
               Left            =   720
               ScaleHeight     =   15
               ScaleWidth      =   375
               TabIndex        =   16
               Top             =   2520
               Width           =   375
            End
            Begin VB.PictureBox pic 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Index           =   3
               Left            =   1200
               ScaleHeight     =   15
               ScaleWidth      =   375
               TabIndex        =   15
               Top             =   2520
               Width           =   375
            End
            Begin VB.PictureBox pic 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Index           =   2
               Left            =   1680
               ScaleHeight     =   15
               ScaleWidth      =   375
               TabIndex        =   14
               Top             =   2520
               Width           =   375
            End
            Begin VB.PictureBox pic 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Index           =   1
               Left            =   2160
               ScaleHeight     =   15
               ScaleWidth      =   375
               TabIndex        =   13
               Top             =   2520
               Width           =   375
            End
            Begin VB.PictureBox pic 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   10
               Index           =   0
               Left            =   2640
               ScaleHeight     =   15
               ScaleWidth      =   375
               TabIndex        =   12
               Top             =   2520
               Width           =   375
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00C0C0C0&
               X1              =   3730
               X2              =   230
               Y1              =   1560
               Y2              =   1560
            End
            Begin VB.Label mes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   2790
               TabIndex        =   32
               Top             =   2550
               Width           =   60
            End
            Begin VB.Label mes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   2310
               TabIndex        =   31
               Top             =   2550
               Width           =   60
            End
            Begin VB.Label mes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   1830
               TabIndex        =   30
               Top             =   2550
               Width           =   60
            End
            Begin VB.Label mes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   3
               Left            =   1350
               TabIndex        =   29
               Top             =   2550
               Width           =   60
            End
            Begin VB.Label mes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   4
               Left            =   870
               TabIndex        =   28
               Top             =   2550
               Width           =   60
            End
            Begin VB.Label mes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   5
               Left            =   390
               TabIndex        =   27
               Top             =   2550
               Width           =   60
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00C0C0C0&
               X1              =   3720
               X2              =   230
               Y1              =   460
               Y2              =   460
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00C0C0C0&
               X1              =   3740
               X2              =   240
               Y1              =   2535
               Y2              =   2535
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "800"
               Height          =   195
               Left            =   3480
               TabIndex        =   26
               Top             =   480
               Width           =   270
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               Height          =   195
               Left            =   3660
               TabIndex        =   25
               Top             =   2355
               Width           =   90
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "400"
               Height          =   195
               Left            =   3480
               TabIndex        =   24
               Top             =   1560
               Width           =   270
            End
            Begin VB.Label lblOS 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   2790
               TabIndex        =   23
               Top             =   240
               Width           =   60
            End
            Begin VB.Label lblOS 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   2310
               TabIndex        =   22
               Top             =   240
               Width           =   60
            End
            Begin VB.Label lblOS 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   1830
               TabIndex        =   21
               Top             =   240
               Width           =   60
            End
            Begin VB.Label lblOS 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   3
               Left            =   1350
               TabIndex        =   20
               Top             =   240
               Width           =   60
            End
            Begin VB.Label lblOS 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   4
               Left            =   870
               TabIndex        =   19
               Top             =   240
               Width           =   60
            End
            Begin VB.Label lblOS 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   5
               Left            =   390
               TabIndex        =   18
               Top             =   240
               Width           =   60
            End
         End
      End
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "frmRelatorioVendas.frx":0ECA
      Left            =   120
      List            =   "frmRelatorioVendas.frx":0EF2
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   5160
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Min             =   1
      Scrolling       =   1
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2775
      Left            =   9840
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   3735
      _Version        =   524288
      _ExtentX        =   6588
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   -2147483638
      Year            =   2011
      Month           =   10
      Day             =   6
      DayLength       =   1
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Left            =   6720
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox cmbAtendente 
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
      ItemData        =   "frmRelatorioVendas.frx":0F5B
      Left            =   1680
      List            =   "frmRelatorioVendas.frx":0F5D
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ano"
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
      Left            =   5160
      TabIndex        =   8
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Mês"
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Atendente"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmRelatorioVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim indice As Integer

Private Sub cabecalho(nome As String, numeroPagina As Integer, paginaTotal As Integer)

Dim mes As String

With Printer
    Printer.DrawWidth = 10
    Printer.Line (0, 100)-(11150, 100)
    .FontSize = 15
    .CurrentY = 110
    .CurrentX = 100
    Printer.Print "Atendente: " & nome
    Printer.Print "Mês referente: " & Combo2.Text & "/" & Combo1.Text
    .CurrentY = 100
    .CurrentX = 6000
    Printer.Print "Página: "; numeroPagina & "/" & paginaTotal
    Printer.Line (0, 780)-(11150, 780)
    Printer.DrawWidth = 5
    Printer.Line (5575, 780)-(5575, 16000)
    Printer.FontSize = 10
End With

End Sub

Private Sub Calendar1_DblClick()
    If indice = 1 Then
        Calendar1.Visible = False
    End If
End Sub

Private Sub cmbAtendente_Click()
    
    If varTipoUsuario <> "Administrador" Then
      If varNomeUsuario <> cmbAtendente.Text Then Exit Sub
    End If
    
    Dim rst As Recordset
    Dim valorOSBaixada As Currency
    Dim quantidadeOsBaixada As Integer
    Dim valorOsNaoBaixada As Currency
    Dim valorOsFaturada As Currency
    Dim quantidadeOsNaoBaixada As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim idUsuario As Integer

    abreConexao

    txtValorOsBaixada = "0,00"
    txtValorOsFaturada.Text = "0,00"
    txtValorOsNaoBaixada.Text = "0,00"
    txtTotal.Text = "0,00"

    rs.Open "SELECT * FROM usuario WHERE nome='" & cmbAtendente.Text & "'", db, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount = 0 And cmbAtendente.Text <> "Todos" Then
        MsgBox "Selecione um nome do atendente"
        cmbAtendente.SetFocus
        Exit Sub
    End If
    
    If cmbAtendente.Text = "Todos" Then
    
        Set rst = New ADODB.Recordset

        rst.Open "SELECT sum(valoros) as Teste FROM os WHERE year(data) =" & Combo1.Text & " AND month(data)=" & Combo2.ListIndex + 1 & "AND baixa='Baixado'", db, adOpenStatic, adLockOptimistic
        If Not IsNull(rst!teste) Then valorOSBaixada = rst!teste
        rst.Close
        
        rst.Open "SELECT sum(valoros) as Teste FROM os WHERE year(data) =" & Combo1.Text & " AND month(data)=" & Combo2.ListIndex + 1 & "AND baixa='Faturado'", db, adOpenStatic, adLockOptimistic
        If Not IsNull(rst!teste) Then valorOsFaturada = rst!teste
        rst.Close
        
        rst.Open "SELECT sum(valoros) as Teste FROM os WHERE year(data) =" & Combo1.Text & " AND month(data)=" & Combo2.ListIndex + 1, db, adOpenStatic, adLockOptimistic
        If Not IsNull(rst!teste) Then
            valorOsNaoBaixada = rst!teste - valorOSBaixada - valorOsFaturada
        End If
        rst.Close
    
    Else
        idUsuario = rs!codigousuario
        
        rs.Close
        
        Set rst = New ADODB.Recordset
        
        rst.Open "SELECT sum(valoros) as Teste FROM os WHERE year(data) =" & Combo1.Text & " AND month(data)=" & Combo2.ListIndex + 1 & "AND baixa='Baixado' AND idusuario=" & idUsuario, db, adOpenStatic, adLockOptimistic
        If Not IsNull(rst!teste) Then valorOSBaixada = rst!teste
        rst.Close
        
        rst.Open "SELECT sum(valoros) as Teste FROM os WHERE year(data) =" & Combo1.Text & " AND month(data)=" & Combo2.ListIndex + 1 & "AND baixa='Faturado' AND idusuario=" & idUsuario, db, adOpenStatic, adLockOptimistic
        If Not IsNull(rst!teste) Then valorOsFaturada = rst!teste
        rst.Close
        
        rst.Open "SELECT sum(valoros) as Teste FROM os WHERE year(data) =" & Combo1.Text & " AND month(data)=" & Combo2.ListIndex + 1 & "AND idusuario=" & idUsuario, db, adOpenStatic, adLockOptimistic
        If Not IsNull(rst!teste) Then
            valorOsNaoBaixada = rst!teste - valorOSBaixada - valorOsFaturada
        End If
        rst.Close
        
    End If
    
    escreveMeses
    escreveOS

    txtValorOsBaixada.Text = Format(valorOSBaixada, "#,###0.00")
    txtValorOsFaturada.Text = Format(valorOsFaturada, "#,###0.00")
    txtValorOsNaoBaixada.Text = Format(valorOsNaoBaixada, "#,###0.00")
    txtTotal.Text = Format(valorOsFaturada + valorOSBaixada, "#,###0.00")
    txtTotalGeral.Text = Format(valorOsNaoBaixada + valorOSBaixada + valorOsFaturada, "#,###0.00")

    Me.MousePointer = 0

End Sub
Private Sub cmdImprimir_Click()

    If cmbAtendente.Text = "Todos" Then
        ImprimirTodos
        Exit Sub
    End If

Dim idUsuario As Integer

abreConexao

rs.Open "SELECT * FROM usuario WHERE nome='" & cmbAtendente.Text & "'", db, adOpenStatic, adLockOptimistic

If rs.RecordCount = 0 Then
    MsgBox "Selecione um nome do atendente"
    cmbAtendente.SetFocus
    Exit Sub
End If

idUsuario = rs!codigousuario

    Dim rs2 As Recordset
    Set rs2 = New ADODB.Recordset

rs.Close

    Dim rst As Recordset
    Dim datainicial As String
    Dim datafinal As String

    Set rst = New ADODB.Recordset
    rst.Open "SELECT * FROM os WHERE year(data) =" & Combo1.Text & " AND month(data)=" & Combo2.ListIndex + 1 & "AND idusuario=" & idUsuario, db, adOpenStatic, adLockOptimistic

    If rst.RecordCount = 0 Then
        MsgBox " Não foi localizado nenhum registro"
        rst.Close
        db.Close
        Exit Sub
    End If

    Dim valorOSBaixada As Currency
    Dim quantidadeOsBaixada As Integer
    Dim valorOsNaoBaixada As Currency
    Dim quantidadeOsNaoBaixada As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim yScale As Long
    Dim xScale As Long
    Dim numeroPagina As Integer
    Dim paginaTotal As Integer
    Dim nRegistros As Long

    yScale = 1000
    numeroPagina = 1
    paginaTotal = Int(rst.RecordCount / 150) + 1
    cabecalho cmbAtendente.Text, numeroPagina, paginaTotal
    
    rst.MoveFirst
    'pgb.Value = rst.RecordCount
    xScale = 20

For X = 1 To rst.RecordCount
        Y = Y + 1
        nRegistros = nRegistros + 1
        
        Printer.CurrentX = xScale
        Printer.CurrentY = yScale
        Printer.Print rst!idOS
        
        Printer.CurrentX = xScale + 730
        Printer.CurrentY = yScale
        
        If IsNull(rst!nomeCliente) Then
            rs2.Open "SELECT nome FROM cliente WHERE idCliente= " & rst!idcliente, db, adOpenStatic, adLockOptimistic
            Printer.Print Mid$(rs2!nome, 1, 30)
            rs2.Close
        Else
            Printer.Print Mid$(rst!nomeCliente, 1, 30)
        End If
        
        Printer.CurrentX = xScale + 4480 - Printer.TextWidth(Format(rst!valorOs, "#,##0.00"))
        Printer.CurrentY = yScale
        Printer.Print Format(rst!valorOs, "#,##0.00")
        
        Printer.CurrentX = xScale + 4580
        Printer.CurrentY = yScale
        
        If IsNull(rst!baixa) Then
            Printer.Print "-"
        Else
            Printer.Print rst!baixa
        End If
        
        yScale = yScale + 200
        
        rst.MoveNext
        
        If nRegistros = 75 Then
            xScale = 5615
            yScale = 1000
            nRegistros = 0
        End If

        If Y = 150 Then
            Printer.NewPage
            numeroPagina = numeroPagina + 1
            cabecalho cmbAtendente.Text, numeroPagina, paginaTotal
            yScale = 1000
            xScale = 20
            Y = 0
            nRegistros = 0
        End If
        'pgb.Value = X
Next

If yScale > 13700 Then
    If xScale < 5615 Then
        yScale = 1000
        xScale = 5616
    Else
        Printer.NewPage
        cabecalho cmbAtendente.Text, numeroPagina, paginaTotal
        yScale = 1000
        xScale = 20
    End If
Else
    yScale = yScale + 300
End If
    
    Printer.CurrentX = xScale
    Printer.CurrentY = yScale
    Printer.Print "OS Baixadas..........................................................."
    Printer.CurrentY = yScale
    Printer.CurrentX = xScale + 5375 - Printer.TextWidth(txtValorOsBaixada.Text)
    Printer.Print txtValorOsBaixada.Text
    
    yScale = yScale + 250
    Printer.CurrentY = yScale
    Printer.CurrentX = xScale
    Printer.Print "OS Faturadas.........................................................."
    Printer.CurrentY = yScale
    Printer.CurrentX = xScale + 5375 - Printer.TextWidth(txtValorOsFaturada.Text)
    Printer.Print txtValorOsFaturada.Text
    
    yScale = yScale + 250
    Printer.CurrentX = xScale
    Printer.CurrentY = yScale
    Printer.Print "Total de vendas......................................................."
    Printer.CurrentY = yScale
    Printer.CurrentX = xScale + 5375 - Printer.TextWidth(txtTotal.Text)
    Printer.Print txtTotal.Text
    
    yScale = yScale + 250
    Printer.CurrentX = xScale
    Printer.CurrentY = yScale
    Printer.Print "OS Não Baixadas...................................................."
    Printer.CurrentY = yScale
    Printer.CurrentX = xScale + 5375 - Printer.TextWidth(txtValorOsNaoBaixada.Text)
    Printer.Print txtValorOsNaoBaixada.Text
    
    yScale = yScale + 250
    Printer.CurrentX = xScale
    Printer.CurrentY = yScale
    Printer.Print "Total Geral..............................................................."
    Printer.CurrentY = yScale
    Printer.CurrentX = xScale + 5375 - Printer.TextWidth(txtTotalGeral.Text)
    Printer.Print txtTotalGeral.Text
    
    Printer.EndDoc
    rst.Close
    db.Close

End Sub









Private Sub Form_Activate()
    Combo2.ListIndex = Month(Date) - 1
    Combo1.AddItem (Year(Date))
    Combo1.AddItem (Year(Date) - 1)
    Combo1.AddItem (Year(Date) - 2)
    Combo1.AddItem (Year(Date) - 3)
    Combo1.ListIndex = 3
End Sub

Private Sub Form_Load()

    abreConexao

    rs.Open "SELECT codigoUsuario, nome FROM usuario", db, adOpenStatic, adLockOptimistic

    rs.MoveFirst
    Dim X As Integer

    For X = 1 To rs.RecordCount
        cmbAtendente.AddItem rs!nome
        rs.MoveNext
    Next
        cmbAtendente.AddItem "Todos"
    rs.Close
    db.Close

End Sub

Private Sub escreveMeses()

    Dim X As Integer
    Dim mesReferente As Integer
    
    mesReferente = Combo2.ListIndex + 1
    
    For X = 0 To 5
        Select Case mesReferente
            Case 1: mes(X).Caption = "JAN"
            Case 2: mes(X).Caption = "FEV"
            Case 3: mes(X).Caption = "MAR"
            Case 4: mes(X).Caption = "ABR"
            Case 5: mes(X).Caption = "MAI"
            Case 6: mes(X).Caption = "JUN"
            Case 7: mes(X).Caption = "JUL"
            Case 8: mes(X).Caption = "AGO"
            Case 9: mes(X).Caption = "SET"
            Case 10: mes(X).Caption = "OUT"
            Case 11: mes(X).Caption = "NOV"
            Case 12: mes(X).Caption = "DEZ"
        End Select
        mesReferente = mesReferente - 1
        If mesReferente = 0 Then mesReferente = 12
    Next
    
End Sub

Private Sub escreveOS()

Dim idUsuario As Integer
Dim X As Integer
Dim mesReferente As Integer
Dim anoReferente As Integer


abreConexao

    Dim rst As Recordset
    Set rst = New ADODB.Recordset

If cmbAtendente.Text = "Todos" Then
    
    mesReferente = Combo2.ListIndex + 1
    anoReferente = Combo1.Text
    
        For X = 0 To 5
            rst.Open "SELECT * FROM os WHERE year(data) =" & anoReferente & " AND month(data)=" & mesReferente, db, adOpenStatic, adLockOptimistic
            tamanhoBarra X, rst.RecordCount
            If rst.RecordCount = 0 Then
                lblOS(X).Caption = 0
            Else
                lblOS(X).Caption = rst.RecordCount
            End If
            rst.Close
            mesReferente = mesReferente - 1
            If mesReferente = 0 Then
                mesReferente = 12
                anoReferente = anoReferente - 1
            End If
        Next

Else
    rs.Open "SELECT * FROM usuario WHERE nome='" & cmbAtendente.Text & "'", db, adOpenStatic, adLockOptimistic
    idUsuario = rs!codigousuario
    rs.Close
    
    mesReferente = Combo2.ListIndex + 1
    anoReferente = Combo1.Text
    
        For X = 0 To 5
            rst.Open "SELECT * FROM os WHERE year(data) =" & anoReferente & " AND month(data)=" & mesReferente & "AND idusuario=" & idUsuario, db, adOpenStatic, adLockOptimistic
            tamanhoBarra X, rst.RecordCount
            If rst.RecordCount = 0 Then
                lblOS(X).Caption = 0
            Else
                lblOS(X).Caption = rst.RecordCount
            End If
            rst.Close
            mesReferente = mesReferente - 1
            If mesReferente = 0 Then
                mesReferente = 12
                anoReferente = anoReferente - 1
            End If
        Next
End If

db.Close
End Sub

Private Sub tamanhoBarra(indice As Integer, quantidade As Integer)

    pic(indice).Height = quantidade * 2.57
    pic(indice).Top = 2055 - pic(indice).Height + 480
    lblOS(indice).Top = pic(indice).Top - 200

End Sub

Private Sub ImprimirTodos()

Dim idUsuario As Integer

abreConexao

    Dim rst As Recordset
    Dim datainicial As String
    Dim datafinal As String

    Set rst = New ADODB.Recordset
    rst.Open "SELECT * FROM os WHERE year(data) =" & Combo1.Text & " AND month(data)=" & Combo2.ListIndex + 1 & " ORDER BY idos ", db, adOpenStatic, adLockOptimistic
    
    Dim valorOSBaixada As Currency
    Dim quantidadeOsBaixada As Integer
    Dim valorOsNaoBaixada As Currency
    Dim quantidadeOsNaoBaixada As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim yScale As Long
    Dim xScale As Long
    Dim numeroPagina As Integer
    Dim paginaTotal As Integer
    Dim nRegistros As Long

    yScale = 1000
    numeroPagina = 1
    paginaTotal = Int(rst.RecordCount / 150) + 1
    cabecalho "Todos", numeroPagina, paginaTotal
    
    rst.MoveFirst
    'pgb.Value = rst.RecordCount
    xScale = 20

For X = 1 To rst.RecordCount
        Y = Y + 1
        nRegistros = nRegistros + 1
        
        Printer.CurrentX = xScale
        Printer.CurrentY = yScale
        Printer.Print rst!idOS
        
        Printer.CurrentX = xScale + 730
        Printer.CurrentY = yScale
        Printer.Print Mid$(rst!nomeCliente, 1, 30)
        
        Printer.CurrentX = xScale + 4480 - Printer.TextWidth(Format(rst!valorOs, "#,##0.00"))
        Printer.CurrentY = yScale
        Printer.Print Format(rst!valorOs, "#,##0.00")
        
        Printer.CurrentX = xScale + 4580
        Printer.CurrentY = yScale
        
        If rst!baixa = Null Then
            Printer.Print "-"
        Else
            Printer.Print rst!baixa
        End If
        
        yScale = yScale + 200
        
        rst.MoveNext
        
        If nRegistros = 75 Then
            xScale = 5615
            yScale = 1000
            nRegistros = 0
        End If

        If Y = 150 Then
            Printer.NewPage
            numeroPagina = numeroPagina + 1
            cabecalho "Todos", numeroPagina, paginaTotal
            yScale = 1000
            xScale = 20
            Y = 0
            nRegistros = 0
        End If
        'pgb.Value = X
Next

If yScale > 13700 Then
    If xScale < 5615 Then
        yScale = 1000
        xScale = 5616
    Else
        Printer.NewPage
        cabecalho "Todos", numeroPagina, paginaTotal
        yScale = 1000
        xScale = 20
    End If
Else
    yScale = yScale + 300
End If
    
    Printer.CurrentX = xScale
    Printer.CurrentY = yScale
    Printer.Print "OS Baixadas..........................................................."
    Printer.CurrentY = yScale
    Printer.CurrentX = xScale + 5375 - Printer.TextWidth(txtValorOsBaixada.Text)
    Printer.Print txtValorOsBaixada.Text
    
    yScale = yScale + 250
    Printer.CurrentY = yScale
    Printer.CurrentX = xScale
    Printer.Print "OS Faturadas.........................................................."
    Printer.CurrentY = yScale
    Printer.CurrentX = xScale + 5375 - Printer.TextWidth(txtValorOsFaturada.Text)
    Printer.Print txtValorOsFaturada.Text
    
    yScale = yScale + 250
    Printer.CurrentX = xScale
    Printer.CurrentY = yScale
    Printer.Print "Total de vendas......................................................."
    Printer.CurrentY = yScale
    Printer.CurrentX = xScale + 5375 - Printer.TextWidth(txtTotal.Text)
    Printer.Print txtTotal.Text
    
    yScale = yScale + 250
    Printer.CurrentX = xScale
    Printer.CurrentY = yScale
    Printer.Print "OS Não Baixadas...................................................."
    Printer.CurrentY = yScale
    Printer.CurrentX = xScale + 5375 - Printer.TextWidth(txtValorOsNaoBaixada.Text)
    Printer.Print txtValorOsNaoBaixada.Text
    
    Printer.EndDoc
    rst.Close
    db.Close




End Sub
