VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRelatorioOs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relat�rio de Vendas por Clientes"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12630
   Icon            =   "frmRelatorioOs.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGeral 
      Caption         =   "Geral"
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
      Left            =   4680
      TabIndex        =   17
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   12
      Top             =   7230
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12621
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
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin MSACAL.Calendar Cal 
      Height          =   2895
      Left            =   5520
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   3735
      _Version        =   524288
      _ExtentX        =   6588
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   12648447
      Year            =   2012
      Month           =   4
      Day             =   25
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   0
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
   Begin VB.Frame Frame1 
      Caption         =   "Filtrar por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7080
      TabIndex        =   6
      Top             =   240
      Width           =   2175
      Begin VB.OptionButton Option1 
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
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "OS n�o baixadas"
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
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
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
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Index           =   1
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
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
      ItemData        =   "frmRelatorioOs.frx":0ECA
      Left            =   120
      List            =   "frmRelatorioOs.frx":0ECC
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
   Begin MSFlexGridLib.MSFlexGrid grdRelatorioOs 
      Height          =   5505
      Left            =   105
      TabIndex        =   8
      Top             =   1680
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   9710
      _Version        =   393216
      Rows            =   19
      Cols            =   7
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.CommandButton cmdCliente 
      Caption         =   "Clientes"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   2070
      Left            =   10920
      Picture         =   "frmRelatorioOs.frx":0ECE
      Top             =   1680
      Visible         =   0   'False
      Width           =   5985
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial"
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
      TabIndex        =   14
      Top             =   960
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Data Final"
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
      Left            =   1800
      TabIndex        =   13
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "C�digo"
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
      Left            =   6120
      TabIndex        =   11
      Top             =   240
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Raz�o Social"
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
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmRelatorioOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public indice As Integer
Public baixa As String
Public somaOs As Currency
Public tipoDeConsulta As String

Private Sub Cal_Click()
    Text1(indice).Text = Format$(Cal.Value, "dd/mm/yy")
    Cal.Visible = False
    Text1(indice).BackColor = &HFFFFFF
End Sub

Private Sub Combo1_Change()
    Text3.Text = ""
End Sub

Private Sub Combo1_Click()
    
    rs.Open "SELECT * FROM cliente WHERE nome='" & Combo1.Text & "'", db, adOpenStatic, adLockOptimistic
    Text3.Text = Format$(rs!idcliente, "###000")
    rs.Close
    
End Sub

Private Sub cmdCliente_Click()

If Text1(0).Text = "" Or Text1(1).Text = "" Then Exit Sub

If "#" & Format$(Text1(0).Text, "mm/dd/yy") & "#" > "#" & Format$(Text1(1).Text, "mm/dd/yy") & "#" Then
   MsgBox "Data inicial maior que data final", vbInformation
   Exit Sub
End If


Dim MyData As Date
Dim OSbaixada As Integer
Dim OSNaoBaixada As Integer
'Dim somaOs As Currency
Dim OStotal As Integer

If Text3.Text = "" Then Exit Sub

'    Select Case baixa
'
'        Case "Baixado"
'            rs.Open "SELECT * FROM os INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE nomecliente='" & Combo1.Text & "' AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND Baixa='Baixado' ORDER BY idos", db, adOpenStatic, adLockOptimistic
'
'        Case "Nao baixado"
'            rs.Open "SELECT * FROM os INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE nomecliente='" & Combo1.Text & "' AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND (Baixa='Faturado' Or IsNull(Baixa)) ORDER BY idos", db, adOpenStatic, adLockOptimistic
'
'        Case "Todos"
'            rs.Open "SELECT * FROM os INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE nomecliente='" & Combo1.Text & "' AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# ORDER BY idos", db, adOpenStatic, adLockOptimistic
'
'    End Select

'Else   'Filtra somente cliente cadastrado

   tipoDeConsulta = "cliente"
   
    Select Case baixa
    
        Case "Baixado"
            rs.Open "SELECT * FROM os INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE idcliente=" & Text3.Text & "AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND baixa='Baixado' ORDER BY idos", db, adOpenStatic, adLockOptimistic
            
        Case "Nao baixado"
            rs.Open "SELECT  * FROM OS INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE IDcliente=" & Text3.Text & " AND Data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# And Data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND (Baixa='Faturado' Or IsNull(Baixa)) ORDER BY idos", db, adOpenStatic, adLockOptimistic
    
        Case "Todos"
            rs.Open "SELECT * FROM os INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE idcliente=" & Text3.Text & "AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# ORDER BY idos", db, adOpenStatic, adLockOptimistic
            
    End Select
    
'End If

If rs.RecordCount = 0 Then
   MsgBox "N�o foi encontrado nenhuma OS neste per�odo.", vbInformation
   rs.Close
   Exit Sub
End If

With grdRelatorioOs
    
    .Clear
    Dim X As Integer
    
    .Rows = rs.RecordCount + 1
    If rs.RecordCount < 19 Then .Rows = 19
        
    For X = 1 To rs.RecordCount
    
        .TextMatrix(X, 0) = X
        .TextMatrix(X, 1) = Format$(rs!idOS, "#####00000")
        .TextMatrix(X, 2) = rs!nomeCliente
        .TextMatrix(X, 3) = Format$(rs!Data, "dd/mm/yy")
        .TextMatrix(X, 4) = Format$(rs!valorOs, "#,##0.00")
        
        If Not IsNull(rs!baixa) Then
            .TextMatrix(X, 5) = rs!baixa
        End If
        .TextMatrix(X, 6) = rs!nome
        
        rs.MoveNext
    Next
    cabecalho
        For X = 1 To 18
        grdRelatorioOs.TextMatrix(X, 0) = X
    Next
End With
    rs.Close

    If Text3.Text = "" Then
        rs.Open "SELECT * FROM os WHERE nomecliente='" & Combo1.Text & "' AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# ORDER BY idos", db, adOpenStatic, adLockOptimistic
        OStotal = rs.RecordCount
        StatusBar1.Panels(1).Text = "Total de OS: " & OStotal
        rs.Close
        
        rs.Open "SELECT * FROM os WHERE nomecliente='" & Combo1.Text & "' AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND baixa='Baixado' ORDER BY idos", db, adOpenStatic, adLockOptimistic
        OSbaixada = rs.RecordCount
        StatusBar1.Panels(2).Text = "O.S. baixada: " & OSbaixada
        rs.Close
        
        rs.Open "SELECT  * FROM OS WHERE nomecliente='" & Combo1.Text & "' AND Data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# And Data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND (Baixa='Faturado' Or IsNull(Baixa)) ORDER BY idos", db, adOpenStatic, adLockOptimistic
        OSbaixada = rs.RecordCount
        StatusBar1.Panels(3).Text = "O.S. n�o baixada: " & OSbaixada
        rs.Close
    Else
        rs.Open "SELECT * FROM os WHERE idcliente=" & Text3.Text & "AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# ORDER BY idos", db, adOpenStatic, adLockOptimistic
        OStotal = rs.RecordCount
        
        rs.Close
        
        rs.Open "SELECT * FROM os WHERE idcliente=" & Text3.Text & "AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND baixa='Baixado' ORDER BY idos", db, adOpenStatic, adLockOptimistic
        OSbaixada = rs.RecordCount
        
        rs.Close
        
        rs.Open "SELECT  * FROM OS WHERE IDcliente=" & Text3.Text & " AND Data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# And Data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND (Baixa='Faturado' Or IsNull(Baixa)) ORDER BY idos", db, adOpenStatic, adLockOptimistic
        OSNaoBaixada = rs.RecordCount
        
        rs.Close
    End If
    
    Select Case baixa
    
        Case "Baixado"
            rs.Open "SELECT sum(valoros) as Teste FROM os WHERE idcliente=" & Text3.Text & "AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND baixa='Baixado'", db, adOpenStatic, adLockOptimistic
            If Not IsNull(rs!teste) Then somaOs = rs!teste
            rs.Close
            
        Case "Nao baixado"
            rs.Open "SELECT sum(valoros) as Teste FROM os WHERE idcliente=" & Text3.Text & "AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND baixa='faturado' AND isNull(Baixa)", db, adOpenStatic, adLockOptimistic
            If Not IsNull(rs!teste) Then somaOs = rs!teste
            rs.Close
    
        Case "Todos"
            rs.Open "SELECT sum(valoros) as Teste FROM os WHERE idcliente=" & Text3.Text & "AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "#", db, adOpenStatic, adLockOptimistic
            If Not IsNull(rs!teste) Then somaOs = rs!teste
            rs.Close
            
    End Select

         StatusBar1.Panels(1).Text = "Total de OS (" & OStotal & ") "
         StatusBar1.Panels(2).Text = "Baixada (" & OSbaixada & ") "
         StatusBar1.Panels(3).Text = "N�o baixada (" & OSNaoBaixada & ") "
         StatusBar1.Panels(4).Text = "Valor total da consulta : R$ " & Format$(somaOs, "#,##0.00")
         cmdImprimir.Enabled = True
         
End Sub

Private Sub cmdImprimir_Click()
    
Dim mes As String
Dim strRecordset As String

With Printer
   
   'Cabe�alho -----------------------------------------------------------------------------
   
   CabecalhoImpressao
   
   'Fim do cabe�alho-----------------------------------------------------------------------

   If tipoDeConsulta = "cliente" Then

      Select Case baixa
      
          Case "Baixado"
              rs.Open "SELECT * FROM os WHERE nomecliente='" & Combo1.Text & "' AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND Baixa='Baixado' ORDER BY idos", db, adOpenStatic, adLockOptimistic
              
          Case "Nao baixado"
              rs.Open "SELECT * FROM os WHERE nomecliente='" & Combo1.Text & "' AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND (Baixa='Faturado' Or IsNull(Baixa)) ORDER BY idos", db, adOpenStatic, adLockOptimistic
              
          Case "Todos"
              rs.Open "SELECT * FROM os WHERE nomecliente='" & Combo1.Text & "' AND data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# ORDER BY idos", db, adOpenStatic, adLockOptimistic
      End Select
   
   Else
   
      Select Case baixa
      
          Case "Baixado"
              rs.Open "SELECT * FROM os WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND Baixa='Baixado' ORDER BY idos", db, adOpenStatic, adLockOptimistic
              
          Case "Nao baixado"
              rs.Open "SELECT * FROM os WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND (Baixa='Faturado' Or IsNull(Baixa)) ORDER BY idos", db, adOpenStatic, adLockOptimistic
              
          Case "Todos"
              rs.Open "SELECT * FROM os WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# ORDER BY idos", db, adOpenStatic, adLockOptimistic
      End Select
   
   
   End If
    
    
   Dim contOS As Integer
   Dim contOSporPagina As Integer
   Dim nPagina As Integer
   Dim totalPagina As Integer
   
   nPagina = 1
   contOSporPagina = 1
   contOS = 1
   'totalPagina = Int(rs.RecordCount / 66)
   
   If Int(rs.RecordCount / 66) / (rs.RecordCount / 66) = 1 Then
      totalPagina = Int(rs.RecordCount / 66)
   Else
      totalPagina = Int(rs.RecordCount / 66) + 1
   End If
   
   Dim xScale As Long
   Dim yScale As Long
   
   xScale = 900
   yScale = 3000
   
   .CurrentX = 500
   .CurrentY = 2100
   .FontBold = True
   Printer.Print "Valor total: R$ "; Format$(somaOs, "#,##0.00")
   .FontBold = False
   
   While Not rs.EOF
      
      .FontSize = 9
      .FontBold = False
      .CurrentX = 11350 - .TextWidth("P�gina: " & nPagina & "/" & totalPagina)
      .CurrentY = 2250
      Printer.Print "P�gina: " & nPagina & "/" & totalPagina
      
      .FontSize = 8
      .CurrentX = xScale
      .CurrentY = yScale
      Printer.Print Format$(contOS, "00")
      
      .CurrentX = xScale + 850
      .CurrentY = yScale
      Printer.Print rs!idOS & ""
      
      .CurrentX = xScale + 2000
      .CurrentY = yScale
      Printer.Print rs!Data
      
      .CurrentX = xScale + 4100 - Printer.TextWidth(Format$(rs!valorOs, "#,##0.00"))
      .CurrentY = yScale
      Printer.Print Format$(rs!valorOs, "#,##0.00")
      
      .CurrentY = yScale
      If Not IsNull(rs!baixa) Then
         .CurrentX = 5850 - (.TextWidth(rs!baixa) / 2)
         Printer.Print rs!baixa
      End If
      
      .CurrentX = xScale + 5700
      .CurrentY = yScale
      Printer.Print rs!nomeCliente
      
      yScale = yScale + 200
      rs.MoveNext
      
      contOS = contOS + 1
      contOSporPagina = contOSporPagina + 1
      
      If contOSporPagina > 66 Then
         Printer.NewPage
         CabecalhoImpressao
         yScale = 3000
         contOSporPagina = 1
         nPagina = nPagina + 1
      End If
   
   Wend
   
   rs.Close
   Printer.EndDoc
End With

End Sub

Private Sub cmdGeral_Click()

If Text1(0).Text = "" Or Text1(1).Text = "" Then Exit Sub

If "#" & Format$(Text1(0).Text, "mm/dd/yy") & "#" > "#" & Format$(Text1(1).Text, "mm/dd/yy") & "#" Then
   MsgBox "Data inicial maior que data final", vbInformation
   Exit Sub
End If

Dim MyData As Date
Dim OSbaixada As Integer
Dim OSNaoBaixada As Integer
'Dim somaOs As Currency
Dim OStotal As Integer

'If Text3.Text = "" Then 'Filtra qualquer cliente

'    Select Case baixa
'
'        Case "Baixado"
'            rs.Open "SELECT * FROM os INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND Baixa='Baixado' ORDER BY idos", db, adOpenStatic, adLockOptimistic
'
'        Case "Nao baixado"
'            rs.Open "SELECT * FROM os INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND (Baixa='Faturado' Or IsNull(Baixa)) ORDER BY idos", db, adOpenStatic, adLockOptimistic
'
'        Case "Todos"
'            rs.Open "SELECT * FROM os INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# ORDER BY idos", db, adOpenStatic, adLockOptimistic
'
'    End Select

'Else   'Filtra somente cliente cadastrado

   tipoDeConsulta = "geral"
   
    Select Case baixa
    
        Case "Baixado"
            rs.Open "SELECT * FROM os INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND baixa='Baixado' ORDER BY idos", db, adOpenStatic, adLockOptimistic
            
        Case "Nao baixado"
            rs.Open "SELECT  * FROM OS INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE Data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# And Data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND (Baixa='Faturado' Or IsNull(Baixa)) ORDER BY idos", db, adOpenStatic, adLockOptimistic
    
        Case "Todos"
            rs.Open "SELECT * FROM os INNER JOIN usuario ON os.idusuario=usuario.codigousuario WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# ORDER BY idos", db, adOpenStatic, adLockOptimistic
            
    End Select
    
'End If

If rs.RecordCount = 0 Then
   MsgBox "N�o foi encontrado nenhuma OS neste per�odo.", vbInformation
   rs.Close
   Exit Sub
End If

With grdRelatorioOs
    
    .Clear
    Dim X As Integer
    
    .Rows = rs.RecordCount + 1
    If rs.RecordCount < 19 Then .Rows = 19
        
    For X = 1 To rs.RecordCount
    
        .TextMatrix(X, 0) = X
        .TextMatrix(X, 1) = Format$(rs!idOS, "#####0000")
        .TextMatrix(X, 2) = rs!nomeCliente
        .TextMatrix(X, 3) = Format$(rs!Data, "dd/mm/yy")
        .TextMatrix(X, 4) = Format$(rs!valorOs, "#,##0.00")
        
        If Not IsNull(rs!baixa) Then
            .TextMatrix(X, 5) = rs!baixa
        End If
        .TextMatrix(X, 6) = rs!nome
        
        rs.MoveNext
    Next
    cabecalho
        For X = 1 To 18
        grdRelatorioOs.TextMatrix(X, 0) = X
    Next
End With
    rs.Close

      rs.Open "SELECT * FROM os WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# ORDER BY idos", db, adOpenStatic, adLockOptimistic
      OStotal = rs.RecordCount
      
      rs.Close
      
      rs.Open "SELECT * FROM os WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND baixa='Baixado' ORDER BY idos", db, adOpenStatic, adLockOptimistic
      OSbaixada = rs.RecordCount
      
      rs.Close
      
      rs.Open "SELECT  * FROM OS WHERE Data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# And Data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND (Baixa='Faturado' Or IsNull(Baixa)) ORDER BY idos", db, adOpenStatic, adLockOptimistic
      OSNaoBaixada = rs.RecordCount
      
      rs.Close
    
    Select Case baixa
    
        Case "Baixado"
            rs.Open "SELECT sum(valoros) as Teste FROM os WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND baixa='Baixado'", db, adOpenStatic, adLockOptimistic
            If Not IsNull(rs!teste) Then somaOs = rs!teste
            rs.Close
            
        Case "Nao baixado"
            rs.Open "SELECT sum(valoros) as Teste FROM os WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "# AND (Baixa='Faturado' Or IsNull(Baixa))", db, adOpenStatic, adLockOptimistic
            If Not IsNull(rs!teste) Then somaOs = rs!teste
            rs.Close
    
        Case "Todos"
            rs.Open "SELECT sum(valoros) as Teste FROM os WHERE data>=#" & Format$(Text1(0).Text, "mm/dd/yy") & "# AND data<=#" & Format$(Text1(1).Text, "mm/dd/yy") & "#", db, adOpenStatic, adLockOptimistic
            If Not IsNull(rs!teste) Then somaOs = rs!teste
            rs.Close
            
    End Select

         StatusBar1.Panels(1).Text = "Total de OS (" & OStotal & ") "
         StatusBar1.Panels(2).Text = "Baixada (" & OSbaixada & ") "
         StatusBar1.Panels(3).Text = "N�o baixada (" & OSNaoBaixada & ") "
         StatusBar1.Panels(4).Text = "Valor total da consulta : R$ " & Format$(somaOs, "#,##0.00")
         cmdImprimir.Enabled = True

End Sub


Private Sub Form_Load()
    
    Me.Width = 12720
    abreConexao
    
    rs.Open "SELECT * FROM cliente", db, adOpenStatic, adLockOptimistic
    Cal.Month = Month(Date)
    Cal.Year = Year(Date)
    Dim X As Integer
    
    For X = 1 To rs.RecordCount
        Combo1.AddItem rs!nome
        rs.MoveNext
    Next
    cabecalho
    rs.Close
    baixa = "Todos"
    
    For X = 1 To 18
        grdRelatorioOs.TextMatrix(X, 0) = X
    Next
    
End Sub


Private Sub Option1_Click(Index As Integer)

    Select Case Index
        Case 0
            baixa = "Baixado"
        Case 1
            baixa = "Nao baixado"
        Case 2
            baixa = "Todos"
    End Select
End Sub

Private Sub Text1_Click(Index As Integer)
    Cal.Visible = True
    Text1(Index).BackColor = &HC0FFFF
    indice = Index
    
    Cal.Top = Text1(Index).Height + Text1(0).Top '1560
    Cal.Left = Text1(Index).Left
End Sub

Public Sub cabecalho()

With grdRelatorioOs
    .FormatString = "^ |^N� OS|<Nome|^Data|>Valor|^Baixa|Atendente"
    
    .ColWidth(0) = 500
    .ColWidth(1) = 1000
    .ColWidth(2) = 3000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 3050
    
End With

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(indice).BackColor = &HFFFFFF
End Sub

Public Sub CabecalhoImpressao()

With Printer
   .CurrentY = 800
   .CurrentX = 11450 - .TextWidth("RELAT�RIO DE OS")
   Printer.Print "RELAT�RIO DE OS"
   
   Printer.Line (500, .CurrentY + 100)-(11450, .CurrentY + 100)
   

   
   .CurrentX = 560
   .CurrentY = .CurrentY + 100
   .FontSize = 14
   Printer.Print Combo1.Text
   
   Printer.Line (500, .CurrentY + 50)-(11450, .CurrentY + 50)
   
   .CurrentX = 500
   .CurrentY = .CurrentY + 100
   .FontSize = 10
   
   If Option1(0).Value = True Then
      Printer.Print "OS baixada no per�odo de " & Text1(0).Text & " � " & Text1(1).Text
   ElseIf Option1(1).Value = True Then
      Printer.Print "OS n�o baixada no per�odo de " & Text1(0).Text & " � " & Text1(1).Text
   Else
      Printer.Print "OS emitida no per�odo de " & Text1(0).Text & " � " & Text1(1).Text
   End If
   
   .FontBold = True
   
   .CurrentX = 650
   .CurrentY = 2600
   Printer.Print "Ordem"
   
   .CurrentX = 1700
   .CurrentY = 2600
   Printer.Print "OS"
   
   .CurrentX = 3100
   .CurrentY = 2600
   Printer.Print "Data"
   
   .CurrentX = 4450
   .CurrentY = 2600
   Printer.Print "Valor"
   
   .CurrentX = 5600
   .CurrentY = 2600
   Printer.Print "Baixa"
   
   .CurrentX = 6600
   .CurrentY = 2600
   Printer.Print "Cliente"
   
   .FontBold = False
   
   Printer.Line (1400, 2500)-(1500, 16200) ' coluna 1
   Printer.Line (2450, 2500)-(2650, 16200) ' coluna 2
   Printer.Line (4050, 2500)-(4350, 16200) ' coluna 3
   Printer.Line (5250, 2500)-(5450, 16200) ' coluna 4
   Printer.Line (6400, 2500)-(6600, 16200) ' coluna 5
   Printer.Line (500, 2500)-(11450, 2500) ' cabe�a1
   Printer.Line (500, 2900)-(11450, 2900) ' cabe�a2
   Printer.Line (500, 16200)-(11450, 16200) ' rodap�
End With

End Sub

