VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSalvar 
   Caption         =   "IndigoWorks"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14790
   Icon            =   "frmSalvar.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10710
   ScaleWidth      =   14790
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cd 
      Left            =   4440
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000C&
      Height          =   11775
      Left            =   7200
      ScaleHeight     =   11715
      ScaleWidth      =   8835
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   8895
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   10935
         Left            =   1680
         ScaleHeight     =   10905
         ScaleWidth      =   8310
         TabIndex        =   5
         Top             =   1440
         Width           =   8340
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Visualizar >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salvar e Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   2400
      Picture         =   "frmSalvar.frx":000C
      Top             =   5760
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   2070
      Left            =   720
      Picture         =   "frmSalvar.frx":09B8
      Top             =   3360
      Visible         =   0   'False
      Width           =   5985
   End
End
Attribute VB_Name = "frmSalvar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modoVisao As Boolean
  Private Function ScalePicPreviewToPrinterInches _
         (picPreview As PictureBox) As Double

         Dim ratio As Double ' raio entre a impressora e a imagem
         Dim LRGap As Double, TBGap As Double
         Dim HeightRatio As Double, WidthRatio As Double
         Dim PgWidth As Double, PgHeight As Double
         Dim smtemp As Long

         ' tamanho da pagina em polegadas
         PgWidth = Printer.Width / 56.693
         PgHeight = Printer.Height / 56.693

         ' tamanho da area que nao é de impressao
         smtemp = Printer.ScaleMode
         Printer.ScaleMode = 6
         LRGap = (PgWidth - Printer.ScaleWidth) / 2
         TBGap = (PgHeight - Printer.ScaleHeight) / 2
         Printer.ScaleMode = smtemp

         ' define o tamanho da imagem na area de impressao
         picPreview.ScaleMode = 6

         ' Compara a altura e o raio para determinar
         ' o tamanho da imagem
         HeightRatio = picPreview.ScaleHeight / PgHeight
         WidthRatio = picPreview.ScaleWidth / PgWidth

         If HeightRatio < WidthRatio Then
            ratio = HeightRatio
            smtemp = picPreview.Container.ScaleMode
            picPreview.Container.ScaleMode = 6
            picPreview.Width = PgWidth * ratio
            picPreview.Container.ScaleMode = smtemp
         Else
            ratio = WidthRatio
            smtemp = picPreview.Container.ScaleMode
            picPreview.Container.ScaleMode = 6
            picPreview.Height = PgHeight * ratio
            picPreview.Container.ScaleMode = smtemp
         End If

         ' define propriedades da imagem
         ' voce pode acrescentar mais itens aqui
         picPreview.Scale (0, 0)-(PgWidth, PgHeight)
         picPreview.Font.Name = Printer.Font.Name
         picPreview.FontSize = Printer.FontSize * ratio
         picPreview.ForeColor = Printer.ForeColor
         picPreview.Cls

         ScalePicPreviewToPrinterInches = ratio
      End Function

Private Sub PrintRoutine(objPrint As Object, Optional ratio As Double = 1)
        
    Dim md As Integer ' Margem direita
    md = 13
    
    
    Dim texto As String
    
    desenhaQuadro objPrint, ratio
    
    With objPrint

    objPrint.PaintPicture Image1.Picture, 17.6, 15.12, 33.782, 11.684
    objPrint.PaintPicture Image2.Picture, 154, 21.216, 42.618, 13.26
    
    End With

 End Sub

Private Sub Command1_Click()

    
    'salvarRegistro
    'Unload Me

End Sub


Private Sub Command2_Click(Index As Integer)
    
    Select Case Index
        Case 0
         Printer.ScaleMode = 6
         PrintRoutine Printer
         Printer.EndDoc
         
        Case 1
            Unload Me
            
    End Select

End Sub

Public Sub Command4_Click()
  
If Command4.Caption = "Visualizar >>" Then
    If modoVisao = False Then
        modoVisao = True
        Picture2.Left = 4560
        Picture2.Top = 0
        
        Picture2.Visible = False
        
    
                 
        Picture2.Height = Me.ScaleHeight
        Picture2.Width = Me.ScaleWidth - Picture2.Left
        
        
        Picture1.Left = Picture2.ScaleWidth / 2 - Picture1.Width / 2
        Picture1.Top = Picture2.ScaleHeight / 2 - Picture1.Height / 2
        
        Me.Width = 13500
        Me.Height = 12150
        
        Me.Top = Screen.Height / 2 - Me.Height / 2
        Me.Left = Screen.Width / 2 - Me.Width / 2
        
        Picture2.Visible = True
        Command4.Caption = "Visualizar <<"
    End If
Else
    modoVisao = False
    Command4.Caption = "Visualizar >>"
    Me.WindowState = 0

    Me.Width = 4695
    Me.Height = 2760
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Picture2.Visible = False
End If
End Sub


Private Sub Form_Load()
modoVisao = False
Command1.Caption = Printer.DeviceName
End Sub

Public Sub Form_Resize()
Picture2.Visible = False
If modoVisao Then

    If Me.Height > 6345 Then
        Picture2.Height = Me.ScaleHeight
    End If

    If Me.Width > 9195 Then
        Picture2.Width = Me.ScaleWidth - Picture2.Left
    End If

    Picture1.Height = Picture2.ScaleHeight - 100
    Picture1.Width = Picture1.Height * 0.7

    Picture1.Left = Picture2.ScaleWidth / 2 - Picture1.Width / 2
    Picture1.Top = Picture2.ScaleHeight / 2 - Picture1.Height / 2 + 50

    Picture2.Visible = True
    montaPreview
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Public Sub salvarRegistro()


'On Error GoTo Trata_Erro
    
    abreConexao
    
' Tabela Cliente
'---------------------------------------------------------------------------
    If Len(frmOrcamento.MaskEdBox1.Text) = 11 Then
    
        rs.Open "SELECT * FROM cliente", db, adOpenStatic, adLockOptimistic
        rs.AddNew
        
        rs!nome = frmOrcamento.txtCliente(0).Text
        rs!telefone = Format$(frmOrcamento.txtTelefone, "(@@)@@@@-@@@@")
        rs!endereco = frmOrcamento.txtCliente(1).Text
        rs!bairro = frmOrcamento.txtCliente(2).Text
        rs!cpf = Format$(frmOrcamento.MaskEdBox1.Text, "@@@.@@@.@@@-@@")
        rs.Update
            
        Dim codigoCliente As String
        codigoCliente = rs!idcliente
        rs.Close
        
    End If


' Tabela OS
'---------------------------------------------------------------------------
        rs.Open "Select * from os", db, adOpenStatic, adLockOptimistic
        rs.AddNew
                
        rs!IDUsuario = varCodigoUsuario
        rs!Data = CDate(frmOrcamento.txtData)
        rs!hora = CDate(frmOrcamento.txtHora)
        rs!alteradopor = frmOrcamento.txtAlteradoPor.Text
        rs!Exemplar = frmOrcamento.txtExemplar.Text
        rs!outros = frmOrcamento.txtAcrescimo.Text
        rs!desconto = frmOrcamento.txtDesconto.Text
                
        If Len(frmOrcamento.MaskEdBox1.Text) = 11 Then
            rs!idcliente = codigoCliente
        Else
            rs!nomecliente = frmOrcamento.txtCliente(0).Text
            rs!telefonecliente = Format$(frmOrcamento.txtTelefone, "(@@)@@@@-@@@@")
        End If
        rs.Update

        Dim codigoOs As Integer
        codigoOs = rs!idos
        rs.Close

' Tabela Plano
'---------------------------------------------------------------------------
        Dim w As Integer
        
        rs.Open "SELECT * FROM plano", db, adOpenStatic, adLockOptimistic
        For w = 1 To frmOrcamento.contadorDePlano
            rs.AddNew
            
            rs!idos = codigoOs
            rs!quantidade = frmOrcamento.gd.TextMatrix(w, 1)
            rs!descricao = frmOrcamento.gd.TextMatrix(w, 2)
            rs!midia = frmOrcamento.gd.TextMatrix(w, 3)
            rs!cores = frmOrcamento.gd.TextMatrix(w, 4)
            rs!Laminacao = frmOrcamento.gd.TextMatrix(w, 5)
            rs!capa = frmOrcamento.gd.TextMatrix(w, 6)
            rs!Wireo = frmOrcamento.gd.TextMatrix(w, 7)
            
            If frmOrcamento.gd.TextMatrix(w, 8) = "" Then
                rs!corte = 0
            Else
                rs!corte = frmOrcamento.gd.TextMatrix(w, 8)
            End If
            
            If frmOrcamento.gd.TextMatrix(w, 9) = "" Then
                rs!meiocorte = 0
            Else
                rs!meiocorte = frmOrcamento.gd.TextMatrix(w, 9)
            End If
            
            If frmOrcamento.gd.TextMatrix(w, 10) = "" Then
                rs!transfer = 0
            Else
                rs!transfer = frmOrcamento.gd.TextMatrix(w, 10)
            End If
            
            rs!valor = frmOrcamento.gd.TextMatrix(w, 11)
            
            rs.Update
        Next
            rs.Close
            db.Close

Exit Sub

'Trata_Erro:
    'MsgBox "houve um erro ai - " & Err

End Sub








Private Sub desenhaQuadro(objPrint As Object, Optional ratio As Double = 1)
    
    Dim w As Integer
    Dim yIni As Double
    Dim yFim As Double
    Dim xIni As Double
    Dim xFim As Double
    Dim linha As Double
    
    
  With objPrint
    
    
    .FontSize = 8 * ratio
    .CurrentX = 35 - objPrint.TextWidth("Rua das Nações, 212 - Bom Pastor") / 2
    .CurrentY = 34.31
    objPrint.Print "Rua das Nações, 212 - Bom Pastor"
    
    objPrint.FontSize = 10 * ratio
    objPrint.CurrentX = 35 - objPrint.TextWidth("(37)3691-7000") / 2
    objPrint.CurrentY = 37.43
    objPrint.Print "(37)3691-7000"

objPrint.FontBold = True
    objPrint.FontSize = 14 * ratio
    objPrint.CurrentX = objPrint.ScaleWidth / 2 - objPrint.TextWidth("Ordem de Serviço Nº 001") / 2
    objPrint.CurrentY = 14
    objPrint.Print "Ordem de Serviço Nº 001"
objPrint.FontBold = False
    
    xIni = 10
    yIni = 48.5
    For w = 1 To 5
    
        Select Case w
            Case 1: xIni = 10
            Case 2: xIni = 73.33
            Case 3: xIni = 136.66
            Case 4
                xIni = 10
                yIni = 98.5
            Case 5
                xIni = 73.33
        End Select
        
        .FontSize = 10 * ratio
        .CurrentX = xIni + 26
        .CurrentY = yIni - 5.2
        objPrint.Print "Plano " & w
        
        .FontSize = 9 * ratio
        .CurrentX = xIni + 1
        .CurrentY = yIni
        objPrint.Print "Quan."
        
        .CurrentX = xIni + 15.5
        .CurrentY = yIni
        objPrint.Print "Descrição"
        
        .CurrentX = xIni + 36.8
        .CurrentY = yIni
        objPrint.Print "V. Unit."
        
        .CurrentX = xIni + 49.3
        .CurrentY = yIni
        objPrint.Print "Sub Total"

    Next
    
    If 0.5 * ratio > 0.5 Then
        objPrint.DrawWidth = 1 * ratio
    End If
    objPrint.Line (73.33, 42.5)-(73.33, 142.65)
    objPrint.Line (136.7, 42.5)-(136.7, 142.65)
    
    ' Moldura
    objPrint.Line (10, 10)-(200, 10)
    objPrint.Line (10, 10)-(10, 142.65)
    objPrint.Line (200, 10)-(200, 142.65)
    
    objPrint.Line (60, 10)-(60, 42.5)
    objPrint.Line (150, 10)-(150, 42.5)
    '-------------------------------------

    yIni = 47.6
    yFim = 87.65
    
    For w = 1 To 2
        objPrint.Line (20.27, yIni)-(20.27, yFim)
        objPrint.Line (44.6, yIni)-(44.6, yFim)
        objPrint.Line (59, yIni)-(59, yFim + 5)
        objPrint.Line (83.6, yIni)-(83.6, yFim)
        objPrint.Line (107.7, yIni)-(107.7, yFim)
        objPrint.Line (121.6, yIni)-(121.6, yFim + 5)
        
        If w < 2 Then
            objPrint.Line (147, yIni)-(147, yFim)
            objPrint.Line (171, yIni)-(171, yFim)
            objPrint.Line (185, yIni)-(185, yFim + 5)
        End If
        
        yIni = 97.63
        yFim = 137.65
        
    Next
    
    linha = 42.5
    xFim = 200
    For w = 1 To 21
        If w > 11 And w < 21 Then
            xFim = 136.66
        Else
            xFim = 200
        End If
        objPrint.Line (10, linha)-(xFim, linha)
        linha = linha + 5
    Next
End With
    
End Sub

Public Sub montaPreview()
    Dim dRatio As Double
    dRatio = ScalePicPreviewToPrinterInches(Picture1)
    PrintRoutine Picture1, dRatio
End Sub
