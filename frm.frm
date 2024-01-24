VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm 
   BackColor       =   &H8000000C&
   Caption         =   "Form1"
   ClientHeight    =   12300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   21795
   Icon            =   "frm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   12300
   ScaleWidth      =   21795
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10080
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm.frx":15C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm.frx":1CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm.frx":2438
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   21795
      _ExtentX        =   38444
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "impressora"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "largura"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "altura"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "cimpressora"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar scrollbar2 
      Height          =   255
      Left            =   0
      Max             =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7680
      Width           =   6015
   End
   Begin VB.VScrollBar Scrollbar1 
      Height          =   3735
      Left            =   6360
      SmallChange     =   500
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3840
      Width           =   270
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   259
      Left            =   9960
      ScaleHeight     =   255
      ScaleWidth      =   270
      TabIndex        =   1
      Top             =   6480
      Width           =   270
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6945
      ScaleWidth      =   4785
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   2070
      Left            =   7320
      Picture         =   "frm.frx":2B32
      Top             =   720
      Visible         =   0   'False
      Width           =   5985
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   11040
      Picture         =   "frm.frx":62C6
      Top             =   4560
      Visible         =   0   'False
      Width           =   2700
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modoLargura As Boolean
Dim modoAltura As Boolean
Dim topInicial As Double
Public numeroCliente As Double
Public numeroOs As Double
Dim impressoraPadrao As String
Dim impressoraEscolhida As String
Private Function calculaWireo(texto As String, modo As Integer) As String

   Dim wireoTexto As String
   
With frmOrcamento
   
   wireoTexto = .preencheCampo(texto, 2)
   
   Select Case modo
      Case 1
         calculaWireo = .preencheCampo(texto, 1)
      Case 2
         calculaWireo = wireoTexto
      Case 3
         Select Case wireoTexto
            Case "1/4", "5/16", "3/8"
                calculaWireo = 1.5
             Case "7/16", "1/2", "9/16"
                calculaWireo = 2
             Case "5/8"
                calculaWireo = 2.5
             Case "7/8", "1"
                calculaWireo = 3
         End Select
   End Select
End With

End Function
Private Function calculaCapa(texto As String, modo As Integer) As String

   Dim capaTexto As String

With frmOrcamento
   
   Select Case modo
      Case 1
         calculaCapa = .preencheCampo(texto, 1)
      Case 2
         calculaCapa = .preencheCampo(texto, 2)
      Case 3
         capaTexto = .preencheCampo(texto, 2)
         Select Case capaTexto
            Case "PPA5"
               calculaCapa = 3
            Case "PPA4"
               calculaCapa = 3
            Case "PPA3"
               calculaCapa = 6
         End Select
   End Select
End With

End Function
Private Function calculaSubTotal(texto As String, cor As String, formato As String) As Double
   
   Dim midia As String
   Dim grama As String
   
   midia = frmOrcamento.preencheCampo(texto, 1)
   grama = frmOrcamento.preencheCampo(texto, 2)
   
   With frmOrcamento
   
   Select Case midia
      
      Case "CL", "CF"
         Select Case grama
            Case "115g", "120g", "150g", "170g"
               calculaSubTotal = 3.5
            Case "250g", "300g"
               calculaSubTotal = 4
         End Select
         
      Case "AP"
         Select Case grama
            Case "75g", "90g", "120g"
               calculaSubTotal = 3
            Case "150g", "180g", "240g"
               calculaSubTotal = 3.5
         End Select
         
      Case "AD"
         calculaSubTotal = 5

      Case "CP"
         calculaSubTotal = 5
                  
      Case "PA"
         calculaSubTotal = 6
      
      Case "K"
         calculaSubTotal = 4
      
      Case "BOPP"
         calculaSubTotal = 6
      
   End Select
   
   
      If cor = "1" Then
         If formato = "A4" Then
            calculaSubTotal = 0.2
         Else
            calculaSubTotal = 0.4
         End If
      End If
      
      If formato = "A4" And cor = "4" Then
         calculaSubTotal = calculaSubTotal / 2
      End If
      
   'If descricao = "Impressos F/V" Then
   '   calculaSubTotal = calculaSubTotal * 2
   'End If
   End With
End Function
Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim l As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
    String, PrinterPort As String)

    Dim iDriver As Integer
    Dim iPort As Integer
    DriverName = ""
    PrinterPort = ""

    ' The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

         ' Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        ' The port name is the second entry after the driver name
        ' separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            ' Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If
End Sub
Private Sub defineImpressoraPadrao(impressora As String)
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim PrinterName As String
    Dim r As Long
    If impressora <> "" Then
        ' Get the printer information for the currently selected
        ' printer in the list. The information is taken from the
        ' WIN.INI file.
        Buffer = Space(1024)
        PrinterName = impressora
        r = GetProfileString("PrinterPorts", PrinterName, "", _
            Buffer, Len(Buffer))

        ' Parse the driver name and port name out of the buffer
        GetDriverAndPort Buffer, DriverName, PrinterPort

           If DriverName <> "" And PrinterPort <> "" Then
               SetDefaultPrinter impressora, DriverName, PrinterPort
               If Printer.DeviceName <> impressora Then
               ' Make sure Printer object is set to the new printer
                  SelectPrinter (impressora)
               End If
           End If
    End If
End Sub
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
   
   Dim texto As String
    
With objPrint
   desenhaQuadro objPrint, ratio
   
   objPrint.FontBold = True
   objPrint.FontSize = 14 * ratio
   objPrint.CurrentX = objPrint.ScaleWidth / 2 - objPrint.TextWidth("Ordem de Serviço Nº " & Format$(numeroOs, "###0000")) / 2
   objPrint.CurrentY = 14
   objPrint.Print "Ordem de Serviço Nº " & Format$(numeroOs, "###0000")
   objPrint.FontBold = False
   
   objPrint.FontSize = 9 * ratio
   .CurrentX = 63
   .CurrentY = 23.3
   objPrint.Print "Nome: " & frmOrcamento.txtCliente(0).Text
   
   .CurrentX = 63
   .CurrentY = 27
   objPrint.Print "Endereço: " & frmOrcamento.txtCliente(1).Text & " - " & frmOrcamento.txtCliente(2).Text
   
   .CurrentX = 63
   .CurrentY = 30.5
   objPrint.Print "Telefone: " & Format$(frmOrcamento.txtTelefone, "(@@)@@@@-@@@@")
   
   .CurrentX = 63
   .CurrentY = 37.5
   objPrint.Print "Atendente: " & frmOrcamento.txtUsuario
   
   .CurrentX = 63
   .CurrentY = 34
   If frmOrcamento.Option1(0).Value = True Then
      objPrint.Print "CPF: " & Format$(frmOrcamento.txtCPF.Text, "@@@.@@@.@@@-@@")
   Else
      objPrint.Print "CNPJ: " & Format$(frmOrcamento.txtCPF.Text, "@@.@@@.@@@/@@@@-@@")
   End If
'Planos
   Dim w As Integer
   Dim z As Integer
   Dim yIni As Double
   Dim xIni As Double
   Dim pos As Double
    
   For w = 1 To frmOrcamento.contadorDePlano
      Select Case w
         Case 1
            yIni = 53.3
            xIni = 19.27
         Case 2
            yIni = 53.3
            xIni = 82.61
         Case 3
            yIni = 53.3
            xIni = 146
      End Select
            
      .FontSize = 9 * ratio
      
      'DESCRIÇÃO IMPRESSOS
      Dim descricao As String
      Dim valorUnitario As Double
      Dim lado As Integer
      Dim subTotal As Double
            
      'armazena todas as colunas
      Dim quantidade As String
      Dim descricao2 As String
      Dim formato As String
      Dim midia As String
      Dim cor As String
      Dim laminacao As String
      Dim encadernacao As String
      Dim capa As String
      Dim wireo As String
      Dim corte As String
      Dim laser As String
      Dim picote As String
      Dim vinco As String
      Dim photobook As Integer

With frmOrcamento.gd
      quantidade = .TextMatrix(w, 1)
      descricao2 = .TextMatrix(w, 2)
      formato = .TextMatrix(w, 3)
      midia = .TextMatrix(w, 4)
      cor = .TextMatrix(w, 5)
      laminacao = .TextMatrix(w, 6)
      encadernacao = .TextMatrix(w, 7)
      capa = .TextMatrix(w, 8)
      wireo = .TextMatrix(w, 9)
      corte = .TextMatrix(w, 10)
      laser = .TextMatrix(w, 11)
      picote = .TextMatrix(w, 12)
      vinco = .TextMatrix(w, 13)
      photobook = frmOrcamento.Check4.Value
End With

'PHOTOBOOK
      If frmOrcamento.gd.TextMatrix(1, 2) = "Capa " Then
                 
         'Quantidade capa
         'pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(1, "##00"))
         objPrint.Print Format$(1, "##00")
         
         'Descrição capa
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         objPrint.Print "C" & ":" & frmOrcamento.gd.TextMatrix(1, 3) & ":" & frmOrcamento.gd.TextMatrix(1, 6)
         
         'Valor Unitário capa
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(12, "#,##0.00"))
         objPrint.Print Format$(12, "#,##0.00")
         
         'Sub total capa
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(12, "#,##0.00"))
         objPrint.Print Format$(12, "#,##0.00")

         'Quantidade miolo
         pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(frmOrcamento.gd.TextMatrix(2, 1), "##00"))
         objPrint.Print Format$(frmOrcamento.gd.TextMatrix(2, 1), "##00")
         
         'Descrição miolo
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         objPrint.Print "M" & ":" & frmOrcamento.gd.TextMatrix(1, 3) & ":" & frmOrcamento.gd.TextMatrix(1, 6)
         
         'Valor Unitário miolo
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(8, "#,##0.00"))
         objPrint.Print Format$(8, "#,##0.00")
         
         'Sub total miolo
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(frmOrcamento.gd.TextMatrix(2, 14), "#,##0.00"))
         objPrint.Print Format$(frmOrcamento.gd.TextMatrix(2, 14), "#,##0.00")

         'Quantidade montagem
         pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(1, "##00"))
         objPrint.Print Format$(1, "##00")
         
         'Descrição montagem
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         objPrint.Print "Montagem"
         
         'Sub total montagem
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(30, "#,##0.00"))
         objPrint.Print Format$(30, "#,##0.00")
         
         GoTo ponto:
      End If
      
'IMPRESSOS
      pos = pos - 5
      If descricao2 = "Impressos" Or descricao2 = "Impressos F/V" Then
         pos = pos + 5
         'yIni = yIni
         .CurrentY = yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(frmOrcamento.gd.TextMatrix(w, 1), "##00"))
         objPrint.Print Format$(frmOrcamento.gd.TextMatrix(w, 1), "##00")
         
         If descricao2 = "Impressos" Then
            descricao = frmOrcamento.gd.TextMatrix(w, 3) & ":" & frmOrcamento.gd.TextMatrix(w, 4)
            valorUnitario = calculaSubTotal(midia, cor, formato)
         ElseIf descricao2 = "Impressos F/V" Then
            descricao = frmOrcamento.gd.TextMatrix(w, 3) & ":" & midia & ":" & "F/V"
            valorUnitario = calculaSubTotal(midia, cor, formato) * 2
         End If
         
         .CurrentY = yIni
         .CurrentX = xIni + 2
         objPrint.Print descricao
         
         'Valor unitário
         .CurrentY = yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(valorUnitario, "#,##0.00"))
         objPrint.Print Format$(valorUnitario, "#,##0.00")
         
         subTotal = valorUnitario * quantidade
         
         'Sub total
         .CurrentY = yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(subTotal, "#,##0.00"))
         objPrint.Print Format$(subTotal, "#,##0.00")
   
      End If

'FOTOS
      If descricao2 = "Fotos" Then
         'Quantidade
         pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(quantidade, "##00"))
         objPrint.Print Format$(quantidade, "##00")
         
         'Descrição
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         objPrint.Print descricao2
         
         'Valor Unitário
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(0.6, "#,##0.00"))
         objPrint.Print Format$(0.6, "#,##0.00")
         
         'Sub total
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(0.6 * quantidade, "#,##0.00"))
         objPrint.Print Format$(0.6 * quantidade, "#,##0.00") 'frmOrcamento.preencheCampo(frmOrcamento.gd.TextMatrix(w, 8), 1), "#,##0.00")
      End If


'BANNER
      If descricao2 = "Banner" Then
         
         pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(quantidade, "##00"))
         objPrint.Print Format$(quantidade, "##00")
         
         'Descrição
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         
         If midia = "Lona" Then
            objPrint.Print "BN" & ":" & formato
         ElseIf midia = "Adesivo L." Then
            objPrint.Print "ADL" & ":" & formato
         Else
            objPrint.Print "ADT" & ":" & formato
         End If
         
         'Sub total
         Dim bnSubtotal As Double
         Dim lado1 As String
         Dim lado2 As String
         
         lado1 = frmOrcamento.preencheCampo("Banner:" & formato, 2)
         lado2 = frmOrcamento.preencheCampo("Banner:" & formato, 3)
         
         bnSubtotal = (lado1 * lado2) / 10000 * 45
         
         'Valor Unitário
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(bnSubtotal, "#,##0.00"))
         objPrint.Print Format$(bnSubtotal, "#,##0.00")
         
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(bnSubtotal * quantidade, "#,##0.00"))
         objPrint.Print Format$(bnSubtotal * quantidade, "#,##0.00")
         
      End If

'LAMINAÇÃO
      If laminacao <> "" And frmOrcamento.gd.TextMatrix(1, 2) <> "Capa " Then
         Dim lamQuantidade As String
         Dim lamTexto As String
         Dim lamFV As String
         
         lamQuantidade = frmOrcamento.preencheCampo(laminacao, 1)
         lamTexto = frmOrcamento.preencheCampo(laminacao, 2)
         lamFV = frmOrcamento.preencheCampo(laminacao, 3)
         
         pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(lamQuantidade, "##00"))
         objPrint.Print Format$(lamQuantidade, "##00")
         
         'Descrição
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         
         If lamFV = "FV" Then
            objPrint.Print lamTexto & ":" & lamFV
         Else
            objPrint.Print lamTexto
         End If
         
         'Valor Unitário
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(valorLami(lamTexto, lamFV, formato), "#,##0.00"))
         objPrint.Print Format$(valorLami(lamTexto, lamFV, formato), "#,##0.00")
         
         'Sub total
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(valorLami(lamTexto, lamFV, formato) * lamQuantidade, "#,##0.00"))
         objPrint.Print Format$(valorLami(lamTexto, lamFV, formato) * lamQuantidade, "#,##0.00")
         
      End If
        
'ENCADERNACAO
      
      If encadernacao <> "" Then
         Dim encQuantidade As String
         Dim encTexto As String
         Dim encEW As String
         Dim encValor As String
         
         encQuantidade = frmOrcamento.preencheCampo(encadernacao, 1)
         encTexto = frmOrcamento.preencheCampo(encadernacao, 2)
         encEW = frmOrcamento.preencheCampo(encadernacao, 3)
         encValor = frmOrcamento.calculaEncadernacao(encadernacao)
         
         pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(encQuantidade, "##00"))
         objPrint.Print Format$(encQuantidade, "##00")
         
         'Descrição
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         
         If encEW = "E" Then
            objPrint.Print "PVC" & ":" & encTexto & ":" & "E"
         Else
            objPrint.Print "PVC" & ":" & encTexto & ":" & "W"
         End If
         
         'Valor Unitário
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(encValor, "#,##0.00"))
         objPrint.Print Format$(encValor, "#,##0.00")
         
         'Sub total
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(encValor * encQuantidade, "#,##0.00"))
         objPrint.Print Format$(encValor * encQuantidade, "#,##0.00")

      End If

'CAPA
      If capa <> "" Then
         
         Dim capaQuantidade As Integer
         Dim capaTexto As String
         Dim capaValorUni As Integer
         
         capaQuantidade = calculaCapa(capa, 1)
         capaTexto = calculaCapa(capa, 2)
         capaValorUni = calculaCapa(capa, 3)
         
         
         pos = pos + 5
         
         'Quantidade
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(capaQuantidade, "##00"))
         objPrint.Print Format$(capaQuantidade, "##00")
         
         'Descrição
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         objPrint.Print capaTexto
         
         'Valor Unitário
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(capaValorUni, "#,##0.00"))
         objPrint.Print Format$(capaValorUni, "#,##0.00")
         
         'Sub total
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(capaQuantidade * capaValorUni, "#,##0.00"))
         objPrint.Print Format$(capaQuantidade * capaValorUni, "#,##0.00")
      End If
      
'WIREO
      If wireo <> "" Then
         
         Dim wireoQuantidade As Integer
         Dim wireoTexto As String
         Dim wireoValorUni As Integer
         
         wireoQuantidade = calculaWireo(wireo, 1)
         wireoTexto = calculaWireo(wireo, 2)
         wireoValorUni = calculaWireo(wireo, 3)
         
         'Quantidade
         pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(wireoQuantidade, "##00"))
         objPrint.Print Format$(wireoQuantidade, "##00")
         
         'Descrição
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         objPrint.Print "Wire-ô " & wireoTexto
         
         'Valor Unitário
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(wireoValorUni, "#,##0.00"))
         objPrint.Print Format$(wireoValorUni, "#,##0.00")
         
         'Sub total
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(wireoQuantidade * wireoValorUni, "#,##0.00"))
         objPrint.Print Format$(wireoQuantidade * wireoValorUni, "#,##0.00")
      End If
      
'CORTE
      If corte <> "" Then
         'Quantidade
         pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(corte, "##00"))
         objPrint.Print Format$(corte, "##00")
         
         'Descrição
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         objPrint.Print "Corte"
         
         'Valor Unitário
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(0.5, "#,##0.00"))
         objPrint.Print Format$(0.5, "#,##0.00")
         
         'Sub total
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(1 * corte, "#,##0.00"))
         objPrint.Print Format$(0.5 * corte, "#,##0.00")
      End If
      
'LASER
      If laser <> "" Then
         'Quantidade
         pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(laser, "##00"))
         objPrint.Print Format$(laser, "##00")
         
         'Descrição
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         objPrint.Print "Corte a Laser"
         
         'Valor Unitário
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(3, "#,##0.00"))
         objPrint.Print Format$(3, "#,##0.00")
         
         'Sub total
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(3 * laser, "#,##0.00"))
         objPrint.Print Format$(3 * laser, "#,##0.00")
      End If
      
'PICOTE
      If picote <> "" Then
         'Quantidade
         pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(picote, "##00"))
         objPrint.Print Format$(picote, "##00")
         
         'Descrição
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         objPrint.Print "Picote"
         
         'Valor Unitário
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(15, "#,##0.00"))
         objPrint.Print Format$(15, "#,##0.00")
         
         'Sub total
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(15, "#,##0.00"))
         objPrint.Print Format$(15, "#,##0.00")
      End If
      
'VINCO
      If vinco <> "" Then
         'Quantidade
         pos = pos + 5
         .CurrentY = pos + yIni
         .CurrentX = xIni - objPrint.TextWidth(Format$(vinco, "##00"))
         objPrint.Print Format$(vinco, "##00")
         
         'Descrição
         .CurrentY = pos + yIni
         .CurrentX = xIni + 2
         objPrint.Print "Vinco"
         
         'Valor Unitário
         .CurrentY = pos + yIni
         .CurrentX = xIni + 38.5 - objPrint.TextWidth(Format$(15, "#,##0.00"))
         objPrint.Print Format$(15, "#,##0.00")
         
         'Sub total
         .CurrentY = pos + yIni
         .CurrentX = xIni + 53 - objPrint.TextWidth(Format$(15, "#,##0.00"))
         objPrint.Print Format$(15, "#,##0.00")
      End If

'TOTAL DO PLANO
      .CurrentX = xIni + 31
      .CurrentY = yIni + 35
      objPrint.Print "Total"
      
      .CurrentY = yIni + 35
      .CurrentX = xIni + 53 - objPrint.TextWidth(frmOrcamento.gd.TextMatrix(w, 14))
      objPrint.Print frmOrcamento.gd.TextMatrix(w, 14)

      pos = 0
   Next

'RESUMO
   .FontSize = 10 * ratio
   
   'Descriçãp dos planos
   .CurrentX = 11 ' - objPrint.TextWidth("V. Planos")
   .CurrentY = 93 ' - (objPrint.TextHeight("V. Planos") / 2)
   objPrint.Print "V. Planos:"
   
ponto:

   'Valor dos planos
   .CurrentX = 47 - objPrint.TextWidth("R$ " & frmOrcamento.txtTotal.Text)
   .CurrentY = 93 ' - (objPrint.TextHeight(frmOrcamento.txtTotal.Text)-48)
   objPrint.Print "R$ " & frmOrcamento.txtTotal.Text
        
'Descriçãp exemplares
   If frmOrcamento.txtExemplar.Text = 1 Then
      descricao = " Exemp.:"
      valorUnitario = frmOrcamento.txtTotal.Text
   Else
      descricao = " Exemp.:"
      valorUnitario = frmOrcamento.txtTotal.Text * frmOrcamento.txtExemplar.Text
   End If
   
   .CurrentX = 49
   .CurrentY = 93
   objPrint.Print frmOrcamento.txtExemplar.Text & descricao
   
   'Valor exemplar
   .CurrentX = 85 - objPrint.TextWidth("R$ " & Format(valorUnitario, "#,##0.00"))
   .CurrentY = 93
   objPrint.Print "R$ " & Format$(valorUnitario, "#,##0.00")
   
   'Taxa entrega
   .CurrentX = 87
   .CurrentY = 93
   objPrint.Print "Entrega:"
   
   '
   ' Verifica se vai entregar
   '
   
   Dim vEntrega As Double
   If frmOrcamento.optEntrega(0).Value = True Then
       vEntrega = 5
   Else
       vEntrega = 0
   End If
       
   'Valor entrega
   .CurrentX = 123 - objPrint.TextWidth("R$ " & Format(vEntrega, "#,##0.00"))
   .CurrentY = 93
   objPrint.Print "R$ " & Format$(vEntrega, "#,##0.00")
   
   'Descrição outros serviços
   .CurrentX = 125
   .CurrentY = 93
   objPrint.Print "Outros:"
       
   'Valor outros serviços
   .CurrentX = 161 - objPrint.TextWidth("R$ " & frmOrcamento.txtAcrescimo.Text)
   .CurrentY = 93
   objPrint.Print "R$ " & frmOrcamento.txtAcrescimo.Text
   
   'Acréscimo
   .CurrentX = 57
   .CurrentY = 98
   objPrint.Print frmOrcamento.txtDescricaoAcrescimo
   
   'Desconto
   .CurrentX = 163
   .CurrentY = 93
   objPrint.Print "Desconto:"
   
   'Valor desconto
   .CurrentX = 199 - objPrint.TextWidth("R$ " & frmOrcamento.txtDesconto.Text)
   .CurrentY = 93
   objPrint.Print "R$ " & frmOrcamento.txtDesconto.Text
   
   'Descrição Desconto
   .CurrentX = 129
   .CurrentY = 98
   objPrint.Print frmOrcamento.txtDescricaoDesconto.Text
   
   objPrint.FontBold = True
   
   'Descrição Total
   .CurrentX = 11
   .CurrentY = 98
   objPrint.Print "V. total serviço:"
   
   'Valor Total
   .CurrentX = 55 - objPrint.TextWidth(frmOrcamento.txtTotalGeral.Text)
   .CurrentY = 98
   objPrint.Print frmOrcamento.txtTotalGeral.Text
   
   objPrint.FontBold = False
   
   
   Dim p As Integer
   Dim indice As Integer
   Dim formadepagamento As String
    
    For p = 0 To 3
      If frmOrcamento.optPagamento.Item(p).Value = True Then indice = p + 1
    Next
    
    
    Select Case indice
      Case 0
         formadepagamento = "Boleto"
      Case 1
         formadepagamento = "Dinheiro"
      Case 2
         formadepagamento = "Cheque"
      Case 3
         formadepagamento = "Cartão"
      Case 4
         formadepagamento = "Depósito Bancário"
    End Select


        
        yIni = 240
      .FontSize = 12 * ratio
      
      .CurrentY = yIni + 6.5
      .CurrentX = 12
      objPrint.Print "Cliente: " & frmOrcamento.txtCliente(0).Text
      
      .CurrentY = yIni + 6.5
      .CurrentX = 140 - .TextWidth("OS Nº " & Format$(numeroOs, "###000")) / 2
      objPrint.Print "OS Nº " & Format$(numeroOs, "###000")
      
      .CurrentY = yIni + 6.5
      .CurrentX = 180 - .TextWidth("Data " & frmOrcamento.txtData) / 2
      objPrint.Print "Data " & frmOrcamento.txtData

      .FontSize = 12 * ratio
      objPrint.FontBold = True
      .CurrentY = yIni + 16
      .CurrentX = 23
      objPrint.Print "Recebi o(s) serviço(s) descrito(s) acima,"
      objPrint.FontBold = False
      
      .CurrentY = .CurrentY + 6
      .CurrentX = 40
      objPrint.Print "Data: _____/_____/____"
      
      .CurrentY = .CurrentY + 1
      .CurrentX = 13
      objPrint.Print "Ass:  _____________________________________"
      
      
    yIni = 100
    
    For w = 1 To 2
      
        If w = 1 Then
            .FontSize = 13 * ratio
            .CurrentY = yIni + 14.5
            .CurrentX = 175 - .TextWidth("OS Nº " & Format$(numeroOs, "###000")) / 2
            objPrint.Print "OS Nº " & Format$(numeroOs, "###000")
            
            .FontSize = 11 * ratio
            .CurrentY = yIni + 60.6
            .CurrentX = 14
            objPrint.Print "Cliente: " & frmOrcamento.txtCliente(0).Text
            
            .CurrentY = yIni + 65.4
            .CurrentX = 14
            objPrint.Print "Data: " & Format$(frmOrcamento.txtData, "dd/mm/yy")
            
            .CurrentY = yIni + 70.2
            .CurrentX = 14
            objPrint.Print "Forma de pagamento: " & formadepagamento
            
            .CurrentY = yIni + 75
            .CurrentX = 14
            objPrint.Print "Atendente: " & frmOrcamento.txtUsuario     'varNomeUsuario
            
            .CurrentY = yIni + 75
            .CurrentX = 198 - objPrint.TextWidth("Valor total serviço: R$ " & frmOrcamento.txtTotalGeral.Text)
            objPrint.Print "Valor total serviço: R$ " & frmOrcamento.txtTotalGeral.Text
            
            yIni = yIni + 89
        
        Else
            .FontSize = 13 * ratio
            .CurrentY = yIni + 14.5
            .CurrentX = 175 - .TextWidth("OS Nº " & Format$(numeroOs, "###000")) / 2
            objPrint.Print "OS Nº " & Format$(numeroOs, "###000")
            
            .FontSize = 11 * ratio
            .CurrentY = yIni + 23
            .CurrentX = 14
            objPrint.Print "Cliente: " & frmOrcamento.txtCliente(0).Text
            
            .CurrentY = yIni + 27.8
            .CurrentX = 14
            objPrint.Print "Data: " & Format$(frmOrcamento.txtData, "dd/mm/yy")
            
            .CurrentY = yIni + 32.6
            .CurrentX = 14
            objPrint.Print "Forma de pagamento: " & formadepagamento
            
            .CurrentY = yIni + 37.4
            .CurrentX = 14
            objPrint.Print "Atendente: " & frmOrcamento.txtUsuario     'varNomeUsuario
            
            .CurrentY = yIni + 37.4
            .CurrentX = 198 - objPrint.TextWidth("Valor total serviço: R$ " & frmOrcamento.txtTotalGeral.Text)
            objPrint.Print "Valor total serviço: R$ " & frmOrcamento.txtTotalGeral.Text

        End If
    Next

    If frmOrcamento.optEntrega(0).Value = True Or frmOrcamento.optEntrega(1).Value = True Then
        objPrint.FontBold = True
        objPrint.FontSize = 12 * ratio
            .CurrentX = 175 - (objPrint.TextWidth("Fazer Entrega") / 2)
            .CurrentY = 17 - (objPrint.TextHeight("Fazer Entrega") / 2)
            objPrint.Print "Fazer Entrega"
        objPrint.FontBold = False
     End If
     
    objPrint.Line (150, 23)-(200, 23)
    
    .CurrentX = 175 - (objPrint.TextWidth("Previsão de Entrega") / 2)
    .CurrentY = 30 - (objPrint.TextHeight("Previsão de Entrega") / 2)
    objPrint.Print "Previsão de Entrega"
    
    .CurrentX = 175 - (objPrint.TextWidth(frmOrcamento.varDataEntrega & " - " & frmOrcamento.varHoraEntrega & "hs") / 2)
    .CurrentY = 37 - (objPrint.TextHeight(frmOrcamento.varDataEntrega & " - " & frmOrcamento.varHoraEntrega & "hs") / 2)
    objPrint.Print frmOrcamento.varDataEntrega & " - " & frmOrcamento.varHoraEntrega & "hs"
    
    .PaintPicture Image1.Picture, 17.6, 15.12, 33.782, 11.684


End With
End Sub


Private Sub desenhaQuadro(objPrint As Object, Optional ratio As Double = 1)
    
    Dim w As Integer
    Dim yIni As Double
    Dim yFim As Double
    Dim xIni As Double
    Dim xFim As Double
    Dim linha As Double
    Dim coluna As Double
    Dim texto As String
    
  With objPrint
    
    .FontSize = 12 * ratio
    objPrint.FontBold = True
    .CurrentX = 35 - objPrint.TextWidth(frmOrcamento.txtData) / 2
    .CurrentY = 29.31
    objPrint.Print frmOrcamento.txtData
    
    .FontSize = 8 * ratio
    objPrint.FontBold = False
    .CurrentX = 35 - objPrint.TextWidth("Rua das Nações, 212 - Bom Pastor") / 2
    .CurrentY = 34.31
    objPrint.Print "Rua das Nações, 212 - Bom Pastor"
    
    objPrint.FontSize = 10 * ratio
    objPrint.CurrentX = 35 - objPrint.TextWidth("(37)3229-8000") / 2
    objPrint.CurrentY = 37.43
    objPrint.Print "(37)3229-8000"

    xIni = 10
    yIni = 48.5
    For w = 1 To 3
    
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
        objPrint.DrawWidth = 0.5 * ratio
    End If
    
    objPrint.Line (73.33, 42.5)-(73.33, 92.66)
    objPrint.Line (136.7, 42.5)-(136.7, 92.66)
    objPrint.Line (48, 92.66)-(48, 97.66)
    objPrint.Line (86, 92.66)-(86, 97.66)
    objPrint.Line (124, 92.66)-(124, 97.66)
    objPrint.Line (162, 92.66)-(162, 97.66)
    
    
    ' Moldura
    objPrint.Line (10, 10)-(200, 10)
    objPrint.Line (10, 10)-(10, 102.66)
    objPrint.Line (200, 10)-(200, 102.66)
    objPrint.Line (10, 102.66)-(200, 102.66)
    objPrint.Line (10, 97.66)-(200, 97.66)
    objPrint.Line (10, 102.66)-(200.1, 102.66)
    
    objPrint.Line (60, 10)-(60, 42.5)
    objPrint.Line (150, 10)-(150, 42.5)
    objPrint.Line (56, 97.66)-(56, 102.66)
    objPrint.Line (128, 97.66)-(128, 102.66)
    
    '-------------------------------------

    yIni = 47.6
    yFim = 87.65
    
    'For w = 1 To 2
        objPrint.Line (20.27, yIni)-(20.27, yFim)
        objPrint.Line (44.6, yIni)-(44.6, yFim)
        objPrint.Line (59, yIni)-(59, yFim + 5)
        objPrint.Line (83.6, yIni)-(83.6, yFim)
        objPrint.Line (107.93, yIni)-(107.93, yFim)
        objPrint.Line (122.3, yIni)-(122.3, yFim + 5)
        
        'If w < 2 Then
            objPrint.Line (147, yIni)-(147, yFim)
            objPrint.Line (171.2, yIni)-(171.2, yFim)
            objPrint.Line (185.6, yIni)-(185.6, yFim + 5)
        'End If
        
        yIni = 97.63
        yFim = 137.65
    'Next
    
    linha = 42.5
    
    For w = 1 To 11
        xFim = 200
        objPrint.Line (10, linha)-(xFim, linha)
        linha = linha + 5
    Next
    
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'If frmOrcamento.txtCodigoCliente <> 0 Then
    yIni = 101
    
    For w = 1 To 3
      If w = 1 Then
         
         objPrint.DrawStyle = 2
         objPrint.Line (10, 104)-(200, 104)
         objPrint.Line (10, 192.5)-(200, 192.5)
         objPrint.Line (10, 242.5)-(200, 242.5)
         objPrint.DrawStyle = 0
   
         'yIni = yIni + 89
         
      ElseIf w = 2 Then
      
            texto = "VIA DO CAIXA"
            
         .PaintPicture Image1.Picture, 16.64, yIni + 6.12, 36.7, 12.6
         .CurrentY = yIni + 8
         .FontSize = 10 * ratio
         .CurrentX = 105 - .TextWidth("Rua das Nações, 212 - Bom Pastor") / 2
         objPrint.Print "Rua das Nações, 212 - Bom Pastor"
         
         .CurrentY = yIni + 13.3
         .CurrentX = 105 - .TextWidth("(37) 3229-8000 - CNPJ: 64.476.872/0001-03") / 2
         objPrint.Print "(37) 3229-8000 - CNPJ: 64.476.872/0001-03"
         
         .FontBold = True
         .FontSize = 25 * ratio
         .CurrentY = yIni + 25
         .CurrentX = 110 - .TextWidth(texto) / 2
         objPrint.Print texto
         .FontBold = False
         
         .FontSize = 7 * ratio
         .CurrentY = yIni + 80.5
         .CurrentX = 105 - .TextWidth("AUTENTICAÇÃO MECÂNICA") / 2
         objPrint.Print "AUTENTICAÇÃO MECÂNICA"
         
         
         objPrint.Line (10, yIni + 5)-(200, yIni + 5)
         objPrint.Line (10, yIni + 20)-(200, yIni + 20)
         objPrint.Line (10, yIni + 80)-(200, yIni + 80)
         objPrint.Line (10, yIni + 89)-(200, yIni + 89)
         
         objPrint.Line (10, yIni + 5)-(10, yIni + 89)
         objPrint.Line (200, yIni + 5)-(200, yIni + 89)
         
         objPrint.Line (60, yIni + 5)-(60, yIni + 20)
         objPrint.Line (150, yIni + 5)-(150, yIni + 20)
         
         objPrint.Line (150, yIni + 12.5)-(200, yIni + 12.5)
         
         yIni = yIni + 89
         
        Else
        
            texto = "Via do Cliente"
        
        .PaintPicture Image1.Picture, 16.64, yIni + 6.12, 36.7, 12.6
         .CurrentY = yIni + 8
         .FontSize = 10 * ratio
         .CurrentX = 105 - .TextWidth("Rua das Nações, 212 - Bom Pastor") / 2
         objPrint.Print "Rua das Nações, 212 - Bom Pastor"
         
         .CurrentY = yIni + 13.3
         .CurrentX = 105 - .TextWidth("(37) 3229-8000 - CNPJ: 64.476.872/0001-03") / 2
         objPrint.Print "(37) 3229-8000 - CNPJ: 64.476.872/0001-03"
         
         .FontSize = 14 * ratio
         .CurrentY = yIni + 6
         .CurrentX = 175 - .TextWidth(texto) / 2
         objPrint.Print texto
         
         .FontSize = 7 * ratio
         .CurrentY = yIni + 41.5
         .CurrentX = 105 - .TextWidth("AUTENTICAÇÃO MECÂNICA") / 2
         objPrint.Print "AUTENTICAÇÃO MECÂNICA"
         
         
         objPrint.Line (10, yIni + 5)-(200, yIni + 5)
         objPrint.Line (10, yIni + 20)-(200, yIni + 20)
         objPrint.Line (10, yIni + 41)-(200, yIni + 41)
         objPrint.Line (10, yIni + 50)-(200, yIni + 50)
         
         objPrint.Line (10, yIni + 5)-(10, yIni + 50)
         objPrint.Line (200, yIni + 5)-(200, yIni + 50)
         
         objPrint.Line (60, yIni + 5)-(60, yIni + 20)
         objPrint.Line (150, yIni + 5)-(150, yIni + 20)
         
         objPrint.Line (150, yIni + 12.5)-(200, yIni + 12.5)
         
         texto = "Caixa"
        End If
   
   Next
        yIni = 240
         objPrint.Line (10, yIni + 5)-(200, yIni + 5)
         objPrint.Line (10, yIni + 40)-(200, yIni + 40)
         
         objPrint.Line (10, yIni + 5)-(10, yIni + 40)
         objPrint.Line (200, yIni + 5)-(200, yIni + 40)
         objPrint.Line (120, yIni + 5)-(120, yIni + 40)
         objPrint.Line (160, yIni + 5)-(160, yIni + 12.5)
         
         objPrint.Line (10, yIni + 12.5)-(200, yIni + 12.5)

End With
    
End Sub
Public Sub montaPreview()
    Dim dRatio As Double
    dRatio = ScalePicPreviewToPrinterInches(picture1)
    PrintRoutine picture1, dRatio
End Sub
Public Sub montaControles()
    Scrollbar1.Top = tbar.Height
    Scrollbar1.Left = frm.ScaleWidth - Scrollbar1.Width
    Scrollbar1.Height = frm.ScaleHeight - tbar.Height - pic2.Height
    
    scrollbar2.Top = frm.ScaleHeight - scrollbar2.Height
    scrollbar2.Left = 0
    scrollbar2.Width = frm.ScaleWidth - pic2.Width
    
    pic2.Top = frm.ScaleHeight - pic2.Height
    pic2.Left = frm.ScaleWidth - pic2.Width
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case Shift
    
        Case 2
        
        If KeyCode = 80 Then
                Screen.MousePointer = vbHourglass
                defineImpressoraPadrao impressoraEscolhida
                Printer.ScaleMode = 6
                Printer.FontName = "Arial"
                Printer.PaperSize = vbPRPSA4
                PrintRoutine Printer
                Printer.EndDoc
                defineImpressoraPadrao impressoraPadrao
                Screen.MousePointer = vbDefault
                Unload Me
        End If
    
    End Select

End Sub

Private Sub Form_Load()
    Unload frmPrevisaodeentrega
    modoAltura = True
    modoLargura = False
    topInicial = picture1.Top
    impressoraPadrao = Printer.DeviceName
End Sub
Private Sub Form_Resize()
    posPagina
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
End Sub
Private Sub Scrollbar1_Change()
    picture1.Top = topInicial - Scrollbar1.Value
End Sub
Private Sub Scrollbar1_Scroll()
    picture1.Top = topInicial - Scrollbar1.Value
End Sub
Private Sub tbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Tag
        Case "largura"
            If modoLargura = False Then
                modoLargura = True
                modoAltura = False
                posPagina
                Scrollbar1.Value = Scrollbar1.Max / 2
            End If
        Case "altura"
            modoAltura = True
            modoLargura = False
            posPagina
            
        Case "impressora"
                Screen.MousePointer = vbHourglass
                defineImpressoraPadrao impressoraEscolhida
                Printer.ScaleMode = 6
                Printer.FontName = "Arial"
                Printer.PaperSize = vbPRPSA4
                PrintRoutine Printer
                Printer.EndDoc
                defineImpressoraPadrao impressoraPadrao
                Screen.MousePointer = vbDefault
                Unload Me

            
        Case "cimpressora"

            On Error GoTo cancelar
                CommonDialog1.CancelError = True
                CommonDialog1.Flags = cdlPDPrintSetup
                CommonDialog1.ShowPrinter
            
                impressoraEscolhida = Printer.DeviceName
                If impressoraEscolhida <> impressoraPadrao Then
                    defineImpressoraPadrao impressoraEscolhida
                End If
                tbar.Buttons(4).ToolTipText = impressoraEscolhida
                Exit Sub
        
cancelar:

    End Select

End Sub
Public Sub posPagina()

montaControles
Dim ratio As Double
Dim areaX As Double
Dim areaY As Double

montaControles

areaX = frm.ScaleWidth - Scrollbar1.Width
areaY = frm.ScaleHeight - tbar.Height - scrollbar2.Height

ratio = areaY / areaX

If modoLargura = False Then
    If ratio > 1.414 Then
        picture1.Width = areaX - 200
        picture1.Height = picture1.Width * 1.414
    Else
        picture1.Height = areaY - 200
        picture1.Width = picture1.Height * 0.707
    End If
    
    picture1.Top = areaY / 2 - picture1.Height / 2 + (tbar.Height / 2 + scrollbar2.Height / 2) + 130
    picture1.Left = areaX / 2 - picture1.Width / 2

Else
    picture1.Width = areaX - 200
    picture1.Height = picture1.Width * 1.414
    If ratio > 1.414 Then
        picture1.Top = areaY / 2 - picture1.Height / 2 + (tbar.Height / 2 + scrollbar2.Height / 2) + 130
    End If
    picture1.Left = areaX / 2 - picture1.Width / 2
End If

If picture1.Height - areaY < 1 Then
    Scrollbar1.Max = 0
Else
    Scrollbar1.Max = picture1.Height - areaY + 200
    Dim cento As Double
    cento = areaY / picture1.Height
    Scrollbar1.LargeChange = 32767 * cento
End If

montaPreview
End Sub
Private Static Function valorImpresso(midia As String, cor As Integer) As Double
    Dim papel As String
    Dim gramatura As String
    Dim valor As Double
    Dim descricao As String
    
    papel = frmOrcamento.preencheCampo(midia, 1)
    gramatura = frmOrcamento.preencheCampo(midia, 2)
    
    If midia = "Fotos" Then
        papel = "Fotos"
    End If
    
    Select Case papel
        Case "CL", "CF"
            If gramatura = "90g" Or gramatura = "115g" Or gramatura = "150g" Or gramatura = "170g" Then valor = 3.5 Else valor = 4
        
        Case "DD"
            valor = 3
        
        Case "AP"
            If gramatura = "180g" Or gramatura = "240g" Then
                valor = 3
            Else
                valor = 2.5
            End If
            
        Case "RC"
            If gramatura = "120g" Then
                valor = 2.5
            Else
                valor = 3
            End If
            
        Case "AD", "BOPP"
            valor = 4
        
        Case "TR"
            valor = 4
            
        Case "Fotos"
            valor = 0.55
            
        Case ""
            valor = 30

    End Select
    
    If cor = 6 Then
        If papel = "CL" Then
            valor = 3
        Else
            valor = valor + 1
        End If
    End If
        
        valorImpresso = valor

End Function
Private Function valorLami(lami As String, FV As String, Optional formato As String) As Double
    
   Select Case lami
      Case "LB", "LF", "VZ"
         valorLami = 1
      Case "PA6"
         valorLami = 1
      Case "PA5"
         valorLami = 1.5
      Case "PA4"
         valorLami = 2
      Case "PA3"
      valorLami = 3.5
   End Select
   
   If FV = "FV" Then
      valorLami = valorLami * 2
   End If
   
   If formato = "A4" Then
      valorLami = valorLami / 2
   End If
   
End Function
Private Function valorCapa(capa As String) As Double
    
    Select Case capa
        Case "PPA5"
            valorCapa = 4
        Case "PPA4"
            valorCapa = 6
        Case "PPA3"
            valorCapa = 12
        Case "PVCA5"
            valorCapa = 1
        Case "PVCA4"
            valorCapa = 2
    End Select

End Function
Private Function valorWireo(wireo As String) As Double

    Select Case wireo
        Case "1/4""", "5/16""", "3/8"""
            valorWireo = 1.5
        Case "7/16""", "1/2""", "9/16"""
            valorWireo = 2
        Case "5/8"""
            valorWireo = 2.5
        Case "7/8""", "1"""
            valorWireo = 3
            
    End Select
    
End Function
