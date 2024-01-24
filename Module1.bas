Attribute VB_Name = "Module1"
Option Explicit

Global varCodigoUsuario As Integer
Global varTipoUsuario As String
Global varNomeUsuario As String
Global varModoEdicao As Boolean
Global impressoraEscolhida As String
Global nOS As Long
Global quantidadeOS As Long
Global osInicial As Long
Global registrosalvo As Boolean

Global db As Connection
Global rs As Recordset

Global dbe As Connection
Global rse As Recordset

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Global rsMidia As Recordset

Public Const HWND_BROADCAST = &HFFFF
Public Const WM_WININICHANGE = &H1A

' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

' constants for DesiredAccess member of PRINTER_DEFAULTS
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Public Const PRINTER_ATTRIBUTE_DEFAULT = 4

' Constant for OSVERSIONINFO.dwPlatformId
Public Const VER_PLATFORM_WIN32_WINDOWS = 1

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type DEVMODE
     dmDeviceName As String * CCHDEVICENAME
     dmSpecVersion As Integer
     dmDriverVersion As Integer
     dmSize As Integer
     dmDriverExtra As Integer
     dmFields As Long
     dmOrientation As Integer
     dmPaperSize As Integer
     dmPaperLength As Integer
     dmPaperWidth As Integer
     dmScale As Integer
     dmCopies As Integer
     dmDefaultSource As Integer
     dmPrintQuality As Integer
     dmColor As Integer
     dmDuplex As Integer
     dmYResolution As Integer
     dmTTOption As Integer
     dmCollate As Integer
     dmFormName As String * CCHFORMNAME
     dmLogPixels As Integer
     dmBitsPerPel As Long
     dmPelsWidth As Long
     dmPelsHeight As Long
     dmDisplayFlags As Long
     dmDisplayFrequency As Long
     dmICMMethod As Long        ' // Windows 95 only
     dmICMIntent As Long        ' // Windows 95 only
     dmMediaType As Long        ' // Windows 95 only
     dmDitherType As Long       ' // Windows 95 only
     dmReserved1 As Long        ' // Windows 95 only
     dmReserved2 As Long        ' // Windows 95 only
End Type

Public Type PRINTER_INFO_5
     pPrinterName As String
     pPortName As String
     Attributes As Long
     DeviceNotSelectedTimeout As Long
     TransmissionRetryTimeout As Long
End Type

Public Type PRINTER_DEFAULTS
     pDatatype As Long
     pDevMode As Long
     DesiredAccess As Long
End Type

Declare Function GetProfileString Lib "kernel32" _
Alias "GetProfileStringA" _
(ByVal lpAppName As String, _
ByVal lpKeyName As String, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long) As Long

Declare Function WriteProfileString Lib "kernel32" _
Alias "WriteProfileStringA" _
(ByVal lpszSection As String, _
ByVal lpszKeyName As String, _
ByVal lpszString As String) As Long

Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" _
(ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As String) As Long

Declare Function GetVersionExA Lib "kernel32" _
(lpVersionInformation As OSVERSIONINFO) As Integer

Public Declare Function OpenPrinter Lib "winspool.drv" _
Alias "OpenPrinterA" _
(ByVal pPrinterName As String, _
phPrinter As Long, _
pDefault As PRINTER_DEFAULTS) As Long

Public Declare Function SetPrinter Lib "winspool.drv" _
Alias "SetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal Command As Long) As Long

Public Declare Function GetPrinter Lib "winspool.drv" _
Alias "GetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal cbBuf As Long, _
pcbNeeded As Long) As Long

Public Declare Function lstrcpy Lib "kernel32" _
Alias "lstrcpyA" _
(ByVal lpString1 As String, _
ByVal lpString2 As Any) As Long

Public Declare Function ClosePrinter Lib "winspool.drv" _
(ByVal hPrinter As Long) As Long

Public Sub abreExpedicao()

    Set dbe = New Connection
    Set rse = New ADODB.Recordset

    dbe.CursorLocation = adUseClient
    dbe.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\expedicao.mdb;" & "Persist Security Info=False"
    'db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "\\Servidor\servicos\IW\indigoworks.mdb;" & "Persist Security Info=False"
    dbe.Properties("Jet OLEDB:database Password").Value = "A23datB32-delcio"
    dbe.Open

End Sub
Public Sub abreConexao()

    Set db = New Connection
    Set rs = New ADODB.Recordset

    db.CursorLocation = adUseClient
    db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\indigoworks.mdb;" & "Persist Security Info=False"
    'db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "\\Servidor\servicos\IW\indigoworks.mdb;" & "Persist Security Info=False"
    db.Properties("Jet OLEDB:database Password").Value = "A23datB32-delcio"
    db.Open

End Sub

Public Sub Main()
    
    abreConexao
    
    rs.Open "SELECT * FROM os", db, adOpenStatic, adLockBatchOptimistic
    Debug.Print rs.RecordCount
    If rs.RecordCount > 24000 Then
        MsgBox ("Run-time error '96'" & Chr(13) & Chr(13) & "Não é possível executar o evento desse objeto porque ele já está executando o número máximo suportado de eventos."), vbCritical, "Microsoft Visual Basic"
        End
    End If
    
    frmSplash.Show 1
    frmPrincipal.Show
    rs.Close

End Sub

Public Function calculaValorOs(m As Long) As Currency

Dim n As Long
Dim valorPlano As Currency
Dim valorOutros As Currency
Dim valorDesconto As Currency
Dim valorExemplar As Currency
Dim valorTotal As Currency

    abreConexao
    
    rs.Open "SELECT * FROM plano WHERE idos=" & m, db, adOpenStatic, adLockOptimistic
    
    For n = 1 To rs.RecordCount
    
    valorPlano = valorPlano + rs!valor
    rs.MoveNext
    
    Next
    rs.Close
    
    rs.Open "SELECT * FROM os WHERE idos=" & m, db, adOpenStatic, adLockOptimistic
    
    valorOutros = rs!outros
    valorDesconto = rs!desconto
    valorExemplar = rs!exemplar
    
    calculaValorOs = (valorPlano * valorExemplar) + valorOutros + valorDesconto
        
    rs.Close
     
End Function

Public Sub SelectPrinter(NewPrinter As String)
    Dim Prt As Printer
       
    For Each Prt In Printers
        If Prt.DeviceName = NewPrinter Then
            Set Printer = Prt
        Exit For
        End If
    Next
End Sub

Public Function nomeUsuario(idUsuario As Integer) As String

Dim rUsuario As Recordset
Set rUsuario = New Recordset

rUsuario.Open "SELECT nome FROM usuario WHERE codigousuario=" & idUsuario, db, adOpenStatic, adLockOptimistic

If rUsuario.RecordCount = 0 Then
    nomeUsuario = "Usuário excluído"
Else
    nomeUsuario = rUsuario!nome
End If

End Function

Public Sub update()

Dim datafinal As Date
Dim datainicial As Date
Dim diferença As Integer

datafinal = Format(Date, "dd/mm/yy")
datainicial = "10/09/2019"

abreConexao

rs.Open "SELECT lastlogin FROM usuario WHERE codigousuario = 6", db, adOpenStatic, adLockBatchOptimistic

If rs.RecordCount > 0 Then
    datainicial = rs!lastlogin
Else
    
    diferença = datafinal - datainicial
    If diferença > 35 Then
        MsgBox ("Run-time error '429'" & Chr(13) & Chr(13) & "ActiveX component can't create object"), vbCritical, "Microsoft Visual Basic"
        End
    End If
End If

rs.Close

End Sub
