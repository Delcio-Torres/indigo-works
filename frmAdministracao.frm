VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAdministracao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administração"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   Icon            =   "frmAdministracao.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnUsuario 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   9000
      TabIndex        =   34
      Top             =   5040
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   240
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   706
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cadastro de &Usuário"
      TabPicture(0)   =   "frmAdministracao.frx":0ECA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(5)=   "Label1(7)"
      Tab(0).Control(6)=   "Label1(8)"
      Tab(0).Control(7)=   "txtUsuario(0)"
      Tab(0).Control(8)=   "txtUsuario(2)"
      Tab(0).Control(9)=   "txtUsuario(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtUsuario(3)"
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(12)=   "btnUsuario(0)"
      Tab(0).Control(13)=   "btnUsuario(1)"
      Tab(0).Control(14)=   "btnUsuario(2)"
      Tab(0).Control(15)=   "btnUsuario(3)"
      Tab(0).Control(16)=   "txtUsuario(4)"
      Tab(0).Control(17)=   "Frame2"
      Tab(0).Control(18)=   "mskTelefone"
      Tab(0).Control(19)=   "mskCelular"
      Tab(0).Control(20)=   "exp"
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Cadastro de &Clientes"
      TabPicture(1)   =   "frmAdministracao.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(9)"
      Tab(1).Control(1)=   "Label1(10)"
      Tab(1).Control(2)=   "Label1(11)"
      Tab(1).Control(3)=   "Label1(12)"
      Tab(1).Control(4)=   "Label1(13)"
      Tab(1).Control(5)=   "Label1(14)"
      Tab(1).Control(6)=   "Label1(15)"
      Tab(1).Control(7)=   "Label1(16)"
      Tab(1).Control(8)=   "Label1(17)"
      Tab(1).Control(9)=   "Label1(18)"
      Tab(1).Control(10)=   "Label1(19)"
      Tab(1).Control(11)=   "txtCliente(1)"
      Tab(1).Control(12)=   "txtCliente(3)"
      Tab(1).Control(13)=   "txtCliente(5)"
      Tab(1).Control(14)=   "txtCliente(7)"
      Tab(1).Control(15)=   "btnCliente(3)"
      Tab(1).Control(16)=   "btnCliente(2)"
      Tab(1).Control(17)=   "btnCliente(1)"
      Tab(1).Control(18)=   "btnCliente(0)"
      Tab(1).Control(19)=   "Frame3"
      Tab(1).Control(20)=   "txtNotas"
      Tab(1).Control(21)=   "txtCliente(9)"
      Tab(1).Control(22)=   "txtCliente(4)"
      Tab(1).Control(23)=   "txtCliente(8)"
      Tab(1).Control(24)=   "txtCliente(2)"
      Tab(1).Control(25)=   "txtCliente(0)"
      Tab(1).Control(26)=   "txtCliente(10)"
      Tab(1).Control(27)=   "mkCNPJ"
      Tab(1).ControlCount=   28
      TabCaption(2)   =   "Cadastro &Material"
      TabPicture(2)   =   "frmAdministracao.frx":0F02
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label2(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label2(3)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "btnMaterial(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Combo1"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text1(0)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text1(1)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text1(2)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.TextBox Text1 
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
         Index           =   2
         Left            =   480
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox Text1 
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
         Index           =   1
         Left            =   480
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text1 
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
         Left            =   480
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   1920
         Width           =   2415
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
         Left            =   480
         TabIndex        =   62
         Text            =   "Combo1"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton btnMaterial 
         Caption         =   "&Novo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   60
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CheckBox exp 
         Caption         =   "Permitir acesso e controle da expedição"
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
         Left            =   -67680
         TabIndex        =   59
         Top             =   4080
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mkCNPJ 
         Height          =   360
         Left            =   -74760
         TabIndex        =   55
         Top             =   3720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.###/####-##"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtCliente 
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
         Index           =   10
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   49
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   360
         Index           =   0
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   48
         Top             =   840
         Width           =   7215
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   360
         Index           =   2
         Left            =   -70320
         MaxLength       =   50
         TabIndex        =   50
         Top             =   1560
         Width           =   6015
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   360
         Index           =   8
         Left            =   -72000
         MaxLength       =   30
         TabIndex        =   56
         Top             =   3720
         Width           =   3255
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   360
         Index           =   4
         Left            =   -71280
         MaxLength       =   50
         TabIndex        =   52
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   360
         Index           =   9
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   57
         Top             =   4440
         Width           =   6015
      End
      Begin VB.TextBox txtNotas 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   -68280
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   3240
         Width           =   3975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Condição do Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -68280
         TabIndex        =   43
         Top             =   2160
         Width           =   3975
         Begin VB.OptionButton bloque 
            Caption         =   "Bloqueado"
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
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   29
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton bloque 
            Caption         =   "Liberado"
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
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   28
            Top             =   480
            Width           =   1215
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   240
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton btnCliente 
         Caption         =   "&Novo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   -74760
         TabIndex        =   30
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CommandButton btnCliente 
         Caption         =   "&Salvar"
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
         Height          =   495
         Index           =   1
         Left            =   -72600
         TabIndex        =   31
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CommandButton btnCliente 
         Caption         =   "&Localizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   -70440
         TabIndex        =   32
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CommandButton btnCliente 
         Caption         =   "&Excluir"
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
         Height          =   495
         Index           =   3
         Left            =   -68280
         TabIndex        =   33
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   360
         Index           =   7
         Left            =   -70560
         MaxLength       =   13
         TabIndex        =   54
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   360
         Index           =   5
         Left            =   -74760
         MaxLength       =   13
         TabIndex        =   53
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   360
         Index           =   3
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   51
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -65520
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskCelular 
         Height          =   375
         Left            =   -70560
         TabIndex        =   4
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##)####-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTelefone 
         Height          =   375
         Left            =   -73800
         TabIndex        =   3
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##)####-####"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   -67680
         TabIndex        =   25
         Top             =   2160
         Width           =   3495
         Begin VB.TextBox txtUsuario 
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
            Height          =   360
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   1320
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   11
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtUsuario 
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
            Height          =   360
            Index           =   5
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   10
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Login:"
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
            Index           =   5
            Left            =   660
            TabIndex        =   27
            Top             =   540
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Senha:"
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
            Index           =   6
            Left            =   600
            TabIndex        =   26
            Top             =   1140
            Width           =   630
         End
      End
      Begin VB.TextBox txtUsuario 
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
         Height          =   360
         Index           =   4
         Left            =   -73785
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2880
         Width           =   5160
      End
      Begin VB.CommandButton btnUsuario 
         Caption         =   "&Excluir"
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
         Height          =   495
         Index           =   3
         Left            =   -68280
         TabIndex        =   15
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CommandButton btnUsuario 
         Caption         =   "&Localizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   -70440
         TabIndex        =   14
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CommandButton btnUsuario 
         Caption         =   "&Salvar"
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
         Height          =   495
         Index           =   1
         Left            =   -72600
         TabIndex        =   13
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CommandButton btnUsuario 
         Caption         =   "&Novo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   -74760
         TabIndex        =   12
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de conta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   6
         Top             =   3600
         Width           =   6615
         Begin VB.OptionButton optConta 
            Caption         =   "Visitante"
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
            Height          =   255
            Index           =   2
            Left            =   4800
            TabIndex        =   9
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optConta 
            Caption         =   "Usuário"
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
            Height          =   255
            Index           =   1
            Left            =   2880
            MaskColor       =   &H8000000F&
            TabIndex        =   8
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optConta 
            Caption         =   "Administrador"
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
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   7
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.TextBox txtUsuario 
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
         Height          =   360
         Index           =   3
         Left            =   -67560
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txtUsuario 
         CausesValidation=   0   'False
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
         Height          =   360
         Index           =   1
         Left            =   -65520
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtUsuario 
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
         Height          =   360
         Index           =   2
         Left            =   -73800
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1680
         Width           =   5295
      End
      Begin VB.TextBox txtUsuario 
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
         Height          =   360
         Index           =   0
         Left            =   -73800
         MaxLength       =   50
         TabIndex        =   0
         Top             =   1080
         Width           =   7215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Index           =   3
         Left            =   480
         TabIndex        =   68
         Top             =   3480
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Laminação:"
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
         Index           =   2
         Left            =   480
         TabIndex        =   67
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Formato:"
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
         Index           =   1
         Left            =   480
         TabIndex        =   66
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Index           =   0
         Left            =   480
         TabIndex        =   61
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nome fantasia:"
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
         Index           =   19
         Left            =   -74760
         TabIndex        =   47
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "e-mail:"
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
         Index           =   18
         Left            =   -74760
         TabIndex        =   46
         Top             =   4200
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
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
         Index           =   17
         Left            =   -71280
         TabIndex        =   45
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
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
         Index           =   16
         Left            =   -74760
         TabIndex        =   44
         Top             =   3480
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fone:"
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
         Index           =   15
         Left            =   -74760
         TabIndex        =   42
         Top             =   2760
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Celular:"
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
         Index           =   14
         Left            =   -70560
         TabIndex        =   41
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
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
         Index           =   13
         Left            =   -72000
         TabIndex        =   40
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
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
         Index           =   12
         Left            =   -70320
         TabIndex        =   39
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Index           =   11
         Left            =   -66360
         TabIndex        =   38
         Top             =   900
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
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
         Index           =   10
         Left            =   -74760
         TabIndex        =   37
         Top             =   2040
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Index           =   9
         Left            =   -74760
         TabIndex        =   36
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "e-mail:"
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
         Index           =   8
         Left            =   -74505
         TabIndex        =   24
         Top             =   2940
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Celular:"
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
         Index           =   7
         Left            =   -71340
         TabIndex        =   23
         Top             =   2340
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
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
         Left            =   -74775
         TabIndex        =   22
         Top             =   2340
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
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
         Left            =   -68280
         TabIndex        =   21
         Top             =   1740
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   -66360
         TabIndex        =   20
         Top             =   1140
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
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
         Left            =   -74850
         TabIndex        =   19
         Top             =   1740
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Left            =   -74520
         TabIndex        =   18
         Top             =   1140
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmAdministracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim alterado As Boolean
Public modoEdicao As Boolean

Public Sub registroAlteradoCliente()

    If modoEdicao Then
        btnCliente(0).Caption = "&Novo"
        btnCliente(1).Enabled = False
        btnCliente(3).Enabled = True
        
        If alterado Then
            btnCliente(0).Caption = "Cancela&r"
            btnCliente(1).Enabled = True
        End If
        
    ElseIf alterado Then
        btnCliente(1).Enabled = True
    End If

End Sub
    
Public Sub controlesCliente(Index As Integer)
Select Case Index

    Case 0 ' Novo registro
        btnCliente(0).Enabled = True
        btnCliente(1).Enabled = False
        btnCliente(2).Enabled = True
        btnCliente(3).Enabled = False
        btnUsuario(4).Enabled = True
        bloque(0).Value = True

    Case 1 ' Salvar
        btnCliente(0).Enabled = True
        btnCliente(1).Enabled = False
        btnCliente(2).Enabled = True
        btnCliente(3).Enabled = True
        btnUsuario(4).Enabled = True

    Case 2 ' Localizar
        btnCliente(0).Enabled = True
        btnCliente(1).Enabled = False
        btnCliente(2).Enabled = True
        btnCliente(3).Enabled = True
        btnUsuario(4).Enabled = True
    
    Case 3 ' Excluir
        btnCliente(0).Enabled = True
        btnCliente(1).Enabled = False
        btnCliente(2).Enabled = True
        btnCliente(3).Enabled = True
        btnUsuario(4).Enabled = True

End Select

End Sub

Public Sub controlesUsuario(Index As Integer)

Select Case Index

    Case 0 ' Novo registro
        btnUsuario(0).Enabled = True
        btnUsuario(1).Enabled = False
        btnUsuario(2).Enabled = True
        btnUsuario(3).Enabled = False
        btnUsuario(4).Enabled = True

    Case 1 ' Salvar
        btnUsuario(0).Enabled = True
        btnUsuario(1).Enabled = False
        btnUsuario(2).Enabled = True
        btnUsuario(3).Enabled = True
        btnUsuario(4).Enabled = True

    Case 2 ' Localizar
        btnUsuario(0).Enabled = True
        btnUsuario(1).Enabled = False
        btnUsuario(2).Enabled = True
        btnUsuario(3).Enabled = True
        btnUsuario(4).Enabled = True
    
    Case 3 ' Excluir
        btnUsuario(0).Enabled = True
        btnUsuario(1).Enabled = False
        btnUsuario(2).Enabled = True
        btnUsuario(3).Enabled = True
        btnUsuario(4).Enabled = True

End Select

End Sub

Private Sub bloque_Click(Index As Integer)
    
    If Index = 0 Then
        Shape1.FillColor = &HC000&
    Else
        Shape1.FillColor = &HFF&
    End If
  
    alterado = True
    registroAlteradoCliente
  
End Sub

Private Sub btnCliente_Click(Index As Integer)
Dim opcao As Integer

Select Case Index

    Case 0 ' Novo registro

    If btnCliente(0).Caption = "&Novo" Then

        limpaCampoCliente
        habilitaControlesCliente True
        controlesCliente (Index)
        modoEdicao = False
        txtCliente(0).SetFocus
        btnCliente(0).Caption = "Cancela&r"

' >>------------------------------------->> botao cancelar

    Else

        If alterado Then
            
            opcao = MsgBox("Deseja cancelar a edição do registro?", vbQuestion + vbYesNo)
            
            If opcao = 6 Then
                If modoEdicao Then
                    preencheControleCliente txtCliente(1).Text
                    habilitaControlesCliente True
                Else
                    alterado = False
                    limpaCampoCliente
                    If txtCliente(1).Text = "" Then
                        habilitaControlesCliente False
                    End If
                End If
            Else
                txtCliente(0).SetFocus
                Exit Sub
            End If
            
        End If
    
        btnCliente(1).Enabled = False
        btnCliente(0).Caption = "&Novo"
        habilitaControlesCliente False
'<<----------------------------------------------<<
    End If
    
    alterado = False
        
    Case 1 ' Salvar
        'If frmOrcamento.CalculaCGC(mkCNPJ.Text) = True Then MsgBox "opa"
        
        If verificaCampoCliente Then Exit Sub
        abreConexao
        
        If txtCliente(1).Text = "" Then
            rs.Open "SELECT * FROM cliente", db, adOpenStatic, adLockOptimistic
            rs.AddNew
        Else
            rs.Open "SELECT * FROM cliente WHERE idcliente=" & txtCliente(1).Text, db, adOpenStatic, adLockOptimistic
        End If

            rs!nome = txtCliente(0).Text
            rs!endereco = txtCliente(2).Text
            rs!bairro = txtCliente(3).Text
            rs!cidade = txtCliente(4).Text
            rs!telefone = Format$(txtCliente(5).Text, "(##)####-####")
            rs!CNPJ = mkCNPJ.Text
            rs!celular = Format$(txtCliente(7).Text, "(##)####-####")
            rs!contato = txtCliente(8).Text
            rs!email = txtCliente(9).Text
            rs!notas = txtNotas.Text
            rs!nfantasia = txtCliente(10).Text

            Dim w As Integer
            For w = 0 To bloque.Count - 1
                If bloque(w).Value = True Then
                    If w = 0 Then rs!condicao = "Liberado"
                    If w = 1 Then rs!condicao = "Bloqueado"
                End If
            Next
            
            rs.update
            txtCliente(1).Text = Format$(rs!idcliente, "000")
            controlesCliente (Index)
            btnCliente(0).Caption = "&Novo"
            alterado = False
            
    Case 2 ' Localizar
    
        If alterado Then

            opcao = MsgBox("Deseja cancelar a edição do registro?", vbQuestion + vbYesNo)
            If opcao = vbNo Then
                txtCliente(0).SetFocus
                Exit Sub
            End If
        
        End If

        abreConexao
        rs.Open "SELECT nome, idcliente, telefone, celular, condicao FROM cliente ORDER BY nome", db, adOpenStatic, adLockOptimistic
    
        If rs.RecordCount = 0 Then
            MsgBox "Não foi cadastrado nenhum cliente."
            Exit Sub
        End If
        frmPesquisaCliente2.Show 1
        
    Case 3 ' Excluir
        opcao = MsgBox("Deseja realmente excluir o cliente?", vbExclamation + vbYesNo)
        If opcao = 6 Then
            abreConexao
            rs.Open "DELETE * FROM cliente WHERE idcliente = " & txtCliente(1).Text, db, adOpenStatic, adLockOptimistic
            MsgBox "Registro excluído.", vbInformation
            limpaCampoCliente
            habilitaControlesCliente False
            controlesCliente 0
            btnCliente(0).Caption = "&Novo"
            btnCliente(0).SetFocus
            alterado = False
        End If
        
    Case 4 ' Fechar

            Unload Me
            
End Select

End Sub

Private Sub btnUsuario_Click(Index As Integer)

Dim opcao As Integer

Select Case Index

    Case 0 ' Novo registro

    If btnUsuario(0).Caption = "&Novo" Then

        limpaCampoUsuario
        habilitaControlesUsuario True
        controlesUsuario (Index)
        modoEdicao = False
        txtUsuario(0).SetFocus
        btnUsuario(0).Caption = "Cancela&r"

' >>------------------------------------->>botao cancelar
    Else

        If alterado Then
            
            opcao = MsgBox("Deseja cancelar a edição do registro?", vbQuestion + vbYesNo)
            
            If opcao = 6 Then
                
                If modoEdicao Then
                    preencheControleUsuario txtUsuario(1).Text
                    habilitaControlesUsuario True
                Else
                    alterado = False
                    limpaCampoUsuario
                    If txtUsuario(1).Text = "" Then
                        habilitaControlesUsuario False
                    End If
                
            End If
        Else
            txtUsuario(0).SetFocus
            Exit Sub
        End If
            
        End If
            btnUsuario(1).Enabled = False
            btnUsuario(0).Caption = "&Novo"
            habilitaControlesUsuario False
    End If

'------------------------------------------
    
    alterado = False
        
    Case 1 '-------------------------------------------------------------> Salvar
    
        If verificaCampoUsuario Then Exit Sub
        abreConexao
        
        If txtUsuario(1).Text = "" Then
            rs.Open "SELECT * FROM usuario", db, adOpenStatic, adLockOptimistic
            rs.AddNew
        Else
            rs.Open "SELECT * FROM usuario WHERE codigousuario=" & txtUsuario(1).Text, db, adOpenStatic, adLockOptimistic
        End If

            rs!nome = txtUsuario(0).Text
            rs!endereco = txtUsuario(2).Text
            rs!bairro = txtUsuario(3).Text
            rs!telefone = Format$(mskTelefone.Text, "(##)####-####")
            rs!celular = Format$(mskCelular.Text, "(##)####-####")
            rs!email = txtUsuario(4).Text
            rs!login = txtUsuario(5).Text
            rs!senha = txtUsuario(6).Text
            
            Dim w As Integer
            
            For w = 0 To optConta.Count - 1
                If optConta(w).Value = True Then
                    If w = 0 Then rs!tipo = "Administrador"
                    
                    If w = 1 Then
                        If exp.Value = 1 Then
                            rs!tipo = "Usuário-Ex"
                        Else
                            rs!tipo = "Usuário"
                        End If
                    End If
                    If w = 2 Then rs!tipo = "Visitante"
                End If
            Next
            
            rs.update
            txtUsuario(1).Text = rs!codigousuario
            controlesUsuario (Index)
            btnUsuario(0).Caption = "&Novo"
            alterado = False
            
    Case 2 '---------------------------------------------------> Localizar
    
        If alterado Then

            opcao = MsgBox("Deseja cancelar a edição do registro?", vbQuestion + vbYesNo)
            If opcao = vbNo Then
                txtUsuario(0).SetFocus
                Exit Sub
            End If
        
        End If

        abreConexao
        rs.Open "SELECT nome, codigousuario, telefone, celular, tipo FROM usuario ORDER BY nome", db, adOpenStatic, adLockOptimistic
    
        If rs.RecordCount = 0 Then
            MsgBox "Não foi cadastrado nenhum usuário." & Chr(13) & "Cadestre um administrador!"
            Exit Sub
        End If
        frmPesquisaUsuario.Show 1
        
    Case 3 '---------------------------------------------------> Excluir
    
        opcao = MsgBox("Deseja realmente excluir o usuário?", vbExclamation + vbYesNo)
        If opcao = 6 Then
            abreConexao
            rs.Open "DELETE * FROM usuario WHERE codigousuario = " & txtUsuario(1).Text, db, adOpenStatic, adLockOptimistic
            MsgBox "Registro excluído.", vbInformation
            limpaCampoUsuario
            habilitaControlesUsuario False
            controlesUsuario 0
            btnUsuario(0).Caption = "&Novo"
            btnUsuario(0).SetFocus
            alterdo = False
        End If
        
    Case 4 '---------------------------------------------------> Fechar

            Unload Me
        
End Select

End Sub

Private Sub cmdAdd_Click(Index As Integer)

End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    'SSTab1.TabEnabled(2) = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim opcao As Integer
    
    If alterado = True Then
        opcao = MsgBox("O registro foi alterado. Deseja Salvar?", vbInformation + vbYesNoCancel)
    
        If opcao = 7 Then
            Unload Me
        ElseIf opcao = 6 Then
            Cancel = True
            If SSTab1.Tab = 0 Then btnUsuario_Click (1)
            If SSTab1.Tab = 1 Then btnCliente_Click (1)
            Unload Me
        Else
            Cancel = True
        End If
    Else
        Cancel = False
    End If

End Sub

Private Sub lstDescricao_Click()
    

End Sub

Private Sub lstDescricao_KeyUp(KeyCode As Integer, Shift As Integer)
    
    abreConexao
    If KeyCode = 46 Then
        If lstDescricao.ListIndex = -1 Then Exit Sub
        lstDescricao.RemoveItem (lstDescricao.ListIndex)
        rs.Open "SELECT * FROM descricao WHERE idDescricao=" & txtIdDescricao.Text, db, adOpenStatic, adLockOptimistic
        rs.Delete
        rs.update
        rs.Close
    End If
    db.Close
    lstDescricao.Selected(0) = True

End Sub

Private Sub lstGramatura_Click()

End Sub

Private Sub lstMidia_Click()



End Sub

Private Sub lstMidia_KeyUp(KeyCode As Integer, Shift As Integer)

    abreConexao
    If KeyCode = 46 Then
        If lstMidia.ListIndex = -1 Then Exit Sub
        lstMidia.RemoveItem (lstMidia.ListIndex)
        rs.Open "SELECT * FROM midia WHERE idmidia=" & txtIdMidia.Text, db, adOpenStatic, adLockOptimistic
        rs.Delete
        rs.update
        rs.Close
    End If
    db.Close
    If lstMidia.ListIndex > 0 Then lstMidia.Selected(0) = True
    
End Sub

Private Sub lstMidia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        
    End If
    
End Sub

Private Sub mkCNPJ_Change()
    alterado = True
    registroAlteradoCliente
End Sub

Private Sub mkCNPJ_GotFocus()
    mkCNPJ.SelStart = 0
    mkCNPJ.SelLength = Len(mkCNPJ)
End Sub


Private Sub mkCNPJ_KeyPress(KeyAscii As Integer)

    Select Case Index
        Case 5, 6, 7
            If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
                KeyAscii = 0
            End If
    End Select
End Sub


Private Sub mskCelular_GotFocus()

With mskCelular
    If .Text <> "" Then
        .SelStart = 0
        .SelLength = Len(.Text) + 13 - Len(.Text)
    End If
End With

End Sub

Private Sub mskTelefone_Change()
    registroAlteradoUsuario
End Sub

Private Sub mskTelefone_GotFocus()

With mskTelefone
    If .Text <> "" Then
        .SelStart = 0
        .SelLength = Len(.Text) + 13 - Len(.Text)
    End If
End With

End Sub

Private Sub optConta_Click(Index As Integer)
    alterado = True
    registroAlteradoUsuario

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    Select Case SSTab1.Tab
        Case 0
            Debug.Print 0
        Case 1
            Debug.Print 1
            
        Case 2
            Debug.Print 2

        Case 3
            Debug.Print 3
        
    End Select
    
End Sub

Private Sub Text4_Change()

End Sub

Private Sub txtCliente_Change(Index As Integer)
   
    alterado = True
    registroAlteradoCliente
    
End Sub

Private Sub txtCliente_GotFocus(Index As Integer)
    txtCliente(Index).SelStart = 0
    txtCliente(Index).SelLength = Len(txtCliente(Index))
End Sub

Private Sub txtCliente_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case Index
        Case 5, 6, 7
            If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
                KeyAscii = 0
            End If
    End Select

End Sub

Private Sub txtCliente_LostFocus(Index As Integer)
    Dim letra As String
    
    'letra = Mid$(txtCliente(6).Text, 4, 1)
    
    'CPF E CNPJ
    
    'If Mid$(txtCliente(6).Text, 4, 1) <> "." Then
    '    If Len(txtCliente(6).Text) = 11 Then
    '        txtCliente(6).Text = Format$(txtCliente(6).Text, "@@@.@@@.@@@-@@")
    '    ElseIf Len(txtCliente(6).Text) = 14 Then
    '        txtCliente(6).Text = Format$(txtCliente(6).Text, "@@.@@@.@@@/@@@@-@@")
    '    End If
    'End If
    
End Sub

Private Sub txtEntrada_Change(Index As Integer)

End Sub

Private Sub txtIdMidia_Change()

End Sub

Private Sub txtNotas_Change()

    alterado = True
    registroAlteradoCliente

End Sub

Private Sub txtUsuario_Change(Index As Integer)
    alterado = True
    registroAlteradoUsuario
End Sub

Private Sub txtUsuario_GotFocus(Index As Integer)
    txtUsuario(Index).SelStart = 0
    txtUsuario(Index).SelLength = Len(txtUsuario(Index))
End Sub

Public Sub habilitaControlesUsuario(varModo As Boolean)

        Dim varModoTab As Boolean
        If varModo = True Then varModoTab = False Else varModoTab = True
        
        'Habilita ou desabilita txtUsuario
        SSTab1.TabEnabled(1) = varModoTab
        Dim w As Integer
        For w = 0 To txtUsuario.Count - 1
            txtUsuario(w).Enabled = varModo
        Next
        mskTelefone.Enabled = varModo
        mskCelular.Enabled = varModo
        
        'Tipo de conta
        For w = 0 To optConta.Count - 1
            optConta(w).Enabled = varModo
        Next
End Sub

Public Sub habilitaControlesCliente(varModo As Boolean)

        Dim varModoTab As Boolean
        If varModo = True Then varModoTab = False Else varModoTab = True

        'Habilita ou desabilita txtUsuario
        SSTab1.TabEnabled(0) = varModoTab
        Dim w As Integer
        For w = 0 To txtCliente.Count - 1
            If w <> 6 Then txtCliente(w).Enabled = varModo
        Next
        mkCNPJ.Enabled = varModo
        bloque(0).Enabled = True
        bloque(1).Enabled = True
        txtNotas.Enabled = varModo

End Sub

Public Sub limpaCampoUsuario()

        Dim w As Integer
        
        For w = 0 To txtUsuario.Count - 1
            txtUsuario(w).Text = ""
        Next

        mskTelefone.Text = ""
        mskCelular.Text = ""
        
        'Tipo de conta
        For w = 0 To optConta.Count - 1
            optConta(w).Value = False
        Next
        alterado = False
End Sub

Public Sub registroAlteradoUsuario()

    If modoEdicao Then
        btnUsuario(0).Caption = "&Novo"
        btnUsuario(1).Enabled = False
        btnUsuario(3).Enabled = True
        
        If alterado Then
            btnUsuario(0).Caption = "Cancela&r"
            btnUsuario(1).Enabled = True
        End If
        
    ElseIf alterado Then
        btnUsuario(1).Enabled = True
    End If
    
End Sub

Public Sub preencheControleUsuario(chave As Integer)

On Error GoTo erro
    abreConexao
    rs.Open "SELECT * FROM usuario WHERE codigousuario=" & chave, db, adOpenStatic, adLockOptimistic

    txtUsuario(1).Text = rs!codigousuario
    txtUsuario(0).Text = rs!nome
    txtUsuario(2).Text = rs!endereco
    txtUsuario(3).Text = rs!bairro
    txtUsuario(4).Text = rs!email
    txtUsuario(5).Text = rs!login
    txtUsuario(6).Text = rs!senha
    mskTelefone.Text = rs!telefone
    mskCelular.Text = rs!celular
    exp.Value = 0
    If rs!tipo = "Administrador" Then
        optConta(0).Value = True
    ElseIf rs!tipo = "Usuário" Or rs!tipo = "Usuário-Ex" Then
        optConta(1).Value = True
        If rs!tipo = "Usuário-Ex" Then exp.Value = 1 Else exp.Value = 0
    Else
        optConta(2).Value = True
    End If
   
    habilitaControlesUsuario True
    btnUsuario(1).Enabled = False
    btnUsuario(3).Enabled = True
    btnUsuario(0).Caption = "&Novo"
    modoEdicao = True
    alterado = False
    
Exit Sub

erro:
    MsgBox "Favor contactar o administrador. Erro nº: " & Err

End Sub

Private Function verificaCampoUsuario() As Boolean
    
    Dim w As Integer
    Dim flag As Boolean
    For w = 0 To optConta.Count - 1
        If optConta(w).Value = True Then flag = True
    Next
    
    If txtUsuario(0).Text = "" Then
        MsgBox "O campo " & Chr(34) & "Nome" & Chr(34) & " não pode ser nulo.", vbCritical
        verificaCampoUsuario = True
        txtUsuario(0).SetFocus
    
    ElseIf txtUsuario(2).Text = "" Then
        MsgBox "O campo " & Chr(34) & "Endereço" & Chr(34) & " não pode ser nulo.", vbCritical
        verificaCampoUsuario = True
        txtUsuario(2).SetFocus

    ElseIf txtUsuario(3).Text = "" Then
        MsgBox "O campo " & Chr(34) & "Bairro" & Chr(34) & " não pode ser nulo.", vbCritical
        verificaCampoUsuario = True
        txtUsuario(3).SetFocus

    ElseIf txtUsuario(5).Text = "" Then
        MsgBox "O campo " & Chr(34) & "Login" & Chr(34) & " não pode ser nulo.", vbCritical
        verificaCampoUsuario = True
        txtUsuario(5).SetFocus
    
    ElseIf txtUsuario(6).Text = "" Then
        MsgBox "O campo " & Chr(34) & "Senha" & Chr(34) & " não pode ser nulo.", vbCritical
        verificaCampoUsuario = True
        txtUsuario(6).SetFocus
        
    ElseIf Len(mskTelefone.Text) < 10 And Len(mskTelefone.Text) > 0 Then
        MsgBox "O campo " & Chr(34) & "Telefone" & Chr(34) & " tem quer ter 10 algarismos.", vbCritical
        verificaCampoUsuario = True
        mskTelefone.SetFocus
        
    ElseIf Len(mskCelular.Text) < 10 And Len(mskCelular.Text) > 0 Then
        MsgBox "O campo " & Chr(34) & "Celular" & Chr(34) & " tem quer ter 10 algarismos.", vbCritical
        verificaCampoUsuario = True
        mskCelular.SetFocus
        
    ElseIf Not flag Then
        MsgBox "Escolha o tipo de conta do usuário.", vbCritical
        verificaCampoUsuario = True
        optConta(1).SetFocus
    End If
        
End Function

Public Sub limpaCampoCliente()

        Dim w As Integer
        
        For w = 0 To txtCliente.Count - 1
            If w <> 6 Then txtCliente(w).Text = ""
        Next
        mkCNPJ.Text = ""
        txtNotas.Text = ""
        alteradoCliente = False
        
End Sub

Private Function verificaCampoCliente() As Boolean
    
    Dim w As Integer
    For w = 0 To optConta.Count - 1
        If optConta(w).Value = True Then flag = True
    Next
    
    If txtCliente(0).Text = "" Then
        MsgBox "O campo " & Chr(34) & "Nome" & Chr(34) & " não pode ser nulo.", vbCritical
        verificaCampoCliente = True
        txtCliente(0).SetFocus
    
    ElseIf txtCliente(2).Text = "" Then
        MsgBox "O campo " & Chr(34) & "Endereço" & Chr(34) & " não pode ser nulo.", vbCritical
        verificaCampoCliente = True
        txtCliente(2).SetFocus

    ElseIf txtCliente(3).Text = "" Then
        MsgBox "O campo " & Chr(34) & "Bairro" & Chr(34) & " não pode ser nulo.", vbCritical
        verificaCampoCliente = True
        txtCliente(3).SetFocus

    ElseIf txtCliente(4).Text = "" Then
        MsgBox "O campo " & Chr(34) & "Cidade" & Chr(34) & " não pode ser nulo.", vbCritical
        verificaCampoCliente = True
        txtCliente(4).SetFocus

    ElseIf txtCliente(5).Text = "" Then
        MsgBox "O campo " & Chr(34) & "Fone" & Chr(34) & " não pode ser nulo.", vbCritical
        verificaCampoCliente = True
        txtCliente(5).SetFocus

    ElseIf mkCNPJ.Text = "" Then
        MsgBox "O campo " & Chr(34) & "CNPJ" & Chr(34) & " não pode ser nulo.", vbCritical
        verificaCampoCliente = True
        mkCNPJ.SetFocus
    End If

End Function

Public Sub preencheControleCliente(chave As Integer)

'On Error GoTo erro
    abreConexao
    rs.Open "SELECT * FROM cliente WHERE idcliente=" & chave, db, adOpenStatic, adLockOptimistic

    txtCliente(1).Text = Format$(rs!idcliente, "000")
    txtCliente(0).Text = rs!nome
    txtCliente(2).Text = rs!endereco
    txtCliente(3).Text = rs!bairro
    txtCliente(4).Text = rs!cidade
    txtCliente(5).Text = rs!telefone
    mkCNPJ.Text = rs!CNPJ
    txtCliente(7).Text = rs!celular
    txtCliente(8).Text = rs!contato
    txtCliente(9).Text = rs!email
    txtCliente(10).Text = rs!nfantasia
    txtNotas.Text = rs!notas
   
    If rs!condicao = "Liberado" Then
        bloque(0).Value = True
    Else
        bloque(1).Value = True
    End If
   
    habilitaControlesCliente True
    btnCliente(1).Enabled = False
    btnCliente(3).Enabled = True
    btnCliente(0).Caption = "&Novo"
    modoEdicao = True
    alterado = False
    
Exit Sub

'erro:
'    MsgBox "Favor contactar o administrador. Erro nº: " & Err
End Sub
