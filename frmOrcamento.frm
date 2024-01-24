VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmOrcamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordem de Serviço"
   ClientHeight    =   12900
   ClientLeft      =   1695
   ClientTop       =   2130
   ClientWidth     =   17160
   Icon            =   "frmOrcamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12900
   ScaleWidth      =   17160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7200
      Top             =   11160
   End
   Begin VB.CheckBox Check2 
      Caption         =   "ACABAMENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   114
      Top             =   3550
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtEntrega 
      Height          =   375
      Left            =   5880
      TabIndex        =   110
      Text            =   "entrega"
      Top             =   11280
      Width           =   1095
   End
   Begin VB.Frame Frame12 
      Height          =   1215
      Left            =   8520
      TabIndex        =   106
      Top             =   2200
      Width           =   2415
      Begin VB.OptionButton optEntrega 
         Caption         =   "Sem entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   109
         Top             =   720
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optEntrega 
         Caption         =   "Entrega Rod."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   108
         Top             =   420
         Width           =   1935
      End
      Begin VB.OptionButton optEntrega 
         Caption         =   "Entrega normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   107
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame Frame9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   11040
      TabIndex        =   100
      Top             =   1600
      Width           =   1815
      Begin VB.OptionButton optPagamento 
         Caption         =   "Boleto"
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
         Index           =   4
         Left            =   240
         TabIndex        =   105
         Top             =   1500
         Width           =   1335
      End
      Begin VB.OptionButton optPagamento 
         Caption         =   "Depósito"
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
         Index           =   3
         Left            =   240
         TabIndex        =   104
         Top             =   1185
         Width           =   1335
      End
      Begin VB.OptionButton optPagamento 
         Caption         =   "Cartão"
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
         Left            =   240
         TabIndex        =   103
         Top             =   870
         Width           =   1215
      End
      Begin VB.OptionButton optPagamento 
         Caption         =   "Cheque"
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
         Left            =   240
         TabIndex        =   102
         Top             =   555
         Width           =   1335
      End
      Begin VB.OptionButton optPagamento 
         Caption         =   "Dinheiro"
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
         Left            =   240
         TabIndex        =   101
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox txtCPF 
      Height          =   420
      Left            =   8400
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   741
      _Version        =   393216
      BackColor       =   16777152
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "###.###.###-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTelefone 
      Height          =   420
      Left            =   8640
      TabIndex        =   1
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   741
      _Version        =   393216
      ClipMode        =   1
      BackColor       =   16777152
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(##)#####-####"
      PromptChar      =   "_"
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
      Height          =   420
      Left            =   4080
      TabIndex        =   95
      Text            =   "Text1"
      Top             =   11400
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Caption         =   "Acabamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   91
      Top             =   5000
      Width           =   10095
      Begin VB.Frame Frame14 
         Caption         =   "Wire-o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   6840
         TabIndex        =   116
         Top             =   1320
         Width           =   3135
         Begin VB.ComboBox cmbWireo 
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
            ItemData        =   "frmOrcamento.frx":0ECA
            Left            =   1200
            List            =   "frmOrcamento.frx":0EE9
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtOrcamento 
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
            Height          =   420
            Index           =   8
            Left            =   120
            MaxLength       =   3
            TabIndex        =   30
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox txtOrcamento 
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
         Height          =   420
         Index           =   7
         Left            =   4800
         MaxLength       =   3
         TabIndex        =   29
         Top             =   1850
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Capa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   6840
         TabIndex        =   44
         Top             =   250
         Width           =   3135
         Begin VB.TextBox txtOrcamento 
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
            Height          =   420
            Index           =   3
            Left            =   120
            MaxLength       =   3
            TabIndex        =   24
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cmbCapa 
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
            ItemData        =   "frmOrcamento.frx":0F1A
            Left            =   1200
            List            =   "frmOrcamento.frx":0F2A
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Encadernação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3480
         TabIndex        =   43
         Top             =   250
         Width           =   3135
         Begin VB.OptionButton Option2 
            Caption         =   "Wire-ô"
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
            Left            =   1560
            TabIndex        =   23
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Espiral"
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
            Left            =   360
            TabIndex        =   22
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtOrcamento 
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
            Height          =   420
            Index           =   0
            Left            =   120
            MaxLength       =   3
            TabIndex        =   20
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cmbEncad 
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
            ItemData        =   "frmOrcamento.frx":0F54
            Left            =   1200
            List            =   "frmOrcamento.frx":0F6D
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Laminação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   42
         Top             =   250
         Width           =   3135
         Begin VB.TextBox txtOrcamento 
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
            Height          =   420
            Index           =   2
            Left            =   120
            MaxLength       =   3
            TabIndex        =   17
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
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
            ItemData        =   "frmOrcamento.frx":0FA3
            Left            =   1200
            List            =   "frmOrcamento.frx":0FC2
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Fente e verso"
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
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   1575
         End
      End
      Begin VB.TextBox txtOrcamento 
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
         Height          =   420
         Index           =   6
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   28
         Top             =   1850
         Width           =   1335
      End
      Begin VB.TextBox txtOrcamento 
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
         Height          =   420
         Index           =   5
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   27
         Top             =   1850
         Width           =   1335
      End
      Begin VB.TextBox txtOrcamento 
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
         Height          =   420
         Index           =   4
         Left            =   240
         MaxLength       =   3
         TabIndex        =   26
         Top             =   1850
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vinco:"
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
         Left            =   5220
         TabIndex        =   115
         Top             =   1600
         Width           =   555
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Picote:"
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
         Left            =   3600
         TabIndex        =   94
         Top             =   1600
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Corte a Laser:"
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
         Left            =   1725
         TabIndex        =   93
         Top             =   1600
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Corte:"
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
         Left            =   585
         TabIndex        =   92
         Top             =   1600
         Width           =   525
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1095
      Left            =   120
      TabIndex        =   83
      Top             =   3840
      Width           =   12735
      Begin VB.ComboBox Combo6 
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
         ItemData        =   "frmOrcamento.frx":102A
         Left            =   5040
         List            =   "frmOrcamento.frx":102C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   17160
         TabIndex        =   111
         Top             =   960
         Width           =   1575
         Begin VB.CheckBox Check5 
            Caption         =   "Montagem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   112
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   13230
         TabIndex        =   8
         Top             =   2310
         Width           =   4095
         Begin VB.OptionButton Option4 
            Caption         =   "Brilho"
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
            Left            =   2160
            TabIndex        =   46
            Top             =   165
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Fosca"
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
            Left            =   3120
            TabIndex        =   47
            Top             =   165
            Width           =   975
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Laminação miolo:"
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
            Left            =   180
            TabIndex        =   48
            Top             =   135
            Width           =   1845
         End
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   13230
         TabIndex        =   96
         Top             =   1875
         Width           =   4095
         Begin VB.OptionButton Option3 
            Caption         =   "Brilho"
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
            Left            =   2160
            TabIndex        =   98
            Top             =   120
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Fosca"
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
            Left            =   3120
            TabIndex        =   97
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Laminação capa:"
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
            Left            =   225
            TabIndex        =   99
            Top             =   120
            Width           =   1800
         End
      End
      Begin VB.ComboBox Combo5 
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
         ItemData        =   "frmOrcamento.frx":102E
         Left            =   11520
         List            =   "frmOrcamento.frx":1030
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "frmOrcamento.frx":1032
         Left            =   9840
         List            =   "frmOrcamento.frx":1034
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "frmOrcamento.frx":1036
         Left            =   6480
         List            =   "frmOrcamento.frx":1038
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frmOrcamento.frx":103A
         Left            =   1200
         List            =   "frmOrcamento.frx":104A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtOrcamento 
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
         Height          =   420
         Index           =   1
         Left            =   120
         MaxLength       =   3
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtOrcamento 
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
         Height          =   420
         Index           =   11
         Left            =   14640
         MaxLength       =   3
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtOrcamento 
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
         Height          =   420
         Index           =   10
         Left            =   13080
         MaxLength       =   3
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.Label label 
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
         Index           =   15
         Left            =   5040
         TabIndex        =   113
         Top             =   240
         Width           =   930
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Cores:"
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
         Index           =   9
         Left            =   11520
         TabIndex        =   90
         Top             =   240
         Width           =   690
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Gramatura:"
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
         Index           =   8
         Left            =   9840
         TabIndex        =   89
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Mídia:"
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
         Index           =   7
         Left            =   6480
         TabIndex        =   88
         Top             =   240
         Width           =   645
      End
      Begin VB.Label label 
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
         Index           =   6
         Left            =   1320
         TabIndex        =   87
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Quant.:"
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
         Index           =   5
         Left            =   105
         TabIndex        =   86
         Top             =   255
         Width           =   735
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Caption         =   "Dimensões do banner:"
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
         Index           =   10
         Left            =   13080
         TabIndex        =   85
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Left            =   14310
         TabIndex        =   84
         Top             =   600
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.CheckBox Check4 
      Caption         =   "PHOTOBOOK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   7
      Top             =   3550
      Width           =   1530
   End
   Begin VB.CheckBox Check3 
      Height          =   375
      Left            =   2280
      TabIndex        =   82
      Top             =   11400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtCodigoCliente 
      Height          =   375
      Left            =   2880
      TabIndex        =   81
      Text            =   "0"
      Top             =   11400
      Width           =   975
   End
   Begin VB.CommandButton btnAdd 
      BackColor       =   &H8000000A&
      Caption         =   "&Salvar - Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   10335
      TabIndex        =   33
      ToolTipText     =   "   F2   "
      Top             =   5640
      Width           =   2535
   End
   Begin VB.TextBox txtTotalGeral 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11318
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Text            =   "0,00"
      Top             =   10335
      Width           =   1455
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11318
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Text            =   "0,00"
      Top             =   8955
      Width           =   1455
   End
   Begin VB.TextBox txtAcrescimo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11318
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Text            =   "0,00"
      Top             =   9420
      Width           =   1455
   End
   Begin VB.TextBox txtExemplar 
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
      Left            =   10080
      MaxLength       =   3
      TabIndex        =   40
      Text            =   "1"
      Top             =   8970
      Width           =   855
   End
   Begin VB.Frame Frame7 
      Height          =   1815
      Left            =   120
      TabIndex        =   45
      Top             =   9000
      Width           =   8415
      Begin VB.TextBox txtDescricaoDesconto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   37
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox txtValorDesconto 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   38
         Text            =   "0,00"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
         Caption         =   "Add"
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
         Left            =   6960
         TabIndex        =   39
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtValorAcrescimo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   36
         Text            =   "0,00"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtDescricaoAcrescimo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   35
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Desconto:"
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
         Index           =   17
         Left            =   360
         TabIndex        =   80
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label label 
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
         Index           =   16
         Left            =   5040
         TabIndex        =   79
         Top             =   960
         Width           =   630
      End
      Begin VB.Label label 
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
         Index           =   12
         Left            =   5040
         TabIndex        =   66
         Top             =   240
         Width           =   630
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Outros serviços:"
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
         Index           =   11
         Left            =   360
         TabIndex        =   65
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtData 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Text            =   "20/03/94"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtHora 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   960
      Width           =   1575
   End
   Begin VB.Frame Frame6 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   58
      Top             =   120
      Width           =   10815
      Begin VB.OptionButton Option1 
         Caption         =   "CNPJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   9480
         TabIndex        =   5
         Top             =   1190
         Width           =   908
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CPF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   8280
         TabIndex        =   4
         Top             =   1190
         Value           =   -1  'True
         Width           =   758
      End
      Begin VB.TextBox txtCliente 
         BackColor       =   &H00FFFFC0&
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
         Index           =   2
         Left            =   5400
         TabIndex        =   3
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtCliente 
         BackColor       =   &H00FFFFC0&
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
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   5175
      End
      Begin VB.TextBox txtCliente 
         BackColor       =   &H00FFFFC0&
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
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   8055
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
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
         Index           =   14
         Left            =   5400
         TabIndex        =   72
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
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
         Index           =   13
         Left            =   120
         TabIndex        =   71
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   690
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
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
         Left            =   8760
         TabIndex        =   59
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.TextBox txtDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   11318
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "-0,00"
      Top             =   9885
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   11280
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   9855
      Width           =   1530
   End
   Begin VB.CommandButton btnAdd 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   10335
      TabIndex        =   34
      ToolTipText     =   "   Esc   "
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Salvar - Im&primir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   10335
      TabIndex        =   41
      ToolTipText     =   "   F2   "
      Top             =   6210
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   53
      Top             =   2200
      Width           =   8295
      Begin VB.TextBox txtAlteradoPor 
         Enabled         =   0   'False
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   480
         Width           =   3970
      End
      Begin VB.TextBox txtUsuario 
         Enabled         =   0   'False
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   480
         Width           =   4000
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Alterado por:"
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
         Left            =   4320
         TabIndex        =   52
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
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
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.CommandButton btnAdd 
      BackColor       =   &H8000000A&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   10335
      TabIndex        =   32
      ToolTipText     =   "   Alt + Enter   "
      Top             =   5040
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid gd 
      Height          =   1250
      Left            =   135
      TabIndex        =   49
      Top             =   7680
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   2196
      _Version        =   393216
      Rows            =   4
      Cols            =   15
      BackColor       =   16777215
      BackColorBkg    =   -2147483633
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   11280
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   9405
      Width           =   1530
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   11280
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   8940
      Width           =   1530
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   11280
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   10320
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   11265
      Left            =   12840
      Picture         =   "frmOrcamento.frx":1077
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1905
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Exemplar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8760
      TabIndex        =   78
      Top             =   9000
      Width           =   1185
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10095
      TabIndex        =   77
      Top             =   9915
      Width           =   1110
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11040
      TabIndex        =   68
      Top             =   9045
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Outros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10425
      TabIndex        =   67
      Top             =   9465
      Width           =   780
   End
   Begin VB.Label label 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
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
      Left            =   11160
      TabIndex        =   64
      Top             =   120
      Width           =   570
   End
   Begin VB.Label label 
      AutoSize        =   -1  'True
      Caption         =   "Hora:"
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
      Index           =   4
      Left            =   11160
      TabIndex        =   63
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total R$"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10260
      TabIndex        =   57
      Top             =   10380
      Width           =   945
   End
End
Attribute VB_Name = "frmOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim bAlterado As Boolean
Dim alt As Boolean
Dim ctrl As Boolean
Dim ctrlAlt As Boolean
Dim queroSair As Boolean
Public contadorDePlano As Integer
Public modoEdicao As Boolean
Public nOS As Long
Public nosFinal As Long
Dim codigoCliente As Long
Public cancelPrevisaoDeEntrega As Boolean
Public varDataEntrega As String
Public varHoraEntrega As String
Dim formCarregado As Boolean
Public linhaInicial As Long
Private Function calculaCapa(texto As String) As Double
   
   Dim quant As Integer
   Dim formato As String
   
   quant = preencheCampo(texto, 1)
   formato = preencheCampo(texto, 2)
   
   Select Case formato
      Case "PPA5"
         calculaCapa = quant * 1
      Case "PPA4"
         calculaCapa = quant * 3
      Case "PPA3"
         calculaCapa = quant * 6
   End Select
   
End Function

Function calculacpf(CPF As String) As Boolean
    'Esta rotina foi adaptada da revista Fórum Access

On Error GoTo Err_CPF
    
    Dim i As Integer 'utilizada nos FOR... NEXT
    Dim strcampo As String 'armazena do CPF que será utilizada para o cálculo
    Dim strCaracter As String 'armazena os digitos do CPF da direita para a esquerda
    Dim intNumero As Integer 'armazena o digito separado para cálculo (uma a um)
    Dim intMais As Integer 'armazena o digito específico multiplicado pela sua base
    Dim lngSoma As Long 'armazena a soma dos digitos multiplicados pela sua base(intmais)
    Dim dblDivisao As Double 'armazena a divisão dos digitos*base por 11
    Dim lngInteiro As Long 'armazena inteiro da divisão
    Dim intResto As Integer 'armazena o resto
    Dim intDig1 As Integer 'armazena o 1º digito verificador
    Dim intDig2 As Integer 'armazena o 2º digito verificador
    Dim strConf As String 'armazena o digito verificador
    
    lngSoma = 0
    intNumero = 0
    intMais = 0
    strcampo = Left(CPF, 9)
    
    'Inicia cálculos do 1º dígito
    
    For i = 2 To 10
        strCaracter = Right(strcampo, i - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * i
        lngSoma = lngSoma + intMais
    Next i
    
    dblDivisao = lngSoma / 11
    
    lngInteiro = Int(dblDivisao) * 11
    intResto = lngSoma - lngInteiro
    If intResto = 0 Or intResto = 1 Then
        intDig1 = 0
    Else
        intDig1 = 11 - intResto
    End If
    
    strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
    lngSoma = 0
    intNumero = 0
    intMais = 0
    
    'Inicia cálculos do 2º dígito
    For i = 2 To 11
        strCaracter = Right(strcampo, i - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * i
        lngSoma = lngSoma + intMais
        
    Next i
    
    dblDivisao = lngSoma / 11
    lngInteiro = Int(dblDivisao) * 11
    
    intResto = lngSoma - lngInteiro
    If intResto = 0 Or intResto = 1 Then
        intDig2 = 0
    Else
        intDig2 = 11 - intResto
    End If
    strConf = intDig1 & intDig2
    'Caso o CPF esteja errado dispara a mensagem
    If strConf <> Right(CPF, 2) Then
        calculacpf = False
    Else
        calculacpf = True
    End If
    Exit Function
    
    
Exit_CPF:
        Exit Function
Err_CPF:
        MsgBox Error$
        Resume Exit_CPF
        
End Function


Public Function CalculaCGC(Numero As String) As String

Dim i As Integer
Dim prod As Integer
Dim mult As Integer
Dim digito As Integer

If Not IsNumeric(Numero) Then
   CalculaCGC = ""
   Exit Function
End If

mult = 2
For i = Len(Numero) To 1 Step -1
  prod = prod + Val(Mid(Numero, i, 1)) * mult
  mult = IIf(mult = 9, 2, mult + 1)
Next

digito = 11 - Int(prod Mod 11)
digito = IIf(digito = 10 Or digito = 11, 0, digito)

CalculaCGC = Trim(Str(digito))

End Function

Public Function calculaEncadernacao(texto As String) As Double
   
   Dim folha As String
   Dim modo As String
   
   folha = preencheCampo(texto, 2)
   modo = preencheCampo(texto, 3)
   
   Select Case modo
      Case "E" 'Espiral
         Select Case folha
            Case "Até 50"
               calculaEncadernacao = 2.5
            Case "50/75"
               calculaEncadernacao = 3
            Case "75/100"
               calculaEncadernacao = 4
            Case "100/125"
               calculaEncadernacao = 5
            Case "125/150"
               calculaEncadernacao = 6
            Case "+ 150"
               calculaEncadernacao = 8
         End Select
      Case "W" 'Wireo
         Select Case folha
            Case "Até 50"
               calculaEncadernacao = 3
            Case "50/75"
               calculaEncadernacao = 4
            Case "75/100"
               calculaEncadernacao = 5
            Case "100/125"
               calculaEncadernacao = 6
            Case "125/150"
               calculaEncadernacao = 7
            Case "+ 150"
               calculaEncadernacao = 9
         End Select
   End Select
End Function

Private Function calculaSubTotal(midia As String, grama As String, quantidade As Double) As Double
   
   Select Case midia
      
      Case "CL", "CF"
         Select Case grama
            Case "115g", "120g", "150g", "170g"
               calculaSubTotal = quantidade * 3.5
            Case "250g", "300g"
               calculaSubTotal = quantidade * 4
         End Select
         
      Case "AP"
         Select Case grama
            Case "75g", "90g", "120g"
               calculaSubTotal = quantidade * 3
            Case "150g", "180g", "240g"
               calculaSubTotal = quantidade * 3.5
         End Select
         
      Case "AD"
         calculaSubTotal = quantidade * 5

      Case "CP", "RC"
         calculaSubTotal = quantidade * 5
                  
      Case "PA"
         calculaSubTotal = quantidade * 6
      
      Case "K"
         calculaSubTotal = quantidade * 4
      
      Case "BOPP"
         calculaSubTotal = quantidade * 6
      
   End Select
   
   'If Combo5.Visible = True Then
      If Combo5.Text = 1 Then
         If Combo6.Text = "A4" Then
            calculaSubTotal = quantidade * 0.2
         Else
            calculaSubTotal = quantidade * 0.4
         End If
      End If
      
      If Combo6.Text = "A4" And Combo5.Text = 4 Then
         calculaSubTotal = calculaSubTotal / 2
      End If
      
   'End If
      
   If Combo2.Text = "Impressos F/V" Then
      calculaSubTotal = calculaSubTotal * 2
   End If

End Function




Private Function calculaWireo(texto As String) As Double
   
   Dim quant As Integer
   Dim formato As String
   
   quant = preencheCampo(texto, 1)
   formato = preencheCampo(texto, 2)
   
   Select Case formato
      Case "1/4", "5/16", "3/8"
         calculaWireo = quant * 1.5
      Case "7/16", "1/2", "9/16"
         calculaWireo = quant * 2
      Case "5/8"
         calculaWireo = quant * 2.5
      Case "7/8", "1"
         calculaWireo = quant * 3
         
   End Select

End Function

Public Function ValidaCGC(CGC As String) As Boolean
If CalculaCGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
   ValidaCGC = False
   Exit Function
End If

If CalculaCGC(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
   ValidaCGC = False
   Exit Function
End If

ValidaCGC = True

End Function
Public Sub salvarRegistro(Optional gridrow As Integer)
    
    
    abreConexao
    
      If modoEdicao Then
         rs.Open "Select * from os WHERE idos=" & nOS, db, adOpenStatic, adLockOptimistic
      Else
         rs.Open "Select * from os", db, adOpenStatic, adLockOptimistic
         rs.AddNew
         rs!idUsuario = varCodigoUsuario
         registrosalvo = True
      End If
      
      rs!Data = Format(txtData, "dd/mm/yy")
      rs!hora = CDate(txtHora)
      rs!alteradopor = txtAlteradoPor.Text
      rs!exemplar = txtExemplar.Text
      rs!outros = txtAcrescimo.Text
      rs!desconto = txtDesconto.Text
      rs!acrescimo = txtDescricaoAcrescimo.Text
      rs!descricaodesconto = txtDescricaoDesconto.Text
      rs!outros = txtValorAcrescimo.Text
      rs!dataEntrega = Format(varDataEntrega, "dd/mm/yy")
      rs!horaEntrega = varHoraEntrega
      rs!Montagem = Check5.Value
      
      Dim i As Integer
      i = 0
      For i = 0 To 2
          If optEntrega.Item(i) = True Then rs!entrega = i
      Next
        
      i = 0
      For i = 0 To 4
         If optPagamento.Item(i) = True Then rs!Pagamento = i
      Next
        
        If Check4.Value = 1 Then rs!photobook = 1 Else rs!photobook = 0
        
        If txtCodigoCliente.Text = 0 Then
            rs!nomeCliente = txtCliente(0).Text
            rs!endereço = txtCliente(1).Text
            rs!bairro = txtCliente(2).Text
            rs!CNPJ = txtCPF.Text
            rs!telefonecliente = txtTelefone.Text
            
        Else
            rs!idcliente = txtCodigoCliente.Text
            rs!nomeCliente = txtCliente(0).Text
        End If
        Dim valor As Currency
        valor = txtTotalGeral
        
        rs!valorOs = valor
        
        If modoEdicao Then rs!alteradopor = varNomeUsuario & " - " & Format$(Date, "dd/mm/yy") & " - " & Format$(Time, "hh:mm")

        rs.update
        
        Dim codigoOs As Double
        codigoOs = rs!idos
        frm.numeroOs = codigoOs
        nOS = rs!idos

        rs.Close

' Tabela Plano
'---------------------------------------------------------------------------
        Dim w As Integer
        
        If modoEdicao Then
            rs.Open "DELETE * FROM plano WHERE IDos=" & nOS, db, adOpenStatic, adLockOptimistic
        End If
        
        rs.Open "SELECT * FROM plano", db, adOpenStatic, adLockOptimistic
      
        For w = 1 To contadorDePlano

            rs.AddNew

            rs!idos = codigoOs
            rs!quantidade = gd.TextMatrix(w, 1)
            rs!descricao = gd.TextMatrix(w, 2)
            rs!formato = gd.TextMatrix(w, 3)
            rs!midia = gd.TextMatrix(w, 4)
            
            If gd.TextMatrix(w, 5) = "" Then
               rs!cores = 0
            Else
               rs!cores = gd.TextMatrix(w, 5)
            End If
            
            If gd.TextMatrix(w, 6) = "" Then
               rs!laminacao = 0
            Else
               rs!laminacao = gd.TextMatrix(w, 6)
            End If
            
            If gd.TextMatrix(w, 7) = "" Then
               rs!encadernacao = 0
            Else
               rs!encadernacao = gd.TextMatrix(w, 7)
            End If
            
            If gd.TextMatrix(w, 8) = "" Then
               rs!capa = 0
            Else
               rs!capa = gd.TextMatrix(w, 8)
            End If
            
            If gd.TextMatrix(w, 9) = "" Then
               rs!wireo = 0
            Else
               rs!wireo = gd.TextMatrix(w, 9)
            End If

            If gd.TextMatrix(w, 10) = "" Then
                rs!corte = 0
            Else
                rs!corte = gd.TextMatrix(w, 10)
            End If
            
            If gd.TextMatrix(w, 11) = "" Then
                rs!meiocorte = 0
            Else
                rs!meiocorte = gd.TextMatrix(w, 11)
            End If
            
            If gd.TextMatrix(w, 12) = "" Then
               
            Else
                rs!picote = gd.TextMatrix(w, 12)
            End If
            
            If gd.TextMatrix(w, 13) = "" Then
                rs!vinco = 0
            Else
               rs!vinco = gd.TextMatrix(w, 13)
            End If
            
            rs!valor = gd.TextMatrix(w, 14)
            
            rs.update
            If modoEdicao Then
                rs.MoveNext
            End If
        Next

        rs.Close

With frmPesquisaOs
'-----------------------------------------------------
      If modoEdicao Then
         .grdUsuario.TextMatrix(frmPesquisaOs.linhaos, 1) = txtCliente(0).Text
         Exit Sub
      End If

      Dim rsCliente As Recordset
      Set rsCliente = New Recordset
      
      rs.Open "SELECT * FROM os ORDER BY idos", db, adOpenStatic, adLockOptimistic
      rs.MoveLast
      
      If .grdUsuario.Rows > 31 Then .grdUsuario.Rows = .grdUsuario.Rows + (codigoOs - osInicial)
      rs.Close
      
      rs.Open "SELECT * FROM os WHERE idOS > " & osInicial & " ORDER BY idos", db, adOpenStatic, adLockOptimistic
      rs.MoveFirst
      
      quantidadeOS = codigoOs - osInicial
      
      Dim h As Integer
      h = linhaInicial + 1

      If modoEdicao Then
        .grdUsuario.TextMatrix(.grdUsuario.Row, 1) = txtCliente(0).Text
        Exit Sub
      End If
      
      While Not rs.EOF
          .grdUsuario.TextMatrix(h, 0) = rs!idos
         .grdUsuario.TextMatrix(h, 1) = rs!nomeCliente
         .grdUsuario.TextMatrix(h, 2) = rs!Data
         .grdUsuario.TextMatrix(h, 3) = rs!hora
         .grdUsuario.TextMatrix(h, 4) = varNomeUsuario
         rs.MoveNext
         h = h + 1
      Wend

   .grdUsuario.Row = rs.RecordCount + quantidadeOS
   
   If .grdUsuario.Rows > 31 Then
      .grdUsuario.TopRow = .grdUsuario.Rows - 29
   End If
   
   .grdUsuario.ColSel = .grdUsuario.Cols - 1
   
End With
   rs.Close
   db.Close

End Sub
Public Sub btnAdd_Click(Index As Integer)

cancelPrevisaoDeEntrega = False

Select Case Index
    
    Case 0 ' ----------------------------------------------------------- botão add
        preenchePlano2
        btnAdd(2).Enabled = True
        
    Case 1, 2  '-------------------------------------------------------- botão salvar
    On Error GoTo erro
        alt = False
        If contadorDePlano < 1 Then
            MsgBox "Voce deve adicionar pelo menos um plano.", vbInformation
            txtOrcamento(1).SetFocus
            Exit Sub
        ElseIf Trim(txtCliente(0).Text) = "" Then
            MsgBox "Voce deve informar o nome do Cliente.", vbInformation
            txtCliente(0).SetFocus
            Exit Sub
        ElseIf txtTelefone.Text = "" Or Len(txtTelefone.Text) < 10 Then
            If txtCodigoCliente.Text = 0 Then
                MsgBox "Voce deve informar o telefone do Cliente.", vbInformation
                txtTelefone.SetFocus
                Exit Sub
            End If
        
        '  VERIFICAR CPF AQUI
        '-------------------------------------------------------------------------
        ElseIf txtCodigoCliente.Text = 0 Then
        
            If Len(txtCPF.Text) > 0 Then
                                
                If Len(txtCPF.Text) < 11 And Option1(0).Value = True Then
                    MsgBox "Preencha o CPF corretamente.", vbInformation
                    txtCPF.SetFocus
                    Exit Sub
                End If
                
                If Len(txtCPF.Text) < 14 And Option1(1).Value = True Then
                    MsgBox "Preencha o CNPJ corretamente.", vbInformation
                    txtCPF.SetFocus
                    Exit Sub
                End If
                
                If Option1(0).Value = True Then
                    If Not calculacpf(txtCPF.Text) Then
                        MsgBox "CPF com DV incorreto !!!"
                        txtCPF.Mask = "###.###.###-##"
                        txtCPF.SetFocus
                        Exit Sub
                    End If
                Else
                    If Not ValidaCGC(txtCPF.Text) Then
                        MsgBox "CNPJ com DV incorreto !!! "
                        txtCPF.Mask = "##-###.###/####-##"
                        txtCPF.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        
        End If
        
        salvarRegistro
        
        frmOrcamento.Caption = "Ordem de Serviço: " & nOS

        If Index = 1 Then
            bAlterado = False
            Unload frmOrcamento
        ElseIf Index = 2 Then   ' Salvar e imprimir
        
            Me.cancelPrevisaoDeEntrega = False
            frmPrevisaodeentrega.Show 1
            modoEdicao = True
            If Me.cancelPrevisaoDeEntrega Then Exit Sub
            
            modoEdicao = True
            Load frm
            frm.posPagina
            frm.Show 1
        End If
        
    Case 3
erro:
    'MsgBox ("Erro bobo")
        Unload Me
End Select

End Sub

Private Sub Check1_Click()
    bAlterado = True
End Sub

Private Sub Check2_Click()
   
   bAlterado = True
   gd.Clear
   txtOrcamento(1).Text = ""
    
   contadorDePlano = 0
   cabeçalhoPlano
   
   If Check2.Value = 1 Then
      
      Frame8.Enabled = False
      Check4.Enabled = False
   Else
      Frame8.Enabled = True
      Check4.Enabled = True
   End If
   
End Sub

Private Sub Check3_Click()
bAlterado = True
End Sub

Private Sub Check4_Click()
            
   bAlterado = True
   gd.Clear
   txtOrcamento(1).Text = ""
    
   contadorDePlano = 0
    
   If Check4.Value = 1 Then
      limpa (1)
      Frame5.Enabled = False
      Check2.Enabled = False
      Frame10.Left = 6500
      Frame11.Left = 6500
      Frame10.Top = 115
      Frame11.Top = 445
      Frame13.Left = 11000
      Frame13.Top = 350
      
      label(5).Caption = "Páginas:"
      label(6).Caption = "Formato:"
      label(7).Left = 14000
      label(8).Left = 14000
      label(9).Left = 14000
      Combo4.Left = 14000
      Combo5.Left = 14000
      Combo3.Left = 14000
   
      Combo6.Visible = False
      label(15).Visible = False
      
      Combo2.Clear
      Combo2.AddItem "Photobook 15x21"
      Combo2.AddItem "Photobook 20x20"
      Combo2.AddItem "Photobook 20x29"
        
   Else

      Frame5.Enabled = True
      Check2.Enabled = True
      Frame10.Left = 14000
      Frame11.Left = 14000
      Frame13.Left = 14000
        
      label(5).Caption = "Quant.:"
      label(6).Caption = "Descrição:"
      
      Combo2.Clear
      Combo2.AddItem "Impressos"
      Combo2.AddItem "Impressos F/V"
      Combo2.AddItem "Banner"
      Combo2.AddItem "Fotos"
      
      label(7).Left = 6480
      label(8).Left = 9840
      label(9).Left = 11520
      Combo4.Left = 9840
      Combo5.Left = 11520
      Combo3.Left = 6480
      Combo6.Visible = True
      label(15).Visible = True
          
   End If
      cabeçalhoPlano
      somaTotal
End Sub

Private Sub Check5_Click()
   
   Dim vMontagem As Integer
   vMontagem = 30
   If Check5.Value = 1 Then
      gd.TextMatrix(3, 1) = 1
      gd.TextMatrix(3, 2) = "Montagem"
      gd.TextMatrix(3, 14) = Format(vMontagem, "#,##0.00")
      contadorDePlano = 3
   Else
      gd.TextMatrix(3, 1) = ""
      gd.TextMatrix(3, 2) = ""
      gd.TextMatrix(3, 14) = ""
      contadorDePlano = 2
   End If
   
   somaTotal
End Sub

Private Sub cmbEncad_Click()
   
   If Option2(0).Value = False And Option2(1).Value = False Then
      Option2(0).Value = True
   End If
End Sub

Private Sub cmbWireo_Change()
   bAlterado = True
End Sub

Private Sub Combo1_Click()

   bAlterado = True
   Check1.Enabled = True
   
   Select Case Combo1.ListIndex
          
      Case 4, 5, 6, 7, 8
         Check1.Enabled = False
         Check1.Value = 0
   
   End Select

End Sub

Private Sub Combo1_GotFocus()
    Combo1.ListIndex = 0
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)

   If KeyCode = 46 Then
      Combo1.ListIndex = -1
   End If

End Sub

Private Sub Combo2_Click()
    
   bAlterado = True
   
   If Check4.Value = 1 Then Exit Sub
   
   Combo3.Enabled = True
   Combo4.Enabled = True
   Combo5.Enabled = True
   Combo6.Enabled = True
   Combo3.Clear
   Combo2.Width = 3735
   
   label(10).Visible = False: txtOrcamento(10).Visible = False: txtOrcamento(11).Visible = False: Label1.Visible = False
   Combo4.Visible = True
   Combo5.Visible = True
   Combo6.Visible = True
   label(7).Visible = True
   label(8).Visible = True
   label(9).Visible = True
   label(15).Visible = True
   txtOrcamento(10).Text = ""
   txtOrcamento(11).Text = ""
   txtOrcamento(10).TabStop = False
   txtOrcamento(11).TabStop = False
   Frame5.Enabled = True
   
   Combo5.AddItem ("4")
   Combo3.Clear
   Combo4.Clear
   Combo5.Clear
   cmbWireo.AddItem "1"
    
   Select Case Combo2.Text
     
   Case "Impressos", "Impressos F/V"
   
      Combo3.AddItem ("Couchê Brilho")
      Combo3.AddItem ("Couchê Fosco")
      Combo3.AddItem ("OffSet")
      Combo3.AddItem ("Adesivo")
      Combo3.AddItem ("Color Plus")
      Combo3.AddItem ("Reciclato")
      Combo3.AddItem ("Aspen")
      Combo3.AddItem ("Kraft")
      Combo3.AddItem ("BOPP")
      Combo3.ListIndex = 0
      Combo5.AddItem ("4")
      Combo5.AddItem ("1")
      Combo5.ListIndex = 0
      Combo6.Clear
      Combo6.AddItem "A3"
      Combo6.AddItem "A4"
      Combo6.ListIndex = 0
   
   Case "Banner"
      Combo2.Width = 5175
      Combo4.Visible = False
      Combo5.Visible = False
      Combo6.Visible = False
      label(8).Visible = False
      label(9).Visible = False
      label(10).Left = 9940: txtOrcamento(10).Left = 10020: txtOrcamento(11).Left = 11380: Label1.Left = 11150
      label(10).Visible = True: txtOrcamento(10).Visible = True: txtOrcamento(11).Visible = True: Label1.Visible = True
      label(15).Visible = False
      txtOrcamento(10).TabStop = True
      txtOrcamento(11).TabStop = True
      Combo3.Clear
      Combo3.AddItem ("Lona")
      Combo3.AddItem ("Adesivo L.")
      Combo3.AddItem ("Adesivo T.")
      Frame5.Enabled = False
      limpa (0)
      Combo3.ListIndex = 0
   
   Case "Fotos"
   
      Combo3.AddItem ("Couchê Brilho")
      Combo3.ListIndex = 0
      Combo5.AddItem ("6")
      Combo5.ListIndex = 0
      Combo4.ListIndex = 3
      Combo4.Visible = True
      Combo5.Visible = True
      Combo6.Clear
      Combo6.AddItem ("A5")
      Combo6.ListIndex = 0
      label(8).Visible = True
      label(9).Visible = True
      txtOrcamento(10).Text = ""
      txtOrcamento(11).Text = ""
      txtOrcamento(10).TabStop = False
      txtOrcamento(11).TabStop = False
      Frame5.Enabled = False
      Combo3.Enabled = False
      Combo4.Enabled = False
      Combo5.Enabled = False
      'Combo6.Enabled = False
      limpa (0)

End Select
 
    
    
End Sub

Private Sub Combo3_Click()
    bAlterado = True
    Combo4.Clear
   
    Select Case Combo3.ListIndex
        
        Case 0, 1 ' Couchê
            Combo4.AddItem ("115g")
            Combo4.AddItem ("150g")
            Combo4.AddItem ("170g")
            Combo4.AddItem ("250g")
            Combo4.AddItem ("300g")
            Combo4.ListIndex = 2
        
        Case 2 ' OffSet
            Combo4.AddItem ("75g")
            Combo4.AddItem ("90g")
            Combo4.AddItem ("120g")
            Combo4.AddItem ("150g")
            Combo4.AddItem ("180g")
            Combo4.AddItem ("240g")
            Combo4.ListIndex = 1
            
        Case 3 ' Adesivo
            
            Combo4.AddItem ("90g")
            Combo4.ListIndex = 0
            Check1.Value = 0
            Check1.Enabled = False
            
        Case 4, 5 ' Color Plus
        
            Combo4.AddItem ("120g")
            Combo4.AddItem ("180g")
            Combo4.AddItem ("240g")
            Combo4.ListIndex = 1
                    
        Case 6 ' Aspen
        
            Combo4.AddItem ("120g")
            Combo4.AddItem ("180g")
            Combo4.AddItem ("250g")
            Combo4.ListIndex = 1
    
        Case 7 ' Kraft
            
            Combo4.AddItem ("120g")
            Combo4.AddItem ("180g")
            Combo4.AddItem ("240g")
            Combo4.ListIndex = 1
            
        Case 8 ' BOPP
        
            Combo4.AddItem ("120g")
            Combo4.ListIndex = 0
            
    End Select
    
    If Combo3.Text = "Adesivo" Then
      Check1.Value = 0
      Check1.Enabled = False
   Else
      Check1.Enabled = True
   End If
   
End Sub

Private Sub Combo5_Click()

   bAlterado = True
   
   'If Check1.Value = 0 Then Exit Sub
   
   If Combo5.ListIndex = 1 Then
      Combo3.ListIndex = 2
      Combo4.ListIndex = 1
      Combo3.Locked = True
      Combo4.Locked = True
   Else
      Combo3.Locked = False
      Combo4.Locked = False
   End If
   
End Sub
Private Sub Combo6_Change()
   bAlterado = True
End Sub

Private Sub Command1_Click()

    txtDesconto.Text = "-" & Format$(txtValorDesconto.Text, "#,##0.00")
    txtAcrescimo.Text = Format$(txtValorAcrescimo.Text, "#,##0.00")
    somaTotal
    
End Sub


Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
    
    bAlterado = False
    
' Se OS estiver baixada: trava tudo

If Check3.Value And varTipoUsuario <> "Administrador" Then
   
   txtCliente(0).Enabled = False
   txtCliente(1).Enabled = False
   txtCliente(2).Enabled = False
   txtTelefone.Enabled = False
   txtCPF.Enabled = False
   Option1(0).Enabled = False
   Option1(1).Enabled = False
   Frame9.Enabled = False
   txtOrcamento(1).Enabled = False
   txtOrcamento(2).Enabled = False
   txtOrcamento(3).Enabled = False
   txtOrcamento(4).Enabled = False
   txtOrcamento(5).Enabled = False
   txtOrcamento(6).Enabled = False
   txtOrcamento(10).Enabled = False
   txtOrcamento(11).Enabled = False
   Combo1.Enabled = False
   cmbEncad.Enabled = False
   cmbCapa.Enabled = False
   Combo2.Enabled = False
   Combo3.Enabled = False
   Combo4.Enabled = False
   Combo5.Enabled = False
   txtDescricaoAcrescimo.Enabled = False
   txtDescricaoDesconto.Enabled = False
   txtExemplar.Enabled = False
   txtValorAcrescimo.Enabled = False
   txtValorDesconto.Enabled = False
   
   btnAdd(0).Enabled = False
   btnAdd(1).Enabled = False
   gd.Enabled = False
   
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 18 Then alt = True
    If KeyCode = 17 Then ctrl = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF3 Then
        If varTipoUsuario = "Administrador" Then
            Load frmBaixarOs
            frmBaixarOs.Show 1
        End If
    End If
    
    If alt Then If KeyCode = 13 Then btnAdd_Click (0)
    
    If KeyCode = 18 Then alt = False
    If KeyCode = 17 Then ctrl = False
    
    
End Sub

Private Sub Form_Load()

abreConexao

txtUsuario.Text = varNomeUsuario

Me.Width = 14745
Me.Height = 11340
numeraPlano
contadorDePlano = 0
cabeçalhoPlano

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim op As Integer
    
    If bAlterado = True Then
        op = MsgBox("A OS foi alterada. Deseja salvar agora?", vbQuestion + vbYesNoCancel, "Salvar")
        Select Case op
            Case 2 'Cancelar
                Cancel = 1
            Case 6 'Sim
                btnAdd_Click (1)
            Case 7 ' Não
                Cancel = 0
        End Select
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    modoEdicao = False
End Sub

Private Sub gd_DblClick()
    
    If Check4.Value = 1 Then Exit Sub
    
With gd
    
    If .TextMatrix(.Row, 1) = "" Then Exit Sub
    
    btnAdd(0).Caption = "Alterar"
    btnAdd(1).Enabled = False
    btnAdd(2).Enabled = False
    btnAdd(3).Enabled = False
    gd.Enabled = False
    
    txtOrcamento(1).Text = .TextMatrix(.Row, 1)
    
'BANNER
    If preencheCampo(.TextMatrix(.Row, 2), 1) = "Banner" Then
        Combo2.Text = preencheCampo(.TextMatrix(.Row, 2), 1)
        Combo3.Text = .TextMatrix(.Row, 4)
        txtOrcamento(10).Text = preencheCampo("Banner" & ":" & .TextMatrix(.Row, 3), 2)
        txtOrcamento(11).Text = preencheCampo("Banner" & ":" & .TextMatrix(.Row, 3), 3)
        Exit Sub
    End If
    
        Combo2.Text = .TextMatrix(.Row, 2)
        Combo6.Text = .TextMatrix(.Row, 3)
        Combo3.ListIndex = inverteMidia(preencheCampo(.TextMatrix(.Row, 4), 1))
        Combo4.ListIndex = inverteGramatura(preencheCampo(.TextMatrix(.Row, 4), 1), preencheCampo(.TextMatrix(.Row, 4), 2))
        Combo5.Text = .TextMatrix(.Row, 5)
    
'LAMINAÇÃO
      txtOrcamento(2).Text = preencheCampo(.TextMatrix(.Row, 6), 1)
      If inverteLami(preencheCampo(.TextMatrix(.Row, 6), 2)) <> "" Then
         Combo1.Text = inverteLami(preencheCampo(.TextMatrix(.Row, 6), 2))
      End If
      If preencheCampo(.TextMatrix(.Row, 6), 3) = "FV" Then
         Check1.Value = 1
      Else
         Check1.Value = 0
      End If
    
'ENCADERNAÇÃO
      If .TextMatrix(.Row, 7) <> "" Then
         txtOrcamento(0).Text = preencheCampo(.TextMatrix(.Row, 7), 1)
         cmbEncad.Text = preencheCampo(.TextMatrix(.Row, 7), 2)
         If preencheCampo(.TextMatrix(.Row, 7), 3) = "E" Then
            Option2(0).Value = True
         Else
            Option2(1).Value = True
         End If
      End If
    
'CAPA
      If .TextMatrix(.Row, 8) <> "" Then
         txtOrcamento(3).Text = preencheCampo(.TextMatrix(.Row, 8), 1)
         cmbCapa.Text = converteCapa(preencheCampo(.TextMatrix(.Row, 8), 2))
      End If
    
'WIREO
      If .TextMatrix(.Row, 9) <> "" Then
         txtOrcamento(8).Text = preencheCampo(.TextMatrix(.Row, 9), 1)
         cmbWireo.Text = preencheCampo(.TextMatrix(.Row, 9), 2)
      End If
      
      
      If .TextMatrix(.Row, 10) <> "" Then 'Corte
          txtOrcamento(4).Text = .TextMatrix(.Row, 10)
      End If
      If .TextMatrix(.Row, 11) <> "" Then 'Laser
          txtOrcamento(5).Text = .TextMatrix(.Row, 11)
      End If
      If .TextMatrix(.Row, 12) <> "" Then 'Picote
          txtOrcamento(6).Text = .TextMatrix(.Row, 12)
      End If
      If .TextMatrix(.Row, 13) <> "" Then 'Vinco
          txtOrcamento(7).Text = .TextMatrix(.Row, 13)
      End If
      
'PHOTOBOOK
      If .TextMatrix(1, 2) = "Capa" Then
         
         MsgBox "Photo Book"
         
      End If

End With
    
End Sub

Private Sub gd_GotFocus()
   
    gd.HighLight = flexHighlightWithFocus
        
End Sub
Private Sub gd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Check4.Value = 1 Then Exit Sub
    
    If KeyCode = 46 Then
        gd.RemoveItem (gd.Row)
        gd.Rows = 4
        numeraPlano
        If contadorDePlano > 0 Then
            contadorDePlano = contadorDePlano - 1
        End If
        somaTotal
        If contadorDePlano = 0 Then btnAdd(2).Enabled = False
    End If
    
End Sub

Private Sub alinhaColunas()
    
With gd
    .ColAlignment(0) = 3
    .ColAlignment(1) = 3
    .ColAlignment(2) = 1
    .ColAlignment(3) = 3
    .ColAlignment(4) = 3
    .ColAlignment(5) = 3
    .ColAlignment(6) = 3
    .ColAlignment(7) = 3
    .ColAlignment(8) = 3
    .ColAlignment(9) = 3
    .ColAlignment(10) = 3
    .ColAlignment(11) = 3
    .ColAlignment(12) = 3
    .ColAlignment(13) = 3
    .ColAlignment(14) = 6
    
End With

End Sub

Private Sub numeraPlano()
    Dim w As Integer
    For w = 1 To gd.Rows - 1
        If Len(Str(w)) < 3 Then
            gd.TextMatrix(w, 0) = "0" & w
        Else
            gd.TextMatrix(w, 0) = w
        End If
    Next
End Sub


Private Sub optEntrega_Click(Index As Integer)

    somaTotal
    
End Sub
Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            txtCPF.Mask = "###.###.###-##"
        Case 1
            txtCPF.Mask = "##.###.###/####-##"
    End Select
End Sub

Private Sub optPagamento_Click(Index As Integer)
   bAlterado = True
End Sub


Private Sub Timer1_Timer()
  If modoEdicao = False Then
      txtData.Text = Format(Date, "dd/mm/yy")
      txtHora.Text = Time
  End If
End Sub

Private Sub txtAcrescimo_Change()
   bAlterado = True
End Sub

Private Sub txtCliente_Change(Index As Integer)
    bAlterado = True
End Sub

Private Sub txtCliente_GotFocus(Index As Integer)
    txtCliente(Index).SelStart = 0
    txtCliente(Index).SelLength = Len(txtCliente(Index).Text)
End Sub

Private Sub txtCliente_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        abreConexao
        rs.Open "SELECT * FROM cliente ORDER BY nome", db, adOpenStatic, adLockOptimistic
        
        If rs.RecordCount = 0 Then
            MsgBox "Nenhum clinte foi cadastrado"
        Else
            frmPesquisaCliente.Show 1
        End If
    End If
    
End Sub

Private Sub txtCPF_Change()
    bAlterado = True
End Sub

Private Sub txtCPF_GotFocus()
    txtCPF.SelStart = 0
    If Option1(0).Value = True Then
        txtCPF.SelLength = Len(txtCPF.Text) + 14 - Len(txtCPF.Text)
    Else
        txtCPF.SelLength = Len(txtCPF.Text) + 18 - Len(txtCPF.Text)
    End If
End Sub

Private Sub txtDesconto_Change()
   bAlterado = True
End Sub

Private Sub txtDescricaoAcrescimo_Change()
    bAlterado = True
End Sub

Private Sub txtDescricaoDesconto_Change()
    bAlterado = True
End Sub

Private Sub txtEntrega_Change()
   bAlterado = True
End Sub

Private Sub txtExemplar_Change()
    bAlterado = True
    If txtExemplar.Text <> "" And txtExemplar.Text <> "0" Then
        somaTotal
    Else
        txtExemplar.Text = 1
        somaTotal
        txtExemplar.SelStart = 0
        txtExemplar.SelLength = Len(txtExemplar.Text)
    End If

End Sub

Private Sub txtExemplar_GotFocus()
    txtExemplar.SelStart = 0
    txtExemplar.SelLength = Len(txtExemplar.Text)
End Sub

Private Sub txtOrcamento_Change(Index As Integer)
   bAlterado = True
   Select Case Index
      Case 2
         If txtOrcamento(2).Text = "0" Or txtOrcamento(2).Text = "" Then
            Combo1.ListIndex = 0
            Check1.Value = 0
         End If
      
      Case 3
         If txtOrcamento(3).Text = "0" Or txtOrcamento(3).Text = "" Then
            cmbEncad.ListIndex = 0
         End If
      
      Case 0
         If txtOrcamento(0).Text = "0" Or txtOrcamento(0).Text = "" Then
            cmbEncad.ListIndex = 0
            Option2(0).Value = False
            Option2(1).Value = False
         End If
         
   End Select
End Sub

Private Sub txtOrcamento_GotFocus(Index As Integer)
    txtOrcamento(Index).SelStart = 0
    txtOrcamento(Index).SelLength = Len(txtOrcamento(Index).Text)
End Sub

Private Sub txtOrcamento_KeyPress(Index As Integer, KeyAscii As Integer)

    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtOrcamento_LostFocus(Index As Integer)
    If txtOrcamento(Index).Text = "0" Then txtOrcamento(Index) = ""
End Sub

Private Function converteMidia(midia As String) As String

Select Case midia
   
   Case "Couchê Brilho"
      converteMidia = "CL"
   Case "Couchê Fosco"
      converteMidia = "CF"
   Case "OffSet"
      converteMidia = "AP"
   Case "Adesivo"
      converteMidia = "AD"
   Case "Color Plus"
      converteMidia = "CP"
   Case "Kraft"
      converteMidia = "K"
   Case "Aspen"
      converteMidia = "PA"
   Case "BOPP"
     converteMidia = "BOPP"
   Case "Reciclato"
      converteMidia = "RC"
End Select

End Function
Private Function converteLami(lami As String) As String

   Select Case lami
   
   Case "BOPP Brilho"
      converteLami = "LB"
   Case "BOPP Fosco"
      converteLami = "LF"
   Case "Polaseal A7"
      converteLami = "PA7"
   Case "Polaseal A6"
      converteLami = "PA6"
   Case "Polaseal A5"
      converteLami = "PA5"
   Case "Polaseal A4"
      converteLami = "PA4"
   Case "Polaseal A3"
      converteLami = "PA3"
   Case "Verniz"
      converteLami = "VZ"
   
   End Select

End Function
Private Function calculaLami() As Double
   
   Dim texto As String
   Dim formato As String
   Dim valor As Double
   Dim quantidade As Integer
   
   
   
   texto = Combo1.Text
   formato = Combo6.Text
   quantidade = txtOrcamento(2).Text
   
   Select Case texto
   
      Case "BOPP Brilho", "BOPP Fosco", "Verniz"
         
         If Check1.Value = 1 Then
            valor = 2
         Else
            valor = 1
         End If
         If formato = "A4" Then valor = valor / 2

      Case "Polaseal A7"
         valor = 1
      Case "Polaseal A6"
         valor = 1.5
      Case "Polaseal A5"
         valor = 2
      Case "Polaseal A4"
         valor = 3
      Case "Polaseal A3"
         valor = 5
   
   End Select
   
   calculaLami = valor * quantidade

End Function

Private Sub limpa(tudo As Integer)

   Dim w As Integer
   
   For w = 2 To 8
      txtOrcamento(w) = ""
   Next
   
   txtOrcamento(0) = ""
   Combo1.ListIndex = -1
   cmbEncad.ListIndex = -1
   cmbCapa.ListIndex = -1
   cmbWireo.ListIndex = -1
   Check1.Value = 0
   Option2(0).Value = False
   Option2(1).Value = False
   
   If tudo = 1 Then
      txtOrcamento(1).Text = ""
      Combo2.ListIndex = -1
      Combo3.ListIndex = -1
      Combo4.ListIndex = -1
      Combo5.ListIndex = -1
      Combo6.Clear
      txtOrcamento(10).Text = ""
      txtOrcamento(11).Text = ""
   End If

End Sub
Private Sub somaTotal()
    Dim soma As Currency
    Dim w As Integer
    soma = 0

    For w = 1 To gd.Rows - 1
        If gd.TextMatrix(w, 14) <> "" Then
            soma = soma + CCur(gd.TextMatrix(w, 14))
        End If
    Next
    
    txtTotal.Text = Format$(soma, "#,##0.00")
    
    Dim valorEntrega As Double
    
    If optEntrega(0).Value = True Then
        valorEntrega = 5
    Else
        valorEntrega = 0
    End If
    
    txtTotalGeral.Text = Format$((CDbl(txtTotal.Text) * txtExemplar.Text) + CDbl(txtDesconto.Text) + CDbl(txtAcrescimo.Text) + valorEntrega, "#,##0.00")

End Sub
Private Function verificaCampoNulo() As Boolean

Dim q As Boolean
Dim Controle As TextBox

If txtOrcamento(1).Text = "" Or txtOrcamento(1).Text = "0" Then 'Verifica campo Quantidade
    Set Controle = txtOrcamento(1)
    q = True
ElseIf Combo2.Text = "" Then 'Verifica Campo Descrição
    alt = False
    MsgBox "Digite um valor válido para este campo.", vbInformation
    Combo2.SetFocus
    contadorDePlano = contadorDePlano - 1
    verificaCampoNulo = True
ElseIf Combo2.Text = "Banner" Then ' Verifica Campo Dimensões do banner
    If txtOrcamento(10).Text = "" Then
        Set Controle = txtOrcamento(10)
        q = True
    ElseIf txtOrcamento(11).Text = "" Then
        Set Controle = txtOrcamento(11)
        q = True
    End If
End If

If q Then
    alt = False
    MsgBox "Digite um valor válido para este campo.", vbInformation
    Controle.SetFocus
    contadorDePlano = contadorDePlano - 1
    verificaCampoNulo = True
End If

End Function
Private Function converteCapa(capa As String) As String

   Select Case capa
      Case "Papelão A5"
         converteCapa = "PPA5"
      Case "Papelão A4"
         converteCapa = "PPA4"
      Case "Papelão A3"
         converteCapa = "PPA3"
      Case "PPA5"
         converteCapa = "Papelão A5"
      Case "PPA4"
         converteCapa = "Papelão A4"
      Case "PPA3"
         converteCapa = "Papelão A3"

   End Select

End Function
Public Function preencheCampo(campo As String, pos As Integer) As String

Dim d As Integer
Dim contador As Integer
Dim resultado As String
Dim texto1 As String
Dim texto2 As String
Dim texto3 As String
contador = 1

For d = 1 To Len(campo) + 1
   If Mid$(campo, d, 1) = ":" Or d = Len(campo) + 1 Then
      d = d + 1
      If contador = 1 Then
         texto1 = resultado
         resultado = ""
         contador = 2
      ElseIf contador = 2 Then
         texto2 = resultado
         contador = 3
         resultado = ""
      ElseIf contador = 3 Then
         texto3 = resultado
      End If
   End If
   resultado = resultado + Mid$(campo, d, 1)
Next

   If texto1 = "Banner" Then
      contador = 1
      campo = texto2
      For d = 1 To Len(campo) + 1
      If Mid$(campo, d, 1) = "x" Then
         d = d + 1
         If contador = 1 Then
         texto2 = resultado
         resultado = ""
         End If
         Else
         texto3 = resultado
      End If
      resultado = resultado + Mid$(campo, d, 1)
      Next
   End If

If pos = 1 Then
    preencheCampo = Trim(texto1)
ElseIf pos = 2 Then
    preencheCampo = Trim(texto2)
ElseIf pos = 3 Then
    preencheCampo = Trim(texto3)
End If

End Function

Private Function inverteMidia(midia As String) As Integer

   Select Case midia
      Case "CL"
         inverteMidia = 0 'Couchê Brilho
      Case "CF"
         inverteMidia = 1 'Couchê Fosco
      Case "AP"
         inverteMidia = 2 'OffSet
      Case "AD"
         inverteMidia = 3 'Adesivo
      Case "CP"
         inverteMidia = 4 'Calor plus
      Case "RC"
         inverteMidia = 5 'Reciclato
      Case "PA"
        inverteMidia = 6 'Aspen
      Case "K"
        inverteMidia = 7 'Adesivo
      Case "BOPP"
        inverteMidia = 8 'BOPP
   End Select

End Function

Private Function inverteLami(lami As String) As String

Select Case lami

    Case "LB"
        inverteLami = "BOPP Brilho"
    Case "LF"
        inverteLami = "BOPP Fosco"
   Case "PA7"
        inverteLami = "Polaseal A7"
    Case "PA6"
        inverteLami = "Polaseal A6"
    Case "PA5"
        inverteLami = "Polaseal A5"
    Case "PA4"
        inverteLami = "Polaseal A4"
    Case "PA3"
        inverteLami = "Polaseal A3"
   Case "VZ"
      inverteLami = "Verniz"
End Select

End Function

Private Function inverteCapa(capa As String) As String

   Select Case capa
   
       Case "PPA5"
           inverteCapa = "Papelão A5"
       Case "PPA4"
           inverteCapa = "Papelão A4"
       Case "PPA3"
           inverteCapa = "Papelão A3"
       Case "PVCA5"
           inverteCapa = "Capa PVC A5"
       Case "PVCA4"
           inverteCapa = "Capa PVC A4"
   End Select

End Function
Private Sub txtTelefone_Change()
    bAlterado = True
End Sub

Private Sub txtTelefone_GotFocus()
    txtTelefone.SelStart = 0
    txtTelefone.SelLength = Len(txtTelefone.Text) + 14
    End Sub

Private Sub txtValorAcrescimo_Change()
   bAlterado = True
   
End Sub

Private Sub txtValorAcrescimo_GotFocus()
    txtValorAcrescimo.SelStart = 0
    txtValorAcrescimo.SelLength = Len(txtValorAcrescimo.Text)
End Sub

Private Sub txtValorAcrescimo_KeyPress(KeyAscii As Integer)
   
   Dim texto As String
   Dim cont As Integer
   
   For cont = 1 To Len(txtValorAcrescimo.Text)
      
   Next
   
   
   If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) And Not (KeyAscii = 44) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtValorAcrescimo_LostFocus()
   txtValorAcrescimo.Text = Format$(txtValorAcrescimo.Text, "#,##0.00")
   
   If txtValorAcrescimo = "" Then txtValorAcrescimo = "0,00"
   
   
End Sub

Private Sub txtValorDesconto_Change()
   bAlterado = True

End Sub

Private Sub txtValorDesconto_GotFocus()
    txtValorDesconto.SelStart = 0
    txtValorDesconto.SelLength = Len(txtValorDesconto.Text)
End Sub

Private Function clienteCadastrado(CPF As String) As Boolean

    Dim sCriterio As String
    
    If Option1(0).Value = True Then
        sCriterio = Format$(CPF, "@@@.@@@.@@@-@@")
    Else
        sCriterio = Format$(CPF, "@@.@@@.@@@/@@@@-@@")
    End If
    
    rs.Open "SELECT * FROM cliente WHERE cpf='" & sCriterio & "'", db, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        clienteCadastrado = True
        codigoCliente = rs!idcliente
    End If
    rs.Close
End Function

Private Sub atualizaGridOs()
    
    Dim linhagrid As Integer
    
With frmPesquisaOs
    linhagrid = .grdUsuario.Row
    
    If modoEdicao Then
        .grdUsuario.TextMatrix(linhagrid, 0) = rs!idos
        .grdUsuario.TextMatrix(linhagrid, 1) = rs!nomeCliente
        .grdUsuario.TextMatrix(linhagrid, 2) = rs!Data
        .grdUsuario.TextMatrix(linhagrid, 3) = rs!hora

    Else
        .grdUsuario.TextMatrix(.grdUsuario.Rows - 9, 0) = rs!idos
        .grdUsuario.TextMatrix(.grdUsuario.Rows - 9, 1) = rs!nomeCliente
        .grdUsuario.TextMatrix(.grdUsuario.Rows - 9, 2) = rs!Data
        .grdUsuario.TextMatrix(.grdUsuario.Rows - 9, 3) = rs!hora
        
        .grdUsuario.Rows = .grdUsuario.Rows + 1
        .grdUsuario.TopRow = .grdUsuario.Rows - 30
    End If
    
    
End With

End Sub

Private Sub cabeçalhoPlano()
With gd
    
    alinhaColunas
    'defineLargura
    
    .ColWidth(0) = 500
    .ColWidth(1) = 500
    .ColWidth(2) = 1700
    .ColWidth(3) = 850
    .ColWidth(4) = 1000
    .ColWidth(5) = 350
    .ColWidth(6) = 1000
    .ColWidth(7) = 1300
    .ColWidth(8) = 1030
    .ColWidth(9) = 600
    .ColWidth(9) = 950
    .ColWidth(10) = 600
    .ColWidth(11) = 600
    .ColWidth(12) = 600
    .ColWidth(13) = 600
'------------Título
    .TextMatrix(0, 1) = "Qtd"
    .TextMatrix(0, 2) = "Descrição"
    .TextMatrix(0, 3) = "For."
    .TextMatrix(0, 4) = "Mídia"
    .TextMatrix(0, 5) = "C"
    .TextMatrix(0, 6) = "Lam."
    .TextMatrix(0, 7) = "Encad."
    .TextMatrix(0, 8) = "Capa"
    .TextMatrix(0, 9) = "Wireo"
    .TextMatrix(0, 10) = "Co"
    .TextMatrix(0, 11) = "Las."
    .TextMatrix(0, 12) = "Pic."
    .TextMatrix(0, 13) = "Vin."
    .TextMatrix(0, 14) = "Sub total"

   numeraPlano
End With
End Sub

Public Sub preenchePlano2()

Dim r As Integer
Dim varQuantidade As Double
Dim varSubTotal As Double
Dim varSubTotalBanner As Double

   If btnAdd(0).Caption = "&Add" Then
      If contadorDePlano < 3 Then
          contadorDePlano = contadorDePlano + 1
      Else
          alt = False
          MsgBox "O quadro de planos está cheio." + Chr(13) + "Delete um plano ou inicie outra OS.", vbInformation
          Exit Sub
       End If
       r = contadorDePlano
   Else
       btnAdd(0).Caption = "&Add"
       btnAdd(1).Enabled = True
       btnAdd(2).Enabled = True
       btnAdd(3).Enabled = True
       gd.Enabled = True
       r = gd.Row
   End If
   
   If verificaCampoNulo Then Exit Sub

'PREÇOS---------------------------------
   
   Dim PB As Double
   Dim A4AP As Double
   Dim A4CL170 As Double
   Dim A4CL300 As Double
   Dim A3CL170 As Double
   Dim A3CL300 As Double
   Dim A3AP As Double
   Dim AD As Double
   Dim BOPP As Double
   Dim BN As Double
   Dim MC As Double
   
   A3CL170 = 3.5
   
'--------------------------------------
   
   varQuantidade = txtOrcamento(1).Text
   
With gd

'IMPRESSOS
   If Combo2.Text = "Impressos" Or Combo2.Text = "Impressos F/V" Then
      'Descrição
      .TextMatrix(r, 1) = varQuantidade
      .TextMatrix(r, 2) = Combo2.Text
      
      varSubTotal = calculaSubTotal(converteMidia(Combo3.Text), Combo4.Text, varQuantidade)
      'Formato
      .TextMatrix(r, 3) = Combo6.Text
      'Midia
      .TextMatrix(r, 4) = converteMidia(Combo3.Text) & ":" & Combo4.Text
   End If

'CORES
   
   .TextMatrix(r, 5) = Combo5.Text

'LAMINAÇÃO

   If txtOrcamento(2).Text = "" Or Combo1.Text = "" Then
      .TextMatrix(r, 6) = ""
   Else
      If Check1.Value = 1 Then
         .TextMatrix(r, 6) = txtOrcamento(2).Text + ":" & converteLami(Combo1.Text) + ":" + "FV"
         varSubTotal = varSubTotal + calculaLami
      Else
         .TextMatrix(r, 6) = txtOrcamento(2).Text + ":" & converteLami(Combo1.Text)
         varSubTotal = varSubTotal + calculaLami
      End If
   End If
   
'ENCADERNAÇÃO
      If txtOrcamento(0).Text = "" Or cmbEncad.Text = "" Then
         .TextMatrix(r, 7) = ""
      Else
         Dim encQuan As Integer
         Dim encText As String
         
         encQuan = txtOrcamento(0).Text
         encText = cmbEncad.Text
         
         If Option2(0).Value = True Then
            .TextMatrix(r, 7) = encQuan & ":" & encText & ":" & "E"
         ElseIf Option2(1).Value = True Then
            .TextMatrix(r, 7) = encQuan & ":" & encText & ":" & "W"
         End If
         
         varSubTotal = varSubTotal + encQuan * calculaEncadernacao(.TextMatrix(r, 7))
      End If
      
'CAPA
      If txtOrcamento(3).Text = "" Or cmbCapa.Text = "" Then
         .TextMatrix(r, 8) = ""
      Else
         .TextMatrix(r, 8) = txtOrcamento(3).Text & ":" & converteCapa(cmbCapa.Text)
         varSubTotal = varSubTotal + calculaCapa(.TextMatrix(r, 8))
      End If

'WIREO
      If txtOrcamento(8).Text = "" Or cmbWireo.Text = "" Then
         .TextMatrix(r, 9) = ""
      Else
         .TextMatrix(r, 9) = txtOrcamento(8).Text & ":" & cmbWireo.Text
         varSubTotal = varSubTotal + calculaWireo(.TextMatrix(r, 9))
      End If
      
'CORTE
      If txtOrcamento(4).Text <> "" Then
         .TextMatrix(r, 10) = txtOrcamento(4).Text
         varSubTotal = varSubTotal + txtOrcamento(4).Text * 0.5
      Else
         .TextMatrix(r, 10) = ""
      End If
   
'LASER
      If txtOrcamento(5).Text <> "" Then
         .TextMatrix(r, 11) = txtOrcamento(5).Text
         varSubTotal = varSubTotal + txtOrcamento(5).Text * 3
      Else
         .TextMatrix(r, 11) = ""
      End If
      
'PICOTE
      If txtOrcamento(6).Text <> "" Then
         .TextMatrix(r, 12) = txtOrcamento(6).Text
         varSubTotal = varSubTotal + 15
      Else
         .TextMatrix(r, 12) = ""
      End If
   
'VINCO
      If txtOrcamento(7).Text <> "" Then
         .TextMatrix(r, 13) = txtOrcamento(7).Text
         varSubTotal = varSubTotal + 15
      Else
         .TextMatrix(r, 13) = ""
      End If
      
'BANNER
   If Combo2.Text = "Banner" Then
      Combo6.Enabled = False
      .TextMatrix(r, 1) = varQuantidade
      .TextMatrix(r, 2) = Combo2.Text
      .TextMatrix(r, 3) = txtOrcamento(10).Text & "x" & txtOrcamento(11).Text
      .TextMatrix(r, 4) = Combo3.Text
      .TextMatrix(r, 5) = 6
      
      varSubTotalBanner = varQuantidade * txtOrcamento(10).Text * txtOrcamento(11).Text * 45 / 10000
      If varSubTotalBanner < 10 Then
         varSubTotal = 10
      Else
         varSubTotal = varSubTotalBanner
      End If
   End If

'PHOTOBOOK
   
   If Check4.Value = 1 Then
      Dim vBookCapa As Double
      Dim vBookMiolo As Double
      Dim vMontagem As Double
      
      vBookCapa = 12
      vBookMiolo = 8 'R$ 8,00 cada A3
      
      If Check5.Value = 1 Then
         vMontagem = 30
      End If
      
      Dim quanPaginas As Integer
      quanPaginas = txtOrcamento(1).Text - 2
      
      If Combo2.Text = "Photobook 15x21" Then
         
         Dim inteiro As Integer
         Dim divisao As Double
         Dim v20x29 As Integer

         If quanPaginas / 4 > Int(quanPaginas / 4) Then
            quanPaginas = Int(quanPaginas / 4) + 1
         Else
            quanPaginas = quanPaginas / 4
         End If
      Else
         quanPaginas = quanPaginas / 2
      End If
      
      Select Case Combo2.Text

         Case "Photobook 15x21"
            gd.TextMatrix(2, 14) = Format(quanPaginas * vBookMiolo, "#,##0.00")
            gd.TextMatrix(1, 3) = "15x21"
            gd.TextMatrix(2, 3) = "15x21"
      
         Case "Photobook 20x20", "Photobook 20x20"
            gd.TextMatrix(2, 14) = Format(quanPaginas * vBookMiolo, "#,##0.00")
            gd.TextMatrix(1, 3) = "20x20"
            gd.TextMatrix(2, 3) = "20x20"
         
         Case "Photobook 20x29"
            gd.TextMatrix(2, 14) = Format(quanPaginas * vBookMiolo, "#,##0.00")
            gd.TextMatrix(1, 3) = "20x29"
            gd.TextMatrix(2, 3) = "20x29"
         
      End Select
      
      gd.TextMatrix(1, 1) = 1
      gd.TextMatrix(1, 2) = "Capa " '& Combo2.Text
      gd.TextMatrix(1, 4) = "CL:170g"
      gd.TextMatrix(1, 5) = "4"
      
      If Option3(0).Value = True Then
         gd.TextMatrix(1, 6) = "LB"
      Else
         gd.TextMatrix(1, 6) = "LF"
      End If
      
      gd.TextMatrix(1, 14) = Format(vBookCapa, "#,##0.00")
      
      'Miolo--------------------------------------------------------------------------------------------
      
      gd.TextMatrix(2, 1) = quanPaginas
      gd.TextMatrix(2, 2) = "Miolo " '& Combo2.Text
      gd.TextMatrix(2, 4) = "CL:250g"
      gd.TextMatrix(2, 5) = "4"
      
      If Check5.Value = 1 Then
         gd.TextMatrix(3, 1) = 1
         gd.TextMatrix(3, 2) = "Montagem"
         gd.TextMatrix(3, 14) = Format(vMontagem, "#,##0.00")
      End If
      
      If Option4(0).Value = True Then
         gd.TextMatrix(2, 6) = "LB "
      Else
         gd.TextMatrix(2, 6) = "LF"
      End If
      
      contadorDePlano = 3
   End If
   
   'fotos
   If Combo2.Text = "Fotos" Then
      If txtOrcamento(1).Text < 8 Or txtOrcamento(1).Text = "" Then
         MsgBox "A quantidade mínima para fotos é de 8 unidades", vbInformation
         txtOrcamento(1).Text = 8
      End If
      
      .TextMatrix(r, 1) = txtOrcamento(1).Text
      .TextMatrix(r, 2) = Combo2.Text
      .TextMatrix(r, 3) = Combo6.Text
      .TextMatrix(r, 4) = "CL250g"
      
      varSubTotal = txtOrcamento(1).Text * 0.6
   
   End If
   
   limpa (1)
   txtOrcamento(1).SetFocus

   If Check4.Value = 0 Then
      .TextMatrix(r, 14) = Format(varSubTotal, "#,##0.00")
   End If
   somaTotal
   
End With

End Sub

Private Function inverteGramatura(midia As String, peso As String) As Integer

Select Case midia
        
        Case "CL", "CF" ' Couchê
            Select Case peso
               Case "120g": inverteGramatura = 0
               Case "150g": inverteGramatura = 1
               Case "170g": inverteGramatura = 2
               Case "250g": inverteGramatura = 3
               Case "300g": inverteGramatura = 4
            End Select
        
        Case "AP" ' OffSet"
            Select Case peso
               Case "75g": inverteGramatura = 0
               Case "90g": inverteGramatura = 1
               Case "120g": inverteGramatura = 2
               Case "150g": inverteGramatura = 3
               Case "180g": inverteGramatura = 4
               Case "240g": inverteGramatura = 5
            End Select
            
        Case "AD" ' Adesivo
            Combo4.AddItem ("90g")
            
        Case "CP" ' Color Plus
            Combo4.AddItem ("120g")
            Combo4.AddItem ("180g")
            Combo4.AddItem ("240g")
                    
        Case "PA" ' Aspen
            Combo4.AddItem ("120g")
            Combo4.AddItem ("180g")
            Combo4.AddItem ("250g")
    
        Case "K" ' Kraft
            Combo4.AddItem ("120g")
            Combo4.AddItem ("180g")
            Combo4.AddItem ("240g")
            
        Case "BOPP" ' BOPP
            Combo4.AddItem ("120g")
            
    End Select
End Function

Private Sub txtValorDesconto_KeyPress(KeyAscii As Integer)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) And Not (KeyAscii = 44) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtValorDesconto_LostFocus()
   txtValorDesconto.Text = Format$(txtValorDesconto.Text, "#,##0.00")
   
   If txtValorDesconto.Text = "" Then txtValorDesconto = "0,00"
End Sub

