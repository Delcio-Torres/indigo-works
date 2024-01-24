VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOrcamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordem de Serviço"
   ClientHeight    =   11025
   ClientLeft      =   1695
   ClientTop       =   2130
   ClientWidth     =   21510
   Icon            =   "formNovo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11025
   ScaleWidth      =   21510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   13200
      TabIndex        =   87
      Text            =   "Text1"
      Top             =   1800
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
      TabIndex        =   80
      Top             =   4560
      Width           =   10095
      Begin VB.Frame Frame4 
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
         Height          =   1215
         Left            =   6840
         TabIndex        =   83
         Top             =   360
         Width           =   3135
         Begin VB.CheckBox Check2 
            Caption         =   "Embutir wire-o"
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
            TabIndex        =   13
            Top             =   840
            Width           =   2295
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
            Index           =   3
            Left            =   120
            MaxLength       =   3
            TabIndex        =   11
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
            Index           =   2
            ItemData        =   "formNovo.frx":000C
            Left            =   1200
            List            =   "formNovo.frx":002E
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Capas"
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
         TabIndex        =   82
         Top             =   360
         Width           =   3135
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
            Index           =   1
            ItemData        =   "formNovo.frx":006B
            Left            =   240
            List            =   "formNovo.frx":0081
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   480
            Width           =   2655
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
         TabIndex        =   81
         Top             =   360
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
            TabIndex        =   7
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
            Index           =   0
            ItemData        =   "formNovo.frx":00C5
            Left            =   1200
            List            =   "formNovo.frx":00DE
            Style           =   2  'Dropdown List
            TabIndex        =   8
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
            TabIndex        =   9
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
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   16
         Top             =   1800
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
         Index           =   5
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   15
         Top             =   1800
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
         Index           =   4
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Picote:"
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
         Left            =   2670
         TabIndex        =   86
         Top             =   1890
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Meio corte:"
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
         Left            =   5160
         TabIndex        =   85
         Top             =   1890
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Corte:"
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
         Left            =   360
         TabIndex        =   84
         Top             =   1890
         Width           =   630
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1095
      Left            =   120
      TabIndex        =   70
      Top             =   3480
      Width           =   15855
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   13230
         TabIndex        =   92
         Top             =   2430
         Width           =   4575
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
            TabIndex        =   94
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
            TabIndex        =   93
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
            TabIndex        =   95
            Top             =   135
            Width           =   1845
         End
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   13230
         TabIndex        =   88
         Top             =   1995
         Width           =   4575
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
            TabIndex        =   90
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
            TabIndex        =   89
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
            TabIndex        =   91
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
         ItemData        =   "formNovo.frx":0131
         Left            =   11520
         List            =   "formNovo.frx":013B
         Style           =   2  'Dropdown List
         TabIndex        =   6
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
         ItemData        =   "formNovo.frx":0145
         Left            =   9840
         List            =   "formNovo.frx":0147
         Style           =   2  'Dropdown List
         TabIndex        =   5
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
         ItemData        =   "formNovo.frx":0149
         Left            =   6480
         List            =   "formNovo.frx":0165
         Style           =   2  'Dropdown List
         TabIndex        =   4
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
         ItemData        =   "formNovo.frx":01C5
         Left            =   1200
         List            =   "formNovo.frx":01D5
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   5175
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
         TabIndex        =   2
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
         TabIndex        =   72
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
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Left            =   14310
         TabIndex        =   73
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
      Left            =   13680
      TabIndex        =   69
      Top             =   1320
      Width           =   1530
   End
   Begin VB.CheckBox Check3 
      Height          =   375
      Left            =   13200
      TabIndex        =   68
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtCodigoCliente 
      Height          =   375
      Left            =   13200
      TabIndex        =   67
      Text            =   "0"
      Top             =   240
      Visible         =   0   'False
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
      TabIndex        =   18
      ToolTipText     =   "   F2   "
      Top             =   5280
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
      TabIndex        =   61
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
      TabIndex        =   59
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
      TabIndex        =   55
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
      TabIndex        =   28
      Text            =   "1"
      Top             =   8970
      Width           =   855
   End
   Begin VB.Frame Frame7 
      Height          =   1695
      Left            =   120
      TabIndex        =   50
      Top             =   9000
      Width           =   8535
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   375
         Left            =   6480
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
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
         TabIndex        =   24
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
         TabIndex        =   25
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
         TabIndex        =   27
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   480
         Width           =   4335
      End
      Begin MSMask.MaskEdBox mk3 
         Height          =   375
         Left            =   6480
         TabIndex        =   26
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtHora 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1080
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
      TabIndex        =   43
      Top             =   120
      Width           =   10815
      Begin VB.OptionButton Option1 
         Caption         =   "CNPJ"
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
         Height          =   255
         Index           =   1
         Left            =   9600
         TabIndex        =   32
         Top             =   1190
         Width           =   908
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CPF"
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
         Height          =   255
         Index           =   0
         Left            =   8400
         TabIndex        =   31
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   30
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
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1440
         Width           =   5295
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
         Index           =   0
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   8535
      End
      Begin MSMask.MaskEdBox txtTelefone 
         Height          =   420
         Left            =   8760
         TabIndex        =   1
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   741
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##)####-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   420
         Left            =   8400
         TabIndex        =   33
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   741
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Left            =   5520
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   45
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
         TabIndex        =   44
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
      TabIndex        =   35
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
      TabIndex        =   41
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
      TabIndex        =   20
      ToolTipText     =   "   Esc   "
      Top             =   6480
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
      TabIndex        =   19
      ToolTipText     =   "   F2   "
      Top             =   5850
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   38
      Top             =   2280
      Width           =   12735
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   480
         Width           =   7095
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
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   480
         Width           =   5295
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
         Left            =   5520
         TabIndex        =   37
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
         TabIndex        =   36
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
      TabIndex        =   17
      ToolTipText     =   "   Alt + Enter   "
      Top             =   4680
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid gd 
      Height          =   1800
      Left            =   135
      TabIndex        =   34
      Top             =   7080
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   3175
      _Version        =   393216
      Rows            =   6
      Cols            =   12
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
      TabIndex        =   56
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
      TabIndex        =   60
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
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   10320
      Width           =   1530
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
      TabIndex        =   64
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
      TabIndex        =   63
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
      TabIndex        =   54
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
      TabIndex        =   53
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
      TabIndex        =   49
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
      TabIndex        =   48
      Top             =   840
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
      TabIndex        =   42
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
Public contadorDePlano As Integer
Public modoEdicao As Boolean
Public nOS As Long
Public nosFinal As Long
Dim codigoCliente As Long
Public cancelPrevisaoDeEntrega As Boolean
Public varDataEntrega As String
Public varHoraEntrega As String

Public Sub salvarRegistro(Optional gridrow As Integer)
    
    abreConexao
    
        If modoEdicao Then
            rs.Open "Select * from os WHERE idos=" & nOS, db, adOpenStatic, adLockOptimistic
        Else
            rs.Open "Select * from os", db, adOpenStatic, adLockOptimistic
            rs.AddNew
            rs!idUsuario = varCodigoUsuario
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
        
        If txtCodigoCliente.Text = 0 Then
            rs!nomeCliente = txtCliente(0).Text
            rs!telefonecliente = txtTelefone.Text 'Format$(txtTelefone.Text, "(##)####-####")
        Else
            rs!idcliente = txtCodigoCliente.Text
        End If
        Dim valor As Currency
        valor = txtTotalGeral
        
        rs!valorOs = valor
        
        If modoEdicao Then rs!alteradopor = varNomeUsuario & " - " & Format$(Date, "dd/mm/yy") & " - " & Format$(Time, "hh:mm")

        rs.Update
        
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
            rs!midia = gd.TextMatrix(w, 3)
            rs!cores = gd.TextMatrix(w, 4)
            rs!Laminacao = gd.TextMatrix(w, 5)
            rs!capa = gd.TextMatrix(w, 6)
            rs!wireo = gd.TextMatrix(w, 7)

            If gd.TextMatrix(w, 8) = "" Then
                rs!corte = 0
            Else
                rs!corte = gd.TextMatrix(w, 8)
            End If
            
            If gd.TextMatrix(w, 9) = "" Then
                rs!meiocorte = 0
            Else
                rs!meiocorte = gd.TextMatrix(w, 9)
            End If
            
            If gd.TextMatrix(w, 10) = "" Then
                rs!transfer = 0
            Else
                rs!transfer = gd.TextMatrix(w, 10)
            End If
            
            rs!valor = gd.TextMatrix(w, 11)
            
            rs.Update
            If modoEdicao Then
                rs.MoveNext
            End If
        Next

        rs.Close

With frmPesquisaOs
'-----------------------------------------------------

        Dim linhagrid As Long
        Dim linhagridfinal As Long
        Dim rsCliente As Recordset
        Set rsCliente = New Recordset
        
        linhagrid = .grdUsuario.Row
        
        '---------------------------------------------------------------------
        '
        '       ERRO NESSA LINHA DE BAIXO
        '
        '
        '
        '
        '
        '
        '
        '
        '             ERRO NESSA LINHA DE BAIXO
        '
        '
        '
        '
        '
        '
        '
        '
        '
        
        '----------------------------------------------------------------------
        
        linhagridfinal = frmPesquisaOs.grdUsuario.TextMatrix(frmPesquisaOs.grdUsuario.Rows - 10, 0)
        
        If modoEdicao Then
            '.grdUsuario.TextMatrix(linhagrid, 0) = rs!idos
            .grdUsuario.TextMatrix(linhagrid, 1) = txtCliente(0).Text
        Else
            rs.Open "SELECT * FROM os WHERE idos>" & linhagridfinal & " ORDER BY idos", db, adOpenStatic, adLockOptimistic '& frmPesquisaOs.nosFinal & " ORDER BY idos", db, adOpenStatic, adLockOptimistic
            rs.MoveFirst
            Dim h As Integer
            
            For h = 1 To rs.RecordCount
                .grdUsuario.TextMatrix(.grdUsuario.Rows - 10 + h, 0) = rs!idos
            
                If IsNull(rs!nomeCliente) Then
                    rsCliente.Open "SELECT nome FROM cliente where idcliente=" & rs!idcliente, db, adOpenStatic, adLockOptimistic
                    .grdUsuario.TextMatrix(.grdUsuario.Rows - 10 + h, 1) = rsCliente!nome
                    rsCliente.Close
                Else
                    .grdUsuario.TextMatrix(.grdUsuario.Rows - 10 + h, 1) = rs!nomeCliente
                End If
                
                .grdUsuario.TextMatrix(.grdUsuario.Rows - 10 + h, 2) = Format(rs!Data, "dd/mm/yy")
                .grdUsuario.TextMatrix(.grdUsuario.Rows - 10 + h, 3) = Format(rs!hora, "hh:mm")
                .grdUsuario.TextMatrix(.grdUsuario.Rows - 10 + h, 4) = varNomeUsuario
                rs.MoveNext
            Next
            
            .grdUsuario.Rows = .grdUsuario.Rows + rs.RecordCount
            .grdUsuario.TopRow = .grdUsuario.Rows - 30
            .grdUsuario.Row = .grdUsuario.Rows - 10
            .grdUsuario.ColSel = .grdUsuario.Cols - 1
            rs.Close
        End If

'----------------------------------------------
End With
        
        db.Close

End Sub
Public Sub btnAdd_Click(Index As Integer)

cancelPrevisaoDeEntrega = False

Select Case Index
    
    Case 0 ' ----------------------------------------------------------- botão add
        preenchePlano
        btnAdd(2).Enabled = True
        
    Case 1, 2  '-------------------------------------------------------- botão salvar
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
        ElseIf Len(MaskEdBox1.Text) < 11 And Len(MaskEdBox1.Text) > 0 Then
            If txtCodigoCliente.Text = 0 Then
                MsgBox "Preencha o CPF corretamente.", vbInformation
                MaskEdBox1.SetFocus
                Exit Sub
            End If
        End If
        
        'If Index = 2 Then
            'frmPrevisaodeentrega.Show 1
            'If cancelPrevisaoDeEntrega Then
                'cancelPrevisaoDeEntrega = False
                'Exit Sub
            'End If
        'End If
        
        salvarRegistro
        
        frmOrcamento.Caption = "Ordem de Serviço: " & nOS

        If Index = 1 Then
            bAlterado = False
            Unload frmOrcamento
        ElseIf Index = 2 Then
            modoEdicao = True
            Load frm
            frm.posPagina
            frm.Show 1
        End If
        
    Case 3
        Unload Me
End Select

End Sub

Private Sub Check1_Click()
    bAlterado = True
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Check2_Click()
bAlterado = True
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Check4_Click()
    
    txtOrcamento(1).Text = ""
    contadorDePlano = 0
    If Check4.Value = 1 Then
        Frame5.Enabled = False
        Frame10.Left = 7200
        Frame11.Left = 7200
        Frame10.Top = 115
        Frame11.Top = 445
        
        label(5).Caption = "Páginas:"
        label(6).Caption = "Formato:"
        label(7).Left = 14000
        label(8).Left = 14000
        label(9).Left = 14000
        Combo4.Left = 14000
        Combo5.Left = 14000
        Combo3.Left = 14000
        
        
        Combo2.Clear
        Combo2.AddItem "Photobook 15x21"
        Combo2.AddItem "Photobook 20x20"
        Combo2.AddItem "Photobook 20x29"
        
    Else

        Frame5.Enabled = True
        Frame10.Left = 14000
        Frame11.Left = 14000
        
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
        

        
    End If
    gd.Clear
    txtOrcamento(1).SetFocus
    somaTotal
End Sub

Private Sub Combo1_Click(Index As Integer)
    If Combo1(Index).Text <> "" Then
        bAlterado = True
    End If
    If Combo1(0).ListIndex > 2 Or Combo1(0).ListIndex = 0 Then
        Check1.Value = 0
        Check1.Enabled = False
    Else
        Check1.Enabled = True
    End If
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
    Combo1(Index).ListIndex = 0
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 1 Then
    If KeyCode = 46 Then
        Combo1(1).ListIndex = -1
    End If
End If
End Sub

Private Sub Combo2_Click()
    
    bAlterado = True
        
    If Check4.Value = 1 Then Exit Sub
    
    Combo3.Enabled = True
    Combo4.Enabled = True
    Combo5.Enabled = True
    Combo3.Clear
    Combo3.AddItem ("Couchê Brilho")
    Combo3.AddItem ("Couchê Fosco")
    Combo3.AddItem ("Duo Design")
    Combo3.AddItem ("Apergaminhado")
    Combo3.AddItem ("Reciclato")
    
    If Combo2.Text = "Impressos" Then
        Combo3.AddItem ("Transfer")
        Combo3.AddItem ("Adesivo")
        Combo3.AddItem ("BOPP")
    End If
    
    label(10).Visible = False: txtOrcamento(10).Visible = False: txtOrcamento(11).Visible = False: Label1.Visible = False
    Combo4.Visible = True
    Combo5.Visible = True
    label(7).Visible = True
    label(8).Visible = True
    label(9).Visible = True
    txtOrcamento(10).Text = ""
    txtOrcamento(11).Text = ""
    txtOrcamento(10).TabStop = False
    txtOrcamento(11).TabStop = False
    Frame5.Enabled = True

    Select Case Combo2.Text
        Case "Banner"
            Combo4.Visible = False
            Combo5.Visible = False
            label(8).Visible = False
            label(9).Visible = False
            label(10).Left = 9940: txtOrcamento(10).Left = 10020: txtOrcamento(11).Left = 11380: Label1.Left = 11150
            label(10).Visible = True: txtOrcamento(10).Visible = True: txtOrcamento(11).Visible = True: Label1.Visible = True
            txtOrcamento(10).TabStop = True
            txtOrcamento(11).TabStop = True
            Combo3.Clear
            Combo3.AddItem ("Lona")
            Combo3.AddItem ("Adesivo L.")
            Combo3.AddItem ("Adesivo T.")
            Frame5.Enabled = False
            limpa (0)

        Case "Fotos"
            
            label(10).Visible = False: txtOrcamento(10).Visible = False: txtOrcamento(11).Visible = False: Label1.Visible = False
            Combo4.Visible = True
            Combo5.Visible = True
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
            limpa (0)

    End Select
 
    Combo3.ListIndex = 0
    
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Combo3_Click()
    bAlterado = True
    Combo4.Clear
    Select Case Combo3.ListIndex
        
        Case 0, 1 'Couche
            Combo4.AddItem ("90g")
            Combo4.AddItem ("115g")
            Combo4.AddItem ("150g")
            Combo4.AddItem ("170g")
            Combo4.AddItem ("250g")
            Combo4.AddItem ("300g")
            
        Case 2 'Duo Design
            Combo4.AddItem ("250g")
            
        Case 3 'Ap
            Combo4.AddItem ("90g")
            Combo4.AddItem ("180g")
            Combo4.AddItem ("240g")
            
        Case 4 'Reciclato
            Combo4.AddItem ("120g")
            Combo4.AddItem ("180g")
            Combo4.AddItem ("240g")
            
        Case 5, 6 'Transfer e Adesivo
            Combo4.AddItem ("90g")
            
        Case 7 'BOPP
            Combo4.AddItem ("120g")
            
    End Select
    
    If Combo3.ListIndex <> -1 Then
        Combo4.ListIndex = 0
    End If

    If Combo2.Text = "Fotos" Then
        Combo4.ListIndex = 4
    End If

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Combo4_Click()
bAlterado = True
If Combo5.ListIndex = -1 Then
    Combo5.ListIndex = 0
End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Combo5_Click()
bAlterado = True
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Command1_Click()

    txtDesconto.Text = "-" & Format$(txtValorDesconto.Text, "#,##0.00")
    txtAcrescimo.Text = Format$(txtValorAcrescimo.Text, "#,##0.00")
    somaTotal
    
End Sub

Private Sub Form_Activate()
    bAlterado = False
    
' Se OS estiver baixada: trava tudo
If Check3.Value And varTipoUsuario <> "Administrador" Then
        txtCliente(0).Enabled = False
        txtCliente(1).Enabled = False
        txtCliente(2).Enabled = False
        txtTelefone.Enabled = False
        MaskEdBox1.Enabled = False
        Option1(0).Enabled = False
        Option1(1).Enabled = False
        txtOrcamento(1).Enabled = False
        txtOrcamento(2).Enabled = False
        txtOrcamento(3).Enabled = False
        txtOrcamento(4).Enabled = False
        txtOrcamento(5).Enabled = False
        txtOrcamento(6).Enabled = False
        txtOrcamento(10).Enabled = False
        txtOrcamento(11).Enabled = False
        Combo1(0).Enabled = False
        Combo1(1).Enabled = False
        Combo1(2).Enabled = False
        Combo2.Enabled = False
        Combo3.Enabled = False
        Combo4.Enabled = False
        Combo5.Enabled = False
        Check2.Enabled = False
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

txtData = Format(Date, "dd/mm/yy")
txtHora = Format(Time, "hh:mm")

abreConexao

txtUsuario.Text = varNomeUsuario

Me.Width = 13095
numeraPlano
contadorDePlano = 0
With gd
    
    alinhaColunas
    'defineLargura
    
    .ColWidth(0) = 500
    .ColWidth(1) = 600
    .ColWidth(2) = 3370
    .ColWidth(3) = 1300
    .ColWidth(4) = 400
    .ColWidth(5) = 1300
    .ColWidth(6) = 1000
    .ColWidth(7) = 1300
    .ColWidth(8) = 600
    .ColWidth(9) = 600
    .ColWidth(10) = 600
    
'------------Título
    .TextMatrix(0, 1) = "Qtd"
    .TextMatrix(0, 2) = "Descrição"
    .TextMatrix(0, 3) = "Mídia"
    .TextMatrix(0, 4) = "C"
    .TextMatrix(0, 5) = "Laminação"
    .TextMatrix(0, 6) = "Capa"
    .TextMatrix(0, 7) = "Wire-o"
    .TextMatrix(0, 8) = "Co"
    .TextMatrix(0, 9) = "MC"
    .TextMatrix(0, 10) = "T"
    .TextMatrix(0, 11) = "Sub total"

End With

'txtCliente(1).Text = frmPesquisaOs.nosFinal

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim op As Integer
    
    If bAlterado = True Then
        op = MsgBox("A OS foi alterada. Deseja salvar agora?", vbQuestion + vbYesNoCancel, "Salvar")
        Select Case op
            Case 2
                Cancel = 1
            Case 6
                btnAdd_Click (1)
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    modoEdicao = False
End Sub

Private Sub gd_DblClick()

With gd

If .TextMatrix(.Row, 1) = "" Then Exit Sub

btnAdd(0).Caption = "Alterar"
btnAdd(1).Enabled = False
btnAdd(2).Enabled = False
btnAdd(3).Enabled = False
gd.Enabled = False

txtOrcamento(1).Text = .TextMatrix(.Row, 1)

'--------------Banner
If preencheCampo(.TextMatrix(.Row, 2), 1) = "Banner" Then
    Combo2.Text = preencheCampo(.TextMatrix(.Row, 2), 1)
    Combo3.Text = .TextMatrix(.Row, 3)
    txtOrcamento(10).Text = preencheCampo(.TextMatrix(.Row, 2), 2)
    txtOrcamento(11).Text = preencheCampo(.TextMatrix(.Row, 2), 3)
    Exit Sub
End If

    Combo2.Text = .TextMatrix(.Row, 2)
    Combo3.Text = inverteMidia(preencheCampo(.TextMatrix(.Row, 3), 1))
    Combo4.Text = preencheCampo(.TextMatrix(.Row, 3), 2)
    Combo5.Text = .TextMatrix(.Row, 4)

'--Laminação---------->
    txtOrcamento(2).Text = preencheCampo(.TextMatrix(.Row, 5), 1)
    If inverteLami(preencheCampo(.TextMatrix(.Row, 5), 2)) <> "" Then
        Combo1(0).Text = inverteLami(preencheCampo(.TextMatrix(.Row, 5), 2))
    End If
    If preencheCampo(.TextMatrix(.Row, 5), 3) = "FV" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If

'--Capa--------------->
    If inverteCapa(.TextMatrix(.Row, 6)) <> "" Then
        Combo1(1).Text = inverteCapa(.TextMatrix(.Row, 6))
    End If

'--Wireo--------------->
    txtOrcamento(3).Text = preencheCampo(.TextMatrix(.Row, 7), 1)
    If preencheCampo(.TextMatrix(.Row, 7), 2) <> "" Then
        Combo1(2).Text = preencheCampo(.TextMatrix(.Row, 7), 2)
    End If
    If preencheCampo(.TextMatrix(.Row, 7), 3) = "EM" Then
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If

If .TextMatrix(.Row, 8) <> "" Then
    txtOrcamento(4).Text = .TextMatrix(.Row, 8)
End If
If .TextMatrix(.Row, 9) <> "" Then
    txtOrcamento(5).Text = .TextMatrix(.Row, 9)
End If
If .TextMatrix(.Row, 10) <> "" Then
    txtOrcamento(6).Text = .TextMatrix(.Row, 10)
End If

End With

End Sub

Private Sub gd_GotFocus()
   
    gd.HighLight = flexHighlightWithFocus
        
End Sub

Private Sub gd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    gd.RemoveItem (gd.Row)
    gd.Rows = 6
    numeraPlano
    If contadorDePlano > 0 Then
        contadorDePlano = contadorDePlano - 1
    End If
    somaTotal
    If contadorDePlano = 0 Then btnAdd(2).Enabled = False
End If
End Sub

Private Sub gd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Y > 20 And Y < gd.RowHeight(0) Then
    If X > 10 And X < 490 Then
        gd.ToolTipText = "Plano"
    ElseIf X > 510 And X < 1090 Then
        gd.ToolTipText = "Quantidade"
    ElseIf X > 1110 And X < 4090 Then
        gd.ToolTipText = "Descrição"
    ElseIf X > 500 And X < 600 Then
        gd.ToolTipText = "Tipo de mídia"
    ElseIf X > 510 And X < 1090 Then
        gd.ToolTipText = "Quantidade"
    ElseIf X > 510 And X < 1090 Then
        gd.ToolTipText = "Quantidade"
    ElseIf X > 510 And X < 1090 Then
        gd.ToolTipText = "Quantidade"
    Else
        gd.ToolTipText = ""
    End If
Else
    gd.ToolTipText = ""
End If

'240
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
    .ColAlignment(11) = 7
    
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

Private Sub MaskEdBox1_GotFocus()
    MaskEdBox1.SelStart = 0
    MaskEdBox1.SelLength = Len(MaskEdBox1.Text)
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub MaskEdBox2_Change()
bAlterado = True
End Sub

Private Sub MaskEdBox2_GotFocus()
    MaskEdBox2.SelStart = 0
    MaskEdBox2.SelLength = Len(MaskEdBox2.Text)
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If

If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) And Not (KeyAscii = 44) Then
    KeyAscii = 0
End If
End Sub

Private Sub MaskEdBox2_LostFocus()
    txtValorAcrescimo.Top = MaskEdBox2.Top
    txtValorAcrescimo.Left = MaskEdBox2.Left
    
    If MaskEdBox2.Text = "" Then
        txtValorAcrescimo.Text = "0,00"
    Else
        txtValorAcrescimo.Text = MaskEdBox2.Text
    End If
    
    txtValorAcrescimo.Visible = True
    MaskEdBox2.Visible = False
End Sub

Private Sub mk3_Change()
bAlterado = True
End Sub

Private Sub mk3_GotFocus()
    mk3.SelStart = 0
    mk3.SelLength = Len(mk3.Text)
End Sub

Private Sub mk3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If

If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) And Not (KeyAscii = 44) Then
    KeyAscii = 0
End If
End Sub

Private Sub mk3_LostFocus()
    txtValorDesconto.Top = mk3.Top
    txtValorDesconto.Left = mk3.Left
    txtValorDesconto.Text = mk3.Text
    
    txtValorDesconto.Visible = True
    mk3.Visible = False
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            MaskEdBox1.Mask = "###.###.###-##"
        Case 1
           MaskEdBox1.Mask = "##.###.###/####-##"
    End Select
End Sub





Private Sub txtCliente_Change(Index As Integer)
    bAlterado = True
End Sub

Private Sub txtCliente_GotFocus(Index As Integer)
    txtCliente(Index).SelStart = 0
    txtCliente(Index).SelLength = Len(txtCliente(Index).Text)
End Sub

Private Sub txtCliente_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
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

Private Sub txtDescricaoAcrescimo_Change()
bAlterado = True
End Sub

Private Sub txtDescricaoAcrescimo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub txtDescricaoDesconto_Change()
bAlterado = True
End Sub

Private Sub txtDescricaoDesconto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
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
            Combo1(0).ListIndex = 0
            Check1.Value = 0
        End If
    Case 3
        If txtOrcamento(3).Text = "0" Or txtOrcamento(3).Text = "" Then
            Combo1(2).ListIndex = 0
            Check2.Value = 0
        End If
End Select

End Sub

Private Sub txtOrcamento_GotFocus(Index As Integer)
    txtOrcamento(Index).SelStart = 0
    txtOrcamento(Index).SelLength = Len(txtOrcamento(Index).Text)
End Sub

Private Sub txtOrcamento_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If

If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
    KeyAscii = 0
End If

End Sub

Private Sub txtOrcamento_LostFocus(Index As Integer)
    If txtOrcamento(Index).Text = "0" Then txtOrcamento(Index) = ""
End Sub

Private Sub txtTelefone_GotFocus()
    txtTelefone.SelStart = 0
    txtTelefone.SelLength = Len(txtTelefone.Text) + 13 - Len(txtTelefone.Text)
End Sub

Private Sub txtTelefone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub preenchePlano()
Dim r As Integer
Dim subtotalbanner As Double
Dim varSubTotal As Double
Dim varQuantidade As Double

If btnAdd(0).Caption = "&Add" Then
    If contadorDePlano < 5 Then
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


If Check4.Value = 0 Then
        ' ----------------------------------------------------- Fotos
        
            If Combo2.Text = "Fotos" Then
                
                If txtOrcamento(1).Text < 8 Then
                    MsgBox "A quantidade mínima para fotos é de 8 unidades", vbInformation
                    txtOrcamento(1).Text = 8
                    contadorDePlano = contadorDePlano - 1
                    Exit Sub
                End If
                
            End If
        
        With gd
        
            If verificaCampoNulo Then Exit Sub
            varQuantidade = txtOrcamento(1).Text
        
            .TextMatrix(r, 1) = varQuantidade
            .TextMatrix(r, 2) = Combo2.Text
            
            If Check4.Value = 0 Then ' NÃO É BOOK
                .TextMatrix(r, 3) = converteMidia(Combo3.Text) & ":" & Combo4.Text
            End If
            
        '-----------------------------------------   Banner
            
            If Combo2.Text = "Banner" Then
                .TextMatrix(r, 2) = Combo2.Text & ": " & txtOrcamento(10).Text & "x" & txtOrcamento(11).Text
                .TextMatrix(r, 3) = Combo3.Text
                .TextMatrix(r, 4) = 6
                
                subtotalbanner = varQuantidade * txtOrcamento(10).Text * txtOrcamento(11).Text * 35 / 10000
                If subtotalbanner < 10 Then
                    varSubTotal = 10
                Else
                    varSubTotal = subtotalbanner
                End If
                
                .TextMatrix(r, 11) = Format(varSubTotal, "#,##0.00")
                limpa (1)
                txtOrcamento(1).SetFocus
                somaTotal
                Exit Sub
            End If
            
            If Combo5.ListIndex = 1 Then
                varSubTotal = varQuantidade
            End If
            
            If Combo2.Text = "Impressos" Then
                varSubTotal = varSubTotal + varQuantidade * 2
            ElseIf Combo2.Text = "Impressos F/V" Then
                varSubTotal = varSubTotal * 2 + varQuantidade * 4
                Dim fv As Double
                fv = 2
            End If
            
            .TextMatrix(r, 4) = Combo5.Text ' CORES
            
        ' ----------------------------------------------------- Laminação
            
            If txtOrcamento(2).Text = "" Or Combo1(0).Text = "" Then
                .TextMatrix(r, 5) = ""
            Else
                Dim valorLami As Currency
                
                Select Case Combo1(0).ListIndex
                    Case 1, 2, 3, 7
                        valorLami = 1
                    Case 4
                        valorLami = 1.5
                    Case 5
                        valorLami = 2
                    Case 6
                        valorLami = 3.5
                End Select
                    
                If Check1.Value = 1 Then
                    .TextMatrix(r, 5) = txtOrcamento(2).Text + ":" & converteLami(Combo1(0).Text) + ":" + "FV"
                    varSubTotal = varSubTotal + (txtOrcamento(2).Text * valorLami * 2)
                Else
                    .TextMatrix(r, 5) = txtOrcamento(2).Text + ":" & converteLami(Combo1(0).Text)
                    varSubTotal = varSubTotal + (txtOrcamento(2).Text * valorLami)
                End If
            End If
            
        ' -----------------------------------------------------Capa
                .TextMatrix(r, 6) = converteCapa(Combo1(1).Text)
                Dim valorCapa As Currency
                
                Select Case Combo1(1).ListIndex
                    Case 1
                        valorCapa = 4
                    Case 2
                        valorCapa = 6
                    Case 3
                        valorCapa = 12
                    Case 4
                        valorCapa = 1
                    Case 5
                        valorCapa = 2
                End Select
                
                varSubTotal = varSubTotal + valorCapa
        
        ' -----------------------------------------------------Wireo
        
            If txtOrcamento(3).Text = "" Or Combo1(2).Text = "" Then
                .TextMatrix(r, 7) = ""
            Else
                Dim valorWireo As Currency
                
                Select Case Combo1(2).ListIndex
                    Case 1, 2, 3
                        valorWireo = txtOrcamento(3).Text * 1.5
                    Case 4, 5, 6
                        valorWireo = txtOrcamento(3).Text * 2
                    Case 7
                        valorWireo = txtOrcamento(3).Text * 2.5
                    Case 8, 9
                        valorWireo = txtOrcamento(3).Text * 3
                End Select
        
                If Check2.Value = 1 Then
                    .TextMatrix(r, 7) = txtOrcamento(3).Text + ":" + Combo1(2).Text + ":" + "EM"
                    varSubTotal = varSubTotal + valorWireo + (txtOrcamento(3).Text * 3.5)
                Else
                    .TextMatrix(r, 7) = txtOrcamento(3).Text + ":" + Combo1(2).Text
                    varSubTotal = varSubTotal + valorWireo
                End If
            End If
        
            ' -----------------------------------------------------Corte
            If txtOrcamento(4).Text = "" Then
                .TextMatrix(r, 8) = ""
            Else
                .TextMatrix(r, 8) = txtOrcamento(4).Text
                varSubTotal = varSubTotal + (txtOrcamento(4).Text * 0.5)
            End If
        ' -----------------------------------------------------Meio Corte
            If txtOrcamento(5).Text = "" Then
                .TextMatrix(r, 9) = ""
            Else
                .TextMatrix(r, 9) = txtOrcamento(5).Text
                varSubTotal = varSubTotal + (txtOrcamento(5).Text * 0.1)
            End If
        ' -----------------------------------------------------Transfer
            If txtOrcamento(6).Text = "" Then
                .TextMatrix(r, 10) = ""
            Else
                .TextMatrix(r, 10) = txtOrcamento(6).Text
                varSubTotal = varSubTotal + (txtOrcamento(6).Text)
            End If
        ' -----------------------------------------------------Sub total
            
            Dim valorMidia As Double
                
                Select Case Combo3.ListIndex
                
                    Case 0, 1 'Couchê liso ou fosco
                        If Combo4.ListIndex > 3 Then
                            valorMidia = 0.5
                        End If
                    Case 2 'Duo Design
                        valorMidia = 0.5
                        
                    Case 3 'AP
                        If Combo4.ListIndex > 0 Then valorMidia = 0.5
                    
                    Case 4 'Reciclato
                        If Combo4.ListIndex > 0 Then valorMidia = 0.5
                    
                    Case 5  'Transfer
                        valorMidia = 2
                        
                    Case 4, 6 'BOPP, Adesivo
                        valorMidia = 1
                    
                End Select
                
            If fv = 2 Then valorMidia = valorMidia * 2 'Se for frente e verso
            
            varSubTotal = varSubTotal + (varQuantidade * valorMidia)
            
            If Combo2.Text = "Fotos" Then varSubTotal = varQuantidade * 0.39
            
            .TextMatrix(r, 11) = Format(varSubTotal, "#,##0.00")
            
        End With
            
            limpa (1)
            txtOrcamento(1).SetFocus
            'somaTotal
Else
    
    '------------------- PHOTOBOOK -------------------------------Capa
    Dim vBookCapa As Double
    Dim vBookMiolo As Double
    
    Select Case Combo2.Text
        Case "Photobook 15x21"
            vBookCapa = 9
            vBookMiolo = 1.4
            
        Case "Photobook 20x20"
            vBookCapa = 12
            vBookMiolo = 2.4
        Case "Photobook 20x29"
            vBookCapa = 12
            vBookMiolo = 2.4
    End Select
    
    gd.TextMatrix(1, 1) = 1
    gd.TextMatrix(1, 2) = "Capa " & Combo2.Text
    gd.TextMatrix(1, 3) = "CL:170g"
    gd.TextMatrix(1, 4) = "4"
    
    If Option3(0).Value = True Then
        gd.TextMatrix(1, 5) = "L:BOPPB"
    Else
        gd.TextMatrix(1, 5) = "L:BOPPF"
    End If
    
    gd.TextMatrix(1, 11) = Format(vBookCapa, "#,##0.00")
    '-------------------------------------------------------------Miolo
    
    gd.TextMatrix(2, 1) = Int(txtOrcamento(1).Text / 2 - 1)
    gd.TextMatrix(2, 2) = "Miolo " & Combo2.Text
    gd.TextMatrix(2, 3) = "DD:250g"
    gd.TextMatrix(2, 4) = "4"
    
    If Option4(0).Value = True Then
        gd.TextMatrix(2, 5) = "L:BOPPB"
    Else
        gd.TextMatrix(2, 5) = "L:BOPPF"
    End If
    
    gd.TextMatrix(2, 11) = Format((txtOrcamento(1).Text - 2) * vBookMiolo, "#,##0.00")
    
    contadorDePlano = 2
    
End If

somaTotal

End Sub
Private Function converteMidia(midia As String) As String

Select Case midia

    Case "Couchê Brilho"
        converteMidia = "CL"
    Case "Couchê Fosco"
        converteMidia = "CF"
    Case "Duo Design"
        converteMidia = "DD"
    Case "Apergaminhado"
        converteMidia = "AP"
    Case "Reciclato"
        converteMidia = "RC"
    Case "Transfer"
        converteMidia = "TR"
    Case "Adesivo"
        converteMidia = "AD"
    Case "BOPP"
        converteMidia = "BOPP"
End Select

End Function
Private Function converteLami(lami As String) As String

Select Case lami

    Case "BOPP Brilho"
        converteLami = "BOPPB"
    Case "BOPP Fosco"
        converteLami = "BOPPF"
    Case "Polaseal A6"
        converteLami = "PA6"
    Case "Polaseal A5"
        converteLami = "PA5"
    Case "Polaseal A4"
        converteLami = "PA4"
    Case "Polaseal A3"
        converteLami = "PA3"

End Select

End Function
Private Sub limpa(tudo As Integer)

Dim w As Integer

For w = 2 To 6
    txtOrcamento(w) = ""
Next

Combo1(0).ListIndex = -1
Combo1(1).ListIndex = -1
Combo1(2).ListIndex = -1
Check1.Value = 0
Check2.Value = 0

If tudo = 1 Then
    txtOrcamento(1).Text = ""
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
    Combo4.ListIndex = -1
    Combo5.ListIndex = -1
    txtOrcamento(10).Text = ""
    txtOrcamento(11).Text = ""
End If

End Sub
Private Sub somaTotal()
    Dim soma As Currency
    Dim w As Integer
    soma = 0

    For w = 1 To gd.Rows - 1
        If gd.TextMatrix(w, 11) <> "" Then
            soma = soma + CCur(gd.TextMatrix(w, 11))
        End If
    Next
    txtTotal.Text = Format$(soma, "#,##0.00")
    
    txtTotalGeral.Text = Format$((CDbl(txtTotal.Text) * txtExemplar.Text) + CDbl(txtDesconto.Text) + CDbl(txtAcrescimo.Text), "#,##0.00")
End Sub
Private Function verificaCampoNulo() As Boolean

Dim q As Boolean
Dim controle As TextBox

If txtOrcamento(1).Text = "" Or txtOrcamento(1).Text = "0" Then 'Verifica campo Quantidade
    Set controle = txtOrcamento(1)
    q = True
ElseIf Combo2.Text = "" Then 'Verifica Campo Descrição
    alt = False
    MsgBox "Digite um valor válido para este campo.", vbInformation
    Combo2.SetFocus
    contadorDePlano = contadorDePlano - 1
    verificaCampoNulo = True
ElseIf Combo2.Text = "Banner" Then ' Verifica Campo Dimensões do banner
    If txtOrcamento(10).Text = "" Then
        Set controle = txtOrcamento(10)
        q = True
    ElseIf txtOrcamento(11).Text = "" Then
        Set controle = txtOrcamento(11)
        q = True
    End If
End If

If q Then
    alt = False
    MsgBox "Digite um valor válido para este campo.", vbInformation
    controle.SetFocus
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
    Case "Capa PVC A5"
        converteCapa = "PVCA5"
    Case "Capa PVC A4"
        converteCapa = "PVCA4"
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

Private Function inverteMidia(midia As String) As String

Select Case midia

    Case "CL"
        inverteMidia = "Couchê Brilho"
    Case "CF"
        inverteMidia = "Couchê Fosco"
    Case "DD"
        inverteMidia = "Duo Design"
    Case "AP"
        inverteMidia = "Apergaminhado"
    Case "RC"
        inverteMidia = "Reciclato"
    Case "TR"
        inverteMidia = "Transfer"
    Case "AD"
        inverteMidia = "Adesivo"
    Case "BOPP"
        inverteMidia = "BOPP"
End Select

End Function

Private Function inverteLami(lami As String) As String

Select Case lami

    Case "BOPPB"
        inverteLami = "BOPP Brilho"
    Case "BOPPF"
        inverteLami = "BOPP Fosco"
    Case "PA6"
        inverteLami = "Polaseal A6"
    Case "PA5"
        inverteLami = "Polaseal A5"
    Case "PA4"
        inverteLami = "Polaseal A4"
    Case "PA3"
        inverteLami = "Polaseal A3"
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

Private Sub txtValorAcrescimo_Change()
    txtValorAcrescimo.Text = Format$(txtValorAcrescimo.Text, "#,##0.00")
End Sub

Private Sub txtValorAcrescimo_GotFocus()
    MaskEdBox2.Top = txtValorAcrescimo.Top
    MaskEdBox2.Left = txtValorAcrescimo.Left
    MaskEdBox2.Text = txtValorAcrescimo.Text
    txtValorAcrescimo.Visible = False
    MaskEdBox2.Visible = True
    MaskEdBox2.SetFocus
End Sub

Private Sub txtValorDesconto_Change()
    txtValorDesconto.Text = Format$(txtValorDesconto.Text, "#,##0.00")
End Sub

Private Sub txtValorDesconto_GotFocus()
    mk3.Top = txtValorDesconto.Top
    mk3.Left = txtValorDesconto.Left
    mk3.Text = txtValorDesconto.Text
    txtValorDesconto.Visible = False
    mk3.Visible = True
    mk3.SetFocus
End Sub

Private Function clienteCadastrado(cpf As String) As Boolean

    Dim sCriterio As String
    
    If Option1(0).Value = True Then
        sCriterio = Format$(cpf, "@@@.@@@.@@@-@@")
    Else
        sCriterio = Format$(cpf, "@@.@@@.@@@/@@@@-@@")
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
