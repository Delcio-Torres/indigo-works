VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Orçamento"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   11640
      Picture         =   "Form1.frx":000C
      ScaleHeight     =   585
      ScaleWidth      =   1770
      TabIndex        =   79
      Top             =   6600
      Width           =   1770
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   360
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "F5 - Sair"
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
      Left            =   8880
      TabIndex        =   78
      Top             =   6960
      Width           =   2880
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "F4 - Imprimir OS"
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
      Left            =   5960
      TabIndex        =   77
      Top             =   6960
      Width           =   2880
   End
   Begin VB.ComboBox text7 
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
      ItemData        =   "Form1.frx":0B2D
      Left            =   1560
      List            =   "Form1.frx":0B43
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   4455
   End
   Begin VB.TextBox txtNumero 
      Height          =   375
      Left            =   8280
      TabIndex        =   74
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.StatusBar stbOrcamento 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   73
      Top             =   7455
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   873
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Vendedor:"
            TextSave        =   "Vendedor:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4939
            MinWidth        =   4939
            Text            =   "Incluído: 17/03/2009 - 20:00h"
            TextSave        =   "Incluído: 17/03/2009 - 20:00h"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Alterado:"
            TextSave        =   "Alterado:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "F3 - Salvar e Sair"
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
      Left            =   3040
      TabIndex        =   50
      Top             =   6960
      Width           =   2880
   End
   Begin VB.Frame Frame10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8400
      TabIndex        =   64
      Top             =   120
      Width           =   3375
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Alterado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   690
         TabIndex        =   72
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   68
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   67
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hora entrada:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   65
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "F2 - Localizar"
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
      Left            =   120
      TabIndex        =   51
      Top             =   6960
      Width           =   2880
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   420
      Left            =   6240
      TabIndex        =   1
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   741
      _Version        =   393216
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
      PromptChar      =   "-"
   End
   Begin VB.TextBox Text5 
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
      Height          =   435
      Left            =   120
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame4 
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
      Height          =   5175
      Left            =   6240
      TabIndex        =   54
      Top             =   1440
      Width           =   5535
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   47
         Text            =   "0"
         Top             =   3653
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   49
         Text            =   "0"
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   48
         Text            =   "0"
         Top             =   4200
         Width           =   975
      End
      Begin VB.Frame Frame9 
         Caption         =   "Wire-o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   3960
         TabIndex        =   57
         Top             =   480
         Width           =   1455
         Begin VB.CheckBox Check1 
            Caption         =   "Embutir"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   4200
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   240
            MaxLength       =   3
            TabIndex        =   36
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "1"""
            Enabled         =   0   'False
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   45
            Top             =   3720
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            Caption         =   "7/8"""
            Enabled         =   0   'False
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   44
            Top             =   3360
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            Caption         =   "5/8"""
            Enabled         =   0   'False
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   43
            Top             =   3000
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "9/16"""
            Enabled         =   0   'False
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   42
            Top             =   2640
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "1/2"""
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   41
            Top             =   2280
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            Caption         =   "7/16"""
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   40
            Top             =   1920
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "3/8"""
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   39
            Top             =   1560
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            Caption         =   "5/16"""
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   38
            Top             =   1200
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "1/4"""
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   37
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Capas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   2040
         TabIndex        =   56
         Top             =   480
         Width           =   1815
         Begin VB.OptionButton Option3 
            Caption         =   "Papelão A5"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Papelão A4"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   1260
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Papelão A3"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   1680
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Capa PVC A5"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   2100
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Capa PVC A4"
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   35
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   360
            MaxLength       =   3
            TabIndex        =   30
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Laminação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   1815
         Begin VB.OptionButton Option1 
            Caption         =   "Polaseal A6"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   195
            TabIndex        =   26
            Top             =   1331
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Polaseal A5"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   195
            TabIndex        =   27
            Top             =   1747
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Polaseal A4"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   195
            TabIndex        =   28
            Top             =   2163
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Polaseal A3"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   195
            TabIndex        =   29
            Top             =   2580
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "BOPP"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   195
            TabIndex        =   25
            Top             =   915
            Width           =   1215
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   360
            MaxLength       =   3
            TabIndex        =   24
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Meio Corte:"
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
         Left            =   855
         TabIndex        =   71
         Top             =   3720
         Width           =   1200
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Transfer aplic.:"
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
         Left            =   495
         TabIndex        =   70
         Top             =   4740
         Width           =   1575
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
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
         Left            =   1455
         TabIndex        =   69
         Top             =   4260
         Width           =   630
      End
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
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   5040
      TabIndex        =   53
      Top             =   2160
      Width           =   975
      Begin VB.OptionButton Option2 
         Caption         =   "6"
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
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "4"
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gramatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3495
      Left            =   3240
      TabIndex        =   52
      Top             =   2160
      Width           =   1575
      Begin VB.OptionButton optGramatura 
         Caption         =   "300g"
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
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   21
         Top             =   3000
         Width           =   855
      End
      Begin VB.OptionButton optGramatura 
         Caption         =   "250g"
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
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   2625
         Width           =   975
      End
      Begin VB.OptionButton optGramatura 
         Caption         =   "180g"
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
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   2250
         Width           =   975
      End
      Begin VB.OptionButton optGramatura 
         Caption         =   "170g"
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
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   1875
         Width           =   975
      End
      Begin VB.OptionButton optGramatura 
         Caption         =   "150g"
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
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   1485
         Width           =   975
      End
      Begin VB.OptionButton optGramatura 
         Caption         =   "120g"
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
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1110
         Width           =   975
      End
      Begin VB.OptionButton optGramatura 
         Caption         =   "115g"
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
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   735
         Width           =   975
      End
      Begin VB.OptionButton optGramatura 
         Caption         =   "90g"
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
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Midia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
      Begin VB.OptionButton optMidia 
         Caption         =   "Banner"
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
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1095
      End
      Begin VB.OptionButton optMidia 
         Caption         =   "BOPP"
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
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   2565
         Width           =   975
      End
      Begin VB.OptionButton optMidia 
         Caption         =   "Adesivo"
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
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   2115
         Width           =   1215
      End
      Begin VB.OptionButton optMidia 
         Caption         =   "Transfer"
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
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optMidia 
         Caption         =   "Couche"
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
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optMidia 
         Caption         =   "Apergaminhado"
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
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   795
         Width           =   2055
      End
      Begin VB.OptionButton optMidia 
         Caption         =   "Reciclato"
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
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1245
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "0"
         Top             =   3000
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "0"
         Top             =   3000
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2050
         TabIndex        =   58
         Top             =   3050
         Visible         =   0   'False
         Width           =   165
      End
   End
   Begin VB.Label labelSoma 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   3735
      TabIndex        =   75
      Top             =   6015
      Width           =   2055
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3720
      TabIndex        =   76
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
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
      Left            =   1560
      TabIndex        =   63
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2235
      TabIndex        =   62
      Top             =   6120
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade"
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
      TabIndex        =   61
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Telefone"
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
      Left            =   6360
      TabIndex        =   60
      Top             =   360
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   120
      TabIndex        =   59
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public valorServico As Currency
Public gramatura As Integer
Public alterado As Boolean
Public altLami As Boolean
Public tecla As Integer
Public totalLami As Currency
Public totalCapa As Currency
Public totalWireo As Currency
Public meiocorte As Currency
Public corte As Currency
Public transfer As Currency


Function calculo()

    valorServico = 0
Select Case gramatura

    Case 0 ' 90g
        If optMidia(0).Value Or optMidia(1).Value = True = True Then
            valorServico = 2
        End If
        
        If optMidia(4).Value = True Or optMidia(3).Value = True Then
            valorServico = 3
        End If
        
    Case 1 ' 115g
        If optMidia(0).Value = True Then
            valorServico = 2
        End If
            
    Case 2 ' 120
        If optMidia(2).Value = True Then
            valorServico = 2
        End If
         If optMidia(5).Value = True Then
            valorServico = 3
        End If
        
    Case 3 ' 150
        If optMidia(0).Value = True Then
            valorServico = 2
        End If
        
    Case 4 ' 170
        If optMidia(0).Value = True Then
            valorServico = 2
        End If
        
    Case 5 ' 180
        If optMidia(1).Value = True Then
            valorServico = 2
        End If
        
    Case 6 ' 250
        If optMidia(0).Value = True Then
            valorServico = 2.5
        End If
    
    Case 7 ' 300
        If optMidia(0).Value = True Then
            valorServico = 2.5
        End If
    
    Case 8 ' banner
        valorServico = 0
End Select

If gramatura < 8 Then
    If Option2(1).Value Then
        valorServico = valorServico + 1
    End If
End If

labelSoma.Caption = (valorServico * Text5.Text) + totalLami + totalCapa + totalWireo + meiocorte + corte + transfer


End Function

Private Sub Check1_Click()
    calculoWireo
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8 Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
    End If

End Sub

Private Sub Command1_Click(Index As Integer)

Select Case Index

    Case 0
        frmSalvar.Show 1

    Case 1
        Dim impressoraPadrão As Printer
        Dim Nova_Impressora
    
        CD.ShowPrinter
        Set impressoraPadrão = Printer
        
        For Each Nova_Impressora In Printers
            If Nova_Impressora.hDC = CD.hDC Then
                Set Printer = Nova_Impressora
            End If
        Next
        
        Printer.Print "Teste!"
        Printer.EndDoc
        'Set Printer = impressoraPadrão

    Case 2
        Debug.Print Printer.DeviceName

    Case 3
        Unload Me

End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        frmListaOrcamento.Show 1
    End If

End Sub

Private Sub Form_Load()
    Label8.Caption = Format(Date, "dd/mm/yy")
    Label9.Caption = Format(Time, "HH:MM")
    alterado = False
    stbOrcamento.Panels(1).Text = "Vendedor: " & varNomeUsuario
    stbOrcamento.Panels(2).Text = "Inclusão: " & Label8.Caption & " - " & Label9.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub labelSoma_Change()
    labelSoma.Caption = Format$(labelSoma.Caption, "#,##0.00")
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub optGramatura_Click(Index As Integer)
    gramatura = Index
    calculo
End Sub

Private Sub Option1_Click(Index As Integer)

If Text6.Text <> 0 Then
    calculoAcabamento
End If

End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub Option2_Click(Index As Integer)
    calculo
End Sub

Private Sub Option2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub Option3_Click(Index As Integer)
If Text8.Text <> 0 Then
    calculoCapa
End If
End Sub

Private Sub Option3_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub Option4_Click(Index As Integer)
If Text4.Text <> 0 Then
    calculoWireo
End If
End Sub


Private Sub Option5_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox KeyCode
End Sub


Private Sub Option4_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Public Sub optMidia_Click(Index As Integer)

alterado = True
Frame2.Enabled = True
Frame3.Enabled = True
Frame4.Enabled = True
Option2(0).Value = True

Select Case Index

    Case 0
        desabilitaMidia (True)
        optGramatura(0).Enabled = True
        optGramatura(1).Enabled = True
        optGramatura(2).Enabled = False
        optGramatura(3).Enabled = True
        optGramatura(4).Enabled = True
        optGramatura(5).Enabled = False
        optGramatura(6).Enabled = True
        optGramatura(7).Enabled = True
        Option2(0).Enabled = True
        Option2(1).Enabled = True
        Option2(0).Value = True
        Text2.Visible = False
        Text3.Visible = False
        lblX.Visible = False
        optGramatura(0).Value = True
        


    Case 1
        optGramatura(0).Enabled = True
        optGramatura(1).Enabled = False
        optGramatura(2).Enabled = False
        optGramatura(3).Enabled = False
        optGramatura(4).Enabled = False
        optGramatura(5).Enabled = True
        optGramatura(6).Enabled = False
        optGramatura(7).Enabled = False
        Option2(0).Enabled = True
        Option2(1).Enabled = True
        Text2.Visible = False
        Text3.Visible = False
        lblX.Visible = False
        optGramatura(0).Value = True
    
    Case 2
    
        optGramatura(0).Enabled = False
        optGramatura(1).Enabled = False
        optGramatura(2).Enabled = True
        optGramatura(3).Enabled = False
        optGramatura(4).Enabled = False
        optGramatura(5).Enabled = False
        optGramatura(6).Enabled = False
        Option2(0).Enabled = True
        Option2(1).Enabled = True
        optGramatura(7).Enabled = False
        Text2.Visible = False
        optGramatura(2).Value = True
        Text3.Visible = False
        lblX.Visible = False
        
    Case 3
    
        optGramatura(0).Enabled = True
        optGramatura(1).Enabled = False
        optGramatura(2).Enabled = False
        optGramatura(3).Enabled = False
        optGramatura(4).Enabled = False
        optGramatura(5).Enabled = False
        optGramatura(6).Enabled = False
        optGramatura(7).Enabled = False
        Option2(0).Enabled = True
        Option2(1).Enabled = True
        Text2.Visible = False
        Text3.Visible = False
        lblX.Visible = False
        optGramatura(0).Value = True
        
    Case 4
    
        optGramatura(0).Enabled = True
        optGramatura(1).Enabled = False
        optGramatura(2).Enabled = False
        optGramatura(3).Enabled = False
        optGramatura(4).Enabled = False
        optGramatura(5).Enabled = False
        optGramatura(6).Enabled = False
        optGramatura(7).Enabled = False
        optGramatura(0).Value = True
        Option2(0).Enabled = True
        Option2(1).Enabled = True
        Text2.Visible = False
        Text3.Visible = False
        lblX.Visible = False
        
    Case 5

        optGramatura(0).Enabled = False
        optGramatura(1).Enabled = False
        optGramatura(2).Enabled = True
        optGramatura(3).Enabled = False
        optGramatura(4).Enabled = False
        optGramatura(5).Enabled = False
        optGramatura(6).Enabled = False
        optGramatura(7).Enabled = False
        optGramatura(2).Value = True
        
        Option2(0).Enabled = True
        Option2(1).Enabled = True
        
        Text2.Visible = False
        Text3.Visible = False
        lblX.Visible = False
        
    Case 6 'Banner
        Dim Y As Integer
        For Y = 0 To optGramatura.Count - 1
            optGramatura(Y).Value = False
            optGramatura(Y).Enabled = True
        Next
        
        gramatura = 8
        Frame2.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
        
        
        Text2.Text = 0
        Text3.Text = 0
        Text6.Text = 0
        Text8.Text = 0
        Text2.Visible = True
        Text3.Visible = True
        lblX.Visible = True
        
        Option2(0).Enabled = False
        Option2(1).Value = True
        
        Exit Sub

End Select
    calculo



End Sub

Private Sub optMidia_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Text10_Change()
    If Text10.Text <> "" Then
        corte = Text10.Text * 0.5
    Else
        Text10.Text = 0
        Text10_GotFocus
    End If
    calculo
End Sub

Private Sub Text10_GotFocus()
    Text10.SelStart = 0
    Text10.SelLength = Len(Text10.Text)
End Sub


Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9") And KeyAscii = Asc(".")) And Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If

End Sub


Private Sub Text11_Change()
    If Text11.Text <> "" Then
        transfer = Text11.Text * 0.5
    Else
        Text11.Text = 0
        Text11_GotFocus
    End If
    calculo
End Sub

Private Sub Text11_GotFocus()
    Text11.SelStart = 0
    Text11.SelLength = Len(Text11.Text)
End Sub


Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If

End Sub


Private Sub Text2_Change()
    Dim vartotal As Integer
    
    If Text2.Text = "" Then
        labelSoma.Caption = 0
        Exit Sub
    Else
       vartotal = Text2.Text * Text3.Text * Text5.Text
       vartotal = vartotal / 10000
       vartotal = vartotal * 35
       labelSoma.Caption = vartotal
    End If
End Sub

Private Sub Text2_GotFocus()

    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_LostFocus()
    If Text2.Text = 0 Or Text2.Text = 0 Then
        Text2.Text = 0
    End If
End Sub

Private Sub Text3_Change()
    Dim vartotal
    
    If Text3.Text = "" Then
        labelSoma.Caption = 0
        Exit Sub
    Else
       vartotal = Text2.Text * Text3.Text * Text5.Text
       vartotal = vartotal / 10000
       vartotal = vartotal * 35
       labelSoma.Caption = vartotal
    End If
    

End Sub


Private Sub Text3_GotFocus()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8 Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text3_LostFocus()
    If Text3.Text = 0 Then
        Text3.Text = 0
    End If
End Sub


Private Sub Text4_Change()
    If Text4.Text <> "" Then
        If Text4.Text <> 0 Then
            Dim Y As Integer
            For Y = 0 To Option4.Count - 1
                Option4(Y).Enabled = True
            Next
        Else
            Y = 0
            For Y = 0 To Option4.Count - 1
                Option4(Y).Enabled = False
                Option4(Y).Value = False
            Next
        End If
    Else
        Text4.Text = 0
        Text4_GotFocus
        Y = 0
        For Y = 0 To Option4.Count - 1
            Option4(Y).Enabled = False
            Option4(Y).Value = False
        Next
    End If
    calculoWireo
End Sub

Private Sub Text4_GotFocus()
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If

End Sub

Private Sub Text5_Change()
    If Text5.Text <> "" Then
        If alterado = True Then
            calculo
        End If
    End If
End Sub

Private Sub Text5_GotFocus()
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text5_LostFocus()

    If Text5.Text = "" Then
            Text5.Text = 0
    End If
    
    If Text5.Text = 0 Then
        desabilitaMidia (False)
        desabilitaGramatura (False)
        desabilitaCores (False)
    Else
        desabilitaMidia (True)
        desabilitaGramatura (True)
        desabilitaCores (True)
    End If

End Sub

Private Sub Text6_Change()
    If Text6.Text <> "" Then
        If Text6.Text <> 0 Then
            Dim Y As Integer
            For Y = 0 To Option1.Count - 1
                Option1(Y).Enabled = True
            Next
        Else
            Y = 0
            For Y = 0 To Option1.Count - 1
                Option1(Y).Enabled = False
                Option1(Y).Value = False
            Next
        End If
    Else
        Text6.Text = 0
        Text6_GotFocus
        Y = 0
        For Y = 0 To Option1.Count - 1
            Option1(Y).Enabled = False
            Option1(Y).Value = False
        Next
    End If
    calculoAcabamento
End Sub

Private Sub Text6_GotFocus()
    Text6.SelStart = 0
    Text6.SelLength = Len(Text6.Text)
End Sub



Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text6_LostFocus()
    If Text6.Text = 0 Then
        Text6.Text = 0
    End If
End Sub

Private Sub text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub Text8_Change()
    If Text8.Text <> "" Then
        If Text8.Text <> 0 Then
            Dim Y As Integer
            For Y = 0 To Option3.Count - 1
                Option3(Y).Enabled = True
            Next
        Else
            Y = 0
            For Y = 0 To Option1.Count - 1
                Option3(Y).Enabled = False
                Option3(Y).Value = False
            Next
        End If
    Else
        Text8.Text = 0
        Text8_GotFocus
        Y = 0
        For Y = 0 To Option3.Count - 1
            Option3(Y).Enabled = False
            Option3(Y).Value = False
        Next
    End If
    calculoCapa
End Sub

Private Sub Text8_GotFocus()
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.Text)
End Sub


Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If

End Sub

Private Sub Text8_LostFocus()
    If Text8.Text = 0 Then
        Text8.Text = 0
    End If
End Sub


Private Sub Text9_Change()
    If Text9.Text <> "" Then
        meiocorte = Text9.Text * 0.1
    Else
        Text9.Text = 0
        Text9_GotFocus
    End If
    calculo
End Sub

Private Sub Text9_GotFocus()
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9.Text)
End Sub


Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtSoma_Change()

End Sub

Public Sub calculoAcabamento()


Dim valor As Currency

If Option1(0).Value = True Or Option1(1).Value = True Then
    valor = 1
End If

If Option1(2).Value = True Then
    valor = 1.5
End If

If Option1(3).Value = True Then
    valor = 2
End If

If Option1(4).Value = True Then
    valor = 3.5
End If

totalLami = valor * Text6.Text
calculo
End Sub

Public Sub calculoCapa()

Dim valor As Currency

If Option3(0).Value = True Then
    valor = 2

ElseIf Option3(1).Value = True Then
    valor = 3

ElseIf Option3(2).Value = True Then
    valor = 6

ElseIf Option3(3).Value = True Then
    valor = 0.5

ElseIf Option3(4).Value = True Then
    valor = 1
End If

totalCapa = valor * Text8.Text
calculo
End Sub

Public Sub calculoWireo()

Dim valor As Currency
Dim embutirwireo As Currency

If Option4(0).Value = True Then
    valor = 1.5

ElseIf Option4(1).Value = True Then
    valor = 1.5

ElseIf Option4(2).Value = True Then
    valor = 1.5

ElseIf Option4(3).Value = True Then
    valor = 2

ElseIf Option4(4).Value = True Then
    valor = 2

ElseIf Option4(5).Value = True Then
    valor = 2

ElseIf Option4(6).Value = True Then
    valor = 2.5

ElseIf Option4(7).Value = True Then
    valor = 3

ElseIf Option4(8).Value = True Then
    valor = 3
End If

If Check1.Value = 1 Then
    embutirwireo = Text4.Text * 3.5
    
End If

totalWireo = valor * Text4.Text + embutirwireo
calculo
End Sub

Public Sub desabilitaMidia(opcao As Boolean)

Dim Y As Integer

For Y = 0 To optMidia.Count - 1
    optMidia(Y).Enabled = opcao
    If Not opcao Then
        optMidia(Y).Value = opcao
    End If
Next

End Sub

Public Sub desabilitaGramatura(opcao As Boolean)

Dim Y As Integer

For Y = 0 To optGramatura.Count - 1
    If Not opcao Then
        optGramatura(Y).Value = opcao
    End If
    optGramatura(Y).Enabled = opcao
Next

End Sub

Public Sub desabilitaCores(opcao As Boolean)

Dim Y As Integer

For Y = 0 To Option2.Count - 1
    
    If Not opcao Then
        Option2(Y).Value = opcao
    End If
    
    Option2(Y).Enabled = opcao
    
Next

End Sub

Private Sub imprimeOs(codigo As Long)


CD.ShowPrinter

    abreConexao
    
    rs.Open "SELECT * FROM servicos WHERE numeroos=" & codigo, db, adOpenStatic, adLockOptimistic
    Printer.FontSize = 10
    Printer.FontName = "Times new roman"
    Printer.Print Spc(20); rs!nome
    Printer.Print rs!quantidade
    Printer.Print rs!descricao
    Printer.Print rs!Telefone
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print rs!Data
    Printer.Print rs!hora
    Printer.Print rs!cores
    Printer.Print
    Printer.EndDoc
End Sub
