VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmSudoku 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Su Doku"
   ClientHeight    =   3735
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox FilPremade 
      Height          =   1650
      Left            =   120
      TabIndex        =   84
      Top             =   4440
      Width           =   2775
   End
   Begin RichTextLib.RichTextBox RTBVerbose 
      Height          =   3495
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6165
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"FrmSudoku.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicFind 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   480
      Picture         =   "FrmSudoku.frx":0080
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   82
      Top             =   3840
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox PicErr 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      Picture         =   "FrmSudoku.frx":02CC
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   81
      Top             =   3840
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   80
      Left            =   3240
      TabIndex        =   80
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   79
      Left            =   2880
      TabIndex        =   79
      Text            =   "7"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   78
      Left            =   2520
      TabIndex        =   78
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   77
      Left            =   2040
      TabIndex        =   77
      Text            =   "8"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   76
      Left            =   1680
      TabIndex        =   76
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   75
      Left            =   1320
      TabIndex        =   75
      Text            =   "5"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   74
      Left            =   840
      TabIndex        =   74
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   73
      Left            =   480
      TabIndex        =   73
      Text            =   "4"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   72
      Left            =   120
      TabIndex        =   72
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   71
      Left            =   3240
      TabIndex        =   71
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   70
      Left            =   2880
      TabIndex        =   70
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   69
      Left            =   2520
      TabIndex        =   69
      Text            =   "9"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   68
      Left            =   2040
      TabIndex        =   68
      Text            =   "6"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   67
      Left            =   1680
      TabIndex        =   67
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   66
      Left            =   1320
      TabIndex        =   66
      Text            =   "2"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   65
      Left            =   840
      TabIndex        =   65
      Text            =   "7"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   64
      Left            =   480
      TabIndex        =   64
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   63
      Left            =   120
      TabIndex        =   63
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   62
      Left            =   3240
      TabIndex        =   62
      Text            =   "2"
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   61
      Left            =   2880
      TabIndex        =   61
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   60
      Left            =   2520
      TabIndex        =   60
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   59
      Left            =   2040
      TabIndex        =   59
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   58
      Left            =   1680
      TabIndex        =   58
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   57
      Left            =   1320
      TabIndex        =   57
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   56
      Left            =   840
      TabIndex        =   56
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   55
      Left            =   480
      TabIndex        =   55
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   54
      Left            =   120
      TabIndex        =   54
      Text            =   "5"
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   53
      Left            =   3240
      TabIndex        =   53
      Text            =   "4"
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   52
      Left            =   2880
      TabIndex        =   52
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   51
      Left            =   2520
      TabIndex        =   51
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   50
      Left            =   2040
      TabIndex        =   50
      Text            =   "1"
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   49
      Left            =   1680
      TabIndex        =   49
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   48
      Left            =   1320
      TabIndex        =   48
      Text            =   "9"
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   840
      TabIndex        =   47
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   480
      TabIndex        =   46
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   120
      TabIndex        =   45
      Text            =   "7"
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   44
      Left            =   3240
      TabIndex        =   44
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   43
      Left            =   2880
      TabIndex        =   43
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   2520
      TabIndex        =   42
      Text            =   "3"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   41
      Left            =   2040
      TabIndex        =   41
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   40
      Left            =   1680
      TabIndex        =   40
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   1320
      TabIndex        =   39
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   840
      TabIndex        =   38
      Text            =   "6"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   37
      Left            =   480
      TabIndex        =   37
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   36
      Left            =   120
      TabIndex        =   36
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   35
      Left            =   3240
      TabIndex        =   35
      Text            =   "6"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   2880
      TabIndex        =   34
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   2520
      TabIndex        =   33
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   2040
      TabIndex        =   32
      Text            =   "7"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   1680
      TabIndex        =   31
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   1320
      TabIndex        =   30
      Text            =   "4"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   840
      TabIndex        =   29
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   480
      TabIndex        =   28
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   120
      TabIndex        =   27
      Text            =   "8"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   3240
      TabIndex        =   26
      Text            =   "1"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   2880
      TabIndex        =   25
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   2520
      TabIndex        =   24
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   2040
      TabIndex        =   23
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   1680
      TabIndex        =   22
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   1320
      TabIndex        =   21
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   840
      TabIndex        =   20
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   480
      TabIndex        =   19
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   120
      TabIndex        =   18
      Text            =   "2"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   3240
      TabIndex        =   17
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   2880
      TabIndex        =   16
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   2520
      TabIndex        =   15
      Text            =   "6"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   2040
      TabIndex        =   14
      Text            =   "5"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   1680
      TabIndex        =   13
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   1320
      TabIndex        =   12
      Text            =   "3"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   840
      TabIndex        =   11
      Text            =   "8"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   480
      TabIndex        =   10
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2880
      TabIndex        =   7
      Text            =   "5"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2040
      TabIndex        =   5
      Text            =   "4"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Text            =   "6"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   83
      Top             =   120
      Width           =   375
   End
   Begin VB.Menu MNUFile 
      Caption         =   "Game"
      Begin VB.Menu MNUHint 
         Caption         =   "Give Hint"
      End
      Begin VB.Menu MNUCheck 
         Caption         =   "Check for Errors"
      End
      Begin VB.Menu MNUSep 
         Caption         =   "-"
      End
      Begin VB.Menu MNU_Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MNU_Gen 
      Caption         =   "Generate"
      Begin VB.Menu MNU_GenUniq 
         Caption         =   "Generate Unique"
      End
      Begin VB.Menu MNUGenCur 
         Caption         =   "Generate from Current"
      End
      Begin VB.Menu MNUSep3 
         Caption         =   "-"
      End
      Begin VB.Menu MNUPremade 
         Caption         =   "Load Premade Square"
      End
      Begin VB.Menu MNUBulk 
         Caption         =   "Bulk Generate"
      End
   End
   Begin VB.Menu MNU_Solve 
      Caption         =   "Solve"
      Begin VB.Menu MNU_SolveEasy 
         Caption         =   "Easy Solve"
      End
      Begin VB.Menu MNU_SolveComplex 
         Caption         =   "Complex Solve"
         Enabled         =   0   'False
      End
      Begin VB.Menu MNUBrute 
         Caption         =   "Brute Force"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MNU_Clear 
      Caption         =   "Clear Grid"
   End
   Begin VB.Menu MNUStop 
      Caption         =   "Stop!"
      Visible         =   0   'False
   End
   Begin VB.Menu MNUHelp 
      Caption         =   "Help"
      Begin VB.Menu MNU_About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FrmSudoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Su Doku Solver and Generater
'By Kevin Pfister

'Have fun and sorry about the Poor Coding

Dim Sudoku_Grid(1 To 9, 1 To 9) As Integer
Dim Grid(1 To 9, 1 To 9, 1 To 9) As Boolean
Dim Grid_Last(1 To 9, 1 To 9, 1 To 9) As Boolean

Dim Gen_Dif(1 To 9, 1 To 9, 1 To 9) As Boolean

Dim Processing As Boolean
Dim AppPath As String

Dim GenGo As Boolean

Sub GRID_Clear()
    'Clear Grid
    For X = 0 To 80
        TxtSudoku(X).Text = ""
    Next
    
    For Y = 1 To 9
        For X = 1 To 9
            Sudoku_Grid(X, Y) = 0
            For Z = 1 To 9
                Grid(X, Y, Z) = False
                Grid_Last(X, Y, Z) = False
            Next
        Next
    Next
End Sub

Private Sub Form_Load()
    
    Me.Show
    DoEvents

    Randomize Timer

    AddText "Welcome to Su Doku Square Solver and Generator"
    AddText "Created by Kevin Pfister"
    
    'Check to see if the Directory exists
    
    If Mid$(App.Path, Len(App.Path) - 1, 1) = "\" Then
        AppPath = App.Path
    Else
        AppPath = App.Path & "\"
    End If
    
    On Error Resume Next
    
    Call MkDir(AppPath & "Premade")
    
    FilPremade.Path = AppPath & "Premade"
    
    DoEvents
    
    MNU_GenUniq_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub MNU_About_Click()
    MsgBox "Created by Kevin Pfister"
End Sub

Private Sub MNU_Clear_Click()

    AddText ""
    AddText "Clearing Grid..."

    Call GRID_Lock
    Call GRID_Clear
    Call GRID_Unlock
End Sub

Private Sub MNU_Exit_Click()
    'Exit Program
    End
End Sub

Sub Txt2Grid()
    Dim Y As Long
    Dim X As Long

    For Y = 1 To 9
        For X = 1 To 9
            If Val(TxtSudoku(Y * 9 - 10 + X).Text) <> 0 Then
                Sudoku_Grid(X, Y) = Val(TxtSudoku(Y * 9 - 10 + X).Text)
            End If
        Next
    Next
End Sub

Sub RenderGrid()
    Dim Y As Long
    Dim X As Long

    For Y = 1 To 9
        For X = 1 To 9
            TxtSudoku(Y * 9 - 10 + X).Text = ""
            For Z = 1 To 9
                If Grid(X, Y, Z) = True Then
                    TxtSudoku(Y * 9 - 10 + X).Text = TxtSudoku(Y * 9 - 10 + X).Text & Z
                End If
            Next
        Next
    Next
    
    DoEvents
End Sub

Sub Grid2Txt()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To 9
        For X = 1 To 9
            If Sudoku_Grid(X, Y) = 0 Then
                TxtSudoku(Y * 9 - 10 + X).Text = ""
            Else
                TxtSudoku(Y * 9 - 10 + X).Text = Sudoku_Grid(X, Y)
            End If
        Next
    Next
End Sub

Function Sudoku_ChkDoubles() As Boolean
    'Check Grid for Sudoku errors
    
    Dim Counted As Integer

    Sudoku_ChkDoubles = False
    
    'Check ->
    
    For Y = 1 To 9
        For Num = 1 To 9
            Counted = 0
            For X = 1 To 9
                If Sudoku_Grid(X, Y) = Num Then
                    Counted = Counted + 1
                End If
            Next
            
            If Counted > 1 Then
               Sudoku_ChkDoubles = True
               Exit Function
            End If
        Next
    Next
    
    'Check \/
    
    For X = 1 To 9
        For Num = 1 To 9
            Counted = 0
            For Y = 1 To 9
                If Sudoku_Grid(X, Y) = Num Then
                    Counted = Counted + 1
                End If
            Next
            
            If Counted > 1 Then
               Sudoku_ChkDoubles = True
               Exit Function
            End If
        Next
    Next
    
    'Check Inner Grids
    
    For X_Outer = 1 To 3
        For Y_Outer = 1 To 3
            
            For Num = 1 To 9
                Counted = 0
                
                For X_Inner = 1 To 3
                    For Y_Inner = 1 To 3
                        
                        If Sudoku_Grid(X_Outer * 3 - 3 + X_Inner, Y_Outer * 3 - 3 + Y_Inner) = Num Then
                            Counted = Counted + 1
                        End If
                        
                    Next
                Next
            Next
            
            If Counted > 1 Then
               Sudoku_ChkDoubles = True
               Exit Function
            End If
        
        Next
    Next
                 
End Function

Private Sub MNU_GenUniq_Click()

    Randomize Timer
    
    AddText ""
    AddText "FIN Generating Square..."
    AddText "...This may take a little while"
    
    GenGo = True
    MNUStop.Visible = True
    
    
    Dim BTime As Long
    BTime = Timer
    
    Call GRID_Lock
    
    Call GRID_Clear
    
    Call Sudoku_GenGrid
    
    If GenGo = False Then
        AddText "ERR Su Doku Generation Stopped"
    Else
    
        AddText "...Su Doku Generated in " & Round(Timer - BTime, 1) & " Seconds"
        AddText "...Saving Grid for Later use"
        Call SaveGrid
    
    End If
    
    Call GRID_Unlock
    
    MNUStop.Visible = False
End Sub

Private Sub MNU_SolveComplex_Click()
    'First Check Grid
    
    If Sudoku_EasySolve = 1 Then
        'Solved, no need for complex Solve
        Exit Sub
    ElseIf Sudoku_EasySolve = -1 Then
        'Unable to Solve, Repeated Square
        Exit Sub
        
    End If
    
    'Start Complex Solve
    
    Call GRID_Unlock
End Sub

Private Sub MNU_SolveEasy_Click()
    
    AddText ""
    AddText "FIN Attempting to Solve Grid..."
    
    Dim BTime As Long
    BTime = Timer

    Call GRID_Lock
    
    For Y = 1 To 9
        For X = 1 To 9
            Sudoku_Grid(X, Y) = 0
            For Z = 1 To 9
                Grid(X, Y, Z) = False
                Grid_Last(X, Y, Z) = False
            Next
        Next
    Next
    
    Call Txt2Grid
    
    'First Check Grid
    
    If Sudoku_ChkDoubles = True Then
        AddText "ERR Su Doku Square has repeated Entry(s)"
        
        Call GRID_Unlock
        
        Exit Sub
    End If
    
    AddText "...Please Wait"
    Possible = Sudoku_EasySolve
    
    Call GRID_Unlock
    
    If Possible = 2 Then
        AddText "ERR Unable to Solve Su Doku Square by Simple Methods, Try using Complex Solve Method"
    ElseIf Possible = 1 Then
        Call RenderGrid
        AddText "...Su Doku Square Solved in " & Abs(Round(Timer - BTime, 4)) & " Seconds"
    End If
    
End Sub

Function Sudoku_EasySolve() As Integer

    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    
    
    Sudoku_EasySolve = 0
    
    'Clear Solving Grid
    For X = 1 To 9
        For Y = 1 To 9
            If Sudoku_Grid(X, Y) > 0 Then
            
                For Z = 1 To 9
                    Grid(X, Y, Z) = False
                    Grid_Last(X, Y, Z) = True
                Next
                
                Grid(X, Y, Sudoku_Grid(X, Y)) = True
            
            Else
            
                For Z = 1 To 9
                    Grid(X, Y, Z) = True
                    Grid_Last(X, Y, Z) = True
                Next
            
            End If
            
        Next
    Next
    
    Dim TotalDone As Boolean
    TotalDone = False
    
    Dim UnSolve As Boolean
    
    Do
        'Major Solve Loop
        
        For Z = 1 To 9
            Call CheckRow(Z)
        Next
    
        For Z = 1 To 9
            Call CheckCol(Z)
        Next

        For Z = 1 To 9
            Call CheckRowVals(Z)
        Next
        
        For Z = 1 To 9
            Call CheckColVals(Z)
        Next

        For X = 1 To 3
            For Y = 1 To 3
                Call CheckMiniSqrs(X, Y)
            Next
        Next
        
        DoEvents
        
        UnSolve = True
        
        For X = 1 To 9
            For Y = 1 To 9
                For Z = 1 To 9
                    If Grid_Last(X, Y, Z) <> Grid(X, Y, Z) Then
                        UnSolve = False
                    End If
                    Grid_Last(X, Y, Z) = Grid(X, Y, Z)
                Next
            Next
        Next
        
        TotalDone = CheckGrid
    Loop Until TotalDone = True Or UnSolve = True
    
    If UnSolve = True Then
        'UnSolvable by simple methods
        Sudoku_EasySolve = 2
    Else
        'Solved
        Sudoku_EasySolve = 1
    End If
    
End Function


Private Sub MNUBrute_Click()
    'First Check Grid
    

    AddText ""
    AddText "FIN Solving Grid by Brute Force"
    
    If Sudoku_EasySolve = 1 Then
        'Solved, no need for complex Solve
        Exit Sub
    ElseIf Sudoku_EasySolve = -1 Then
        'Unable to Solve, Repeated Square
        Exit Sub
        
    End If
    
    'Start Brute Force
    
    Dim PosTries As Double
    
    'First Calculate number of Possible Tries
    
    PosTries = 1
    
    AddText "Calculating Possible Combinations"
    
    For Y = 1 To 9
        For X = 1 To 9
            Sudoku_Grid(X, Y) = 0
            For Z = 1 To 9
                Grid(X, Y, Z) = False
                Grid_Last(X, Y, Z) = False
            Next
        Next
    Next
    
    Call Txt2Grid
    
    For X = 1 To 9
        For Y = 1 To 9
            If Sudoku_Grid(X, Y) > 0 Then
            
                For Z = 1 To 9
                    Grid(X, Y, Z) = False
                    Grid_Last(X, Y, Z) = True
                Next
                
                Grid(X, Y, Sudoku_Grid(X, Y)) = True
            
            Else
            
                For Z = 1 To 9
                    Grid(X, Y, Z) = True
                    Grid_Last(X, Y, Z) = True
                Next
            
            End If
            
        Next
    Next
    
    For X = 1 To 9
        For Y = 1 To 9
            SquareTry = 0
            For Z = 1 To 9
                If Grid(X, Y, Z) = True Then
                    SquareTry = SquareTry + 1
                End If
            Next
            
            PosTries = PosTries * SquareTry
        
        Next
    Next
    
    AddText "FIN " & PosTries & " Possible Combinations"
    
    Call GRID_Unlock
End Sub

Private Sub MNUBulk_Click()

    Ask = InputBox("How Many Squares would you like Generated?", , 5)
    
    If Val(Ask) = 0 Then
        Exit Sub
    End If
    
    AddText ""
    AddText "FIN Generating " & Ask & " Squares..."
    AddText "...This may take a while"
    
    GenGo = True
    MNUStop.Visible = True
    
    Call GRID_Lock
    
    Dim TTime As Long
    Dim BTime As Long
    TTime = Timer
    
    
    For Z = 1 To Ask
    
        Randomize Timer
        
        BTime = Timer
        
        Call GRID_Clear
        
        Call Sudoku_GenGrid
        
        If GenGo = False Then
            AddText "ERR Su Doku Generation Stopped"
            
            Call GRID_Unlock
            Exit Sub
        Else
            
            AddText "...[" & Z & "] Su Doku Generated in " & Round(Timer - BTime, 1) & " Seconds"
            AddText "...Total Time :: " & Round(Timer - TTime, 1) & " Seconds"
            
            Call SaveGrid
            
        End If
    Next
    
    AddText "...Finished"
    
    Call GRID_Unlock
    
    MNUStop.Visible = False
    
End Sub

Private Sub MNUCheck_Click()
    Call Txt2Grid

    AddText ""
    AddText "FIN Checking Su Doku Square"

    If Sudoku_ChkDoubles = False Then
        AddText "Su Doku Square is currently Valid"
    Else
        AddText "ERR Su Doku Square is currently Not Valid"
    End If
End Sub

Private Sub MNUGenCur_Click()

    Dim X As Long
    Dim Y As Long
    Dim X1 As Long
    Dim Y1 As Long
    Dim Z As Long
    Dim TempVal As Long
    
    AddText ""
    AddText "FIN Generating New Square from Current"
    
    Dim BTime As Long
    BTime = Timer

    Call GRID_Lock
    
    For Y = 1 To 9
        For X = 1 To 9
            Sudoku_Grid(X, Y) = 0
            For Z = 1 To 9
                Grid(X, Y, Z) = False
                Grid_Last(X, Y, Z) = False
            Next
        Next
    Next
    
    Call Txt2Grid
    
    For X = 0 To 80
        TxtSudoku(X).Text = ""
    Next
    
    'First Check Grid
    
    If Sudoku_ChkDoubles = True Then
        AddText "ERR Unable to Generate new Square from current Layout"
        
        Call GRID_Unlock
        
        Exit Sub
    End If
    
    AddText "...Please Wait"
    Possible = Sudoku_EasySolve
    
    If Possible = 2 Then
        AddText "ERR Unable to Generate new Square from current Layout"
    ElseIf Possible = 1 Then
        Z = 0
        
        Dim Unable As Long
        
        
        For X = 1 To 9
            For Y = 1 To 9
                For Z = 1 To 9
                    Gen_Dif(X, Y, Z) = Grid(X, Y, Z)
                    Grid_Last(X, Y, Z) = False
                Next
            Next
        Next
        
        Dim Counter As Long
        
        Counter = 0
        Unable = 0
        
        Do
            X = Int(Rnd * 9) + 1
            Y = Int(Rnd * 9) + 1
            
            TempVal = WhatVal1(X, Y)
            
            If TempVal <> 0 Then
            
            
                For X1 = 1 To 9
                    For Y1 = 1 To 9
                        For Z1 = 1 To 9
                            Grid(X1, Y1, Z1) = Gen_Dif(X1, Y1, Z1)
                            Grid_Last(X1, Y1, Z1) = False
                        Next
                        Sudoku_Grid(X1, Y1) = WhatVal(X1, Y1)
                        DoEvents
                    Next
                Next
                
                For Z1 = 1 To 9
                    Grid(X, Y, Z1) = False
                    Gen_Dif(X, Y, Z1) = False
                Next
                
                Sudoku_Grid(X, Y) = 0
                
                DoEvents
                
                Answer = Sudoku_EasySolve
                
                If Answer = 2 Then
                    Grid(X, Y, TempVal) = True
                    Gen_Dif(X, Y, TempVal) = True
                    Sudoku_Grid(X, Y) = TempVal
                    Unable = Unable + 1
                Else
                    Counter = Counter + 1
                    Unable = 0
                End If
                
            End If
            
        Loop Until Counter = 80 Or Unable = 200
        
        For X = 1 To 9
            For Y = 1 To 9
                For Z = 1 To 9
                    Grid(X, Y, Z) = Gen_Dif(X, Y, Z)
                Next
            Next
        Next
        
        Call RenderGrid
        
        AddText "...Differnt Square Generated in " & Abs(Round(Timer - BTime, 4)) & " Seconds"
        AddText "...Saving Grid for Later use"
        Call SaveGrid
    End If
    Call GRID_Unlock
End Sub

Private Sub MNUHint_Click()
    
    AddText ""
    AddText "FIN Giving a hint..."
    
    Call GRID_Lock
    
    For Y = 1 To 9
        For X = 1 To 9
            Sudoku_Grid(X, Y) = 0
            For Z = 1 To 9
                Grid(X, Y, Z) = False
                Grid_Last(X, Y, Z) = False
            Next
        Next
    Next
    
    Call Txt2Grid
    
    'First Check Grid
    
    If Sudoku_ChkDoubles = True Then
        AddText "ERR Unable to give hint for current Square"
        
        Call GRID_Unlock
        
        Exit Sub
    End If
    
    Possible = Sudoku_EasySolve
    
    Call GRID_Unlock
    
    If Possible = 2 Then
        AddText "ERR Unable to give hint for current Square"
    ElseIf Possible = 1 Then
        
        Dim GivenHint As Boolean
        GivenHint = False
        
        Do
            X = Int(Rnd * 9) + 1
            Y = Int(Rnd * 9) + 1
            
            If TxtSudoku(Y * 9 - 10 + X).Text = "" Then
                For Z = 1 To 9
                    If Grid(X, Y, Z) = True Then
                        TxtSudoku(Y * 9 - 10 + X).Text = TxtSudoku(Y * 9 - 10 + X).Text & Z
                    End If
                Next
                GivenHint = True
            End If
        Loop Until GivenHint = True
    End If
    
End Sub

Private Sub MNUPremade_Click()

    AddText ""
    AddText "Loading a premade Su Doku Square"
    
    
    If FilPremade.ListCount = 0 Then
        AddText "ERR No Premade Squares available, Try generating some"
        Exit Sub
    End If
    
    Premade = Int(Rnd * FilPremade.ListCount) + 1
    
    Dim Filename As String
    
    Filename = FilPremade.List(Premade)
    
    Filename = Mid$(Filename, 1, Len(Filename) - 4)
    
    If Len(Filename) <> 81 Then
        AddText "ERR Premade File is Corrupt"
        Exit Sub
    End If
    
    
    Call GRID_Lock
    Call GRID_Clear
    
    For Z = 1 To 81
        If Mid$(Filename, Z, 1) <> "0" Then
            TxtSudoku(Z - 1).Text = Mid$(Filename, Z, 1)
        End If
    Next
    
    AddText "...Square Loaded"
    
    Call GRID_Unlock
End Sub


Private Sub MNUStop_Click()
    GenGo = False
End Sub

Private Sub TxtSudoku_Change(Index As Integer)
    'Check for Long Entry

    If Processing = False Then
    
        If Len(TxtSudoku(Index).Text) > 1 Then
            TxtSudoku(Index).Text = Mid$(TxtSudoku(Index).Text, 1, 1)
            Beep
        End If
        
        
        'Check for Invalid Entries
        If Len(TxtSudoku(Index).Text) = 1 Then
            If Val(TxtSudoku(Index).Text) = 0 Then
                TxtSudoku(Index).Text = ""
                Beep
            End If
        End If
        
    End If
End Sub


Sub GRID_Lock()
    'Lock the Grid from editing
    
    Processing = True
    
    For X = 1 To 9
        For Y = 1 To 9
            TxtSudoku(Y * 9 - 10 + X).Locked = True
        Next
    Next
    
    MNUFile.Enabled = False
    MNU_Gen.Enabled = False
    MNU_Clear.Enabled = False
    MNU_Solve.Enabled = False
    
End Sub

Sub GRID_Unlock()
    'UnLock the grid for editing

    Processing = False
    
    For X = 1 To 9
        For Y = 1 To 9
            TxtSudoku(Y * 9 - 10 + X).Locked = False
        Next
    Next
    
    MNUFile.Enabled = True
    MNU_Gen.Enabled = True
    MNU_Clear.Enabled = True
    MNU_Solve.Enabled = True
End Sub

Sub Sudoku_GenGrid()
    Dim Gen As Boolean
    
    Dim X As Long
    Dim X1 As Long
    Dim X2 As Long
    
    Dim Y As Long
    Dim Y1 As Long
    Dim Y2 As Long
    
    Dim Z As Long
    
    
    Gen = False
    
    Processing = True
    
    For Y = 1 To 9
        For X = 1 To 9
            Sudoku_Grid(X, Y) = 0
            For Z = 1 To 9
                Grid(X, Y, Z) = False
                Grid_Last(X, Y, Z) = False
            Next
        Next
    Next
    
    Dim Adding As Boolean
    Dim AddNum(1 To 9) As Boolean
    
    Adding = True
    
    Dim SquareFilled As Long
    SquareFilled = 0
        
    Do
    
        X = Int(Rnd * 9) + 1
        Y = Int(Rnd * 9) + 1
        
        For Z = 1 To 9
            AddNum(Z) = True
        Next
        
        If Sudoku_Grid(X, Y) = 0 Then
            If Adding = True Then
                
                'Check Row/Col to see what Number can be used
                For Z = 1 To 9
                    If Sudoku_Grid(Z, Y) > 0 Then
                        AddNum(Sudoku_Grid(Z, Y)) = False
                    End If
                    If Sudoku_Grid(X, Z) > 0 Then
                        AddNum(Sudoku_Grid(X, Z)) = False
                    End If
                Next
                
                'Check mini Square
                
                If X < 4 Then
                    X1 = 1
                ElseIf X < 7 Then
                    X1 = 2
                Else
                    X1 = 3
                End If
                
                If Y < 4 Then
                    Y1 = 1
                ElseIf Y < 7 Then
                    Y1 = 2
                Else
                    Y1 = 3
                End If
                
                For X2 = 1 To 3
                    For Y2 = 1 To 3
                        If Sudoku_Grid(X1 * 3 - 3 + X2, Y1 * 3 - 3 + Y2) > 0 Then
                            AddNum(Sudoku_Grid(X1 * 3 - 3 + X2, Y1 * 3 - 3 + Y2)) = False
                        End If
                    Next
                Next
                
                SqrSol = 0
                
                For Z = 1 To 9
                    If AddNum(Z) = True Then
                        SqrSol = SqrSol + 1
                    End If
                Next
                
                If SqrSol > 0 Then
                    Do
                        PosNum = Int(Rnd * 9) + 1
                        If AddNum(PosNum) = True Then
                            Sudoku_Grid(X, Y) = PosNum
                            
                            TxtSudoku(Y * 9 - 10 + X).Text = PosNum
                        End If
                        
                        DoEvents
                    Loop Until Sudoku_Grid(X, Y) <> 0
                End If
            End If
        Else
            If Adding = False Then
                Sudoku_Grid(X, Y) = 0
                TxtSudoku(Y * 9 - 10 + X).Text = ""
            End If
        
        End If
        
        
        If Adding = True Then
            SquareFilled = SquareFilled + 1
            If SquareFilled = 24 Then
                Adding = False
            End If
        Else
            SquareFilled = SquareFilled - 1
            If SquareFilled = 12 Then
                Adding = True
            End If
        End If
        
        For Y = 1 To 9
            For X = 1 To 9
                For Z = 1 To 9
                    Grid(X, Y, Z) = False
                    Grid_Last(X, Y, Z) = False
                Next
            Next
        Next
        
        If Sudoku_EasySolve = 1 Then
            Gen = True
        End If
    
        DoEvents
        
    Loop Until Gen = True Or GenGo = False
    
    Processing = False
    
    
    If GenGo = False Then
        Call ClearGrid
        Call GRID_Clear
    Else
        Call Grid2Txt
    End If
End Sub

Sub ClearGrid()
    For X = 1 To 9
        For Y = 1 To 9
            Sudoku_Grid(X, Y) = 0
        Next
    Next
End Sub

Function CheckGrid() As Boolean
    Dim X As Long
    Dim Y As Long
    

    CheckGrid = True
    For Y = 1 To 9
        For X = 1 To 9
            If IsVal(X, Y) = False Then
                CheckGrid = False
                Exit Function
            End If
        Next
    Next
End Function

Sub CheckRow(Y As Long)
    Dim X As Long
    Dim Z As Long

    For X = 1 To 9
        If IsVal(X, Y) = True Then
            ChkVal = WhatVal(X, Y)
            For Z = 1 To 9
                If Z <> X Then
                    Grid(Z, Y, ChkVal) = False
                End If
            Next
        End If
    Next
End Sub

Sub CheckCol(X As Long)
    Dim Y As Long
    Dim Z As Long

    For Y = 1 To 9
        If IsVal(X, Y) = True Then
            ChkVal = WhatVal(X, Y)
            For Z = 1 To 9
                If Z <> Y Then
                    Grid(X, Z, ChkVal) = False
                End If
            Next
        End If
    Next
End Sub

Sub CheckRowVals(Y As Long)
    Dim X As Long
    Dim Z As Long
    Dim Vals As Long
    Dim Pos As Long
    Dim Z1 As Long

    For Z = 1 To 9
        Vals = 0
        Pos = 0
        For X = 1 To 9
            If Grid(X, Y, Z) = True Then
                Vals = Vals + 1
                Pos = X
            End If
        Next
        
        If Vals = 1 Then
            For Z1 = 1 To 9
                If Z1 <> Z Then
                    Grid(Pos, Y, Z1) = False
                End If
            Next
        End If
    Next
End Sub

Sub CheckColVals(X As Long)
    Dim Y As Long
    Dim Z As Long
    Dim Z1 As Long
    
    Dim Vals As Long
    Dim Pos As Long
    
    For Z = 1 To 9
        Vals = 0
        Pos = 0
        For Y = 1 To 9
            If Grid(X, Y, Z) = True Then
                Vals = Vals + 1
                Pos = Y
            End If
        Next
        
        If Vals = 1 Then
            For Z1 = 1 To 9
                If Z1 <> Z Then
                    Grid(X, Pos, Z1) = False
                End If
            Next
        End If
    Next
End Sub

Sub CheckMiniSqrs(X As Long, Y As Long)
    Dim X1 As Long
    Dim Y1 As Long
    Dim X2 As Long
    Dim Y2 As Long
    Dim ChkVal As Long

    For X1 = 1 To 3
        For Y1 = 1 To 3
            If IsVal(X * 3 - 3 + X1, Y * 3 - 3 + Y1) = True Then
                ChkVal = WhatVal(X * 3 - 3 + X1, Y * 3 - 3 + Y1)
                For X2 = 1 To 3
                    For Y2 = 1 To 3
                        If X * 3 - 3 + X1 <> X * 3 - 3 + X2 And Y * 3 - 3 + Y1 <> Y * 3 - 3 + Y2 Then
                            Grid(X * 3 - 3 + X2, Y * 3 - 3 + Y2, ChkVal) = False
                        End If
                    Next
                Next
            End If
        Next
    Next
End Sub

Function IsVal(X As Long, Y As Long) As Boolean

    Dim Vals As Long
    Dim Z As Long
    
    For Z = 1 To 9
        If Grid(X, Y, Z) = True Then
            Vals = Vals + 1
        End If
    Next
    
    If Vals = 1 Then
        IsVal = True
    Else
        IsVal = False
    End If
End Function

Function WhatVal(X As Long, Y As Long)
    Dim Z As Long
    
    WhatVal = 0
    
    For Z = 1 To 9
        If Grid(X, Y, Z) = True Then
            WhatVal = Z
            Exit Function
        End If
    Next
End Function

Function WhatVal1(X As Long, Y As Long)
    Dim Z As Long
    
    WhatVal1 = 0
    
    For Z = 1 To 9
        If Gen_Dif(X, Y, Z) = True Then
            WhatVal1 = Z
            Exit Function
        End If
    Next
End Function

Sub MakeVal(Y As Long, X As Long, Value As Long)
    Dim Z As Long
    
    For Z = 1 To 9
        Grid(X, Y, Z) = False
    Next
    
    Grid(X, Y, Value) = True
End Sub

Sub SaveGrid()
    'Save the Grid to a File
    
    Dim Filename As String
    Dim X As Long
    Dim Y As Long
    
    
    For X = 1 To 9
        For Y = 1 To 9
            If TxtSudoku(Y * 9 - 10 + X).Text = "" Then
                Filename = Filename & "0"
            Else
                Filename = Filename & TxtSudoku(Y * 9 - 10 + X).Text
            End If
        Next
    Next
    
    Filename = Filename
    
    Close
    Open AppPath & "Premade\" & Filename & ".txt" For Output As #1
    
    Print #1, Filename
    
    Close
    
End Sub
