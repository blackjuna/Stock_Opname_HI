VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form Progress SO"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20160
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh Data"
      Height          =   495
      Left            =   5280
      TabIndex        =   58
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Laporan"
      Height          =   495
      Left            =   1680
      TabIndex        =   57
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   5520
      Top             =   7800
   End
   Begin VB.Label Blank 
      AutoSize        =   -1  'True
      Caption         =   "BLANK"
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
      Left            =   1800
      TabIndex        =   113
      Top             =   6360
      Width           =   780
   End
   Begin VB.Label ltr_blank 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   112
      Top             =   6360
      Width           =   1740
   End
   Begin VB.Label lti_blank 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   111
      Top             =   6360
      Width           =   1740
   End
   Begin VB.Label lc_blank 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   110
      Top             =   6360
      Width           =   1740
   End
   Begin VB.Label lc_asset 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   109
      Top             =   6360
      Width           =   1740
   End
   Begin VB.Label lc_cc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   108
      Top             =   5880
      Width           =   1740
   End
   Begin VB.Label lc_g 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   107
      Top             =   5400
      Width           =   1740
   End
   Begin VB.Label lc_wh_main 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   106
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Label lc_csd_room 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   105
      Top             =   4440
      Width           =   1740
   End
   Begin VB.Label lc_container 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   104
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label lti_asset 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   103
      Top             =   6360
      Width           =   1740
   End
   Begin VB.Label lti_cc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   102
      Top             =   5880
      Width           =   1740
   End
   Begin VB.Label lti_g 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   101
      Top             =   5400
      Width           =   1740
   End
   Begin VB.Label lti_wh_main 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   100
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Label lti_csd_room 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   99
      Top             =   4440
      Width           =   1740
   End
   Begin VB.Label lti_container 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   98
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label ltr_asset 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   97
      Top             =   6360
      Width           =   1740
   End
   Begin VB.Label ltr_cc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   96
      Top             =   5880
      Width           =   1740
   End
   Begin VB.Label ltr_g 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   95
      Top             =   5400
      Width           =   1740
   End
   Begin VB.Label ltr_wh_main 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   94
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Label ltr_csd_room 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   93
      Top             =   4440
      Width           =   1740
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "ASSET"
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
      Left            =   9960
      TabIndex        =   92
      Top             =   6360
      Width           =   795
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "CONT. CONS"
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
      Left            =   9600
      TabIndex        =   91
      Top             =   5880
      Width           =   1440
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "GAS AREA"
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
      Left            =   9720
      TabIndex        =   90
      Top             =   5400
      Width           =   1260
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "WHMAIN "
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
      Left            =   9840
      TabIndex        =   89
      Top             =   4920
      Width           =   1065
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Area C"
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
      Left            =   9600
      TabIndex        =   88
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "CONTAINER"
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
      Left            =   9600
      TabIndex        =   87
      Top             =   3960
      Width           =   1395
   End
   Begin VB.Label lc_u 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   86
      Top             =   5880
      Width           =   1740
   End
   Begin VB.Label lti_u 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   85
      Top             =   5880
      Width           =   1740
   End
   Begin VB.Label ltr_u 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   84
      Top             =   5880
      Width           =   1740
   End
   Begin VB.Label lc_t 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   83
      Top             =   5400
      Width           =   1740
   End
   Begin VB.Label lti_t 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   82
      Top             =   5400
      Width           =   1740
   End
   Begin VB.Label ltr_t 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   81
      Top             =   5400
      Width           =   1740
   End
   Begin VB.Label lc_j 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   80
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Label lti_j 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   79
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Label ltr_j 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   78
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "J"
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
      Left            =   2160
      TabIndex        =   77
      Top             =   4920
      Width           =   120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "U"
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
      Left            =   2160
      TabIndex        =   76
      Top             =   5880
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "T"
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
      Left            =   2160
      TabIndex        =   75
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label ltr_rluar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   74
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label ltr_k 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   73
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label lti_rluar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   72
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lti_k 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   71
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label lc_rluar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   70
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lc_k 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   69
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "FG"
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
      Left            =   9600
      TabIndex        =   68
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "K"
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
      Left            =   10200
      TabIndex        =   67
      Top             =   3480
      Width           =   150
   End
   Begin VB.Label lc_i 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   66
      Top             =   4440
      Width           =   1740
   End
   Begin VB.Label lc_h 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   65
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label lti_i 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   64
      Top             =   4440
      Width           =   1740
   End
   Begin VB.Label lti_h 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   63
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label ltr_i 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   62
      Top             =   4440
      Width           =   1740
   End
   Begin VB.Label ltr_h 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   61
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "I"
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
      Left            =   2160
      TabIndex        =   60
      Top             =   4440
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "H"
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
      Left            =   2160
      TabIndex        =   59
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label lbl_jam 
      AutoSize        =   -1  'True
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
      Left            =   6840
      TabIndex        =   56
      Top             =   7080
      Width           =   2010
   End
   Begin VB.Label lbl_tanggal 
      AutoSize        =   -1  'True
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
      Left            =   3120
      TabIndex        =   55
      Top             =   7080
      Width           =   2010
   End
   Begin VB.Label ltr_container 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   54
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Line Line9 
      X1              =   9480
      X2              =   17640
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line8 
      X1              =   9480
      X2              =   17640
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label Label110 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Area"
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
      Left            =   9720
      TabIndex        =   53
      Top             =   600
      Width           =   1305
   End
   Begin VB.Label Label109 
      AutoSize        =   -1  'True
      Caption         =   "% Complete"
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
      Left            =   15720
      TabIndex        =   52
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label Label108 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Input"
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
      Left            =   13560
      TabIndex        =   51
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label107 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Released"
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
      Left            =   11400
      TabIndex        =   50
      Top             =   600
      Width           =   1770
   End
   Begin VB.Label Label106 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "PT. HIDROFLEX"
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
      Left            =   12645
      TabIndex        =   49
      Top             =   0
      Width           =   2145
   End
   Begin VB.Label Label105 
      AutoSize        =   -1  'True
      Caption         =   "S"
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
      Left            =   10080
      TabIndex        =   48
      Top             =   1080
      Width           =   165
   End
   Begin VB.Label Label103 
      AutoSize        =   -1  'True
      Caption         =   "R2"
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
      Left            =   10080
      TabIndex        =   47
      Top             =   2040
      Width           =   315
   End
   Begin VB.Label Label102 
      AutoSize        =   -1  'True
      Caption         =   "R3"
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
      Left            =   10080
      TabIndex        =   46
      Top             =   2520
      Width           =   315
   End
   Begin VB.Line Line7 
      X1              =   9480
      X2              =   9480
      Y1              =   480
      Y2              =   7680
   End
   Begin VB.Line Line6 
      X1              =   9480
      X2              =   17640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      X1              =   11160
      X2              =   11160
      Y1              =   480
      Y2              =   7680
   End
   Begin VB.Line Line4 
      X1              =   9480
      X2              =   17640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   13320
      X2              =   13320
      Y1              =   480
      Y2              =   7680
   End
   Begin VB.Label ltr_r3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   45
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label ltr_r2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   44
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label ltr_s 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   43
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label lti_r3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   42
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lti_r2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   41
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lti_s 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   40
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Line Line2 
      X1              =   15480
      X2              =   15480
      Y1              =   480
      Y2              =   7680
   End
   Begin VB.Line Line1 
      X1              =   17640
      X2              =   17640
      Y1              =   480
      Y2              =   7680
   End
   Begin VB.Label lc_r3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   39
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lc_r2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   38
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lc_s 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   37
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label ltotal_tr 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   36
      Top             =   7200
      Width           =   1740
   End
   Begin VB.Label ltotal_ti 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   35
      Top             =   7200
      Width           =   1740
   End
   Begin VB.Label ltotal_c 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   34
      Top             =   7200
      Width           =   1785
   End
   Begin VB.Label Label74 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      Left            =   9960
      TabIndex        =   33
      Top             =   7200
      Width           =   765
   End
   Begin VB.Line Line17 
      X1              =   1680
      X2              =   9240
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label73 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Area"
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
      Left            =   1920
      TabIndex        =   32
      Top             =   600
      Width           =   705
   End
   Begin VB.Label Label72 
      AutoSize        =   -1  'True
      Caption         =   "% Complete"
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
      Left            =   7320
      TabIndex        =   31
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label Label71 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Input"
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
      Left            =   5160
      TabIndex        =   30
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Released"
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
      Left            =   3000
      TabIndex        =   29
      Top             =   600
      Width           =   1770
   End
   Begin VB.Label Label69 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "PT. HIDROFLEX"
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
      Left            =   4425
      TabIndex        =   28
      Top             =   0
      Width           =   2145
   End
   Begin VB.Label Label68 
      AutoSize        =   -1  'True
      Caption         =   "A"
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
      Left            =   2160
      TabIndex        =   27
      Top             =   1080
      Width           =   165
   End
   Begin VB.Label Label65 
      AutoSize        =   -1  'True
      Caption         =   "B"
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
      Left            =   2160
      TabIndex        =   26
      Top             =   1560
      Width           =   165
   End
   Begin VB.Label Label64 
      AutoSize        =   -1  'True
      Caption         =   "R1"
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
      Left            =   10080
      TabIndex        =   25
      Top             =   1560
      Width           =   315
   End
   Begin VB.Line Line16 
      X1              =   1680
      X2              =   1680
      Y1              =   480
      Y2              =   6840
   End
   Begin VB.Label Label61 
      AutoSize        =   -1  'True
      Caption         =   "C"
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
      Left            =   2160
      TabIndex        =   24
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label Label60 
      AutoSize        =   -1  'True
      Caption         =   "D"
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
      Left            =   2160
      TabIndex        =   23
      Top             =   2520
      Width           =   180
   End
   Begin VB.Label Label59 
      Caption         =   "E"
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
      Left            =   2160
      TabIndex        =   22
      Top             =   3000
      Width           =   240
   End
   Begin VB.Label Label57 
      AutoSize        =   -1  'True
      Caption         =   "F"
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
      Left            =   2160
      TabIndex        =   21
      Top             =   3480
      Width           =   150
   End
   Begin VB.Line Line15 
      X1              =   1680
      X2              =   9240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line14 
      X1              =   2760
      X2              =   2760
      Y1              =   480
      Y2              =   6840
   End
   Begin VB.Line Line13 
      X1              =   1680
      X2              =   9240
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line12 
      X1              =   4920
      X2              =   4920
      Y1              =   480
      Y2              =   6840
   End
   Begin VB.Label ltr_e 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   20
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label ltr_d 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   19
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label ltr_c 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   18
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label ltr_r1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   11400
      TabIndex        =   17
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label ltr_b 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   16
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label ltr_a 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   15
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label ltr_f 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   3000
      TabIndex        =   14
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label lti_e 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   13
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lti_d 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   12
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lti_c 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   11
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lti_r1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   10
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lti_b 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   9
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lti_a 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   8
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label lti_f 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5160
      TabIndex        =   7
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Line Line11 
      X1              =   7080
      X2              =   7080
      Y1              =   480
      Y2              =   6840
   End
   Begin VB.Line Line10 
      X1              =   9240
      X2              =   9240
      Y1              =   480
      Y2              =   6840
   End
   Begin VB.Label lc_e 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   6
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lc_d 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   5
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lc_c 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   4
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lc_r1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   15720
      TabIndex        =   3
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lc_b 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   2
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lc_a 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   1
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label lc_f 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   7320
      TabIndex        =   0
      Top             =   3480
      Width           =   1740
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub current_progress_hidroflex()
    'Gudang A
    strsql = "select count(tag_no) AS input_a from tag_stock_opname where left(tag_no,3)='TRA' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_a.Caption = Val(rs_so!input_a)
    strsql = "select count(tag_no) AS release_a from tag_stock_opname where left(tag_no,3)='TRA'"
    Set rs_so = conn.Execute(strsql)
    ltr_a.Caption = Val(rs_so!release_a)
    If Val(rs_so!release_a) = 0 Then lc_a.Caption = 0 Else _
    lc_a.Caption = Round((Val(lti_a.Caption) / Val(ltr_a.Caption) * 100), 2)
    
    'Gudang B
    strsql = "select count(tag_no) AS input_b from tag_stock_opname where left(tag_no,3)='TRB' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_b.Caption = Val(rs_so!input_b)
    strsql = "select count(tag_no) AS release_b from tag_stock_opname where left(tag_no,3)='TRB'"
    Set rs_so = conn.Execute(strsql)
    ltr_b.Caption = Val(rs_so!release_b)
    If Val(rs_so!release_b) = 0 Then lc_b.Caption = 0 Else _
    lc_b.Caption = Round((Val(lti_b.Caption) / Val(ltr_b.Caption) * 100), 2)
    
    'Gudang C
    strsql = "select count(tag_no) AS input_c from tag_stock_opname where left(tag_no,3)='TRC' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_c.Caption = Val(rs_so!input_c)
    strsql = "select count(tag_no) AS release_c from tag_stock_opname where left(tag_no,3)='TRC'"
    Set rs_so = conn.Execute(strsql)
    ltr_c.Caption = Val(rs_so!release_c)
    If Val(rs_so!release_c) = 0 Then lc_c.Caption = 0 Else _
    lc_c.Caption = Round((Val(lti_c.Caption) / Val(ltr_c.Caption) * 100), 2)
    
    'Gudang D
    strsql = "select count(tag_no) AS input_d from tag_stock_opname where left(tag_no,3)='TRD' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_d.Caption = Val(rs_so!input_d)
    strsql = "select count(tag_no) AS release_d from tag_stock_opname where left(tag_no,3)='TRD'"
    Set rs_so = conn.Execute(strsql)
    ltr_d.Caption = Val(rs_so!release_d)
    If Val(rs_so!release_d) = 0 Then lc_d.Caption = 0 Else _
    lc_d.Caption = Round((Val(lti_d.Caption) / Val(ltr_d.Caption) * 100), 2)
    
    'Gudang E
    strsql = "select count(tag_no) AS input_e from tag_stock_opname where left(tag_no,3)='TRE' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_e.Caption = Val(rs_so!input_e)
    strsql = "select count(tag_no) AS release_e from tag_stock_opname where left(tag_no,3)='TRE'"
    Set rs_so = conn.Execute(strsql)
    ltr_e.Caption = Val(rs_so!release_e)
    If Val(rs_so!release_e) = 0 Then lc_e.Caption = 0 Else _
    lc_e.Caption = Round((Val(lti_e.Caption) / Val(ltr_e.Caption) * 100), 2)
    
    'Gudang F
    strsql = "select count(tag_no) AS input_F from tag_stock_opname where left(tag_no,3)='TRF' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_f.Caption = Val(rs_so!input_F)
    strsql = "select count(tag_no) AS release_f from tag_stock_opname where left(tag_no,3)='TRF'"
    Set rs_so = conn.Execute(strsql)
    ltr_f.Caption = Val(rs_so!release_f)
    If Val(rs_so!release_f) = 0 Then lc_f.Caption = 0 Else _
    lc_f.Caption = Round((Val(lti_f.Caption) / Val(ltr_f.Caption) * 100), 2)
    
    'Gudang H
    strsql = "select count(tag_no) AS input_h from tag_stock_opname where (left(tag_no,3)='TRH' OR left(tag_no,3)='TRP') and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_h.Caption = Val(rs_so!input_h)
    strsql = "select count(tag_no) AS release_h from tag_stock_opname where (left(tag_no,3)='TRH' OR left(tag_no,3)='TRP')"
    Set rs_so = conn.Execute(strsql)
    ltr_h.Caption = Val(rs_so!release_h)
    If Val(rs_so!release_h) = 0 Then lc_h.Caption = 0 Else _
    lc_h.Caption = Round((Val(lti_h.Caption) / Val(ltr_h.Caption) * 100), 2)
    
    'Gudang i
    strsql = "select count(tag_no) AS input_i from tag_stock_opname where left(tag_no,3)='TRI' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_i.Caption = Val(rs_so!input_i)
    strsql = "select count(tag_no) AS release_i from tag_stock_opname where left(tag_no,3)='TRI'"
    Set rs_so = conn.Execute(strsql)
    ltr_i.Caption = Val(rs_so!release_i)
    If Val(rs_so!release_i) = 0 Then lc_i.Caption = 0 Else _
    lc_i.Caption = Round((Val(lti_i.Caption) / Val(ltr_i.Caption) * 100), 2)
    
    'Gudang J
    strsql = "select count(tag_no) AS input_j from tag_stock_opname where left(tag_no,3)='TRJ' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_j.Caption = Val(rs_so!input_j)
    strsql = "select count(tag_no) AS release_j from tag_stock_opname where left(tag_no,3)='TRJ'"
    Set rs_so = conn.Execute(strsql)
    ltr_j.Caption = Val(rs_so!release_j)
    If Val(rs_so!release_j) = 0 Then lc_j.Caption = 0 Else _
    lc_j.Caption = Round((Val(lti_j.Caption) / Val(ltr_j.Caption) * 100), 2)
    
    'Gudang T
    strsql = "select count(tag_no) AS input_t from tag_stock_opname where left(tag_no,3)='TRT' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_t.Caption = Val(rs_so!input_t)
    strsql = "select count(tag_no) AS release_t from tag_stock_opname where left(tag_no,3)='TRT'"
    Set rs_so = conn.Execute(strsql)
    ltr_t.Caption = Val(rs_so!release_t)
    If Val(rs_so!release_t) = 0 Then lc_t.Caption = 0 Else _
    lc_t.Caption = Round((Val(lti_t.Caption) / Val(ltr_t.Caption) * 100), 2)
    
    'Gudang u
    strsql = "select count(tag_no) AS input_u from tag_stock_opname where left(tag_no,3)='TRU' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_u.Caption = Val(rs_so!input_u)
    strsql = "select count(tag_no) AS release_u from tag_stock_opname where left(tag_no,3)='TRU'"
    Set rs_so = conn.Execute(strsql)
    ltr_u.Caption = Val(rs_so!release_u)
    If Val(rs_so!release_u) = 0 Then lc_u.Caption = 0 Else _
    lc_u.Caption = Round((Val(lti_u.Caption) / Val(ltr_u.Caption) * 100), 2)
    
    'Gudang S
    strsql = "select count(tag_no) AS input_s from tag_stock_opname where left(tag_no,3)='TRS' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_s.Caption = Val(rs_so!input_s)
    strsql = "select count(tag_no) AS release_s from tag_stock_opname where left(tag_no,3)='TRS'"
    Set rs_so = conn.Execute(strsql)
    ltr_s.Caption = Val(rs_so!release_s)
    If Val(rs_so!release_s) = 0 Then lc_s.Caption = 0 Else _
    lc_s.Caption = Round((Val(lti_s.Caption) / Val(ltr_s.Caption) * 100), 2)
    
    'Gudang R3
    strsql = "select count(tag_no) AS input_r3 from tag_stock_opname where left(tag_no,3)='TR3' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_r3.Caption = Val(rs_so!input_r3)
    strsql = "select count(tag_no) AS release_r3 from tag_stock_opname where left(tag_no,3)='TR3'"
    Set rs_so = conn.Execute(strsql)
    ltr_r3.Caption = Val(rs_so!release_r3)
    If Val(rs_so!release_r3) = 0 Then lc_r3.Caption = 0 Else _
    lc_r3.Caption = Round((Val(lti_r3.Caption) / Val(ltr_r3.Caption) * 100), 2)
    
    'Gudang R1
    strsql = "select count(tag_no) AS input_r1 from tag_stock_opname where left(tag_no,3)='TSR' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_r1.Caption = Val(rs_so!input_r1)
    strsql = "select count(tag_no) AS release_r1 from tag_stock_opname where left(tag_no,3)='TSR'"
    Set rs_so = conn.Execute(strsql)
    ltr_r1.Caption = Val(rs_so!release_r1)
    If Val(rs_so!release_r1) = 0 Then lc_r1.Caption = 0 Else _
    lc_r1.Caption = Round((Val(lti_r1.Caption) / Val(ltr_r1.Caption) * 100), 2)
    
    'Gudang R2
    strsql = "select count(tag_no) AS input_r2 from tag_stock_opname where left(tag_no,3)='TR2' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_r2.Caption = Val(rs_so!input_r2)
    strsql = "select count(tag_no) AS release_r2 from tag_stock_opname where left(tag_no,3)='TR2'"
    Set rs_so = conn.Execute(strsql)
    ltr_r2.Caption = Val(rs_so!release_r2)
    If Val(rs_so!release_r2) = 0 Then lc_r2.Caption = 0 Else _
    lc_r2.Caption = Round((Val(lti_r2.Caption) / Val(ltr_r2.Caption) * 100), 2)
    
    'Gudang FG
    strsql = "select count(tag_no) AS input_rl from tag_stock_opname where left(tag_no,3)='TFG' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_rluar.Caption = Val(rs_so!input_rl)
    strsql = "select count(tag_no) AS release_rl from tag_stock_opname where left(tag_no,3)='TFG'"
    Set rs_so = conn.Execute(strsql)
    ltr_rluar.Caption = Val(rs_so!release_rl)
    If Val(rs_so!release_rl) = 0 Then lc_rluar.Caption = 0 Else _
    lc_rluar.Caption = Round((Val(lti_rluar.Caption) / Val(ltr_rluar.Caption) * 100), 2)
    
    'Gudang K
    strsql = "select count(tag_no) AS input_k from tag_stock_opname where left(tag_no,3)='TSK' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_k.Caption = Val(rs_so!input_k)
    strsql = "select count(tag_no) AS release_k from tag_stock_opname where left(tag_no,3)='TSK'"
    Set rs_so = conn.Execute(strsql)
    ltr_k.Caption = Val(rs_so!release_k)
    If Val(rs_so!release_k) = 0 Then lc_k.Caption = 0 Else _
    lc_k.Caption = Round((Val(lti_k.Caption) / Val(ltr_k.Caption) * 100), 2)
    
'    'Gudang Container
'    strsql = "select count(tag_no) AS input_co from tag_stock_opname where left(tag_no,3)='TCC' and status='OK'"
'    Set rs_so = conn.Execute(strsql)
'    lti_container.Caption = Val(rs_so!input_co)
'    strsql = "select count(tag_no) AS release_co from tag_stock_opname where left(tag_no,3)='TCC'"
'    Set rs_so = conn.Execute(strsql)
'    ltr_container.Caption = Val(rs_so!release_co)
'    If Val(rs_so!release_co) = 0 Then lc_container.Caption = 0 Else _
'    lc_container.Caption = Round((Val(lti_container.Caption) / Val(ltr_container.Caption) * 100), 2)
    
    'Gudang Closed Room
    strsql = "select count(tag_no) AS input_csd from tag_stock_opname where left(tag_no,3)='TSC' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_csd_room.Caption = Val(rs_so!input_csd)
    strsql = "select count(tag_no) AS release_csd from tag_stock_opname where left(tag_no,3)='TSC'"
    Set rs_so = conn.Execute(strsql)
    ltr_csd_room.Caption = Val(rs_so!release_csd)
    If Val(rs_so!release_csd) = 0 Then lc_csd_room.Caption = 0 Else _
    lc_csd_room.Caption = Round((Val(lti_csd_room.Caption) / Val(ltr_csd_room.Caption) * 100), 2)
    
    'Gudang Main WH
    strsql = "select count(tag_no) AS input_mwh from tag_stock_opname where left(tag_no,3)='TCM' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_wh_main.Caption = Val(rs_so!input_mwh)
    strsql = "select count(tag_no) AS release_mwh from tag_stock_opname where left(tag_no,3)='TCM'"
    Set rs_so = conn.Execute(strsql)
    ltr_wh_main.Caption = Val(rs_so!release_mwh)
    If Val(rs_so!release_mwh) = 0 Then lc_wh_main.Caption = 0 Else _
    lc_wh_main.Caption = Round((Val(lti_wh_main.Caption) / Val(ltr_wh_main.Caption) * 100), 2)
    
    'Gudang Gases
    strsql = "select count(tag_no) AS input_g from tag_stock_opname where left(tag_no,3)='TCG' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_g.Caption = Val(rs_so!input_g)
    strsql = "select count(tag_no) AS release_g from tag_stock_opname where left(tag_no,3)='TCG'"
    Set rs_so = conn.Execute(strsql)
    ltr_g.Caption = Val(rs_so!release_g)
    If Val(rs_so!release_g) = 0 Then lc_g.Caption = 0 Else _
    lc_g.Caption = Round((Val(lti_g.Caption) / Val(ltr_g.Caption) * 100), 2)
    
    'Gudang Cont. Cons
    strsql = "select count(tag_no) AS input_cc from tag_stock_opname where left(tag_no,3)='TCC' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_cc.Caption = Val(rs_so!input_cc)
    strsql = "select count(tag_no) AS release_cc from tag_stock_opname where left(tag_no,3)='TCC'"
    Set rs_so = conn.Execute(strsql)
    ltr_cc.Caption = Val(rs_so!release_cc)
    If Val(rs_so!release_cc) = 0 Then lc_cc.Caption = 0 Else _
    lc_cc.Caption = Round((Val(lti_cc.Caption) / Val(ltr_cc.Caption) * 100), 2)
    
    'Gudang Asset
    strsql = "select count(tag_no) AS input_asset from tag_stock_opname where left(tag_no,3)='TAS' and status='OK'"
    Set rs_so = conn.Execute(strsql)
    lti_asset.Caption = Val(rs_so!input_asset)
    strsql = "select count(tag_no) AS release_asset from tag_stock_opname where left(tag_no,3)='TAS'"
    Set rs_so = conn.Execute(strsql)
    ltr_asset.Caption = Val(rs_so!release_asset)
    If Val(rs_so!release_asset) = 0 Then lc_asset.Caption = 0 Else _
    lc_asset.Caption = Round((Val(lti_asset.Caption) / Val(ltr_asset.Caption) * 100), 2)
     
    'TAG BLANK
    qry_input_tb = "select count(tag_no) AS input_tb from tag_stock_opname where left(tag_no,2)='TB' and status='OK'"
    Set rs_so = conn.Execute(qry_input_tb)
    lti_blank.Caption = Val(rs_so!input_tb)
    qry_release_tb = "select count(tag_no) AS release_TB from tag_stock_opname where left(tag_no,2)='TB'"
    Set rs_so = conn.Execute(qry_release_tb)
    ltr_blank.Caption = Val(rs_so!release_TB)
    If Val(rs_so!release_TB) = 0 Then lc_blank.Caption = 0 Else _
    lc_blank.Caption = Round((Val(lti_blank.Caption) / Val(ltr_blank.Caption) * 100), 2)
    
    'Total
    ltotal_tr.Caption = Val(ltr_a.Caption) + Val(ltr_b.Caption) _
        + Val(ltr_c.Caption) + Val(ltr_d.Caption) + Val(ltr_e.Caption) _
        + Val(ltr_f.Caption) + Val(ltr_h.Caption) + Val(ltr_i.Caption) + Val(ltr_j.Caption) + _
        Val(ltr_t.Caption) + Val(ltr_u.Caption) + Val(ltr_s.Caption) + Val(ltr_r3.Caption) + _
        Val(ltr_r1.Caption) + Val(ltr_r2.Caption) + Val(ltr_rluar.Caption) + Val(ltr_k.Caption) + _
         Val(ltr_csd_room.Caption) + Val(ltr_wh_main.Caption) + _
        Val(ltr_g.Caption) + Val(ltr_cc.Caption) + Val(ltr_asset.Caption) + Val(ltr_blank.Caption)
    ltotal_ti.Caption = Val(lti_a.Caption) + Val(lti_b.Caption) _
        + Val(lti_c.Caption) + Val(lti_d.Caption) + Val(lti_e.Caption) _
        + Val(lti_f.Caption) + Val(lti_h.Caption) + Val(lti_i.Caption) + Val(lti_j.Caption) + _
        Val(lti_t.Caption) + Val(lti_u.Caption) + Val(lti_s.Caption) + Val(lti_r3.Caption) + _
        Val(lti_r1.Caption) + Val(lti_r2.Caption) + Val(lti_rluar.Caption) + Val(lti_k.Caption) + _
        Val(lti_csd_room.Caption) + Val(lti_wh_main.Caption) + _
        Val(lti_g.Caption) + Val(lti_cc.Caption) + Val(lti_asset.Caption) + Val(lti_blank.Caption)
    If Val(ltotal_ti.Caption) = 0 Then ltotal_c.Caption = 0 Else _
        ltotal_c.Caption = Round(((Val(ltotal_ti) / Val(ltotal_tr) * 100)), 2)
    
End Sub

Private Sub Command1_Click()
    Printer.PaperSize = vbPRPSLegal
    Printer.Orientation = vbPRORLandscape
    Form2.PrintForm
End Sub

Private Sub Command2_Click()
    Call current_progress_hidroflex
End Sub

Private Sub Form_Load()
Timer1.Interval = 500
Timer1.Enabled = True
Call db
If rs_so.State = 1 Then rs_so.Close
rs_so.Open "Select * from tag_stock_opname", conn
Set rscompletion_slip = Nothing
Call current_progress_hidroflex
lbl_tanggal = Date
End Sub

Private Sub Timer1_Timer()
lbl_jam.Caption = Format(Time, "hh:mm:ss")
End Sub
