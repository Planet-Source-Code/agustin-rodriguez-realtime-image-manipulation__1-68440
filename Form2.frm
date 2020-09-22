VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Help"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "Form2"
   ScaleHeight     =   5925
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK with BUTTON 2 to set TRANSPARENT COLOR"
      Height          =   195
      Index           =   7
      Left            =   135
      TabIndex        =   15
      Top             =   4185
      Width           =   3930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESC Key to END"
      Height          =   195
      Index           =   15
      Left            =   150
      TabIndex        =   14
      Top             =   5535
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F4 Key = SHOW/HIDE OLEDragDrop FEATURE"
      Height          =   195
      Index           =   14
      Left            =   120
      TabIndex        =   13
      Top             =   3195
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DRAG the Objeto to MOVE"
      Height          =   195
      Index           =   13
      Left            =   135
      TabIndex        =   12
      Top             =   5085
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOUBLE CLICK to OPEN NEW IMAGE"
      Height          =   195
      Index           =   12
      Left            =   105
      TabIndex        =   11
      Top             =   4620
      Width           =   2820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F5 Key = CREATE NEW INSTANCE"
      Height          =   195
      Index           =   11
      Left            =   135
      TabIndex        =   10
      Top             =   3690
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F3 Key = Step 10"
      Height          =   195
      Index           =   10
      Left            =   2985
      TabIndex        =   9
      Top             =   2745
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F2 Key = Step 5"
      Height          =   195
      Index           =   9
      Left            =   1545
      TabIndex        =   8
      Top             =   2745
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F1 Key = Step 1 "
      Height          =   195
      Index           =   8
      Left            =   105
      TabIndex        =   7
      Top             =   2730
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+ H Key = HUE"
      Height          =   195
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+ S Key = SATURATION"
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Top             =   2205
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+ L Key = LUMINANCE"
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   1890
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+ C Key = CONTRAST"
      Height          =   195
      Index           =   5
      Left            =   1440
      TabIndex        =   3
      Top             =   1230
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+ CRTL = TRANSLUCENCY"
      Height          =   195
      Index           =   4
      Left            =   1440
      TabIndex        =   2
      Top             =   945
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+ SHIFT = SCALE"
      Height          =   195
      Index           =   2
      Left            =   1440
      TabIndex        =   1
      Top             =   660
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOUSE WHEEL Alone       =  ROTATION"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   345
      Width           =   3000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
