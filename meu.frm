VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Realtime Image Rotation"
   ClientHeight    =   8310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   568
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll4 
      Height          =   240
      Left            =   1125
      Max             =   255
      Min             =   10
      TabIndex        =   14
      Top             =   7725
      Value           =   255
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000FF&
      ForeColor       =   &H00400040&
      Height          =   375
      Left            =   150
      OLEDropMode     =   1  'Manual
      TabIndex        =   13
      Text            =   "Drag to Here"
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7620
      Top             =   7590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   270
      Index           =   2
      Left            =   4470
      Max             =   360
      TabIndex        =   11
      Top             =   7425
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   270
      Index           =   1
      Left            =   4470
      Max             =   360
      TabIndex        =   9
      Top             =   7140
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   270
      Index           =   0
      Left            =   4470
      Max             =   360
      TabIndex        =   4
      Top             =   6825
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   270
      Left            =   615
      Max             =   300
      Min             =   10
      TabIndex        =   3
      Top             =   7410
      Value           =   200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   270
      Left            =   615
      Max             =   100
      TabIndex        =   2
      Top             =   7110
      Value           =   100
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5115
      Left            =   2295
      Picture         =   "meu.frx":0000
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   1
      Top             =   1245
      Visible         =   0   'False
      Width           =   4185
   End
   Begin VB.HScrollBar vScroll1 
      Height          =   270
      Left            =   615
      Max             =   360
      TabIndex        =   0
      Top             =   6810
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Translucency"
      Height          =   225
      Index           =   6
      Left            =   60
      TabIndex        =   15
      Top             =   7710
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Lum"
      Height          =   225
      Index           =   5
      Left            =   3990
      TabIndex        =   12
      Top             =   7455
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Sat"
      Height          =   225
      Index           =   4
      Left            =   3975
      TabIndex        =   10
      Top             =   7170
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Hue"
      Height          =   225
      Index           =   3
      Left            =   3975
      TabIndex        =   8
      Top             =   6855
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Scale"
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   7425
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Alpha"
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   6
      Top             =   7155
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Angle"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Top             =   6840
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HTCAPTION As Integer = 2
Private Const WM_NCLBUTTONDOWN As Integer = &HA1

Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const LWA_ALPHA As Long = &H2&
Private Const LWA_COLORKEY As Integer = &H1

Private Key_Press As Byte
Private Incr As Byte

Private Const Rad As Currency = 1.74532925199433E-02
Private Const Pi As Currency = 3.14159265358979

Private Const WM_MOUSEWHEEL       As Long = &H20A
Private Const WM_WINDOWPOSCHANGED As Long = &H47
Private sc          As cSuperClass
Implements iSuperClass
Private Declare Function rotatedc Lib "Rotate.Lib" Alias "rotatedc@60" (ByVal aHDC As Long, ByVal Angle As Single, ByVal x As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, ByVal PicDC As Long, Optional ByVal SrcX As Long = 0, Optional ByVal SrcY As Long = 0, Optional ByVal pScale As Single = 1, Optional ByVal TraspColor As Long = -1, Optional ByVal Alpha As Single = 1, Optional ByVal Hue As Single = 0, Optional ByVal Sat As Single = 0, Optional ByVal Lum As Single = 0) As Long
Private T As Long

Private Sub Form_DblClick()
    On Error GoTo erro
    
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "All Supported Images|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur|Bitmaps|*.bmp;*.dib|JPEG Images|*.jpg|Metafiles|*.wmf;*.emf|Icons|*.ico;*.cur"
    CommonDialog1.ShowOpen
        
    If CommonDialog1.FileName <> "" Then
        S.Picture = LoadPicture(CommonDialog1.FileName)
        Draw
    End If

exit_sub:
    Exit Sub
    
erro:
    Resume exit_sub
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Key_Press = KeyCode
    
    Select Case KeyCode
      Case 112
        Incr = 1
      Case 113
        Incr = 5
      Case 114
        Incr = 10
      Case 115
        Text1.Visible = Text1.Visible Xor -1
      Case 116
        Dim x As New Form1
        x.Show
      Case 27
        End
    End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Key_Press = 0

End Sub

Private Sub Form_Load()

  Dim Ret As Long
  Dim sSave As String
  Dim x() As Byte
    
    Form2.Show
    
    Incr = 10
    sSave = Space(255)

    Ret = GetSystemDirectory(sSave, 255)

    sSave = Left$(sSave, Ret)

    If Dir(sSave & "\Rotate.Lib") = "" Then
        x = LoadResData(101, "CUSTOM")
        Open sSave & "\Rotate.Lib" For Binary As 1
        Put #1, 1, x
        Close 1
    End If

    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Me.hWnd, 0, 255, LWA_COLORKEY Or LWA_ALPHA

    Set sc = New cSuperClass
    With sc
        .AddMsg WM_MOUSEWHEEL
        '.AddMsg WM_WINDOWPOSCHANGED
        .Subclass hWnd, Me
    End With

    T = &HFAEED0

    Draw

End Sub

Private Sub Draw()

    Me.Cls
    rotatedc Me.hDC, 2 * Pi - vScroll1.Value * Rad, Me.ScaleWidth / 2, Me.ScaleHeight / 2, S.Width, S.Height, S.hDC, 0, 0, HScroll2.Value / 200, T, HScroll1.Value / 100, HScroll3(0).Value / 360, HScroll3(1).Value / 360, HScroll3(2).Value / 360
    Me.Refresh

End Sub

Private Sub me_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Button = 2 Then
        T = Me.Point(x, Y)
        Command2.BackColor = T
        Draw
    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Button = 2 Then
        T = Point(x, Y)
        Draw
        Exit Sub
    End If

    ReleaseCapture

    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0

End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Data.Files.Count > 0 Then
        Select Case LCase(Right(Data.Files(1), 3))
          Case "bmp", "dib", "gif", "jpg", "wmf", "emf", "ico", "cur"
            Text1.Text = Data.Files(1)
            S.Picture = LoadPicture(Text1.Text)
            Draw
        End Select
    End If

End Sub

Private Sub HScroll4_Change()

    SetLayeredWindowAttributes Me.hWnd, 0, HScroll4, LWA_COLORKEY Or LWA_ALPHA

End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Data.Files.Count > 0 Then
        Select Case LCase(Right(Data.Files(1), 3))
          Case "bmp", "dib", "gif", "jpg", "wmf", "emf", "ico", "cur"
            S.Picture = LoadPicture(Data.Files(1))
            Draw
        End Select
    End If

End Sub

Private Sub HScroll1_Change()

    Draw

End Sub

Private Sub HScroll1_Scroll()

    Draw

End Sub

Private Sub HScroll2_Change()

    Draw

End Sub

Private Sub HScroll2_Scroll()

    Draw

End Sub

Private Sub vScroll1_Change()

    Draw

End Sub

Private Sub vScroll1_Scroll()

    Draw

End Sub

Private Sub HScroll3_Change(Index As Integer)

    Draw

End Sub

Private Sub HScroll3_Scroll(Index As Integer)

    Draw

End Sub

Private Sub iSuperClass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

  Dim x As Integer

    Select Case uMsg
     'Case WM_WINDOWPOSCHANGED

      Case WM_MOUSEWHEEL
        
        Select Case wParam
          
          Case 7864320 ' UP
            
            Select Case Key_Press
              
              Case 72 '+H Key = Hue
                x = HScroll3(0)
                x = x + Incr
                If x < HScroll3(0).Max Then
                    HScroll3(0) = x
                End If
                Exit Sub
                
              Case 83 '+S Key = Saturation
                x = HScroll3(1)
                x = x + Incr
                If x < HScroll3(1).Max Then
                    HScroll3(1) = x
                End If
                Exit Sub

              Case 76 '+L Key = Luminance
                x = HScroll3(2)
                x = x + Incr
                If x < HScroll3(2).Max Then
                    HScroll3(2) = x
                End If
                Exit Sub

              Case 67 '+C Key = Luminance
                x = HScroll1
                x = x + Incr
                If x < HScroll1.Max Then
                    HScroll1 = x
                End If
                Exit Sub
            
            End Select
            
            'ROTATION TO RIGHT
            x = vScroll1
            x = x + Incr
            If x > vScroll1.Max Then
                vScroll1 = vScroll1.Min + Incr
              Else
                vScroll1 = x
            End If
            
          Case -7864320 'DOWN
            
            Select Case Key_Press
              
              Case 72 '+H Key = Hue
                x = HScroll3(0)
                x = x - Incr
                If x > HScroll3(0).Min Then
                    HScroll3(0) = x
                End If
                Exit Sub

              Case 83 '+S Key= Saturation
                x = HScroll3(1)
                x = x - Incr
                If x > HScroll3(1).Min Then
                    HScroll3(1) = x
                End If
                Exit Sub

              Case 76 '+L Key= Luminance
                x = HScroll3(2)
                x = x - Incr
                If x > HScroll3(2).Min Then
                    HScroll3(2) = x
                End If
                Exit Sub

              Case 67 '+C Key = Contrast
                x = HScroll1
                x = x - Incr
                If x > HScroll1.Min Then
                    HScroll1 = x
                End If
                Exit Sub

            End Select
            
            'ROTATION TO LEFT
            x = vScroll1
            x = x - Incr
            If x < vScroll1.Min Then
                vScroll1 = vScroll1.Max - Incr
              Else
                vScroll1 = x
            End If

          Case 7864328 'UP + CTRL = +TRANSLUCENCY
            x = HScroll4
            x = x + Incr
            If x < HScroll4.Max Then
                HScroll4 = x
            End If

          Case -7864312 'DOWN + CRTL = -TRANSLUCENCY
            x = HScroll4
            x = x - Incr
            If x > HScroll4.Min Then
                HScroll4 = x
            End If

          Case -7864316 'UP + SHIFT = + SCALE
            x = HScroll2
            x = x + Incr
            If x < HScroll2.Max Then
                HScroll2 = x
            End If
          Case 7864324 'DOWN + SHIFT = - SCALE
            x = HScroll2
            x = x - Incr
            If x > HScroll2.Min Then
                HScroll2 = x
            End If
        End Select

    End Select

End Sub

