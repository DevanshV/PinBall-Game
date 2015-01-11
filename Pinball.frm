VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Pinball 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   2820
   ClientTop       =   1395
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   14865
   Begin VB.Timer Paddle 
      Interval        =   1000
      Left            =   120
      Top             =   4920
   End
   Begin VB.Timer ColorChange 
      Interval        =   1000
      Left            =   120
      Top             =   5400
   End
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   5880
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   6360
   End
   Begin VB.Label lblPrize 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   30
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label BlueTR 
      BackColor       =   &H00C0C000&
      Height          =   1575
      Left            =   12720
      TabIndex        =   28
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label BlueTL 
      BackColor       =   &H00C0C000&
      Height          =   1575
      Left            =   1320
      TabIndex        =   27
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label BlueM 
      BackColor       =   &H00C0C000&
      Height          =   375
      Left            =   6240
      TabIndex        =   26
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblPoints 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   7320
      TabIndex        =   25
      Top             =   3120
      Width           =   855
   End
   Begin VB.Shape eyeright 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape eyeleft 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label SB2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   10440
      TabIndex        =   24
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label SB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label RRight 
      BackColor       =   &H000000FF&
      Height          =   1215
      Left            =   10200
      TabIndex        =   22
      Top             =   960
      Width           =   375
   End
   Begin VB.Label RLeft 
      BackColor       =   &H000000FF&
      Height          =   1215
      Left            =   4560
      TabIndex        =   21
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblNow 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOW"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6840
      TabIndex        =   20
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblleave 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LEAVE"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6840
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblEYER 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7800
      TabIndex        =   18
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblEYEL 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label PrizeL 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label YSix 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9480
      TabIndex        =   15
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label YFive 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8880
      TabIndex        =   14
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label YFour 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Ythree 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Ytwo 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Yone 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label PrizeR 
      BackColor       =   &H000000FF&
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   10560
      TabIndex        =   9
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label LeftEndwall 
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   7455
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   375
   End
   Begin VB.Line PaddleRight 
      BorderColor     =   &H000000C0&
      BorderWidth     =   5
      X1              =   7800
      X2              =   10080
      Y1              =   8040
      Y2              =   7440
   End
   Begin VB.Line PaddleLeft 
      BorderColor     =   &H000000C0&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      X1              =   4680
      X2              =   7080
      Y1              =   7440
      Y2              =   8040
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   2000
      Y2              =   0
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10800
      TabIndex        =   5
      Top             =   7320
      Width           =   480
   End
   Begin VB.Label lblSeconds 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12840
      TabIndex        =   4
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   12240
      TabIndex        =   3
      Top             =   7320
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   330
   End
   Begin VB.Shape ball 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   11400
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label topwall 
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   14655
   End
   Begin VB.Label RightEndwall 
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   7455
      Left            =   14400
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Label RightBlock 
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   8
      Top             =   7320
      Width           =   4215
   End
   Begin VB.Label LeftBlock 
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   7320
      Width           =   4095
   End
   Begin WMPLibCtl.WindowsMediaPlayer Music 
      Height          =   360
      Left            =   3480
      TabIndex        =   29
      Top             =   8280
      Visible         =   0   'False
      Width           =   540
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   953
      _cy             =   635
   End
End
Attribute VB_Name = "Pinball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private movex, movey As Integer   'ball movement horizontally and vertically
Private howfar As Integer
Private PaddleLeftPosition As Integer
Private PaddleRightPosition As Integer
Private Points As Integer


Private Sub ColorChange_Timer()

    If ball.FillColor = vbGreen Then
    ball.FillColor = vbBlue
    ElseIf ball.FillColor = vbBlue Then
    ball.FillColor = vbMagenta
    ElseIf ball.FillColor = vbGreen Then
    ball.FillColor = vbCyan
    ElseIf ball.FillColor = vbWhite Then
    ball.FillColor = vbGreen
    ElseIf ball.FillColor = vbRed Then
    ball.FillColor = vbBlue
    ElseIf ball.FillColor = vbMagenta Then
    ball.FillColor = vbGreen
    ElseIf ball.FillColor = vbCyan Then
    ball.FillColor = vbWhite
    ElseIf ball.FillColor = vbRed Then
    ball.FillColor = vbGreen
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyLeft
    
    'for
        PaddleLeft.Y2 = PaddleLeft.Y1
        PaddleLeftPosition = 1
    Case vbKeyRight
        PaddleRight.Y1 = PaddleRight.Y2
        PaddleRightPosition = 1
End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Paddlemovement = 0
PaddleLeft.Y2 = 8040
PaddleRight.Y1 = 8040
PaddleLeftPosition = 0
PaddleRightPosition = 0
End Sub

Private Sub Form_Load()
ball.FillColor = vbRed
Music.URL = "C:\Users\User\Documents\Final\epicmusic.mp3"
Music.Controls.play
movex = 0
movey = 0
End Sub

Private Sub Label3_Click()

End Sub

Private Sub lblStart_Click()

Dim random As Integer
random = Int(2 * Rnd) + 1

lblPrize.Caption = ""
lblSeconds.Caption = ""
 
If random = 1 Then
    ball.Left = 2400
Else
    ball.Left = 12240
End If

ball.Top = Int(501 * Rnd) + 5000

movex = 75
movey = 75
tmrClock.Enabled = True
Timer1.Enabled = True
Points = 0

End Sub


Private Sub obstacle_Click(Index As Integer)

End Sub

Private Sub Timer1_Timer()


lblPoints = Points

Call LeftPaddle

Call RightPaddle

If ball.Top >= 8000 Then
    MsgBox "GAME OVER"
    Timer1.Enabled = False
    tmrClock.Enabled = False
End If


'Add call statements here
Call Collision(Yone)
Call Collision(Ytwo)
Call Collision(Ythree)
Call Collision(YFour)
Call Collision(YFive)
Call Collision(YSix)
Call Collision(RLeft)
Call Collision(RRight)
Call CollisionP(SB)
Call CollisionP(SB2)
Call Collision(lblEYEL)
Call Collision(lblEYER)
Call CollisionR(PrizeL)
Call CollisionR(PrizeR)
Call Collision(RightBlock)
Call Collision(LeftBlock)
Call CollisionP(BlueM)
Call CollisionP(BlueTL)
Call CollisionP(BlueTR)

If lblleave.Visible = True Then
        
        Call Collision(lblleave)
    
             ElseIf lblNow.Visible = True Then
    
        Call Collision(lblNow)
        
    End If

'move the ball
ball.Left = ball.Left + movex
ball.Top = ball.Top + movey

'ball collision with leftendwall
If ball.Left < LeftEndwall.Left + LeftEndwall.Width Then
    movex = -1 * movex
End If

'ball collision with rightendwall
If ball.Left + ball.Width > RightEndwall.Left Then
    movex = -1 * movex
End If

'ball collision with topwall
If ball.Top < topwall.Top + topwall.Height Then
    movey = -1 * movey
End If

'check to see if game is over...
If ball.Left <= 0 Then
    MsgBox "Game Over!"
    Timer1.Enabled = False
    tmrClock.Enabled = False
    lblleave.Caption = "I WIN"
    lblNow.Caption = "I WIN"
End If


If ball.Left < lblEYEL.Left Then
        eyeright.Left = 7920
        eyeleft.Left = 6360
ElseIf ball.Left > lblEYER.Left + lblEYER.Width Then
        eyeright.Left = 8400
        eyeleft.Left = 6840
Else
        eyeright.Left = 8160
        eyeleft.Left = 6600
End If
    
End Sub

Sub LeftPaddle()

'Paddle raised
If PaddleLeftPosition = 1 Then
    If ball.Left >= PaddleLeft.X1 And ball.Left <= PaddleLeft.X2 Then
        If ball.Top + ball.Height >= PaddleLeft.Y1 Then
            ball.Top = ball.Top - 4 * movey
            ball.Left = ball.Left - 4 * movex
            
            If movex > 0 Then
                movex = -1 * movex
                movey = -1 * movey
            Else
                movey = -movey
            End If
    
        End If
    End If
End If

'Paddle in standard position
'calculates a linear equation for the paddle and determines the height of the ball at collision
If PaddleLeftPosition <> 1 Then

Dim LeftSlope As Double
    LeftSlope = (PaddleLeft.Y2 - PaddleLeft.Y1) / (PaddleLeft.X2 - PaddleLeft.X1) 'Calculate slope of paddle

Dim YintLeft As Double
    YintLeft = PaddleLeft.Y1 - (LeftSlope * PaddleLeft.X1) 'Find y-intercept for y=mx+b
          
Static LeftChecker As Integer
    LeftChecker = (LeftSlope * ball.Left) + YintLeft
    
    If LeftChecker <= ball.Top + ball.Height Then
        If ball.Left >= PaddleLeft.X1 And ball.Left <= PaddleLeft.X2 Then
            ball.Top = ball.Top - movey
            ball.Left = ball.Left - movex
            
            movey = -1 * movey
        End If
    End If
    
End If

End Sub

Sub RightPaddle()

'Paddle Raised
If PaddleRightPosition = 1 Then
    If ball.Left >= PaddleRight.X1 And ball.Left <= PaddleRight.X2 Then
        If ball.Top + ball.Height >= PaddleRight.Y1 Then
            ball.Top = ball.Top - 4 * movey
            ball.Left = ball.Left - 4 * movex
            
            If movex < 0 Then
                movey = -movey
                movex = -movex
            Else
                movey = -movey
            End If
            
        End If
    End If
End If

'Paddle in standard position
'calculates a linear equation for the paddle and determines the height of the ball at collision

If PaddleRightPosition <> 1 Then

Dim RightSlope As Double
    RightSlope = (PaddleRight.Y2 - PaddleRight.Y1) / (PaddleRight.X2 - PaddleRight.X1) 'Calculate slope of paddle

Dim YintRight As Double
    YintRight = PaddleRight.Y1 - (RightSlope * PaddleRight.X1) 'Find y-intercept for y=mx+b
          
Static RightChecker As Integer
    RightChecker = (RightSlope * ball.Left) + YintRight
    
    If RightChecker <= ball.Top + ball.Height Then
        If ball.Left >= PaddleRight.X1 And ball.Left <= PaddleRight.X2 Then
            ball.Top = ball.Top - movey
            ball.Left = ball.Left - movex
            
            movey = -1 * movey
        End If
    End If
    
End If

End Sub
Function XValue(ByVal RecentPosition As Integer, ByVal Box As Label)

If RecentPosition + ball.Width <= Box.Left Then
    XValue = -movex
ElseIf RecentPosition >= Box.Left + Box.Width Then
    XValue = -movex
Else
    XValue = movex
End If

End Function

Function YValue(ByVal RecentPosition As Integer, ByVal Box As Label)

If RecentPosition + ball.Width >= Box.Left And RecentPosition <= Box.Left + Box.Width Then
    YValue = -movey
Else
    YValue = movey
End If

End Function

Sub Collision(ByVal Box As Label)

Dim RecentPosition As Integer
RecentPosition = ball.Left - (2 * movex)

'Debug.Print "movex"; movex
'Debug.Print "recent"; RecentPosition

If ball.Left <= Box.Left + Box.Width Then 'to the left of box right
   If ball.Left + ball.Width >= Box.Left Then   'to the right of box left
        If ball.Top <= Box.Top + Box.Height Then 'above box bottom
            If ball.Top + ball.Height >= Box.Top Then 'below box top
                        
                        movex = XValue(RecentPosition, Box)
                        movey = YValue(RecentPosition, Box)
                        
                    End If
                End If
            End If
        End If

End Sub

'This sub used for barriers that alot points
Sub CollisionR(ByVal Box As Label)

Dim RecentPosition As Integer
RecentPosition = ball.Left - (movex)

If ball.Left <= Box.Left + Box.Width Then 'to the left of box right
  
   If ball.Left + ball.Width >= Box.Left Then   'to the right of box left
       
        If ball.Top <= Box.Top + Box.Height Then 'above box bottom
         
            If ball.Top + ball.Height >= Box.Top Then 'below box top
                        Points = Points + 10
                        movex = XValue(RecentPosition, Box)
                        movey = YValue(RecentPosition, Box)
                        
                        Dim prizenum As Integer
                        prizenum = Int(4 * Rnd) + 1
                        
                        If prizenum = 2 Then
                            Points = Points + 33
                            lblPrize.Caption = "You just won 33 points"
                           
                        ElseIf prizenum = 3 Then
                            Points = Points + 99
                            lblPrize.Caption = "You just won 99 points"
                            
                        'ElseIf prizenum = 1 Then
                           ' movex = movex + 33
                           ' lblPrize.Caption = "SUPER SPEED"
                        End If
    
                        lblPoints = Points
                     
                    End If
                
                End If

            End If
         
        End If

End Sub

Sub CollisionP(ByVal Box As Label)


Dim RecentPosition As Integer
RecentPosition = ball.Left - (movex)

If ball.Left <= Box.Left + Box.Width Then 'to the left of box right
  
   If ball.Left + ball.Width >= Box.Left Then   'to the right of box left
       
        If ball.Top <= Box.Top + Box.Height Then 'above box bottom
         
            If ball.Top + ball.Height >= Box.Top Then 'below box top
                        Points = Points + 5
                        movex = XValue(RecentPosition, Box)
                        movey = YValue(RecentPosition, Box)
                         lblPoints = Points
                     
                    End If
                
                End If

            End If
         
        End If

End Sub

Private Sub tmrClock_Timer()
Static seconds As Integer
'show number of seconds gone by...
seconds = seconds + 1
lblSeconds = seconds

    If lblleave.Visible = False Then
        lblleave.Visible = True
        lblNow.Visible = False
    Else
        lblleave.Visible = False
        lblNow.Visible = True
    End If
  

End Sub

