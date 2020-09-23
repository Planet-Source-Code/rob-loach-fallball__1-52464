VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFB044&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fall Ball"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrScore 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   120
   End
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4920
      Top             =   120
   End
   Begin VB.Label lblBestPlayerName 
      BackStyle       =   0  'Transparent
      Caption         =   "First Place"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblBestPlayerScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   13
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblBestPlayerName 
      BackStyle       =   0  'Transparent
      Caption         =   "Second Place"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblBestPlayerScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   11
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto-Invinciballs:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblAutoInvinciballs 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblInvinciballs 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Invinciballs:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblBestPlayerScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   6
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblBestPlayerName 
      BackStyle       =   0  'Transparent
      Caption         =   "Third Place"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Created By:  Rob Loach"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "robl@loacheng.on.ca"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblMessage 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Shape Board 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   9
      Left            =   3120
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Shape Board 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   8
      Left            =   0
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Shape Board 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   7
      Left            =   3120
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Shape Board 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   6
      Left            =   0
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Shape Board 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   5
      Left            =   3120
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Shape Board 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   4
      Left            =   0
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Shape Board 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   3120
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Shape Board 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   0
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Shape Board 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   3120
      Top             =   120
      Width           =   2295
   End
   Begin VB.Shape Board 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   2580
      Shape           =   3  'Circle
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ArrowLeft As Byte = 37
Private Const ArrowRight As Byte = 39
Private Const ArrowUp As Byte = 38
Private Const ArrowDown As Byte = 40
Private Const HowFastBallMovesHoriz As Byte = 20
Private Const HowFastBallMovesVertUp As Byte = 5
Private Const HowFastBallMovesVertDown As Byte = 8
Private Const EachLevelUp As Byte = 10
Private Const InvinsableTime As Byte = 3
Private Const InvinsableColour As String = &HFF&
Private Const NormalColour As String = &HFFFF&

Dim boolPaused As Boolean
Dim strCheat As String
Dim boolCheated As Boolean
Dim lngLevel As Long
Dim lngScore As Long
Dim boolArrowLeft As Boolean
Dim boolArrowRight As Boolean
Dim boolArrowUp As Boolean
Dim boolArrowDown As Boolean
Dim BoardsUpEachTick As Long
Dim x As Long, y As Long
Dim lngInvinsableTimeLeft As Long
Dim strTemp As String


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case ArrowLeft
        boolArrowLeft = True
    Case ArrowRight
        boolArrowRight = True
    Case ArrowDown
        boolArrowDown = True
    Case ArrowUp
        boolArrowUp = True
    Case 80 'P
        Screen.MousePointer = 11
        If lblMessage.Caption <> "Press 'N' To Start A New Game" Then
            If boolPaused = True Then
                boolPaused = False
                tmrGame.Enabled = True
                tmrScore.Enabled = True
                lblMessage.Caption = ""
            Else
                boolPaused = True
                tmrGame.Enabled = False
                tmrScore.Enabled = False
                lblMessage.Caption = "Press 'P' To Un-Pause Game"
            End If
        End If
        Screen.MousePointer = 0
    Case 123 'F12
        boolPaused = True
        tmrGame.Enabled = False
        tmrScore.Enabled = False
        lblMessage.Caption = "Press 'P' To Un-Pause Game"
        MsgBox "You have found the Cheat Menu by pressing F12!" & vbCrLf & "Type the following cheats ingame:" & vbCrLf & vbCrLf & "GODMODE - Makes you invincible for 20 seconds" & vbCrLf & "GATE - Gets rid of walls" & vbCrLf & "SLOW - Slows down game" & vbCrLf & "BALL - Gives you an Auto-Invinciball" & vbCrLf & vbCrLf & "Press BACKSPACE to start typing the cheat again", vbOKOnly + vbInformation, App.ProductName & " - Cheat Menu"
    Case 71 'g
        strCheat = strCheat & "g"
    Case 79 'o
        strCheat = strCheat & "o"
    Case 68 'd
        strCheat = strCheat & "d"
    Case 77 'm
        strCheat = strCheat & "m"
    Case 65 'a
        strCheat = strCheat & "a"
    Case 66 'b
        strCheat = strCheat & "b"
    Case 84 't
        strCheat = strCheat & "t"
    Case 69 'e
        strCheat = strCheat & "e"
        If strCheat = "gate" Then
            boolCheated = True
            strCheat = ""
            For x = 0 To 9
                Board(x).Width = 100
            Next
        End If
        If strCheat = "godmode" Then
            boolCheated = True
            strCheat = ""
            If Ball.BackColor = NormalColour Then
                Ball.BackColor = InvinsableColour
                lngInvinsableTimeLeft = 20
            Else
                lngInvinsableTimeLeft = 1
            End If
        End If
    Case 83 's
        strCheat = strCheat & "s"
    Case 76 'l
        strCheat = strCheat & "l"
        If strCheat = "ball" Then
            boolCheated = True
            strCheat = ""
            If lblMessage.Caption = "" Then
                lblAutoInvinciballs.Caption = CLng(lblAutoInvinciballs.Caption) + 1
            End If
        End If
    Case 87 'w
        strCheat = strCheat & "w"
        If strCheat = "slow" Then
            boolCheated = True
            strCheat = ""
            BoardsUpEachTick = 7
        End If
    Case 46 'DELETE
        strCheat = ""
    Case 8 'backspace
        strCheat = ""
        
            
            
            
    Case 78 'N
        Screen.MousePointer = 11
        boolPaused = True
        tmrGame.Enabled = False
        tmrScore.Enabled = False
        lblMessage.Caption = "Press 'P' When Ready"
        lngScore = 0
        lblScore.Caption = "0"
        Ball.Left = 2520
        Ball.Top = 120
        Ball.BackColor = InvinsableColour
        lngInvinsableTimeLeft = InvinsableTime
        BoardsUpEachTick = 8
        lngLevel = 0
        lblInvinciballs.Caption = "0"
        boolArrowLeft = False
        boolArrowRight = False
        boolArrowUp = False
        boolArrowDown = False
        Randomize CDbl(Right(Str(CDbl(Timer) * CInt(Day(Now)) * 11), 2)) * lblBestPlayerScore.Item(1).Caption
        lblInvinciballs.Caption = 0
        boolCheated = False
        Screen.MousePointer = 0
    Case 81 'Q
        End
    Case 27 'ESC
        End
    Case 32 'space
        x = CLng(lblInvinciballs.Caption)
        If Ball.BackColor <> InvinsableColour Then
            If x > 0 Then
                Ball.BackColor = InvinsableColour
                lngInvinsableTimeLeft = InvinsableTime
                lblInvinciballs.Caption = x - 1
            End If
        End If
    Case 13 'enter
        x = CLng(lblInvinciballs.Caption)
        y = CLng(lblAutoInvinciballs.Caption)
        If x >= 3 Then
            x = x - 3
            y = y + 1
            lblInvinciballs.Caption = x
            lblAutoInvinciballs.Caption = y
        End If
        
    Case 112 'f1
        boolPaused = True
        tmrGame.Enabled = False
        tmrScore.Enabled = False
        lblMessage.Caption = "Press 'P' To Un-Pause Game"
        MsgBox "Guide your ball through the endless maze of walls." & vbCrLf & "Don't let it hit a wall though!" & vbCrLf & vbCrLf & "F1 - Display this help menu" & vbCrLf & "Arrows - Move ball" & vbCrLf & "N - New Game" & vbCrLf & "P - Pause" & vbCrLf & "SPACE - Use Invinciball" & vbCrLf & "ENTER - Convert 3 Invinciball Into an Auto-Invinciball", vbInformation + vbOKOnly, App.ProductName & " - Help"
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case ArrowLeft
        boolArrowLeft = False
    Case ArrowRight
        boolArrowRight = False
    Case ArrowUp
        boolArrowUp = False
    Case ArrowDown
        boolArrowDown = False
    End Select
End Sub

Private Sub tmrGame_Timer()
    Dim LeftWidth As Long
    Dim Opening As Long

    For x = 0 To 9
        Board(x).Top = Board(x).Top - BoardsUpEachTick
    Next
    For x = 0 To 9 Step 2
        If Board(x).Top + Board(x).Height < 0 Then
            LeftWidth = RandomNumber(500, 3500)
            Board(x).Top = Me.Height - 200
            Board(x + 1).Top = Me.Height - 200
            Board(x).Width = LeftWidth
            Opening = RandomNumber(600, 1500)
            Board(x + 1).Width = Me.Width - Opening - LeftWidth - 100
            Board(x + 1).Left = Board(x).Width + Opening
        End If
    Next
    If boolArrowLeft Then
        Ball.Left = Ball.Left - HowFastBallMovesHoriz
    ElseIf boolArrowRight Then
        Ball.Left = Ball.Left + HowFastBallMovesHoriz
    End If
    If boolArrowUp Then
        Ball.Top = Ball.Top - HowFastBallMovesVertUp
    ElseIf boolArrowDown Then
        Ball.Top = Ball.Top + HowFastBallMovesVertDown
    End If
    If Ball.BackColor <> InvinsableColour Then
        For x = 0 To 9
            If CollisionDetect(Board(x)) = True Then
            
                If CLng(lblAutoInvinciballs.Caption) >= 1 Then
                    lblAutoInvinciballs.Caption = CLng(lblAutoInvinciballs.Caption) - 1
                    Ball.BackColor = InvinsableColour
                    lngInvinsableTimeLeft = InvinsableTime
                    Exit Sub
                End If
                Screen.MousePointer = 11
                tmrGame.Enabled = False
                tmrScore.Enabled = False
                lblMessage.Caption = "Press 'N' To Start A New Game"
                For y = 1 To 3
                    If lngScore > CLng(lblBestPlayerScore(y).Caption) Then
                        strTemp = App.Path & "\fallball.dat"
                        Select Case y
                            Case 1
                                File_INIWrite "FallBall", "Score3", File_INIRead("FallBall", "Score2", strTemp), strTemp
                                File_INIWrite "FallBall", "Score2", File_INIRead("FallBall", "Score1", strTemp), strTemp
                                File_INIWrite "FallBall", "Name3", File_INIRead("FallBall", "Name2", strTemp), strTemp
                                File_INIWrite "FallBall", "Name2", File_INIRead("FallBall", "Name1", strTemp), strTemp
                                lblBestPlayerName(3).Caption = lblBestPlayerName(2).Caption
                                lblBestPlayerScore(3).Caption = lblBestPlayerScore(2).Caption
                                lblBestPlayerName(2).Caption = lblBestPlayerName(1).Caption
                                lblBestPlayerScore(2).Caption = lblBestPlayerScore(1).Caption
                            Case 2
                                File_INIWrite "FallBall", "Score3", File_INIRead("FallBall", "Score2", strTemp), strTemp
                                File_INIWrite "FallBall", "Name3", File_INIRead("FallBall", "Name2", strTemp), strTemp
                                lblBestPlayerName(3).Caption = lblBestPlayerName(2).Caption
                                lblBestPlayerScore(3).Caption = lblBestPlayerScore(2).Caption
                        End Select
                        lblBestPlayerName(y).Caption = Trim(InputBox("Congratulations!" & vbCrLf & "You are #" & y & " on the high-score list!" & vbCrLf & vbCrLf & "What is your name?", App.ProductName & " - High-Score", " Player"))
                        If lblBestPlayerName(y).Caption = "" Then lblBestPlayerName(y).Caption = "Player"
                        If boolCheated = True Then lblBestPlayerName(y).Caption = lblBestPlayerName(y).Caption & " (Cheat)"
                        lblBestPlayerScore(y).Caption = lngScore
                        File_INIWrite "FallBall", "Score" & y, Encrypt(lblBestPlayerScore(y).Caption), App.Path & "\fallball.dat"
                        File_INIWrite "FallBall", "Name" & y, Encrypt(lblBestPlayerName(y).Caption), App.Path & "\fallball.dat"
                        Exit For
                    End If
                Next
            End If
        Next
    End If
    Screen.MousePointer = 0
End Sub


Private Sub Form_Load()
    Screen.MousePointer = 11
    Randomize CDbl(Right(Str(CDbl(Timer) * CInt(Day(Now)) * 11), 3))
    Me.Caption = App.ProductName & "  -  v" & App.Major & "." & App.Minor
    boolPaused = True
    BoardsUpEachTick = 8
    lngLevel = 0
    boolCheated = False
    tmrGame.Enabled = False
    tmrScore.Enabled = False
    lblMessage.Caption = "Press 'P' When Ready"
    Ball.Left = 2580
    Ball.Top = 120
    Ball.BackColor = InvinsableColour
    lngInvinsableTimeLeft = InvinsableTime
    lblScore.Caption = "0"
    If File_INIRead("FallBall", "Score1", App.Path & "\fallball.dat") = "" Then
        File_INIWrite "FallBall", "Score1", Encrypt("20"), App.Path & "\fallball.dat"
        File_INIWrite "FallBall", "Score2", Encrypt("10"), App.Path & "\fallball.dat"
        File_INIWrite "FallBall", "Score3", Encrypt("5"), App.Path & "\fallball.dat"
        File_INIWrite "FallBall", "Name1", Encrypt("First Place"), App.Path & "\fallball.dat"
        File_INIWrite "FallBall", "Name2", Encrypt("Second Place"), App.Path & "\fallball.dat"
        File_INIWrite "FallBall", "Name3", Encrypt("Third Place"), App.Path & "\fallball.dat"
        MsgBox "Welcome to Fall Ball!" & vbCrLf & vbCrLf & "If you ever need any help about playing the game," & vbCrLf & "press F1.  This will pop up the help menu.", vbOKOnly + vbInformation, App.ProductName & " - Welcome"
    Else
        For x = 1 To 3
            lblBestPlayerName(x).Caption = Decrypt(File_INIRead("FallBall", "Name" & x, App.Path & "\fallball.dat"))
            lblBestPlayerScore(x).Caption = Decrypt(File_INIRead("FallBall", "Score" & x, App.Path & "\fallball.dat"))
        Next

    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub tmrScore_Timer()
    lngScore = lngScore + 1
    lblScore.Caption = lngScore
    If Ball.Top < 0 Then Ball.Top = 10
    If Ball.Top > Me.Height - Ball.Height - 200 Then Ball.Top = Me.Height - Ball.Height - 600
    If Ball.Left < 0 Then Ball.Left = 15
    If Ball.Left > Me.Width - Ball.Width Then Ball.Left = Me.Width - Ball.Width - 100
    x = lngScore / EachLevelUp
    If x > lngLevel Then
        lngLevel = lngLevel + 1
        BoardsUpEachTick = BoardsUpEachTick + 1
        lblInvinciballs.Caption = lblInvinciballs.Caption + 1
    End If
    If lngInvinsableTimeLeft > 0 Then
        lngInvinsableTimeLeft = lngInvinsableTimeLeft - 1
    Else
        Ball.BackColor = NormalColour
    End If
End Sub

Public Function CollisionDetect(ByRef Board As Shape) As Boolean
    If Ball.Left + Ball.Width > Board.Left + 30 Then 'check left
        If Ball.Top + Ball.Height > Board.Top + 30 Then 'check top
            If Ball.Top < Board.Top + Board.Height - 30 Then 'check bottem
                If Ball.Left < Board.Left + Board.Width - 30 Then 'check right
                    CollisionDetect = True
                End If
            End If
        End If
    End If
End Function









