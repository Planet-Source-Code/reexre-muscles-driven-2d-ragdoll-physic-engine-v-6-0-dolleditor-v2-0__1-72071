VERSION 5.00
Begin VB.Form frmEditor 
   Caption         =   "Doll Editor"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   740
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUNDO1 
      Caption         =   "Undo"
      Height          =   615
      Left            =   9360
      TabIndex        =   34
      Top             =   1680
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10440
      Top             =   2040
   End
   Begin VB.CheckBox RunTest 
      Caption         =   "TEST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   840
      Width           =   975
   End
   Begin VB.HScrollBar LoadScale 
      Height          =   255
      Left            =   7560
      Max             =   200
      Min             =   35
      TabIndex        =   28
      Top             =   7440
      Value           =   100
      Width           =   3375
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   7560
      TabIndex        =   27
      Top             =   5520
      Width           =   3375
   End
   Begin VB.CheckBox chShowNum 
      Caption         =   "Show Numbers"
      Height          =   495
      Left            =   9360
      TabIndex        =   26
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame fLink 
      Caption         =   "Link Options"
      Height          =   2535
      Left            =   7560
      TabIndex        =   6
      Top             =   2880
      Width           =   3375
      Begin VB.PictureBox PicThick 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   55
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   95
         TabIndex        =   31
         Top             =   840
         Width           =   1455
      End
      Begin VB.PictureBox PicColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   1680
         ScaleHeight     =   55
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   47
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.VScrollBar sB 
         Height          =   1335
         Left            =   3000
         Max             =   255
         TabIndex        =   15
         Top             =   480
         Width           =   255
      End
      Begin VB.VScrollBar sG 
         Height          =   1335
         Left            =   2760
         Max             =   255
         TabIndex        =   14
         Top             =   480
         Value           =   255
         Width           =   255
      End
      Begin VB.VScrollBar sR 
         Height          =   1335
         Left            =   2520
         Max             =   255
         TabIndex        =   13
         Top             =   480
         Width           =   255
      End
      Begin VB.HScrollBar sThickness 
         Height          =   255
         Left            =   120
         Max             =   12
         Min             =   1
         TabIndex        =   10
         Top             =   480
         Value           =   3
         Width           =   1455
      End
      Begin VB.Frame frHT 
         Caption         =   "Draw Type"
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   3135
         Begin VB.OptionButton oSticky 
            Caption         =   "Sticky"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton oFilledCircle 
            Caption         =   "Filled Circle"
            Height          =   255
            Left            =   1080
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton oSmile 
            Caption         =   "Smile"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2280
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Label Label2 
         Caption         =   "R G B"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Thickness 1"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   21
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer TimerM 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10440
      Top             =   2520
   End
   Begin VB.CommandButton SaveDoll 
      Caption         =   "Save Doll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame fPoint 
      Caption         =   "Point Options"
      Height          =   735
      Left            =   7560
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
      Begin VB.CheckBox chUnMovable 
         Caption         =   "UnMovable"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add Entity"
      Height          =   1215
      Left            =   7560
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
      Begin VB.OptionButton oMuscle 
         Caption         =   "Muscle"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton oLink 
         Caption         =   "Link"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton oPOINT 
         Caption         =   "Point"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Lmuscle 
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Llink 
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Lpoint 
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004000&
      DrawStyle       =   5  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   479
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   479
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
   Begin VB.Frame fMuscle 
      Caption         =   "Muscle Options"
      Height          =   495
      Left            =   7560
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Use Mouse to Interact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   7560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Load/Save Scale"
      Height          =   255
      Left            =   7560
      TabIndex        =   30
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      Height          =   255
      Left            =   9000
      TabIndex        =   29
      Top             =   7800
      Width           =   495
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author : Creator Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
Option Explicit

Private DOLL       As New OBJphysic
Private FPointSelected As Boolean
Private Closest    As Integer
Private Closest2   As Integer
Private Color      As Long
Private FLinkSelected As Boolean
Private defStren   As Double
Private x1         As Single
Private y1         As Single
Private x2         As Single
Private y2         As Single


Private InteractWith As Integer
Private PtoMove    As Integer
Private mouseX     As Single
Private mouseY     As Single


Private Sub Check1_Click()

End Sub

Private Sub chShowNum_Click()
    DRAWDOLL

End Sub

Private Sub cmdClearAll_Click()

    RunTest.Value = Unchecked

    PIC.Cls

    DOLL.DestroyMe
    FPointSelected = False
    FLinkSelected = False
    oPOINT = True
    Lpoint = 0
    Llink = 0
    Lmuscle = 0

End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdUNDO1_Click()
    If oPOINT Then
        If DOLL.Npoints > 0 Then DOLL.Npoints = DOLL.Npoints - 1
    End If
    If oLink Then
        If DOLL.Nlinks > 0 Then DOLL.Nlinks = DOLL.Nlinks - 1
    End If
    If oMuscle Then
        If DOLL.NMuscles > 0 Then DOLL.NMuscles = DOLL.NMuscles - 1
    End If

    DRAWDOLL

End Sub

Private Sub File1_DblClick()
    DOLL.DestroyMe
    DOLL.OBJ_LOADandPlace File1, PIC.Width \ 2, PIC.Height \ 2, LoadScale / 100
    DRAWDOLL
    Lpoint = DOLL.Npoints
    Llink = DOLL.Nlinks
    Lmuscle = DOLL.NMuscles

End Sub

Private Sub Form_Activate()
    SelectChange


    defStren = 0.03 * 1.2
    DOLL.GlobalMAXStrength = defStren
    DOLL.CurrentMAXStrength = defStren
    DOLL.MaxX = PIC.Width - 2
    DOLL.MaxY = PIC.Height - 2


    Gravity = 0.035
    Doll_Air_Resistence = 0.994

End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " V" & App.Major & "." & App.Minor

    File1.Path = App.Path
    File1.Filename = "*.doll"


End Sub

Private Sub LoadScale_Change()
    Label3 = LoadScale / 100
    If File1 <> "" Then
        DOLL.OBJ_LOADandPlace File1, PIC.Width \ 2, PIC.Height \ 2, LoadScale / 100
        DRAWDOLL
    End If
End Sub

Private Sub LoadScale_Scroll()
    Label3 = LoadScale / 100
    If File1 <> "" Then
        DOLL.OBJ_LOADandPlace File1, PIC.Width \ 2, PIC.Height \ 2, LoadScale / 100
        DRAWDOLL
    End If
End Sub

Private Sub oFilledCircle_Click()
    If oSticky Then TmpDrawShape = sLine
    If oFilledCircle Then TmpDrawShape = sFillCircle
    If oSmile Then TmpDrawShape = sFace
End Sub

Private Sub oMuscle_Click()
    If DOLL.Nlinks < 2 Then oLink = True
    SelectChange
End Sub

Private Sub oSmile_Click()
    If oSticky Then TmpDrawShape = sLine
    If oFilledCircle Then TmpDrawShape = sFillCircle
    If oSmile Then TmpDrawShape = sFace
End Sub

Private Sub oSticky_Click()
    If oSticky Then TmpDrawShape = sLine
    If oFilledCircle Then TmpDrawShape = sFillCircle
    If oSmile Then TmpDrawShape = sFace
End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InteractWith = 0
End Sub

Private Sub RunTest_Click()
    Dim I          As Long

    If RunTest.Value = Checked Then
        For I = 1 To DOLL.Npoints
            DOLL.PointVX(I) = 0
            DOLL.PointVY(I) = 0
        Next

        DOLL.OBJ_SAVE "zztmp.doll"
        Timer1.Enabled = True
        Label5.Visible = True
    Else
        Timer1.Enabled = False
        DOLL.OBJ_LOADandPlace "zztmp.doll", PIC.Width \ 2, PIC.Height \ 2, 1    'LoadScale / 100
        DRAWDOLL
        Label5.Visible = False

    End If
End Sub

Private Sub SaveDoll_Click()

    RunTest.Value = Unchecked

    Dim S          As String
    S = "DollName.doll"
    S = InputBox("Type Doll Name", , S)
    If Right$(S, 5) <> ".doll" Then S = S & ".doll"


    DOLL.OBJ_SAVE S
    MsgBox S & " saved"

    DOLL.OBJ_LOADandPlace S, PIC.Width \ 2, PIC.Height \ 2
    DRAWDOLL

    File1.Refresh



End Sub

Private Sub sB_Change()
    ColorChange
End Sub

Private Sub sB_Scroll()
    ColorChange
End Sub

Private Sub sG_Change()
    ColorChange
End Sub

Private Sub sG_Scroll()
    ColorChange
End Sub

Private Sub sR_Change()
    ColorChange
End Sub

Private Sub sR_Scroll()
    ColorChange
End Sub

Private Sub sThickness_Change()
    Label1 = "Thickness " & sThickness
    ColorChange
    DoEvents

End Sub

Private Sub sThickness_Scroll()
    Label1 = "Thickness " & sThickness
    ColorChange
    DoEvents
End Sub

Private Sub oLink_Click()
    If DOLL.Npoints < 2 Then oPOINT = True

    SelectChange
End Sub

Private Sub oPOINT_Click()
    SelectChange
End Sub

Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Pmouse     As tPoint
    Dim Pdoll      As tPoint
    Dim D          As Single
    Dim Dmin       As Single
    Dim Mmid       As tPoint
    Dim I          As Long

    If Timer1.Enabled = False Then
        Pmouse.X = X
        Pmouse.Y = Y

        ' Point
        '---------------------------------------------------------------
        If oPOINT Then
            DOLL.ADDpoint X, Y, IIf(chUnMovable = Checked, True, False)
            Lpoint = DOLL.Npoints
            cmdUNDO1.Enabled = True
        End If

        'Link
        '---------------------------------------------------------------
        If oLink And Button = 1 Then
            If FPointSelected = False Then

                Dmin = 999999999
                For I = 1 To DOLL.Npoints
                    Pdoll.X = DOLL.PointX(I)
                    Pdoll.Y = DOLL.PointY(I)
                    D = Distance(Pmouse, Pdoll)
                    If D < Dmin Then Dmin = D: Closest = I
                Next I

                FPointSelected = True
                Exit Sub
            End If
        End If

        If oLink And Button = 2 Then
            FPointSelected = False
            DRAWDOLL
        End If
        If oLink And Button = 1 Then
            If FPointSelected = True Then

                Dmin = 999999999
                For I = 1 To DOLL.Npoints
                    Pdoll.X = DOLL.PointX(I)
                    Pdoll.Y = DOLL.PointY(I)
                    D = Distance(Pmouse, Pdoll)
                    If D < Dmin Then Dmin = D: Closest2 = I
                Next I
                If Closest <> Closest2 Then
                    DOLL.ADDLink Closest, Closest2, sThickness, Color, TmpDrawShape, DOLL.Nlinks + 1
                    Llink = DOLL.Nlinks
                    FPointSelected = False
                    cmdUNDO1.Enabled = True
                End If
            End If
        End If

        'Muscle
        '----------------------------------------------------------------------
        If oMuscle And Button = 1 Then
            If FLinkSelected = False Then
                '        Stop

                Dmin = 999999999
                For I = 1 To DOLL.Nlinks
                    With DOLL
                        Mmid.X = (.PointX(.Link_P1(I)) + .PointX(.Link_P2(I))) / 2
                        Mmid.Y = (.PointY(.Link_P1(I)) + .PointY(.Link_P2(I))) / 2
                        D = Distance(Pmouse, Mmid)
                        If D < Dmin Then Dmin = D: Closest = I: x1 = Mmid.X: y1 = Mmid.Y
                    End With
                Next I
                TimerM.Enabled = True
                FLinkSelected = True
                Exit Sub

            End If
        End If
        If oMuscle And Button = 2 Then
            FLinkSelected = False
            DRAWDOLL
            TimerM.Enabled = False
        End If
        If oMuscle And Button = 1 Then
            If FLinkSelected = True Then

                Dmin = 999999999
                For I = 1 To DOLL.Nlinks
                    With DOLL
                        Mmid.X = (.PointX(.Link_P1(I)) + .PointX(.Link_P2(I))) / 2
                        Mmid.Y = (.PointY(.Link_P1(I)) + .PointY(.Link_P2(I))) / 2
                        D = Distance(Pmouse, Mmid)
                        If D < Dmin Then Dmin = D: Closest2 = I: x2 = Mmid.X: y2 = Mmid.Y
                    End With
                Next I
                If Closest <> Closest2 Then
                    DOLL.ADDMuscle Closest, Closest2, defStren
                    Lmuscle = DOLL.NMuscles
                    FLinkSelected = False
                    TimerM.Enabled = False
                    cmdUNDO1.Enabled = True
                End If
            End If
        End If



        '-------------------------------------------------

        DRAWDOLL

    Else



        Dim P1     As tPoint
        Dim P2     As tPoint
        If Button = 1 Then
            Dmin = 1E+19
            P2.X = X
            P2.Y = Y
            For I = 1 To DOLL.Npoints
                P1.X = DOLL.PointX(I)
                P1.Y = DOLL.PointY(I)
                If Distance(P1, P2) < Dmin Then Dmin = Distance(P1, P2): PtoMove = I
            Next
            mouseX = X
            mouseY = Y
            InteractWith = 1
        Else

        End If

    End If



End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Timer1.Enabled = False Then


        If FPointSelected And oLink Then
            DRAWDOLL
            FastLine PIC.hdc, DOLL.PointX(Closest), DOLL.PointY(Closest), X \ 1, Y \ 1, 1, vbGreen
        End If


        If FLinkSelected And oMuscle Then

            DRAWDOLL
            x2 = X
            y2 = Y
        End If

    Else
        If Button <> 0 Then
            mouseX = X
            mouseY = Y
        End If
    End If


End Sub
Sub DRAWDOLL()
    Dim X0         As Double
    Dim Y0         As Double
    Dim x1         As Double
    Dim y1         As Double
    Dim x2         As Double
    Dim y2         As Double
    Dim A1         As Double
    Dim A2         As Double
    Dim A          As Double
    Dim L1         As Double

    Dim L2         As Double
    Dim xx1        As Long
    Dim yy1        As Long

    Dim I          As Long

    'PIC.Cls
    BitBlt PIC.hdc, 0, 0, PIC.Width, PIC.Height, PIC.hdc, 0, 0, vbBlackness
    DOLL.DRAW PIC.hdc, IIf(chShowNum.Value = Checked, True, False)


    '---------------------------------------------------------------
    For I = 1 To DOLL.Npoints
        MyCircle PIC.hdc, DOLL.PointX(I), DOLL.PointY(I), 3, 2, IIf(DOLL.PointIsFix(I), vbRed, vbGreen)
    Next

    With DOLL
        For I = 1 To DOLL.NMuscles

            'xx1 = (.PointX((.Muscle_P0(I))) + .PointX((.Muscle_P1(I)))) / 2
            'yy1 = (.PointY((.Muscle_P0(I))) + .PointY((.Muscle_P1(I)))) / 2
            'xX2 = (.PointX((.Muscle_P0(I))) + .PointX((.Muscle_P2(I)))) / 2
            'yY2 = (.PointY((.Muscle_P0(I))) + .PointY((.Muscle_P2(I)))) / 2
            'FastLine PIC.hdc, xx1, yy1, xX2, yY2, 5, vbYellow
            'FastLine PIC.hdc, xx1, yy1, xX2, yY2, 1, vbMagenta
            X0 = .PointX(.Muscle_P0(I))
            Y0 = .PointY(.Muscle_P0(I))
            x1 = .PointX(.Muscle_P1(I))
            y1 = .PointY(.Muscle_P1(I))
            x2 = .PointX(.Muscle_P2(I))
            y2 = .PointY(.Muscle_P2(I))
            A1 = Atan2(x1 - X0, y1 - Y0)
            A2 = Atan2(x2 - X0, y2 - Y0)

            If A1 > A2 Then
                A = A1
                A1 = A2
                A2 = A
            End If


            L1 = .Link_MainL(.Muscle_L1(I)) / 2
            L2 = .Link_MainL(.Muscle_L2(I)) / 2

            If L1 > L2 Then L1 = L2


            For A = A1 To A2 Step 3 / L1
                xx1 = (X0 + Cos(A) * L1)
                yy1 = (Y0 + Sin(A) * L1)
                FastLine PIC.hdc, xx1, yy1, xx1, yy1, 3, vbYellow
            Next


        Next
    End With
    PIC.Refresh

    DoEvents
End Sub

Sub ColorChange()
    Color = RGB(sR, sG, sB)
    PicColor.Line (0, 0)-(PicColor.Width, PicColor.Height), Color, BF
    PicColor.Refresh
    PicThick.Cls
    FastLine PicThick.hdc, 10, 10, PicThick.ScaleWidth - 10, PicThick.ScaleHeight - 10, sThickness, Color

End Sub

Sub SelectChange()
    If oPOINT Then
        fPoint.Visible = True: fLink.Visible = False: fMuscle.Visible = False
    End If
    If oLink Then
        ColorChange
        Label1 = "Thickness " & sThickness
        fPoint.Visible = False: fLink.Visible = True: fMuscle.Visible = False
    End If
    If oMuscle Then
        fPoint.Visible = False: fLink.Visible = False: fMuscle.Visible = True
    End If

    cmdUNDO1.Enabled = False

End Sub

Private Sub Timer1_Timer()

    Dim D          As Long
    Dim ii         As Long

    '-------------------------- DRAW
    BitBlt PIC.hdc, 0, 0, PIC.Width, PIC.Height, PIC.hdc, 0, 0, vbBlackness
    'DrawObstacles
    DOLL.DRAW PIC.hdc
    PIC.Refresh

    '-------------------------- Doll-obsatcle collisions

    'For I = 1 To Nobs
    '    With Obstacle(I)
    '        For D = 1 To UBound(DOLL)
    '            'DOLL(D).CheckCircleShapedCollision .P.X, .P.Y, .P.vX, .P.vY, .R, .IsMotionLess
    '            For ii = 1 To DOLL(D).Npoints
    '                DOLL(D).CheckCircleShapedCollision2 ii, .P.x, .P.y, .P.vX, .P.vY, .R, .IsMotionLess
    '            Next ii
    '        Next
    '    End With
    'Next

    'Stop

    DOLL.DoPHYSICS True
    DoMouseForces InteractWith
End Sub

Private Sub TimerM_Timer()


    FastLine PIC.hdc, x1 \ 1, y1 \ 1, x2 \ 1, y2 \ 1, 2, RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)

End Sub


Sub DoMouseForces(DollOrObstacle)

    If DollOrObstacle = 1 Then
        DOLL.PointVX(PtoMove) = DOLL.PointVX(PtoMove) - (DOLL.PointX(PtoMove) - mouseX) * 0.015
        DOLL.PointVY(PtoMove) = DOLL.PointVY(PtoMove) - (DOLL.PointY(PtoMove) - mouseY) * 0.015
        'PIC.Line (Doll.PointX(PtoMove), Doll.PointY(PtoMove))-(mouseX, mouseY), vbYellow
        FastLine PIC.hdc, DOLL.PointX(PtoMove), DOLL.PointY(PtoMove), mouseX \ 1, mouseY \ 1, 2, vbYellow
    ElseIf DollOrObstacle = 2 Then
        'With Obstacle(PtoMove)
        '    .P.vX = .P.vX - (.P.X - mouseX) * 0.003
        '    .P.vY = .P.vY - (.P.Y - mouseY) * 0.003
        '    FastLine PIC.hdc, .P.X, .P.Y, mouseX, mouseY, 2, vbYellow
        'End With
    End If



End Sub

