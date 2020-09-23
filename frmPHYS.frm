VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPHYS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Muscles Driven 2D Ragdoll Physic Engine "
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   ControlBox      =   0   'False
   Icon            =   "frmPHYS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   581
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar RopeStrengthBar 
      Height          =   255
      Left            =   7920
      TabIndex        =   23
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Frame frHT 
      Caption         =   "Head Type"
      Height          =   615
      Left            =   7920
      TabIndex        =   18
      Top             =   6840
      Width           =   3615
      Begin VB.OptionButton oSmile 
         Caption         =   "Smile"
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton oFilledCircle 
         Caption         =   "Filled Circle"
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton oSticky 
         Caption         =   "Sticky"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton ReStart 
      Caption         =   "ReStart"
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
      Left            =   10680
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar NRGbar 
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   7440
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9120
      Top             =   600
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
      Left            =   10680
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chGravity 
      Caption         =   "Gravity"
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
      Left            =   7920
      TabIndex        =   12
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton SavePos 
      Caption         =   "Save Pose"
      Height          =   375
      Left            =   10080
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   9840
      TabIndex        =   10
      ToolTipText     =   "Click to load Pose"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.HScrollBar DollStrengthBar 
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   3720
      Width           =   2655
   End
   Begin VB.HScrollBar MuscleANG 
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   7
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CheckBox ApplyMuscle 
      Caption         =   "Doll Uses Muscles"
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Only Lines"
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   1320
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton ShowPointLinkNumbers 
      Caption         =   "Show Point Link Numbers"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004000&
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   479
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   503
      TabIndex        =   1
      Top             =   120
      Width           =   7575
   End
   Begin VB.Timer TIMER1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8520
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RUN"
      Height          =   615
      Left            =   10200
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   $"frmPHYS.frx":030A
      ForeColor       =   &H80000011&
      Height          =   975
      Left            =   6240
      TabIndex        =   25
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "'Rope' Muscles Max Strength"
      Height          =   495
      Left            =   10680
      TabIndex        =   24
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      Caption         =   $"frmPHYS.frx":03A6
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3360
      TabIndex        =   22
      Top             =   7680
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "-Left Button- Interact with the Doll. -Right Button- Interact with Obstacles."
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   7680
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Muscles ENERGY"
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
      Left            =   6120
      TabIndex        =   16
      Top             =   7440
      Width           =   1725
   End
   Begin VB.Label mDESC 
      Caption         =   "Label3"
      Height          =   255
      Index           =   0
      Left            =   10680
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Muscles Driven 2D Ragdoll Physic Engine.  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Doll Muscles Max Strength"
      Height          =   495
      Left            =   10680
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "frmPHYS"
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

'
Option Explicit

Dim InteractWith   As Long
Dim PtoMove        As Long
Dim mouseX         As Single
Dim mouseY         As Single

Dim Obstacle()     As TObstacle
Dim Nobs           As Long




Private Sub chGravity_Click()
    If chGravity.Value = Checked Then
        Gravity = 0.035           '0.035
    Else
        Gravity = 0

    End If

End Sub

Private Sub cmdExit_Click()
    End

End Sub

Private Sub Command1_Click()

    Dim S          As Single
    Dim M          As Long
    Dim I          As Long
    Dim I2         As Long


    Randomize Timer

    ReDim DOLL(4)
    Njoints = 0
    ReDim Joint(0)


    Command1.Visible = False
    SavePos.Visible = True
    cmdExit.Visible = True
    ReStart.Visible = True

    Doll_Air_Resistence = 0.994   '0.99
    Obstacle_Air_Resistence = 0.998

    chGravity_Click

    S = 0.03 * 1.2                '* 0.9 '* 1.2 '* 0.9
    DOLL(1).GlobalMAXStrength = S
    DOLL(1).CurrentMAXStrength = S
    DOLL(1).MaxX = PIC.Width - 2
    DOLL(1).MaxY = PIC.Height - 2

    NRGbar.Max = DOLL(1).GlobalMAXStrength * 100
    NRGbar.Value = NRGbar.Max

    ''''''''''''''''''''''''''''''''''''''''
    'RagDoll
    DOLL(1).ADDpoint 110, 200
    DOLL(1).ADDpoint 110, 170
    DOLL(1).ADDpoint 120, 140
    DOLL(1).ADDpoint 130, 170
    DOLL(1).ADDpoint 130, 200
    DOLL(1).ADDpoint 120, 110

    DOLL(1).ADDpoint 100, 110
    DOLL(1).ADDpoint 90, 130

    DOLL(1).ADDpoint 140, 110
    DOLL(1).ADDpoint 150, 130

    DOLL(1).ADDpoint 120, 90 - 2  '-20

    'Links
    DOLL(1).ADDLink 1, 2, 9, RGB(110, 130, 235), , 5, 2.5
    DOLL(1).ADDLink 2, 3, 9, RGB(115, 135, 255), , 1, 2.5
    DOLL(1).ADDLink 5, 4, 9, RGB(105, 125, 230), , 2, 2.5
    DOLL(1).ADDLink 4, 3, 9, RGB(110, 130, 250), , 3, 2.5

    DOLL(1).ADDLink 6, 3, 12, RGB(180, 0, 0), , 4, 2.5

    DOLL(1).ADDLink 8, 7, 6, RGB(200, 200, 50), , 6, 2.5
    DOLL(1).ADDLink 7, 6, 6, RGB(200, 90, 90), , 7, 2.5

    DOLL(1).ADDLink 10, 9, 6, RGB(200, 200, 50), , 8, 2.5
    DOLL(1).ADDLink 9, 6, 6, RGB(200, 90, 90), , 9, 2.5
    'Sticky Head
    'doll(1).ADDLink 11, 6, 6, RGB(180, 220, 50), , 9, 2.5
    'Circled Head
    'doll(1).ADDLink 11, 6, 6, RGB(180, 220, 50), sCircle, , 10, 2.5
    'FilledCircle Head
    DOLL(1).ADDLink 11, 6, 6, RGB(180, 220, 50), sFace, 10, 2.5


    DOLL(1).ADDMuscle 1, 2, S
    DOLL(1).ADDMuscle 3, 4, S
    DOLL(1).ADDMuscle 2, 5, S
    DOLL(1).ADDMuscle 4, 5, S

    DOLL(1).ADDMuscle 6, 7, S
    DOLL(1).ADDMuscle 8, 9, S
    DOLL(1).ADDMuscle 7, 5, S
    DOLL(1).ADDMuscle 9, 5, S

    DOLL(1).ADDMuscle 10, 5, S * 0.9

    DOLL(1).OBJ_SAVE
    DOLL(1).OBJ_LOADandPlace "obj.doll", PIC.Width / 2, 250
    DOLL(1).SetLinkDrawShape 10, sFillCircle



    DollStrengthBar.Min = 0
    DollStrengthBar.Max = DOLL(1).GlobalMAXStrength * 1000
    DollStrengthBar.Value = DollStrengthBar.Max


    For M = 2 To DOLL(1).NMuscles
        Load MuscleANG(M - 1)
        MuscleANG(M - 1).Visible = True
        MuscleANG(M - 1).Top = MuscleANG(M - 2).Top + MuscleANG(M - 2).Height
        Load mDESC(M - 1)
        mDESC(M - 1).Top = MuscleANG(M - 1).Top
        mDESC(M - 1).Visible = True
    Next

    For M = 1 To DOLL(1).NMuscles
        MuscleANG(M - 1).Min = -PI * 200
        MuscleANG(M - 1).Max = PI * 200
        MuscleANG(M - 1).Value = DOLL(1).MUSCLE_MainANG(M) * 100
    Next

    'knee , hip, elbow, shoulder, head
    mDESC(0) = "L - Knee"
    mDESC(1) = "R - Knee"
    mDESC(2) = "L - Hip"
    mDESC(3) = "R - Hip"
    mDESC(4) = "L - Elbow"
    mDESC(5) = "R - Elbow"
    mDESC(6) = "L - Shoulder"
    mDESC(7) = "R - Shoulder"
    mDESC(8) = "Head"

    GetPosesLIST
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'BALLS
    'Obstacles
    '

    Nobs = 8                      '7 '10 '5
    ReDim Obstacle(Nobs)

    Dim P          As tPoint
    For I = 1 To Nobs
        Obstacle(I).R = Rnd * 10 + 27
        If I = Nobs Then Obstacle(I).R = 50

NewP:
        Obstacle(I).P.X = Rnd * (PIC.Width - Obstacle(I).R * 2) + Obstacle(I).R
        Obstacle(I).P.Y = Rnd * PIC.Height * 2 / 3 + PIC.Height / 3 - 50
        Obstacle(I).P.vX = (Rnd * 3 - 1)    '/ 2
        Obstacle(I).P.vY = (Rnd * 3 - 1)    ' / 2

        For I2 = 1 To DOLL(1).Npoints
            P.X = DOLL(1).PointX(I2)
            P.Y = DOLL(1).PointY(I2)
            If Distance(P, Obstacle(I).P) < Obstacle(I).R Then I2 = 99999999
        Next I2
        If I2 = 99999999 + 1 Then GoTo NewP

        Obstacle(I).MaxX = PIC.Width - Obstacle(I).R - 1
        Obstacle(I).MaxY = PIC.Height - Obstacle(I).R - 1


    Next

    Obstacle(1).IsMotionLess = True
    Obstacle(1).P.vX = 0
    Obstacle(1).P.vY = 0
    Obstacle(2).IsMotionLess = True
    Obstacle(2).P.vX = 0
    Obstacle(2).P.vY = 0

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    '------------------------------------------------------------------
    'ROPE


    DOLL(2).GlobalMAXStrength = S
    DOLL(2).CurrentMAXStrength = S
    DOLL(2).MaxX = PIC.Width - 2
    DOLL(2).MaxY = PIC.Height - 2

    DOLL(2).ADDpoint PIC.Width / 2, 50, True
    For I = 2 To 9
        DOLL(2).ADDpoint PIC.Width / 2 + (I - 1) * 10, 50 + ((I + 1) Mod 2) * 30
    Next

    DOLL(2).ADDLink 1, 2
    DOLL(2).ADDLink 2, 3
    DOLL(2).ADDLink 3, 4
    DOLL(2).ADDLink 4, 5
    DOLL(2).ADDLink 5, 6
    DOLL(2).ADDLink 6, 7
    DOLL(2).ADDLink 7, 8
    DOLL(2).ADDLink 8, 9


    DOLL(2).ADDMuscle 1, 2, DOLL(2).CurrentMAXStrength
    DOLL(2).ADDMuscle 2, 3, DOLL(2).CurrentMAXStrength
    DOLL(2).ADDMuscle 3, 4, DOLL(2).CurrentMAXStrength
    DOLL(2).ADDMuscle 4, 5, DOLL(2).CurrentMAXStrength
    DOLL(2).ADDMuscle 5, 6, DOLL(2).CurrentMAXStrength
    DOLL(2).ADDMuscle 6, 7, DOLL(2).CurrentMAXStrength
    DOLL(2).ADDMuscle 7, 8, DOLL(2).CurrentMAXStrength

    RopeStrengthBar.Min = 0
    RopeStrengthBar.Max = DOLL(2).GlobalMAXStrength * 1000
    RopeStrengthBar.Value = RopeStrengthBar.Max
    RopeStrengthBar.Value = 0

    DOLL(2).OBJ_SAVE "Rope.doll"


    '------------------------------------------------------------------
    'CUBE
    With DOLL(3)

        .GlobalMAXStrength = S
        .CurrentMAXStrength = S
        .MaxX = PIC.Width - 2
        .MaxY = PIC.Height - 2

        .ADDpoint 100, 50, True
        .ADDpoint 150, 50
        .ADDpoint 150, 100
        .ADDpoint 100, 100

        .ADDLink 1, 2, , , , 1
        .ADDLink 2, 3, , , , 2
        .ADDLink 3, 4, , , , 3
        .ADDLink 4, 1, , , , 4

        .ADDMuscle 1, 2, S
        .ADDMuscle 2, 3, S
        .ADDMuscle 3, 4, S
        .ADDMuscle 4, 1, S

        .OBJ_SAVE "Cube.doll"
    End With

    '------------------------------------------------------------------
    'from Doll Editor
    With DOLL(4)
        .MaxX = PIC.Width - 2
        .MaxY = PIC.Height - 2

        .OBJ_LOADandPlace "DollName.doll", PIC.Width - 100, 50, 0.75

    End With




    '--------------------------------------------------------------------

    List1.ListIndex = 1
    List1_Click

    TIMER1.Enabled = True
End Sub





Private Sub Form_Load()
    Me.Caption = Me.Caption & "  V" & App.Major & "." & App.Minor & "  "

    Command1_Click

End Sub



Private Sub DollStrengthBar_Change()

    DOLL(1).CurrentMAXStrength = DollStrengthBar.Value / 1000

End Sub

Private Sub DollStrengthBar_Scroll()
    'doll(1).GlobalMAXStrength = DollStrengthBar.Value / 1000
    DOLL(1).CurrentMAXStrength = DollStrengthBar.Value / 1000
    'For M = 1 To doll(1).NMuscles
    '    doll(1).MUSCLE_SetStrength(M) = DollStrengthBar.Value / 1000
    'Next

End Sub

Private Sub hSticky_Click()


End Sub

Private Sub GlobStrength_Change()

End Sub



Private Sub List1_Click()
    Dim M          As Long

    DOLL(1).OBJ_LoadPose (List1)
    For M = 1 To DOLL(1).NMuscles
        MuscleANG(M - 1).Value = DOLL(1).MUSCLE_MainANG(M) * 100
    Next M

End Sub

Private Sub MuscleANG_Change(Index As Integer)
    MuscleToDraw = Index + 1
    DOLL(1).MUSCLE_MainANG(Index + 1) = CDbl(MuscleANG(Index).Value / 100)


End Sub

Private Sub MuscleANG_Scroll(Index As Integer)
    MuscleToDraw = Index + 1
    DOLL(1).MUSCLE_MainANG(Index + 1) = CDbl(MuscleANG(Index).Value / 100)
End Sub



Private Sub oFilledCircle_Click()
    If oSticky Then TmpDrawShape = sLine
    If oFilledCircle Then TmpDrawShape = sFillCircle
    If oSmile Then TmpDrawShape = sFace
    DOLL(1).SetLinkDrawShape 10, TmpDrawShape

End Sub

Private Sub oSmile_Click()
    If oSticky Then TmpDrawShape = sLine
    If oFilledCircle Then TmpDrawShape = sFillCircle
    If oSmile Then TmpDrawShape = sFace
    DOLL(1).SetLinkDrawShape 10, TmpDrawShape

End Sub

Private Sub oSticky_Click()
    If oSticky Then TmpDrawShape = sLine
    If oFilledCircle Then TmpDrawShape = sFillCircle
    If oSmile Then TmpDrawShape = sFace
    DOLL(1).SetLinkDrawShape 10, TmpDrawShape

End Sub

Private Sub PIC_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim D          As Long
    Dim P          As Long
    Dim D2         As Long
    Dim P2         As Long

    'Me.Caption = KeyCode

    If KeyCode = 65 Then          'A
        D = GetClosestDoll(1, 8)
        MakeJoint 1, D, 8, GetClosestPointOfDoll(1, D, 8)
    End If

    If KeyCode = 68 Then          'D
        D = GetClosestDoll(1, 10)
        MakeJoint 1, D, 10, GetClosestPointOfDoll(1, D, 10)
    End If

    If KeyCode = 32 Then          'Space
        DOLL(1).BreakAtPoint 3    'Int(Rnd * DOLL(1).Npoints) + 1
    End If
    '-------------------------------------
    If KeyCode = 81 Then          'Q
        D = GetDollJoinedTo(1, 8)

        If D <> 0 Then
            P = GetPointJoinedTo(1, 8)
            D2 = GetClosestDoll(D, P, 1)
            If MakeJoint(D, D2, P, GetClosestPointOfDoll(D, D2, P)) Then
                D = GetClosestDoll(1, 8)
                MakeJoint 1, D, 8, GetClosestPointOfDoll(1, D, 8)
            End If
        End If

    End If

    If KeyCode = 69 Then          'E
        D = GetDollJoinedTo(1, 10)

        If D <> 0 Then
            P = GetPointJoinedTo(1, 10)
            D2 = GetClosestDoll(D, P, 1)
            If MakeJoint(D, D2, P, GetClosestPointOfDoll(D, D2, P)) Then
                D = GetClosestDoll(1, 10)
                MakeJoint 1, D, 10, GetClosestPointOfDoll(1, D, 10)
            End If
        End If

    End If


    'Z and C  maybe have little bug
    If KeyCode = 90 Then          'Z
        D = GetDollJoinedTo(1, 8)
        P = GetPointJoinedTo(1, 8)
        If D = 0 Then

            D = GetClosestDoll(1, 8)
            P = GetClosestPointOfDoll(1, D, 8)
        End If
        If D <> 0 Then
            D2 = GetDollJoinedTo(D, P)
            If D2 <> 0 Then

                P2 = GetPointJoinedTo(D, P)

                MakeJoint D, D2, P, P2
            End If
        End If
    End If

    If KeyCode = 67 Then          'C
        D = GetDollJoinedTo(1, 10)
        P = GetPointJoinedTo(1, 10)
        If D = 0 Then
            D = GetClosestDoll(1, 10)
            P = GetClosestPointOfDoll(1, D, 10)
        End If
        If D <> 0 Then
            D2 = GetDollJoinedTo(D, P)
            If D2 <> 0 Then
                P2 = GetPointJoinedTo(D, P)
                MakeJoint D, D2, P, P2
            End If
        End If
    End If

End Sub

Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    Dim Dmin       As Single
    Dim P1         As tPoint
    Dim P2         As tPoint

    Dim I          As Long

    If Button = 1 Then
        Dmin = 1E+19
        P2.X = X
        P2.Y = Y
        For I = 1 To DOLL(1).Npoints
            P1.X = DOLL(1).PointX(I)
            P1.Y = DOLL(1).PointY(I)
            If Distance(P1, P2) < Dmin Then Dmin = Distance(P1, P2): PtoMove = I
        Next
        mouseX = X
        mouseY = Y
        InteractWith = 1
    Else
        Dmin = 1E+19
        P2.X = X
        P2.Y = Y
        For I = 1 To Nobs
            If Not (Obstacle(I).IsMotionLess) Then
                If Distance(Obstacle(I).P, P2) < Dmin Then Dmin = Distance(Obstacle(I).P, P2): PtoMove = I
            End If
        Next
        mouseX = X
        mouseY = Y
        InteractWith = 2
    End If

End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 0 Then
        mouseX = X
        mouseY = Y
    End If

End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InteractWith = 0
End Sub

Private Sub ReStart_Click()
    Dim UB         As Long
    Dim I          As Long

    UB = MuscleANG.Count - 1

    For I = UB To 1 Step -1
        Unload MuscleANG(I)
        Unload mDESC(I)
    Next

    List1.Clear
    For I = 1 To UBound(DOLL)
        DOLL(I).DestroyMe
    Next
    Command1_Click
End Sub

Private Sub RopeStrengthBar_Change()

    DOLL(2).CurrentMAXStrength = RopeStrengthBar.Value / 1000

End Sub

Private Sub RopeStrengthBar_Scroll()

    DOLL(2).CurrentMAXStrength = RopeStrengthBar.Value / 1000

End Sub

Private Sub SavePos_Click()
    DOLL(1).OBJ_SavePose ("POS" & List1.ListCount & ".txt")
    List1.AddItem "POS" & List1.ListCount & ".txt"

End Sub

Private Sub Timer1_Timer()

    Dim D          As Long
    Dim ii         As Long
    Dim I          As Long

    '-------------------------- DRAW
    BitBlt PIC.hdc, 0, 0, PIC.Width, PIC.Height, PIC.hdc, 0, 0, vbBlackness
    DRAWObstacles
    DOLL(1).DRAW PIC.hdc, ShowPointLinkNumbers
    For D = 2 To UBound(DOLL)
        DOLL(D).DRAW PIC.hdc
    Next D
    PIC.Refresh

    '-------------------------- Doll-obsatcle collisions
    For I = 1 To Nobs
        With Obstacle(I)
            For D = 1 To UBound(DOLL)
                'DOLL(D).CheckCircleShapedCollision .P.X, .P.Y, .P.vX, .P.vY, .R, .IsMotionLess
                For ii = 1 To DOLL(D).Npoints
                    DOLL(D).CheckCircleShapedCollision2 ii, .P.X, .P.Y, .P.vX, .P.vY, .R, .IsMotionLess
                Next ii
            Next
        End With
    Next


    DOLL(1).DoPHYSICS frmPHYS.ApplyMuscle = Checked
    For D = 2 To UBound(DOLL)
        DOLL(D).DoPHYSICS True
    Next

    MoveObstacles

    DoMouseForces InteractWith

    DoJoints

    frmPHYS.NRGbar.Value = DOLL(1).MUSCLE_GetStrength(1) * 100

End Sub

Sub GetPosesLIST()
    Dim D          As String

    D = Dir(App.Path & "\Pos" & "*.txt")
    While D <> ""
        List1.AddItem D
        D = Dir
    Wend

End Sub

Sub DoMouseForces(DollOrObstacle)

    If DollOrObstacle = 1 Then
        DOLL(1).PointVX(PtoMove) = DOLL(1).PointVX(PtoMove) - (DOLL(1).PointX(PtoMove) - mouseX) * 0.015
        DOLL(1).PointVY(PtoMove) = DOLL(1).PointVY(PtoMove) - (DOLL(1).PointY(PtoMove) - mouseY) * 0.015
        'PIC.Line (doll(1).PointX(PtoMove), doll(1).PointY(PtoMove))-(mouseX, mouseY), vbYellow
        FastLine PIC.hdc, DOLL(1).PointX(PtoMove), DOLL(1).PointY(PtoMove), mouseX \ 1, mouseY \ 1, 2, vbYellow
    ElseIf DollOrObstacle = 2 Then
        With Obstacle(PtoMove)
            .P.vX = .P.vX - (.P.X - mouseX) * 0.003
            .P.vY = .P.vY - (.P.Y - mouseY) * 0.003
            FastLine PIC.hdc, .P.X \ 1, .P.Y \ 1, mouseX \ 1, mouseY \ 1, 2, vbYellow
        End With
    End If



End Sub

Private Sub Timer2_Timer()

    'For M = 1 To doll(1).NMuscles
    'doll(1).MUSCLE_MainANG(M) = doll(1).MUSCLE_MainANG(M) + 0.05 * IIf(M Mod 2, -1, 1)
    'Next M

End Sub

Sub DRAWObstacles()
    Dim O          As Long


    For O = 1 To Nobs
        With Obstacle(O)
            MyCircle PIC.hdc, .P.X \ 1, .P.Y \ 1, .R \ 1, 2, BallsColor    '85675 'vbCyan

            If .IsMotionLess Then MyCircle PIC.hdc, .P.X \ 1, .P.Y \ 1, 5, 2, vbRed
            'Filled [Flicks to Much :-( ]
            'MyCircle PIC.hdc, Obstacle(o).P.x, Obstacle(o).P.y, Obstacle(o).R / 2, Obstacle(o).R, BallsColor'85675 ' vbCyan
            'PIC.Circle (Obstacle(O).P.X, Obstacle(O).P.y), Obstacle(O).R, vbCyan
        End With
    Next O

End Sub
Sub MoveObstacles()
    Dim O          As Long
    Dim O2         As Long

    For O = 1 To Nobs

        With Obstacle(O)

            If .IsMotionLess = False Then
                .P.vX = .P.vX * Obstacle_Air_Resistence
                .P.vY = .P.vY * Obstacle_Air_Resistence
                .P.vY = .P.vY + Gravity
                .P.X = .P.X + .P.vX
                .P.Y = .P.Y + .P.vY

                For O2 = 1 To Nobs
                    If O2 <> O Then ChangeVelocities Obstacle(O), Obstacle(O2)
                Next

                If .P.X > .MaxX Then
                    .P.X = .MaxX  'PIC.Width - .R
                    .P.vX = -.P.vX * CollLostEnergy
                End If
                If .P.X < .R Then
                    .P.X = .R
                    .P.vX = -.P.vX * CollLostEnergy
                End If

                If .P.Y > .MaxY Then
                    .P.Y = .MaxY
                    .P.vY = -.P.vY * CollLostEnergy
                End If
                'If .P.y < .R Then
                '    .P.y = .R
                '    .P.Vy = -.P.Vy*CollLostEnergy
                'End If

            End If

        End With

    Next O

End Sub

