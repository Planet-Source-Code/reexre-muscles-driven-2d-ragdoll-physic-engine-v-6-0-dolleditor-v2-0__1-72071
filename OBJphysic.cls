VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OBJphysic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Public Npoints
Private Point()    As tPoint

Public Nlinks
Private Link()     As tLink

Public NMuscles
Private MUSCLE()   As tMuscle


Public GlobalMAXStrength As Double
Public CurrentMAXStrength As Double


Public MaxX        As Long
Public MaxY        As Long
Public UseMuscles  As Boolean


Private LinkDrawOrder() As Long


Public Function MUSCLE_AddToMainANG(M, ValADD)
    MUSCLE(M).MainA = MUSCLE(M).MainA + ValADD
End Function

Public Property Let MUSCLE_SetStrength(M, Strength)
    MUSCLE(M).f = Strength
End Property
Public Property Get MUSCLE_GetStrength(M)
    If (M < 0) Or (M > NMuscles) Then MUSCLE_GetStrength = 0: Exit Property
    MUSCLE_GetStrength = MUSCLE(M).f
End Property
Public Property Get MUSCLE_MainANG(M)
    MUSCLE_MainANG = MUSCLE(M).MainA
End Property

Public Property Let MUSCLE_MainANG(M, Value)
    MUSCLE(M).MainA = Value
End Property

Public Property Get PointVX(P)
    PointVX = Point(P).vX
End Property
Public Property Get PointVY(P)
    PointVY = Point(P).vY
End Property
Public Property Get PointX(P)
    PointX = Point(P).X
End Property
Public Property Get PointY(P)
    PointY = Point(P).Y
End Property

Public Property Let PointVX(P, V)
    Point(P).vX = V
End Property
Public Property Let PointVY(P, V)
    Point(P).vY = V
End Property
Public Property Let PointX(P, V)
    Point(P).X = V
End Property
Public Property Let PointY(P, V)
    Point(P).Y = V
End Property

Public Property Get PointIsFix(P)
    PointIsFix = Point(P).IsMotionLess
End Property

Public Property Get Link_P1(L)
    Link_P1 = Link(L).P1
End Property
Public Property Get Link_P2(L)
    Link_P2 = Link(L).P2
End Property
Public Property Get Link_MainL(L)
    Link_MainL = Link(L).MainL
End Property


Public Property Get Muscle_L1(M)
    Muscle_L1 = MUSCLE(M).L1
End Property

Public Property Get Muscle_L2(M)
    Muscle_L2 = MUSCLE(M).L2
End Property

Public Property Get Muscle_P0(M)
    Muscle_P0 = MUSCLE(M).P0
End Property
Public Property Get Muscle_P1(M)
    Muscle_P1 = MUSCLE(M).P1
End Property
Public Property Get Muscle_P2(M)
    Muscle_P2 = MUSCLE(M).P2
End Property

Public Sub SetPointVelXY(P, vX, vY)
    Point(P).vX = vX
    Point(P).vY = vY

End Sub

Public Function ADDpoint(X, Y, Optional IsMotionLess = False)

    Npoints = Npoints + 1
    ReDim Preserve Point(Npoints)

    Point(Npoints).X = X
    Point(Npoints).Y = Y
    Point(Npoints).IsMotionLess = IsMotionLess

End Function

Public Function ADDLink(P1, P2, Optional Thickness = 2, _
                        Optional Color = vbGreen, Optional Shape As tDrawShape = sLine, _
                        Optional ByVal LinkToDraw = 0, Optional BreakMaxStress = 1.5)
    Dim P          As Long
    Dim L          As Long

    Nlinks = Nlinks + 1

    If LinkToDraw = 0 Then LinkToDraw = Nlinks

    ReDim Preserve Link(Nlinks)
    ReDim Preserve LinkDrawOrder(Nlinks)

    With Link(Nlinks)

        .P1 = P1
        .P2 = P2
        .MainL = Distance(Point(P1), Point(P2))
        .Thickness = Thickness
        .Color = Color
        .Shape = Shape
        .MaxStress = BreakMaxStress

        .InvMass = 150 / (.MainL * .Thickness)


        LinkDrawOrder(Nlinks) = LinkToDraw


        Debug.Print Nlinks & "   " & .MainL

    End With


    For P = 1 To Npoints
        Point(P).InvMass = 0
        For L = 1 To Nlinks
            If Link(L).P1 = P Or Link(L).P2 = P Then
                Point(P).InvMass = Point(P).InvMass + 1 / Link(L).InvMass
            End If
        Next
        If Point(P).InvMass <> 0 Then Point(P).InvMass = 1 / Point(P).InvMass
    Next

End Function

Public Function ADDMuscle(L1, L2, Strength) As Boolean

    NMuscles = NMuscles + 1
    ReDim Preserve MUSCLE(NMuscles)

    ADDMuscle = False

    With MUSCLE(NMuscles)
        .L1 = L1
        .L2 = L2
        .f = Strength
        .isNotBroken = True

        If Link(L1).P1 = Link(L2).P1 Then
            .P0 = Link(L1).P1
            .P1 = Link(L1).P2
            .P2 = Link(L2).P2
            ADDMuscle = True
        End If

        If Link(L1).P1 = Link(L2).P2 Then
            .P0 = Link(L1).P1
            .P1 = Link(L1).P2
            .P2 = Link(L2).P1
            ADDMuscle = True
        End If

        If Link(L1).P2 = Link(L2).P1 Then
            .P0 = Link(L1).P2
            .P1 = Link(L1).P1
            .P2 = Link(L2).P2
            ADDMuscle = True
        End If

        If Link(L1).P2 = Link(L2).P2 Then
            .P0 = Link(L1).P2
            .P1 = Link(L1).P1
            .P2 = Link(L2).P1
            ADDMuscle = True
        End If
        'Stop

    End With

    If ADDMuscle = False Then
        NMuscles = NMuscles - 1
        ReDim Preserve MUSCLE(NMuscles)
        Exit Function
    End If



    With MUSCLE(NMuscles)
        .MainA = Atan2(Point(.P1).X - Point(.P0).X, Point(.P1).Y - Point(.P0).Y) - _
                 Atan2(Point(.P2).X - Point(.P0).X, Point(.P2).Y - Point(.P0).Y)
    End With


End Function




Public Sub DRAW(ByRef PicHdc As Long, Optional Extended = False)
    Dim x1         As Long
    Dim x2         As Long
    Dim y1         As Long
    Dim y2         As Long
    Dim X3         As Long
    Dim Y3         As Long

    Dim Xc         As Long
    Dim Yc         As Long
    Dim R          As Long

    Dim A          As Double

    Dim Smile      As Double

    Dim Color2     As Long

    Dim I          As Long

    'P.Circle (Point(MUSCLE(MuscleToDraw).P0).X, (Point(MUSCLE(MuscleToDraw).P0).y)), 5, vbRed
    'P.Line (Point(MUSCLE(MuscleToDraw).P1).X, (Point(MUSCLE(MuscleToDraw).P1).y))- _
     (Point(MUSCLE(MuscleToDraw).P2).X, (Point(MUSCLE(MuscleToDraw).P2).y)), vbRed

    For I = 1 To Npoints
        If Point(I).IsMotionLess Then MyCircle PicHdc, Point(I).X \ 1, Point(I).Y \ 1, 5, 2, vbRed
    Next I


    For I = 1 To Nlinks           'Nlinks To 1 Step -1
        'For i = Nlinks To 1 Step -1
        With Link(LinkDrawOrder(I))



            x1 = Point(.P1).X
            y1 = Point(.P1).Y
            x2 = Point(.P2).X
            y2 = Point(.P2).Y
            'P.Line (X1, Y1)-(X2, Y2), vbGreen
            '    Stop

            Select Case .Shape

                Case sLine
                    FastLine PicHdc, x1, y1, x2, y2, .Thickness \ 1, .Color
                Case sFillCircle

                    Xc = (x1 + x2) * 0.5
                    Yc = (y1 + y2) * 0.5
                    R = .MainL * 0.24    '0.25
                    MyCircle PicHdc, Xc, Yc, R, R * 2, .Color
                Case sFace
                    Color2 = Not (.Color)
                    Xc = (x1 + x2) * 0.5
                    Yc = (y1 + y2) * 0.5
                    R = .MainL * 0.24    '0.25
                    MyCircle PicHdc, Xc, Yc, R, R * 2, .Color

                    A = Atan2(x2 - x1, y2 - y1)
                    x2 = Xc - Cos(A + 0.85) * .MainL * 0.24
                    y2 = Yc - Sin(A + 0.85) * .MainL * 0.24
                    MyCircle PicHdc, x2, y2, 2, 1, Color2
                    x2 = Xc - Cos(A - 0.85) * .MainL * 0.24
                    y2 = Yc - Sin(A - 0.85) * .MainL * 0.24
                    MyCircle PicHdc, x2, y2, 2, 1, Color2

                    x1 = Xc + Cos(A + 1.2) * .MainL * 0.36
                    y1 = Yc + Sin(A + 1.2) * .MainL * 0.36
                    x2 = Xc + Cos(A - 1.2) * .MainL * 0.36
                    y2 = Yc + Sin(A - 1.2) * .MainL * 0.36


                    Smile = MUSCLE(1).f / (CurrentMAXStrength + 0.0001) - 0.5
                    X3 = Xc + Cos(A) * .MainL * (0.08 + Smile * 0.5)
                    Y3 = Yc + Sin(A) * .MainL * (0.08 + Smile * 0.5)

                    FastLine PicHdc, x1, y1, X3, Y3, 1, Color2
                    FastLine PicHdc, X3, Y3, x2, y2, 1, Color2


                Case sCircle
                    Xc = (x1 + x2) * 0.5
                    Yc = (y1 + y2) * 0.5
                    R = .MainL * 0.5
                    MyCircle PicHdc, Xc, Yc, R, .Thickness \ 1, .Color

            End Select


            If Extended Then
                'x1 = Point(Link(I).P1).X
                'y1 = Point(Link(I).P1).Y
                'x2 = Point(Link(I).P2).X
                'y2 = Point(Link(I).P2).Y

                'P.Circle (x1, y1), 2, vbYellow
                'P.CurrentX = x1
                'P.CurrentY = y1 - 10
                'P.ForeColor = vbYellow '
                'P.Print Link(I).P1

                'P.Circle (x2, y2), 2, vbYellow
                'P.CurrentX = x2
                'P.CurrentY = y2 - 10
                'P.Print Link(I).P2

                'x1 = (x1 + x2) / 2
                'y1 = (y1 + y2) / 2
                'P.CurrentX = x1
                'P.CurrentY = y1 - 10
                'P.ForeColor = vbGreen
                'P.Print "-" & I & "-" '
            End If
        End With

    Next I

End Sub


Public Sub DoPHYSICS(MustUseMuscles As Boolean)
    UseMuscles = MustUseMuscles
    'Limiter
    Dim I          As Long



    For I = 1 To Npoints
        With Point(I)
            .OldX = .X
            .OldY = .Y
        End With
    Next


    DoForces
    DoMuscles

    ''''Update X Y   and Collision
    For I = 1 To Npoints
        With Point(I)


            If .IsMotionLess Then
                .vX = 0
                .vY = 0
            Else
                .vX = .vX * Doll_Air_Resistence
                .vY = .vY * Doll_Air_Resistence
                .vY = .vY + Gravity
                .X = .X + .vX
                .Y = .Y + .vY
            End If

            If .Y > MaxY Then

                .Y = .OldY
                .vY = -.vY * 1.05    ' 0.9
                .vX = .vX * 0.5
            End If

            'If .y < 50 Then: .y = 50: .vx = 0: .vy = .vy = 0

            If .X < 3 Or .X > MaxX Then
                .X = .OldX
                .vX = -.vX * 0.9
            End If
        End With
    Next I

End Sub

Private Sub DoForces()
    'VERLET
    Dim Mx         As Double
    Dim My         As Double

    Dim P1x        As Double
    Dim P1y        As Double
    Dim P2x        As Double
    Dim P2y        As Double
    Dim A          As Double

    Dim L          As Long


    For L = 1 To Nlinks
        With Link(L)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If Distance(Point(.P1), Point(.P2)) / .MainL > .MaxStress Then
                ''''''''''''''''''''''BreakAtPoint IIf(Rnd < 0.5, .P1, .P2)
                ''BreakAtPoint .P2
            End If



            Mx = (Point(.P1).X + Point(.P2).X) * 0.5
            My = (Point(.P1).Y + Point(.P2).Y) * 0.5



            'mx,my middle point of link L

            A = Atan2(Point(.P2).X - Point(.P1).X, Point(.P2).Y - Point(.P1).Y)

            'To preserve Link Lenght
            P1x = Mx - Cos(A) * .MainL * 0.5
            P1y = My - Sin(A) * .MainL * 0.5
            P2x = Mx + Cos(A) * .MainL * 0.5
            P2y = My + Sin(A) * .MainL * 0.5

            'Find Velocities
            Point(.P1).vX = Point(.P1).vX + (P1x - Point(.P1).X)
            Point(.P1).vY = Point(.P1).vY + (P1y - Point(.P1).Y)
            Point(.P2).vX = Point(.P2).vX + (P2x - Point(.P2).X)
            Point(.P2).vY = Point(.P2).vY + (P2y - Point(.P2).Y)


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End With
    Next L
End Sub

Private Sub DoMuscles()
    Dim x1         As Double
    Dim y1         As Double
    Dim x2         As Double
    Dim y2         As Double
    Dim AM         As Double
    Dim Dx1        As Double
    Dim Dy1        As Double
    Dim Dx2        As Double
    Dim Dy2        As Double

    Dim Dx1Move    As Double
    Dim Dy1Move    As Double
    Dim Dx2Move    As Double
    Dim Dy2Move    As Double

    Dim M          As Long
    Dim A1         As Double
    Dim A2         As Double

    If UseMuscles Then

        For M = 1 To NMuscles

            With MUSCLE(M)

                If .isNotBroken Then
                    A1 = Atan2(Point(.P1).X - Point(.P0).X, Point(.P1).Y - Point(.P0).Y)
                    A2 = Atan2(Point(.P2).X - Point(.P0).X, Point(.P2).Y - Point(.P0).Y)

                    AM = (A1 - A2) - .MainA

                    x1 = Point(.P0).X + Cos(A1 - AM) * Link(.L1).MainL
                    y1 = Point(.P0).Y + Sin(A1 - AM) * Link(.L1).MainL
                    x2 = Point(.P0).X + Cos(A2 + AM) * Link(.L2).MainL
                    y2 = Point(.P0).Y + Sin(A2 + AM) * Link(.L2).MainL

                    Dx1 = (Point(.P1).X - x1)
                    Dy1 = (Point(.P1).Y - y1)
                    Dx2 = (Point(.P2).X - x2)
                    Dy2 = (Point(.P2).Y - y2)

                    Dx1Move = Dx1 * Link(.L1).InvMass
                    Dy1Move = Dy1 * Link(.L1).InvMass
                    Dx2Move = Dx2 * Link(.L2).InvMass
                    Dy2Move = Dy2 * Link(.L2).InvMass

                    Point(.P1).vX = Point(.P1).vX - .f * Dx1Move
                    Point(.P1).vY = Point(.P1).vY - .f * Dy1Move
                    Point(.P2).vX = Point(.P2).vX - .f * Dx2Move
                    Point(.P2).vY = Point(.P2).vY - .f * Dy2Move

                    Point(.P0).vX = Point(.P0).vX + .f * (Dx1Move + Dx2Move)
                    Point(.P0).vY = Point(.P0).vY + .f * (Dy1Move + Dy2Move)


                    'Applay a Decrement of Muscles Energy due to the "work" Done
                    ChangeStrengthsBy -0.018 * .f * (Sqr(Dx1 * Dx1 + Dy1 * Dy1) + Sqr(Dx2 * Dx2 + Dy2 * Dy2))
                    ' -0.018
                    '0.05 .f  Sqr(Dx1 * Dx1 + Dy1 * Dy1)

                End If
            End With

        Next M

    End If

    'Constantly Increase Muscles Energy (untill a Max Value)
    'as we do breathing
    ChangeStrengthsBy 0.009 * NMuscles
    '0.01

End Sub




Public Function LINK_GetWorldAngle(L) As Double

    Dim P1         As Long
    Dim P2         As Long

    P1 = Link(L).P1
    P2 = Link(L).P2

    LINK_GetWorldAngle = Atan2(Point(P2).X - Point(P1).X, Point(P2).Y - Point(P1).Y)


End Function

Public Sub OBJ_SAVE(Optional Filename = "Obj.doll")
    Dim Cx         As Double
    Dim Cy         As Double
    Dim P          As Long
    Dim L          As Long
    Dim M          As Long


    Cx = 0
    Cy = 0
    For P = 1 To Npoints
        With Point(P)
            Cx = Cx + .X
            Cy = Cy + .Y
        End With
    Next P
    Cx = Cx / Npoints
    Cy = Cy / Npoints

    Open App.Path & "\" & Filename For Output As 1

    Print #1, Replace(GlobalMAXStrength, ",", ".")
    Print #1, Replace(CurrentMAXStrength, ",", ".")

    Print #1, Npoints
    Print #1, Nlinks
    Print #1, NMuscles

    Print #1, "Points:"
    For P = 1 To Npoints
        Print #1, Replace(Point(P).X - Cx, ",", ".")
        Print #1, Replace(Point(P).Y - Cy, ",", ".")
        Print #1, IIf(Point(P).IsMotionLess, "True", "False")
    Next

    Print #1, "Links:"
    For L = 1 To Nlinks
        Print #1, Link(L).P1
        Print #1, Link(L).P2
        Print #1, Link(L).Thickness
        Print #1, Link(L).Color
        Print #1, Link(L).Shape
        Print #1, LinkDrawOrder(L)
        Print #1, Replace(Link(L).MaxStress, ",", ".")

    Next

    Print #1, "Muscles:"
    For M = 1 To NMuscles
        Print #1, MUSCLE(M).L1
        Print #1, MUSCLE(M).L2
        Print #1, Replace(MUSCLE(M).f, ",", ".")
    Next

    Close 1

End Sub

Public Sub OBJ_LOADandPlace(Filename As String, CenterX, CenterY, Optional dScale = 1)
    Dim S          As String
    Dim S2         As String
    Dim P          As Long
    Dim L          As Long
    Dim M          As Long

    Dim V1         As Double
    Dim V2         As Double
    Dim V3         As Double
    Dim V4         As Double
    Dim V5         As tDrawShape
    Dim V6         As Long
    Dim V7         As String

    Dim NP         As Long
    Dim nL         As Long
    Dim nM         As Long

    Open App.Path & "\" & Filename For Input As 1

    Input #1, GlobalMAXStrength
    Input #1, CurrentMAXStrength



    Input #1, NP                  'Npoints
    Input #1, nL                  'Nlinks
    Input #1, nM                  'NMuscles

    Npoints = 0
    Nlinks = 0
    NMuscles = 0


    Input #1, S
    For P = 1 To NP
        Input #1, V1              'Point(p).x
        Input #1, V2              'Point(p).y
        Input #1, S2
        ADDpoint V1 * dScale + CenterX, V2 * dScale + CenterY, IIf(S2 = "True", True, False)
    Next

    Input #1, S
    For L = 1 To nL
        Input #1, V1              'Link(l).P1
        Input #1, V2              'Link(l).P2
        Input #1, V3
        Input #1, V4
        Input #1, V5
        Input #1, V6
        Input #1, V7

        ADDLink V1, V2, V3, V4, V5, V6, Val(Replace(V7, ",", "."))


    Next

    Input #1, S
    For M = 1 To nM
        Input #1, V1              'MUSCLE(M).L1
        Input #1, V2              'MUSCLE(M).L2
        Input #1, V3              'MUSCLE(M).F
        ADDMuscle V1, V2, V3
    Next

    Close 1

End Sub

Public Sub OBJ_SavePose(Filename As String)
    Dim I          As Long

    Open App.Path & "\" & Filename For Output As 1
    Print #1, NMuscles
    For I = 1 To NMuscles
        Print #1, Replace(CStr(MUSCLE(I).MainA), ",", ".")
    Next I

    Close 1

End Sub
Public Sub OBJ_LoadPose(Filename As String)
    Dim nM         As Long
    Dim I          As Long


    Open App.Path & "\" & Filename For Input As 1
    Input #1, nM
    If nM <> NMuscles Then Close 1: Exit Sub
    For I = 1 To NMuscles
        Input #1, MUSCLE(I).MainA
    Next I

    Close 1

End Sub


Public Sub CheckCircleShapedCollision(ByRef x1, ByRef y1, ByRef vX, ByRef vY, Radius, Optional IsMotionLess = False)
    Dim P          As tPoint
    Dim A          As Double
    Dim Obstacle_Mass As Double
    Dim I          As Long

    Const Div1000  As Single = 1 / 1000


    P.X = x1
    P.Y = y1


    For I = 1 To Npoints

        If Distance(Point(I), P) < Radius Then


            'This is not Realistic physic. Anyway works good.
            ' Replaced with CheckCircleShapedCollision2


            A = Atan2(Point(I).X - P.X, Point(I).Y - P.Y)
            Obstacle_Mass = Radius * Radius


            If Not (Point(I).IsMotionLess) Then

                Point(I).X = Point(I).OldX + Cos(A) * 0.5 + vX * Obstacle_Mass * Div1000
                Point(I).Y = Point(I).OldY + Sin(A) * 0.5 + vY * Obstacle_Mass * Div1000

                Point(I).vX = -Point(I).vX + vX * Obstacle_Mass * Div1000
                Point(I).vY = -Point(I).vY + vY * Obstacle_Mass * Div1000


                If Not IsMotionLess Then
                    'X1 = X1 - Vx
                    'Y1 = Y1 - Vy
                    vX = vX - Point(I).vX * 500 / (Obstacle_Mass)
                    vY = vY - Point(I).vY * 500 / (Obstacle_Mass)
                End If

            Else
                If Not IsMotionLess Then
                    x1 = x1 - vX
                    y1 = y1 - vY
                    vX = -vX
                    vY = -vY
                End If
            End If

        End If
    Next

End Sub

Public Sub CheckCircleShapedCollision2(wP, ByRef BallX, ByRef BallY, ByRef BallVX, ByRef BallVY, Radius, Optional IsMotionLess As Boolean = False)
    'Do Real Physic

    Dim BallPoint  As tPoint

    BallPoint.X = BallX
    BallPoint.Y = BallY
    BallPoint.vX = BallVX
    BallPoint.vY = BallVY
    If Distance(Point(wP), BallPoint) < Radius Then
        CollisionReact Point(wP), BallPoint, Radius, IsMotionLess
        BallX = BallPoint.X
        BallY = BallPoint.Y
        BallVX = BallPoint.vX
        BallVY = BallPoint.vY
    End If
End Sub

Private Function ChangeStrengthsBy(PercOfMaxStrength)


    Dim I          As Long
    Dim Value      As Double


    If NMuscles = 0 Then Exit Function

    Value = GlobalMAXStrength * PercOfMaxStrength

    MUSCLE(1).f = MUSCLE(1).f + Value
    If MUSCLE(1).f < 0 Then MUSCLE(1).f = 0
    If MUSCLE(1).f > CurrentMAXStrength Then MUSCLE(1).f = CurrentMAXStrength
    For I = 2 To NMuscles
        MUSCLE(I).f = MUSCLE(1).f
    Next

End Function


Public Function DestroyMe()

    Npoints = 0
    Nlinks = 0
    NMuscles = 0

End Function



Public Sub SetLinkDrawShape(L, hType As tDrawShape)
    Link(L).Shape = hType

End Sub
Private Sub BreakMuscleAtPoint(P)
    'This Remove Muscles with .P0=P
    'This Works.

    Dim I          As Long
    Dim I2         As Long

    If NMuscles > 0 Then
        I = 1
        Do
            If MUSCLE(I).P0 = P Then
                'mode 1
                'For I2 = I To NMuscles - 1
                '    MUSCLE(I2) = MUSCLE(I2 + 1)
                'Next I2
                'NMuscles = NMuscles - 1
                'I = I - 1

                'mode2 ' better
                MUSCLE(I).isNotBroken = False

            End If
            I = I + 1
        Loop While I <= NMuscles
        ReDim Preserve MUSCLE(NMuscles)
    End If
End Sub

Public Sub BreakAtPoint(P)
    Dim tmpNlinks
    Dim I          As Long
    Dim Im         As Long
    Dim tmpP       As tPoint

    Dim IsOrphan() As Boolean
    Dim FirstPoint

    Dim L1         As Long
    Dim L2         As Long


    BreakMuscleAtPoint P

    tmpNlinks = Nlinks

    tmpP = Point(P)

    '-----------------------------------------------
    'BUG
    'To Better develope: Don't know , but it seems that in Some Cases
    'this sub creates useless/unused point(s).
    'Can you Help Me?
    '
    FirstPoint = True
    For I = 1 To tmpNlinks
        If Link(I).P1 = P And FirstPoint = True Then
            Point(P) = tmpP
            FirstPoint = False
        End If
        If Link(I).P2 = P And FirstPoint = True Then
            Point(P) = tmpP
            FirstPoint = False
        End If
        If Link(I).P1 = P And FirstPoint = False Then
            Npoints = Npoints + 1
            ReDim Preserve Point(Npoints)
            Point(Npoints) = tmpP
            Link(I).P1 = Npoints
        End If
        If Link(I).P2 = P And FirstPoint = False Then
            Npoints = Npoints + 1
            ReDim Preserve Point(Npoints)
            Point(Npoints) = tmpP
            Link(I).P2 = Npoints
        End If

    Next I
    '-----------------------------------------------
    'frmPHYS.Caption = Npoints



    'this is
    'ok
    For Im = 1 To NMuscles
        With MUSCLE(Im)
            L1 = .L1
            L2 = .L2
            If Link(L1).P1 = Link(L2).P1 Then
                .P0 = Link(L1).P1
                .P1 = Link(L1).P2
                .P2 = Link(L2).P2
            End If

            If Link(L1).P1 = Link(L2).P2 Then
                .P0 = Link(L1).P1
                .P1 = Link(L1).P2
                .P2 = Link(L2).P1
            End If

            If Link(L1).P2 = Link(L2).P1 Then
                .P0 = Link(L1).P2
                .P1 = Link(L1).P1
                .P2 = Link(L2).P2
            End If

            If Link(L1).P2 = Link(L2).P2 Then
                .P0 = Link(L1).P2
                .P1 = Link(L1).P1
                .P2 = Link(L2).P1
            End If
        End With
    Next Im

End Sub

