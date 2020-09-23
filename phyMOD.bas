Attribute VB_Name = "phyMOD"
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

Public Enum tDrawShape
    sLine
    sFace
    sFillCircle
    sCircle
End Enum

Public Type tPoint
    X              As Double
    Y              As Double
    OldX           As Double
    OldY           As Double
    vX             As Double
    vY             As Double

    IsMotionLess   As Boolean

    InvMass        As Double

End Type



Public Type tLink
    P1             As Long        '     Point 1
    P2             As Long        '     Point 2
    MainL          As Double      '   Distance Between P1 and P2
    'For Drawing:
    Color          As Long        'Color
    Thickness      As Integer     'Line Thickness
    Shape          As tDrawShape  'Draw Shape [As word says, only shape, not collision. Collision is managed as 'Line' ]

    MaxStress      As Double

    InvMass        As Double

End Type

Public Type tMuscle
    L1             As Long        '     Link1
    L2             As Long        '     Link2
    MainA          As Double      '   Angle that should be between L1 and L2
    P0             As Long        '     Common point of L1 and L2
    P1             As Long        '     Other point on L1
    P2             As Long        '     Other point on L2
    f              As Double      '       Muscle Force(strength)

    isNotBroken    As Boolean

End Type


Public Type TObstacle
    P              As tPoint
    R              As Double
    IsMotionLess   As Boolean
    MaxX           As Double
    MaxY           As Double
End Type


Public Const PI2   As Double = 6.28318530717959
Public Const PI    As Double = 3.14159265358979
Public Const PIh   As Double = 1.5707963267949


Public Doll_Air_Resistence As Double
Public Obstacle_Air_Resistence As Double
Public Gravity     As Double


Public MuscleToDraw As Long


Public TmpDrawShape As tDrawShape

Public DOLL()      As New OBJphysic

Public Const CollLostEnergy = 0.97    ' 0.975

Public Function Distance(P1 As tPoint, P2 As tPoint) As Double
    Dim dX         As Double
    Dim dY         As Double

    dX = P1.X - P2.X
    dY = P1.Y - P2.Y

    Distance = Sqr(dX * dX + dY * dY)

End Function
Public Function Atan2(X As Double, Y As Double) As Double
    If X Then
        Atan2 = -PI + Atn(Y / X) - (X > 0) * PI
    Else
        Atan2 = -PIh - (Y > 0) * PI
    End If

    ' While Atan2 < 0: Atan2 = Atan2 + Pi2: Wend
    ' While Atan2 > Pi2: Atan2 = Atan2 - Pi2: Wend

    If Atan2 < 0 Then Atan2 = Atan2 + PI2

End Function

'Public Function Atan2(ByVal dX As Double, ByVal dY As Double) As Double
'    'This Should return Angle
'
'    Dim theta      As Double
'
'    If (Abs(dX) < 0.0000001) Then
'        If (Abs(dY) < 0.0000001) Then
'            theta = 0#
'        ElseIf (dY > 0#) Then
'            theta = 1.5707963267949
'            'theta = PI / 2
'        Else
'            theta = -1.5707963267949
'            'theta = -PI / 2
'        End If
'    Else
'        theta = Atn(dY / dX)
'
'        If (dX < 0) Then
'            If (dY >= 0#) Then
'                theta = PI + theta
'            Else
'                theta = theta - PI
'            End If
'        End If
'    End If'
'
'    Atan2 = theta
'End Function



Public Function GetClosestPointOfDoll(Doll_1, TargetDoll, PointOfDoll_1)
    Dim D          As Double
    Dim Dmin       As Double
    Dim P1         As tPoint
    Dim P2         As tPoint
    Dim I          As Long

    P1.X = DOLL(Doll_1).PointX(PointOfDoll_1)
    P1.Y = DOLL(Doll_1).PointY(PointOfDoll_1)

    Dmin = 9999999999999#
    For I = 1 To DOLL(TargetDoll).Npoints
        P2.X = DOLL(TargetDoll).PointX(I)
        P2.Y = DOLL(TargetDoll).PointY(I)
        D = Distance(P1, P2)
        If D < Dmin Then Dmin = D: GetClosestPointOfDoll = I
    Next

End Function

Public Function GetClosestDoll(Doll_from, PointFrom, Optional ExludedDollFromSearch = 0)

    Dim D          As Double
    Dim Dmin       As Double
    Dim P1         As tPoint
    Dim P2         As tPoint
    Dim I          As Long
    Dim TargetDoll As Long

    P1.X = DOLL(Doll_from).PointX(PointFrom)
    P1.Y = DOLL(Doll_from).PointY(PointFrom)

    Dmin = 9999999999999#
    For TargetDoll = 1 To UBound(DOLL)
        If TargetDoll <> Doll_from And TargetDoll <> ExludedDollFromSearch Then
            For I = 1 To DOLL(TargetDoll).Npoints
                P2.X = DOLL(TargetDoll).PointX(I)
                P2.Y = DOLL(TargetDoll).PointY(I)
                D = Distance(P1, P2)
                If D < Dmin Then Dmin = D: GetClosestDoll = TargetDoll
            Next I
        End If
    Next TargetDoll

End Function


Public Sub CollisionReact(ByRef A As tPoint, ByRef B As tPoint, B_R, BisMotionless As Boolean)
    'A is a point on Doll
    'B is circled shaped obstacle

    Dim Dis        As Double
    Dim Angle      As Double
    Dim Vx1        As Double
    Dim Vy1        As Double
    Dim Vx2        As Double
    Dim Vy2        As Double
    Dim Ret_VX1    As Double
    Dim Ret_VY1    As Double
    Dim Ret_VX2    As Double
    Dim Ret_VY2    As Double

    Dim A_Mass     As Double
    Dim B_Mass     As Double

    A.X = A.OldX
    A.Y = A.OldY


    A_Mass = 200 / A.InvMass
    'Increase 200 to make Doll Less React to Obstacles
    'Decrease 200 to make Doll More React to Obstacles


    B_Mass = IIf(BisMotionless, 999999999999#, B_R * B_R)

    Separate_Point_Ball A, B, B_R
    'get the angle between the positions of the balls
    Angle = Atan2(B.X - A.X, B.Y - A.Y)

    Vx1 = A.vX
    Vy1 = A.vY
    Vx2 = B.vX
    Vy2 = B.vY
    'resolve the velocitis such that they are along the line of collision
    Ret_VX1 = Vx1 * Cos(-Angle) - Vy1 * Sin(-Angle)
    Ret_VY1 = Vx1 * Sin(-Angle) + Vy1 * Cos(-Angle)
    Ret_VX2 = Vx2 * Cos(-Angle) - Vy2 * Sin(-Angle)
    Ret_VY2 = Vx2 * Sin(-Angle) + Vy2 * Cos(-Angle)
    'swap the horizontal components of the velocities
    '(do any momemtum calculations here)
    Vx1 = (Ret_VX1 * (A_Mass - B_Mass) + (Ret_VX2 * 2 * B_Mass)) / (A_Mass + B_Mass)
    Vx2 = ((Ret_VX1 * 2 * A_Mass) + Ret_VX2 * (A_Mass - B_Mass)) / (A_Mass + B_Mass)
    'keep the vertical component the same
    Vy1 = Ret_VY1
    Vy2 = Ret_VY2
    'resolve back the velocities to their normal coordinates
    Ret_VX1 = Vx1 * Cos(Angle) - Vy1 * Sin(Angle)
    Ret_VY1 = Vx1 * Sin(Angle) + Vy1 * Cos(Angle)
    Ret_VX2 = Vx2 * Cos(Angle) - Vy2 * Sin(Angle)
    Ret_VY2 = Vx2 * Sin(Angle) + Vy2 * Cos(Angle)
    'set the velocities of the ball
    A.vX = Ret_VX1 * CollLostEnergy
    A.vY = Ret_VY1 * CollLostEnergy
    B.vX = Ret_VX2 * CollLostEnergy
    B.vY = Ret_VY2 * CollLostEnergy


    If BisMotionless Then B.vX = 0: B.vY = 0

End Sub
Public Sub Separate_Point_Ball(ByRef A As tPoint, ByRef B As tPoint, B_R)
    'A is a point on Doll
    'B is circled shaped obtsacle

    Dim dX         As Double
    Dim dY         As Double
    Dim L          As Double
    Dim G          As Double
    Dim DeltaX     As Double
    Dim DeltaY     As Double

    If A.IsMotionLess Then

        dX = (B.X - A.X)
        dY = (B.Y - A.Y)
        L = Sqr(dX * dX + dY * dY)
        G = (B_R + 0.5) - L
        DeltaX = (G / L) * dX
        DeltaY = (G / L) * dY
        B.X = B.X + DeltaX
        B.Y = B.Y + DeltaY

    Else

        dX = (B.X - A.X)
        dY = (B.Y - A.Y)
        L = Sqr(dX * dX + dY * dY)
        G = (B_R + 0.5) - L

        DeltaX = (G / L) * dX
        DeltaY = (G / L) * dY
        A.X = A.X - DeltaX
        A.Y = A.Y - DeltaY

    End If

End Sub
