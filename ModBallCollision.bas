Attribute VB_Name = "ModBALL"
'By nirpaudyal@hotmail.com
'Modified By reexre@gmail.com

'This module contains all the important momemtum and forces stuff
'You may use any part of the code here in your own program
'Contact me at nirpaudyal@hotmail.com if you don't understand any part
'Include me in the about box if you felt that i have helped you!
'
Option Explicit

Public Const Mu = 0.0225          '0.025 '0.0035 '0.01      'coefficent of friction



Public Const BallsColor = 9226380    '7910520  '4953770 '4951180 '1341560#

'This sub deals with collision dection and bouncing of balls
Public Sub ChangeVelocities(A As TObstacle, B As TObstacle)
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

    A_Mass = IIf(A.IsMotionLess, 999999999999#, A.R * A.R)
    B_Mass = IIf(B.IsMotionLess, 999999999999#, B.R * B.R)

    'Get the distance between the two balls
    Dis = Sqr((A.P.X - B.P.X) ^ 2 + (A.P.Y - B.P.Y) ^ 2)
    'check to see if a collision has occured
    If Dis > A.R + B.R Then Exit Sub
    'if collision occurs then seperate the balls
    Seperate_Balls A, B
    'get the angle between the positions of the balls
    '    angle = Atn((Ret_VY2 - Ret_VY1) / (Ret_VX2 - Ret_VX1))
    '''If (Ret_VX2 - Ret_VX1) <> 0 Then angle = Atn((Ret_VY2 - Ret_VY1) / (Ret_VX2 - Ret_VX1)) Else angle = PI / 2
    Angle = Atan2(B.P.X - A.P.X, B.P.Y - A.P.Y)

    Vx1 = A.P.vX
    Vy1 = A.P.vY
    Vx2 = B.P.vX
    Vy2 = B.P.vY
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
    A.P.vX = Ret_VX1 * CollLostEnergy
    A.P.vY = Ret_VY1 * CollLostEnergy
    B.P.vX = Ret_VX2 * CollLostEnergy
    B.P.vY = Ret_VY2 * CollLostEnergy

    If A.IsMotionLess Then A.P.vX = 0: A.P.vY = 0
    If B.IsMotionLess Then B.P.vX = 0: B.P.vY = 0

    ''Dim AB As TObstacle
    ''AB.P.Vx = Abs(A.P.Vx) + Abs(B.P.Vx)
    ''AB.P.Vy = Abs(A.P.Vy) + Abs(B.P.Vy)
    ''V = Sqr(AB.P.Vx * AB.P.Vx + AB.P.Vy * AB.P.Vy)
    ''Stop
    ''PlayHitSound CSng(V / 2)

End Sub
Public Sub Seperate_Balls(A As TObstacle, B As TObstacle)
    'reset the position of the balls so that they dont overlap
    'this process is achieved using similar triangles
    Dim Tmp        As TObstacle
    Dim dX         As Double
    Dim dY         As Double
    Dim L          As Double
    Dim G          As Double
    Dim DeltaX     As Double
    Dim DeltaY     As Double

    If A.IsMotionLess Then
        Tmp = A
        A = B
        B = Tmp
    End If


    dX = (B.P.X - A.P.X)
    dY = (B.P.Y - A.P.Y)
    L = Sqr(dX * dX + dY * dY)
    G = (A.R + B.R) - L
    DeltaX = (G / L) * dX
    DeltaY = (G / L) * dY
    A.P.X = A.P.X - DeltaX
    A.P.Y = A.P.Y - DeltaY

End Sub
'-----------------------------------------------------------------------

'Public Sub HandleFriction(o As TObstacle)
''get the speed of the ball
'With o
'    V = Sqr(.P.Vx * .P.Vx + .P.Vy * .P.Vy)
'
'    ''friction doesn't act while ball is not in motion
'    'If V < 0 Then Exit Sub
'
'    'if speed is really low then set it to zero
'    If V < 0.01 Then '001
'        .P.Vx = 0
'        .P.Vy = 0
'        'o.V = 0
'
'        Exit Sub
'    End If
'    Dim fx As Single
'    Dim fy As Single
'
'
'    'calculate the friction
'    Friction = Mu * (.R * .R) * Abs(Gravity)
'    If .P.Vx = 0 Then ANg = 0 Else ANg = Atan2(.P.Vx, .P.Vy) 'ANg = Atn(.P.Vy / .P.Vx)
'    'get the components of frictions in the two directions
'    fx = Abs(Friction * Cos(ANg))
'    fy = Abs(Friction * Sin(ANg))
'    'ensure that the friction is opposing the direction of motion
'    If .P.Vx > 0 Then fx = -fx
'    If .P.Vy > 0 Then fy = -fy
'    'apply the force
'   ' ApplyForce o, fx, fy, 0.1
'
'    'o.Vy = o.Vy + GravityY
'End With
''End Sub


'Sub ApplyForce(o As TObstacle, ForceX As Single, ForceY As Single, Time_Of_Force As Single)
''Use F= (mv-mu)/t to find v, the new velocity of the ball once the force is applied
'
'o.P.Vx = o.P.Vx + (ForceX * Time_Of_Force / (o.R * o.R))
'o.P.Vy = o.P.Vy + (ForceY * Time_Of_Force / (o.R * o.R))
'
'End Sub
