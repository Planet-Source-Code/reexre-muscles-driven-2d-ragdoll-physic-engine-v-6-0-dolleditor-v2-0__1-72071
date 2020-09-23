Attribute VB_Name = "mJOINT"
'
Option Explicit

Type tJoint
    D1             As Long
    D2             As Long
    Pd1            As Long
    pd2            As Long
End Type


Public Njoints     As Long
Private Joint()    As tJoint


Public Function MakeJoint(wDoll1, wDoll2, PointOfDoll1, PointOfDoll2) As Boolean


    Dim P1         As tPoint
    Dim P2         As tPoint
    Dim I          As Long

    'If wDoll2 < wDoll1 Then Exit Function

    For I = 1 To Njoints
        If Joint(I).D1 = wDoll1 And Joint(I).D2 = wDoll2 And Joint(I).Pd1 = PointOfDoll1 And Joint(I).pd2 = PointOfDoll2 Then
            RemoveJoint I
            Exit Function
        End If
    Next I


    P1.X = DOLL(wDoll1).PointX(PointOfDoll1)
    P1.Y = DOLL(wDoll1).PointY(PointOfDoll1)
    P2.X = DOLL(wDoll2).PointX(PointOfDoll2)
    P2.Y = DOLL(wDoll2).PointY(PointOfDoll2)

    '30
    If Distance(P1, P2) < 40 Then

        MakeJoint = True

        Njoints = Njoints + 1

        ReDim Preserve Joint(Njoints)
        Joint(Njoints).D1 = wDoll1
        Joint(Njoints).D2 = wDoll2
        Joint(Njoints).Pd1 = PointOfDoll1
        Joint(Njoints).pd2 = PointOfDoll2

    End If

End Function

Public Sub RemoveJoint(w)
    Dim I          As Long

    For I = w To Njoints - 1
        Joint(I) = Joint(I + 1)
    Next
    Njoints = Njoints - 1
    ReDim Preserve Joint(Njoints)
End Sub
Sub DoJoints()
    Dim X          As Double
    Dim Y          As Double
    Dim vX         As Double
    Dim vY         As Double
    'Stop
    Dim I          As Long

    For I = 1 To Njoints
        With Joint(I)


            X = (DOLL(.D1).PointX(.Pd1) + DOLL(.D2).PointX(.pd2)) * 0.5
            Y = (DOLL(.D1).PointY(.Pd1) + DOLL(.D2).PointY(.pd2)) * 0.5
            vX = (DOLL(.D1).PointVX(.Pd1) + DOLL(.D2).PointVX(.pd2)) * 0.5
            vY = (DOLL(.D1).PointVY(.Pd1) + DOLL(.D2).PointVY(.pd2)) * 0.5

            If Not (DOLL(.D1).PointIsFix(.Pd1)) Then

                DOLL(.D1).PointX(.Pd1) = X
                DOLL(.D1).PointY(.Pd1) = Y
                DOLL(.D1).PointVX(.Pd1) = vX
                DOLL(.D1).PointVY(.Pd1) = vY

            End If

            If Not (DOLL(.D2).PointIsFix(.pd2)) Then

                DOLL(.D2).PointX(.pd2) = X
                DOLL(.D2).PointY(.pd2) = Y
                DOLL(.D2).PointVX(.pd2) = vX
                DOLL(.D2).PointVY(.pd2) = vY
            End If

        End With

    Next I

End Sub

Public Function GetDollJoinedTo(Doll1, P1)
    Dim I          As Long

    For I = 1 To Njoints

        If (Joint(I).D1 = Doll1 And Joint(I).Pd1 = P1) _
           Then

            GetDollJoinedTo = Joint(I).D2
            Exit Function
        End If

    Next

End Function

Public Function GetPointJoinedTo(Doll1, P1)
    Dim I          As Long

    For I = 1 To Njoints

        If (Joint(I).D1 = Doll1 And Joint(I).Pd1 = P1) _
           Or (Joint(I).D2 = Doll1 And Joint(I).pd2 = P1) Then
            GetPointJoinedTo = Joint(I).pd2
            Exit Function
        End If
    Next

End Function
