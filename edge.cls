VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "edge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public start_pt As vertex, end_pt As vertex
Public opaque As Boolean
Public color As ColorConstants
Public neighbor As cell


Public Function constructor(sv As vertex, ev As vertex, col As ColorConstants, op As Boolean)
    Set start_pt = sv
    Set end_pt = ev
    color = col
    opaque = op

End Function


Public Function Point_Side(x As Double, y As Double) As Integer

    '// Compute the determinant: | xs ys 1 |
    '//                          | xe ye 1 |
    '//                          | x  y  1 |
    '// Use its sign to get the answer.

    Dim det As Double

    det = start_pt.x * (end_pt.y - y) _
        - start_pt.y * (end_pt.x - x) _
    + end_pt.x * y - end_pt.y * x
    
    If (det = 0#) Then
        Point_Side = 0
    ElseIf (det > 0#) Then
        Point_Side = 1
    Else
        Point_Side = -1
    End If
End Function


Public Function Cross_Param(e As Edge) As Double
    Dim dx1, dy1, dx2, dy2 As Double
    Dim denom, s As Double

    '// This computation comes from writing each segment in parametric form,
    '// and solving a simulataneous equation to determine the parameter
    '// value of the intersection point on this LineSeg.

    dx1 = e.end_pt.x - e.start_pt.x
    dy1 = e.end_pt.y - e.start_pt.y
    dx2 = end_pt.x - start_pt.x
    dy2 = end_pt.y - start_pt.y


    denom = dx2 * dy1 - dy2 * dx1
    If (denom = 0#) Then
        '// Parallel segments.
        Cross_Param = 0
    Else
        s = (e.start_pt.x - start_pt.x) * dy1 - (e.start_pt.y - start_pt.y) * dx1
        Cross_Param = s / denom
    End If

End Function
