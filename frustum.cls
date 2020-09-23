VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "frustum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public p_left As vertex, p_right As vertex, p_eye As vertex
Public left_plane As edge, right_plane As edge


Public Function constructor(l As vertex, r As vertex, e As vertex)
    Set p_left = l
    Set p_right = r
    Set p_eye = e
    Set left_plane = New edge
    Call left_plane.constructor(p_eye, p_left, vbBlack, False)
    Set right_plane = New edge
    Call right_plane.constructor(p_eye, p_right, vbBlack, False)
End Function

Public Function clip_left(e As edge) As edge
   
    Dim t As Double
    't = as_seg[LEFT]->Cross_Param(*ls);
    t = e.cross_param(left_plane)

    'clipped edge
    Dim tx As Double, ty As Double
    'tx = ls->start[0] + (t*(ls->end[0] - ls->start[0]));
    tx = e.start_pt.x + (t * (e.end_pt.x - e.start_pt.x))
    'ty = ls->start[1] + (t*(ls->end[1] - ls->start[1]));
    ty = e.start_pt.y + (t * (e.end_pt.y - e.start_pt.y))
    

    'Vertex *vc = new Vertex(0,tx,ty);
    Dim vc As New vertex
    Set vc = New vertex
    Call vc.construct(tx, ty, 0)

    'figure out which point is to the right of the left clip plane
    Dim sx As Double, sy As Double, ex As Double, ey As Double
    sx = e.start_pt.x
    sy = e.start_pt.y
    ex = e.end_pt.x
    ey = e.end_pt.y

    Dim te As edge

    
    'if start is left side of left clip plane
    If (left_plane.Point_Side(sx, sy) = 1 And _
        left_plane.Point_Side(ex, ey) = -1) Then
    
        'return clipped edge and END endpoint
        Dim ve As New vertex
        Set ve = New vertex
        Call ve.construct(ex, ey, 0)
        Set te = New edge
        Call te.constructor(vc, ve, e.color, e.opaque)
        
        Set clip_left = te
    ElseIf (left_plane.Point_Side(sx, sy) = -1 And _
        left_plane.Point_Side(ex, ey) = 1) Then
    'if start is right side of left clip plane
        
        'return clipped edge and start endpoint
        Dim vs As New vertex
        Set vs = New vertex
        Call vs.construct(sx, sy, 0)
        Set te = New edge
        Call te.constructor(vs, vc, e.color, e.opaque)
        
        Set clip_left = te

    Else
        Set clip_left = e
    End If
    
    
End Function


Public Function clip_right(e As edge) As edge
    Dim t As Double
    't = as_seg[LEFT]->Cross_Param(*ls);
    t = e.cross_param(right_plane)

    'clipped edge
    Dim tx As Double, ty As Double
    tx = e.start_pt.x + (t * (e.end_pt.x - e.start_pt.x))
    ty = e.start_pt.y + (t * (e.end_pt.y - e.start_pt.y))

    'Vertex *vc = new Vertex(0,tx,ty);
    Dim vc As New vertex
    Set vc = New vertex
    Call vc.construct(tx, ty, 0)

    'figure out which point is to the left of the right clip plane
    Dim sx As Double, sy As Double, ex As Double, ey As Double
    sx = e.start_pt.x
    sy = e.start_pt.y
    ex = e.end_pt.x
    ey = e.end_pt.y
    
    If (right_plane.Point_Side(sx, sy) = -1 And _
        right_plane.Point_Side(ex, ey) = 1) Then
        'return clipped edge and END endpoint
        Dim ve As New vertex
        Set ve = New vertex
        Call ve.construct(ex, ey, 0)
        Set te = New edge
        Call te.constructor(vc, ve, e.color, e.opaque)
        
        Set clip_right = te
    ElseIf (right_plane.Point_Side(sx, sy) = 1 And _
        right_plane.Point_Side(ex, ey) = -1) Then
        'return clipped edge and start endpoint
        Dim vs As New vertex
        Set vs = New vertex
        Call vs.construct(sx, sy, 0)
        Set te = New edge
        Call te.constructor(vc, vs, e.color, e.opaque)
        
        Set clip_right = te
    Else
        Set clip_right = e
    End If

End Function
