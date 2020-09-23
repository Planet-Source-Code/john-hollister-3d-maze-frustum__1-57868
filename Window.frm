VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Maze"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   11280
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4800
      Width           =   1935
   End
   Begin VB.PictureBox floor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   10440
      Picture         =   "Window.frx":0000
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox map 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   10440
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
   Begin VB.PictureBox screen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      ForeColor       =   &H80000008&
      Height          =   10335
      Left            =   0
      ScaleHeight     =   687
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   0
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DON'T LET THE LENGTH OF THIS CODE FOOL YOU
'95% IS USED TO WRITE THE VERTICES FOR THE MAZE

'IF ANYONE KNOWS A GOOD/EASY MAZE ALGORITHM TO IMPLEMENT
'FOR THIS STYLE (CELLS & PORTALS) OF MAZE, I'D LIKE
'TO HEAR IT.

'ONCE A GOOD MAZE ALGO IS GIVEN, THEN IT WOULD BE WORTH
'WHILE TO DO SOME FILE I/O

'WRITTEN BY JOHN HOLLISTER
'12/2004


'thanks to professor chenney at uw-madison
'for giving me the concept of this maze program


Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private t As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private viewer_posn(0 To 2) As Double
Private viewer_dir As Double
Private Filler, OffX, OffY As Long
' Constants
Private Const ALTERNATE = 1
Private Const WINDING = 2
Private hRgn, res As Long

Public cur_cell As cell

Private fov As Double
Private focal_dist As Long

Private viewer_frust As frustum

Private maze_width As Integer, maze_height As Integer
Private cell_array() As New cell
Private orig_x As Double, orig_y As Double
Private m_orig_x As Double, m_orig_y As Double

Private Const PI = 3.14159



Private Sub Form_Load()

    'focal distance doesn't work all that well for some reason...
    'you have to change an offset when you're creating the frustum
    'it's 5 for 45 degrees and something like 11 for 60 degrees
    'and it gets really messed up the higher you go
    '45 degrees still looks nice though
    fov = 45
    
    'offset to center of screen
    orig_x = screen.ScaleWidth / 2
    orig_y = screen.ScaleHeight / 2
    
    m_orig_x = 0
    m_orig_y = map.ScaleHeight - 1

    
    focal_dist = (screen.ScaleWidth) / (2 * Tan(fov * PI / 180) * 0.5)

    

    
    
    
    Me.Show
    
   
    viewer_dir = 90 + fov / 2
    viewer_posn(0) = 3
    viewer_posn(1) = 3
    viewer_posn(2) = 1

    
    'need to implement file I/O to read in maze
    'load maze cell by cell
    
    'need to specify cells, neighbors (edges)
    ' specify maze dimensions
    'track viewer
    
    'implement exact visibility recursion
 
    Call build_maze

    'forms pointers between cells
    Call connect_maze

    Set cur_cell = cell_array(0, 0)
    
    Call mainloop

End Sub


Public Function mainloop()
    Do
        t = GetTickCount + 30
        
        Call getControls
        Call find_viewer
        Call draw
        
        
        Text1.Text = cur_cell.x & vbNewLine & cur_cell.y
       
        While GetTickCount < t
            DoEvents
        Wend
    Loop Until False
End Function

Public Function getControls()
    'get user controls and collision detect
    Dim tx As Double, ty As Double
    'walk
    If GetAsyncKeyState(vbKeyUp) <> 0 Then
        tx = viewer_posn(0) + Cos(viewer_dir * (PI / 180)) / 5
        ty = viewer_posn(1) + Sin(viewer_dir * (PI / 180)) / 5
        If valid_move(tx, ty) Then
            viewer_posn(0) = tx
            viewer_posn(1) = ty
        End If
    ElseIf GetAsyncKeyState(vbKeyDown) <> 0 Then
        tx = viewer_posn(0) - Cos(viewer_dir * (PI / 180)) / 5
        ty = viewer_posn(1) - Sin(viewer_dir * (PI / 180)) / 5
        If valid_move(tx, ty) Then
            viewer_posn(0) = tx
            viewer_posn(1) = ty
        End If
    End If
    'strafe
    If GetAsyncKeyState(188) <> 0 Then
        tx = viewer_posn(0) + Cos((viewer_dir + 90) * (PI / 180)) / 5
        ty = viewer_posn(1) + Sin((viewer_dir + 90) * (PI / 180)) / 5
        If valid_move(tx, ty) Then
            viewer_posn(0) = tx
            viewer_posn(1) = ty
        End If
    ElseIf GetAsyncKeyState(190) <> 0 Then
        tx = viewer_posn(0) + Cos((viewer_dir - 90) * (PI / 180)) / 5
        ty = viewer_posn(1) + Sin((viewer_dir - 90) * (PI / 180)) / 5
        If valid_move(tx, ty) Then
            viewer_posn(0) = tx
            viewer_posn(1) = ty
        End If
    End If
    'turn
    If GetAsyncKeyState(vbKeyLeft) <> 0 Then
        viewer_dir = (viewer_dir + 5) Mod 360
    ElseIf GetAsyncKeyState(vbKeyRight) <> 0 Then
        viewer_dir = (viewer_dir - 5)
        If viewer_dir < 0 Then viewer_dir = 355
    End If
End Function

Public Function draw()
    'create initial frustum
    Dim ve As vertex, vl As vertex, vr As vertex
    Set ve = New vertex
    Call ve.construct(viewer_posn(0), viewer_posn(1), 0)
    Set vl = New vertex
    Call vl.construct(viewer_posn(0) + Cos((PI / 180) * (viewer_dir + (fov / 2) + 5)), _
            viewer_posn(1) + Sin((PI / 180) * (viewer_dir + (fov / 2) + 5)), 0)
    Set vr = New vertex
    Call vr.construct(viewer_posn(0) + Cos((PI / 180) * (viewer_dir - (fov / 2) - 5)), _
            viewer_posn(1) + Sin((PI / 180) * (viewer_dir - (fov / 2) - 5)), 0)
    Set viewer_frust = New frustum
    Call viewer_frust.constructor(vl, vr, ve)

    screen.Cls
    'draw floor
    StretchBlt screen.hDC, 0, screen.ScaleHeight / 2, screen.ScaleWidth, screen.ScaleHeight / 2, floor.hDC, 0, 0, floor.ScaleWidth, floor.ScaleHeight, vbSrcCopy

    map.Cls
    map.PSet (orig_x + viewer_posn(0) * 5, orig_y - viewer_posn(1) * 5)
    
    'left
    map.Line (m_orig_x + viewer_posn(0) * 5, m_orig_y - viewer_posn(1) * 5)- _
            (m_orig_x + viewer_posn(0) * 5 + 10 * Cos((PI / 180) * (viewer_dir + (fov / 2) + 5)), _
            (m_orig_y - viewer_posn(1) * 5 - 10 * Sin((PI / 180) * (viewer_dir + (fov / 2) + 5))))
    
    'right
    map.Line (m_orig_x + viewer_posn(0) * 5, m_orig_y - viewer_posn(1) * 5)- _
            (m_orig_x + viewer_posn(0) * 5 + 10 * Cos((PI / 180) * (viewer_dir - (fov / 2) - 5)), _
            (m_orig_y - viewer_posn(1) * 5 - 10 * Sin((PI / 180) * (viewer_dir - (fov / 2) - 5))))
    
    'map.Line (m_orig_x + viewer_posn(0) * 5, m_orig_y - viewer_posn(1) * 5)- _
    '        (m_orig_x + viewer_posn(0) * 5 + 10 * Cos(viewer_dir * (PI / 180)), _
    '        m_orig_y - viewer_posn(1) * 5 - 10 * Sin(viewer_dir * (PI / 180)))

    Call draw_recurse(viewer_frust, cur_cell, cur_cell)

End Function



Public Function draw_recurse(frust As frustum, view_cell As cell, last_cell As cell)
        
    Dim te() As New Edge
    te = view_cell.getEdges
    Dim temp_edge As New Edge
    
    
    For a = 0 To 3
        'do the clipping
        Set temp_edge = New Edge
        Call temp_edge.constructor(te(a).start_pt, te(a).end_pt, te(a).color, te(a).opaque)
        Set temp_edge = frust.clip_left(temp_edge)
        Set temp_edge = frust.clip_right(temp_edge)
        Set temp_edge.neighbor = te(a).neighbor
        
        'draw wall if opaque and within frustum
        If frust.left_plane.Point_Side(temp_edge.start_pt.x, temp_edge.start_pt.y) < 0 _
            And frust.right_plane.Point_Side(temp_edge.end_pt.x, temp_edge.end_pt.y) > 0 Then
            
            'draw wall if opaque
            If temp_edge.opaque Then
                Dim sx As Double, sy As Double, sz As Double
                Dim ex As Double, ey As Double, ez As Double
                sx = temp_edge.start_pt.x
                sy = temp_edge.start_pt.y
                sz = temp_edge.start_pt.z
                
                ex = temp_edge.end_pt.x
                ey = temp_edge.end_pt.y
                ez = temp_edge.end_pt.z
                
                map.Line (m_orig_x + 5 * temp_edge.start_pt.x, m_orig_y - 5 * temp_edge.start_pt.y)- _
                       (m_orig_x + 5 * temp_edge.end_pt.x, m_orig_y - 5 * temp_edge.end_pt.y), temp_edge.color
               
                'offset viewer angle and position
                Dim u As Double
                u = (sx - viewer_posn(0)) * Cos((-viewer_dir) * (PI / 180)) - _
                    (sy - viewer_posn(1)) * Sin((-viewer_dir) * (PI / 180))
                sy = (sx - viewer_posn(0)) * Sin((-viewer_dir) * (PI / 180)) + _
                    (sy - viewer_posn(1)) * Cos((-viewer_dir) * (PI / 180))
                sx = u
    
                u = (ex - viewer_posn(0)) * Cos((-viewer_dir) * (PI / 180)) - _
                    (ey - viewer_posn(1)) * Sin((-viewer_dir) * (PI / 180))
                ey = (ex - viewer_posn(0)) * Sin((-viewer_dir) * (PI / 180)) + _
                    (ey - viewer_posn(1)) * Cos((-viewer_dir) * (PI / 180))
                ex = u
    
                'rotate the world space
                Dim rtx1, rty1, rty2, rtx2, rtz1, rtz2 As Double
    
                'rotate -90 about z
                rtx1 = sx * Cos(-90 * PI / 180) - sy * Sin(-90 * PI / 180)
                rty1 = sx * Sin(PI / 180 * (-90)) + sy * Cos(PI / 180 * (-90))
                rtx2 = ex * Cos(PI / 180 * (-90#)) - ey * Sin(PI / 180 * (-90#))
                rty2 = ex * Sin(PI / 180 * (-90#)) + ey * Cos(PI / 180 * (-90#))
                sx = rtx1
                ex = rtx2
                sy = rty1
                ey = rty2
    
                'then rotate -90 about x
                rty1 = sy * Cos(PI / 180 * (-90#))
                rtz1 = sy * Sin(PI / 180 * (-90#))
                rty2 = ey * Cos(PI / 180 * (-90#))
                rtz2 = ey * Sin(PI / 180 * (-90#))
                    
                Dim depth1 As Double, depth2 As Double
                    
                'negate x axis now
                sx = -sx
                ex = -ex
                sy = rty1
                ey = rty2
                depth1 = rtz1
                depth2 = rtz2
            
                'i should now have x,y, and z coords
                'use z coord and focal distance to obtain
                'perspective projection of wall
                '(focal_dist / depth)*point x or y
                sx = (focal_dist / depth1) * sx
                ex = (focal_dist / depth2) * ex
                sy = (focal_dist / depth1) * sy
                ey = (focal_dist / depth2) * ey
                sz = (focal_dist / depth1) * 2
                ez = (focal_dist / depth2) * 2
            
                'points to draw "polygon" with
                Dim points(0 To 3) As POINTAPI
                points(0).x = orig_x + sx
                points(0).y = orig_y - sy - (viewer_posn(2) / (depth1 / focal_dist))
                points(1).x = orig_x + sx
                points(1).y = orig_y - sy + sz - (viewer_posn(2) / (depth1 / focal_dist))
                points(2).x = orig_x + ex
                points(2).y = orig_y - ey + ez - (viewer_posn(2) / (depth2 / focal_dist))
                points(3).x = orig_x + ex
                points(3).y = orig_y - ey - (viewer_posn(2) / (depth2 / focal_dist))
                
                'fill "polygon"
                Dim hBrush As Long
                hBrush = CreateSolidBrush(te(a).color)
                hRgn = CreatePolygonRgn(points(0), 4, ALTERNATE)
                Call FillRgn(screen.hDC, hRgn, hBrush)
                Call DeleteObject(hRgn)
                Call DeleteObject(hBrush)
            Else
                'if transparent and within frustum then
                '-create new frustum defined by clipped transparent wall endpoints
                Dim ve As vertex, vl As vertex, vr As vertex
                Set ve = New vertex
                Call ve.construct(viewer_posn(0), viewer_posn(1), 0)
                Set vr = New vertex
                Call vr.construct(temp_edge.start_pt.x, temp_edge.start_pt.y, 0)
                Set vl = New vertex
                Call vl.construct(temp_edge.end_pt.x, temp_edge.end_pt.y, 0)
                Dim next_frust As New frustum
                Set next_frust = New frustum
                Call next_frust.constructor(vl, vr, ve)
                
                '-update current cell to be transparentedge.neighbor
                Dim next_cell As New cell
                Set next_cell = temp_edge.neighbor
                               
                'recurse (make sure you're not going back to a cell
                'you've already been to

                If (next_cell.x <> last_cell.x Or next_cell.y <> last_cell.y) Then
                    Call draw_recurse(next_frust, next_cell, view_cell)
                End If
            End If
        End If
    Next a
End Function


Public Function find_viewer()
    'check cur_cell then all of it's neighbors
    'set cur_cell accordingly
    'which ever wall it is to the outside of, go
    'to that neighboring cell
    '0 = north
    '1 = east
    '2 = south
    '3 = west
    Dim te() As New Edge
    te = cur_cell.getEdges
    If (te(0).Point_Side(viewer_posn(0), viewer_posn(1)) = -1) Then
        Set cur_cell = cur_cell.north
    ElseIf (te(1).Point_Side(viewer_posn(0), viewer_posn(1)) = -1) Then
        Set cur_cell = cur_cell.east
    ElseIf (te(2).Point_Side(viewer_posn(0), viewer_posn(1)) = -1) Then
        Set cur_cell = cur_cell.south
    ElseIf (te(3).Point_Side(viewer_posn(0), viewer_posn(1)) = -1) Then
        Set cur_cell = cur_cell.west
    End If
End Function


Public Function connect_maze()

    '0 = north
    '1 = east
    '2 = south
    '3 = west

    'establish pointers in the cell array
    Dim this_cell As New cell

    'set locations first
    For a = 0 To maze_width - 1
        For b = 0 To maze_height - 1
            Set this_cell = cell_array(a, b)
            this_cell.x = a
            this_cell.y = b
            Set cell_array(a, b) = this_cell
        Next b
    Next a
    
    For a = 0 To maze_width - 1
        For b = 0 To maze_height - 1
            Set this_cell = cell_array(a, b)
            'north pointer
            If b = 0 Then
                Call this_cell.setNeighbor(2, Nothing)
            Else
                Call this_cell.setNeighbor(2, cell_array(a, b - 1))
            End If
            'south
            If b = maze_height - 1 Then
                Call this_cell.setNeighbor(0, Nothing)
            Else
                Call this_cell.setNeighbor(0, cell_array(a, b + 1))
            End If
            'west pointer
            If a = 0 Then
                Call this_cell.setNeighbor(3, Nothing)
            Else
                Call this_cell.setNeighbor(3, cell_array(a - 1, b))
            End If
            'east
            If a = maze_width - 1 Then
                Call this_cell.setNeighbor(1, Nothing)
            Else
                Call this_cell.setNeighbor(1, cell_array(a + 1, b))
            End If
            Set cell_array(a, b) = this_cell
        Next b
    Next a

End Function


Public Function valid_move(tx As Double, ty As Double) As Boolean
    Dim te() As Edge
    te = cur_cell.getEdges
    valid_move = True
    For a = 0 To 3
        If te(a).Point_Side(tx, ty) <= 0 And te(a).opaque Then
            'hit
            valid_move = False
            a = 4
        End If
    Next a


End Function


Public Function build_maze()

'EMERGENCY

'GOTTA DO FILE I/O!!!!!!!!!!!111one.
'does anyone know any good maze building algos!?






    maze_width = 9
    maze_height = 5
    
    'for connecting the maze
    ReDim cell_array(-1 To maze_width + 1, -1 To maze_height + 1) As New cell

    Dim this_cell As New cell
    Dim this_edge As New Edge
    Dim e(0 To 3) As Edge
    Dim v1 As vertex
    Dim v2 As vertex
        
    'start cell
    'north
    Set v1 = New vertex
    Call v1.construct(12, 6, 10)
    Set v2 = New vertex
    Call v2.construct(6, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(0) = this_edge
    
    'east
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 0, 10)
    Call v2.construct(12, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, False)
    Set e(1) = this_edge

    'south
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 0, 10)
    Call v2.construct(12, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(2) = this_edge

    'west
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 6, 10)
    Call v2.construct(6, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge

    Set this_cell = New cell
    Call this_cell.constructor(e())
   
    Set cell_array(1, 0) = this_cell
    
    
    'west cell
    Set v1 = New vertex
    Call v1.construct(6, 6, 10)
    Set v2 = New vertex
    Call v2.construct(0, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 0, 10)
    Call v2.construct(6, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(0, 0, 10)
    Call v2.construct(6, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(0, 6, 10)
    Call v2.construct(0, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(0, 0) = this_cell


    
    
     'east cell
    Set v1 = New vertex
    Call v1.construct(18, 6, 10)
    Set v2 = New vertex
    Call v2.construct(12, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 0, 10)
    Call v2.construct(18, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 0, 10)
    Call v2.construct(18, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 6, 10)
    Call v2.construct(12, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(2, 0) = this_cell
    
'--------------------------------------------
    'X = 3 Y = 0 cell
    
    Set v1 = New vertex
    Call v1.construct(24, 6, 10)
    Set v2 = New vertex
    Call v2.construct(18, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 0, 10)
    Call v2.construct(24, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 0, 10)
    Call v2.construct(24, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 6, 10)
    Call v2.construct(18, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(3, 0) = this_cell
    
    
    
'--------------------------------------------
    'X = 4 Y = 0 cell
    
    Set v1 = New vertex
    Call v1.construct(30, 6, 10)
    Set v2 = New vertex
    Call v2.construct(24, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 0, 10)
    Call v2.construct(30, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 0, 10)
    Call v2.construct(30, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 6, 10)
    Call v2.construct(24, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(4, 0) = this_cell
    
    
 '--------------------------------------------
    'X = 5 Y = 0 cell
    
    Set v1 = New vertex
    Call v1.construct(36, 6, 10)
    Set v2 = New vertex
    Call v2.construct(30, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 0, 10)
    Call v2.construct(36, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 0, 10)
    Call v2.construct(36, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 6, 10)
    Call v2.construct(30, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(5, 0) = this_cell
    

 '--------------------------------------------
    'X = 6 Y = 0 cell
    
    Set v1 = New vertex
    Call v1.construct(42, 6, 10)
    Set v2 = New vertex
    Call v2.construct(36, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 0, 10)
    Call v2.construct(42, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 0, 10)
    Call v2.construct(42, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 6, 10)
    Call v2.construct(36, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(6, 0) = this_cell
    
 '--------------------------------------------
    'X = 7 Y = 0 cell
    
    Set v1 = New vertex
    Call v1.construct(48, 6, 10)
    Set v2 = New vertex
    Call v2.construct(42, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 0, 10)
    Call v2.construct(48, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 0, 10)
    Call v2.construct(48, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 6, 10)
    Call v2.construct(42, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(7, 0) = this_cell
    
  '--------------------------------------------
    'X = 8 Y = 0 cell
    
    Set v1 = New vertex
    Call v1.construct(54, 6, 10)
    Set v2 = New vertex
    Call v2.construct(48, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(54, 0, 10)
    Call v2.construct(54, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 0, 10)
    Call v2.construct(54, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 6, 10)
    Call v2.construct(48, 0, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(8, 0) = this_cell
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    '---------------------------------------
    ' row y = 1
    
    '-------------------
    'x = 0  Y = 1
    
    Set v1 = New vertex
    Call v1.construct(6, 12, 10)
    Set v2 = New vertex
    Call v2.construct(0, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 6, 10)
    Call v2.construct(6, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(0, 6, 10)
    Call v2.construct(6, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(0, 12, 10)
    Call v2.construct(0, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(0, 1) = this_cell
    
    
    'north
    Set v1 = New vertex
    Call v1.construct(12, 12, 10)
    Set v2 = New vertex
    Call v2.construct(6, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(0) = this_edge
    
    'east
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 6, 10)
    Call v2.construct(12, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, False)
    Set e(1) = this_edge

    'south
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 6, 10)
    Call v2.construct(12, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(2) = this_edge

    'west
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 12, 10)
    Call v2.construct(6, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge

    Set this_cell = New cell
    Call this_cell.constructor(e())
   
    Set cell_array(1, 1) = this_cell
    
    
'--------------------------------
    'X = 2 Y = 1 cell
    Set v1 = New vertex
    Call v1.construct(18, 12, 10)
    Set v2 = New vertex
    Call v2.construct(12, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 6, 10)
    Call v2.construct(18, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 6, 10)
    Call v2.construct(18, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 12, 10)
    Call v2.construct(12, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(2, 1) = this_cell

'-------------------------------------------
    'X = 3 Y = 1 cell
    
    Set v1 = New vertex
    Call v1.construct(24, 12, 10)
    Set v2 = New vertex
    Call v2.construct(18, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 6, 10)
    Call v2.construct(24, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 6, 10)
    Call v2.construct(24, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 12, 10)
    Call v2.construct(18, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(3, 1) = this_cell
    


'--------------------------------------------
    'X = 4 Y = 1 cell
    
    Set v1 = New vertex
    Call v1.construct(30, 12, 10)
    Set v2 = New vertex
    Call v2.construct(24, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 6, 10)
    Call v2.construct(30, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 6, 10)
    Call v2.construct(30, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 12, 10)
    Call v2.construct(24, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(4, 1) = this_cell

 '--------------------------------------------
    'X = 5 Y = 1 cell
    
    Set v1 = New vertex
    Call v1.construct(36, 12, 10)
    Set v2 = New vertex
    Call v2.construct(30, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 6, 10)
    Call v2.construct(36, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlack, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 6, 10)
    Call v2.construct(36, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 12, 10)
    Call v2.construct(30, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(5, 1) = this_cell

 '--------------------------------------------
    'X = 6 Y = 1 cell
    
    Set v1 = New vertex
    Call v1.construct(42, 12, 10)
    Set v2 = New vertex
    Call v2.construct(36, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 6, 10)
    Call v2.construct(42, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 6, 10)
    Call v2.construct(42, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 12, 10)
    Call v2.construct(36, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(6, 1) = this_cell

 '--------------------------------------------
    'X = 7 Y = 1 cell
    
    Set v1 = New vertex
    Call v1.construct(48, 12, 10)
    Set v2 = New vertex
    Call v2.construct(42, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 6, 10)
    Call v2.construct(48, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 6, 10)
    Call v2.construct(48, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 12, 10)
    Call v2.construct(42, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(7, 1) = this_cell

  '--------------------------------------------
    'X = 8 Y = 1 cell
    
    Set v1 = New vertex
    Call v1.construct(54, 12, 10)
    Set v2 = New vertex
    Call v2.construct(48, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(54, 6, 10)
    Call v2.construct(54, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 6, 10)
    Call v2.construct(54, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 12, 10)
    Call v2.construct(48, 6, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(8, 1) = this_cell

























    '---------------------------------------
    ' row y = 2
    
    
    '---------------------------
    ' X = 0 'Y = 2
    
    Set v1 = New vertex
    Call v1.construct(6, 18, 10)
    Set v2 = New vertex
    Call v2.construct(0, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 12, 10)
    Call v2.construct(6, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(0, 12, 10)
    Call v2.construct(6, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(0, 18, 10)
    Call v2.construct(0, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(0, 2) = this_cell
    
    
    
    
    '-----------------------------
    ' X = 1  Y = 2
    
    'north
    Set v1 = New vertex
    Call v1.construct(12, 18, 10)
    Set v2 = New vertex
    Call v2.construct(6, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(0) = this_edge
    
    'east
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 12, 10)
    Call v2.construct(12, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, False)
    Set e(1) = this_edge

    'south
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 12, 10)
    Call v2.construct(12, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(2) = this_edge

    'west
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 18, 10)
    Call v2.construct(6, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(3) = this_edge

    Set this_cell = New cell
    Call this_cell.constructor(e())
   
    Set cell_array(1, 2) = this_cell
    
    
'--------------------------------
    'X = 2 Y = 2 cell
    Set v1 = New vertex
    Call v1.construct(18, 18, 10)
    Set v2 = New vertex
    Call v2.construct(12, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 12, 10)
    Call v2.construct(18, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 12, 10)
    Call v2.construct(18, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 18, 10)
    Call v2.construct(12, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(2, 2) = this_cell

'-------------------------------------------
    'X = 3 Y = 2 cell
    
    Set v1 = New vertex
    Call v1.construct(24, 18, 10)
    Set v2 = New vertex
    Call v2.construct(18, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 12, 10)
    Call v2.construct(24, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 12, 10)
    Call v2.construct(24, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 18, 10)
    Call v2.construct(18, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(3, 2) = this_cell
    


'--------------------------------------------
    'X = 4 Y = 2 cell
    
    Set v1 = New vertex
    Call v1.construct(30, 18, 10)
    Set v2 = New vertex
    Call v2.construct(24, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 12, 10)
    Call v2.construct(30, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 12, 10)
    Call v2.construct(30, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 18, 10)
    Call v2.construct(24, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(4, 2) = this_cell

 '--------------------------------------------
    'X = 5 Y = 2 cell
    
    Set v1 = New vertex
    Call v1.construct(36, 18, 10)
    Set v2 = New vertex
    Call v2.construct(30, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 12, 10)
    Call v2.construct(36, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlack, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 12, 10)
    Call v2.construct(36, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 18, 10)
    Call v2.construct(30, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(5, 2) = this_cell

 '--------------------------------------------
    'X = 6 Y = 2 cell
    
    Set v1 = New vertex
    Call v1.construct(42, 18, 10)
    Set v2 = New vertex
    Call v2.construct(36, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 12, 10)
    Call v2.construct(42, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 12, 10)
    Call v2.construct(42, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 18, 10)
    Call v2.construct(36, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(6, 2) = this_cell

 '--------------------------------------------
    'X = 7 Y = 2 cell
    
    Set v1 = New vertex
    Call v1.construct(48, 18, 10)
    Set v2 = New vertex
    Call v2.construct(42, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 12, 10)
    Call v2.construct(48, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 12, 10)
    Call v2.construct(48, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 18, 10)
    Call v2.construct(42, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(7, 2) = this_cell

  '--------------------------------------------
    'X = 8 Y = 2 cell
    
    Set v1 = New vertex
    Call v1.construct(54, 18, 10)
    Set v2 = New vertex
    Call v2.construct(48, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(54, 12, 10)
    Call v2.construct(54, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 12, 10)
    Call v2.construct(54, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 18, 10)
    Call v2.construct(48, 12, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(8, 2) = this_cell




















    '---------------------------------------
    ' row y = 3
    
    
    '---------------------------
    ' X = 0 'Y = 3
    
    Set v1 = New vertex
    Call v1.construct(6, 24, 10)
    Set v2 = New vertex
    Call v2.construct(0, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 18, 10)
    Call v2.construct(6, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(0, 18, 10)
    Call v2.construct(6, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(0, 24, 10)
    Call v2.construct(0, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(0, 3) = this_cell
    
    
    
    
    '-----------------------------
    ' X = 1  Y = 3
    
    'north
    Set v1 = New vertex
    Call v1.construct(12, 24, 10)
    Set v2 = New vertex
    Call v2.construct(6, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(0) = this_edge
    
    'east
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 18, 10)
    Call v2.construct(12, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(1) = this_edge

    'south
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 18, 10)
    Call v2.construct(12, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(2) = this_edge

    'west
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 24, 10)
    Call v2.construct(6, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(3) = this_edge

    Set this_cell = New cell
    Call this_cell.constructor(e())
   
    Set cell_array(1, 3) = this_cell
    
    
'--------------------------------
    'X = 2 Y = 3 cell
    Set v1 = New vertex
    Call v1.construct(18, 24, 10)
    Set v2 = New vertex
    Call v2.construct(12, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 18, 10)
    Call v2.construct(18, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 18, 10)
    Call v2.construct(18, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 24, 10)
    Call v2.construct(12, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(2, 3) = this_cell

'-------------------------------------------
    'X = 3 Y = 3 cell
    
    Set v1 = New vertex
    Call v1.construct(24, 24, 10)
    Set v2 = New vertex
    Call v2.construct(18, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 18, 10)
    Call v2.construct(24, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 18, 10)
    Call v2.construct(24, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 24, 10)
    Call v2.construct(18, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(3, 3) = this_cell
    


'--------------------------------------------
    'X = 4 Y = 3 cell
    
    Set v1 = New vertex
    Call v1.construct(30, 24, 10)
    Set v2 = New vertex
    Call v2.construct(24, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 18, 10)
    Call v2.construct(30, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 18, 10)
    Call v2.construct(30, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 24, 10)
    Call v2.construct(24, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(4, 3) = this_cell

 '--------------------------------------------
    'X = 5 Y = 3 cell
    
    Set v1 = New vertex
    Call v1.construct(36, 24, 10)
    Set v2 = New vertex
    Call v2.construct(30, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 18, 10)
    Call v2.construct(36, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlack, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 18, 10)
    Call v2.construct(36, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 24, 10)
    Call v2.construct(30, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(5, 3) = this_cell

 '--------------------------------------------
    'X = 6 Y = 3 cell
    
    Set v1 = New vertex
    Call v1.construct(42, 24, 10)
    Set v2 = New vertex
    Call v2.construct(36, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 18, 10)
    Call v2.construct(42, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 18, 10)
    Call v2.construct(42, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 24, 10)
    Call v2.construct(36, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(6, 3) = this_cell

 '--------------------------------------------
    'X = 7 Y = 3 cell
    
    Set v1 = New vertex
    Call v1.construct(48, 24, 10)
    Set v2 = New vertex
    Call v2.construct(42, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 18, 10)
    Call v2.construct(48, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 18, 10)
    Call v2.construct(48, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 24, 10)
    Call v2.construct(42, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(7, 3) = this_cell

  '--------------------------------------------
    'X = 8 Y = 3 cell
    
    Set v1 = New vertex
    Call v1.construct(54, 24, 10)
    Set v2 = New vertex
    Call v2.construct(48, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(54, 18, 10)
    Call v2.construct(54, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 18, 10)
    Call v2.construct(54, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 24, 10)
    Call v2.construct(48, 18, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(8, 3) = this_cell


















    '---------------------------------------
    ' row y = 4
    
    
    '---------------------------
    ' X = 0 'Y = 4
    
    Set v1 = New vertex
    Call v1.construct(6, 30, 10)
    Set v2 = New vertex
    Call v2.construct(0, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 24, 10)
    Call v2.construct(6, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(0, 24, 10)
    Call v2.construct(6, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(0, 30, 10)
    Call v2.construct(0, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(0, 4) = this_cell
    
    
    
    
    '-----------------------------
    ' X = 1  Y = 4
    
    'north
    Set v1 = New vertex
    Call v1.construct(12, 30, 10)
    Set v2 = New vertex
    Call v2.construct(6, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(0) = this_edge
    
    'east
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 24, 10)
    Call v2.construct(12, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(1) = this_edge

    'south
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 24, 10)
    Call v2.construct(12, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, False)
    Set e(2) = this_edge

    'west
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(6, 30, 10)
    Call v2.construct(6, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(3) = this_edge

    Set this_cell = New cell
    Call this_cell.constructor(e())
   
    Set cell_array(1, 4) = this_cell
    
    
'--------------------------------
    'X = 2 Y = 4 cell
    Set v1 = New vertex
    Call v1.construct(18, 30, 10)
    Set v2 = New vertex
    Call v2.construct(12, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 24, 10)
    Call v2.construct(18, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 24, 10)
    Call v2.construct(18, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(12, 30, 10)
    Call v2.construct(12, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(2, 4) = this_cell

'-------------------------------------------
    'X = 3 Y = 4 cell
    
    Set v1 = New vertex
    Call v1.construct(24, 30, 10)
    Set v2 = New vertex
    Call v2.construct(18, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 24, 10)
    Call v2.construct(24, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 24, 10)
    Call v2.construct(24, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(18, 30, 10)
    Call v2.construct(18, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(3, 4) = this_cell
    


'--------------------------------------------
    'X = 4 Y = 4 cell
    
    Set v1 = New vertex
    Call v1.construct(30, 30, 10)
    Set v2 = New vertex
    Call v2.construct(24, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 24, 10)
    Call v2.construct(30, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 24, 10)
    Call v2.construct(30, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(24, 30, 10)
    Call v2.construct(24, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(4, 4) = this_cell

 '--------------------------------------------
    'X = 5 Y = 4 cell
    
    Set v1 = New vertex
    Call v1.construct(36, 30, 10)
    Set v2 = New vertex
    Call v2.construct(30, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 24, 10)
    Call v2.construct(36, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlack, False)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 24, 10)
    Call v2.construct(36, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(30, 30, 10)
    Call v2.construct(30, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(5, 4) = this_cell

 '--------------------------------------------
    'X = 6 Y = 4 cell
    
    Set v1 = New vertex
    Call v1.construct(42, 30, 10)
    Set v2 = New vertex
    Call v2.construct(36, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 24, 10)
    Call v2.construct(42, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 24, 10)
    Call v2.construct(42, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(36, 30, 10)
    Call v2.construct(36, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, False)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(6, 4) = this_cell

 '--------------------------------------------
    'X = 7 Y = 4 cell
    
    Set v1 = New vertex
    Call v1.construct(48, 30, 10)
    Set v2 = New vertex
    Call v2.construct(42, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 24, 10)
    Call v2.construct(48, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 24, 10)
    Call v2.construct(48, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbYellow, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(42, 30, 10)
    Call v2.construct(42, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbGreen, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(7, 4) = this_cell

  '--------------------------------------------
    'X = 8 Y = 4 cell
    
    Set v1 = New vertex
    Call v1.construct(54, 30, 10)
    Set v2 = New vertex
    Call v2.construct(48, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbMagenta, True)
    Set e(0) = this_edge
    
    
    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(54, 24, 10)
    Call v2.construct(54, 30, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbCyan, True)
    Set e(1) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 24, 10)
    Call v2.construct(54, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbRed, True)
    Set e(2) = this_edge


    Set v1 = New vertex
    Set v2 = New vertex
    Call v1.construct(48, 30, 10)
    Call v2.construct(48, 24, 10)
    Set this_edge = New Edge
    Call this_edge.constructor(v1, v2, vbBlue, True)
    Set e(3) = this_edge


    Set this_cell = New cell
    Call this_cell.constructor(e())

    Set cell_array(8, 4) = this_cell


    
    


End Function



Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


