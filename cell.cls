VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private edges(0 To 3) As Edge
Public north As New cell, south As New cell, east As New cell, west As New cell
Public x As Integer, y As Integer


Public Function constructor(e() As Edge)
    For a = 0 To 3
        Set edges(a) = e(a)
    Next a
    Set north = Nothing
    Set south = Nothing
    Set east = Nothing
    Set west = Nothing
    
    
    
End Function

Public Function getEdges() As Edge()
    getEdges = edges
End Function

Public Function setNeighbor(side As Integer, c As cell)
    '0 = north
    '1 = east
    '2 = south
    '3 = west
    If side = 0 Then
        Set north = c
        Set edges(0).neighbor = c
    ElseIf side = 1 Then
        Set east = c
        Set edges(1).neighbor = c
    ElseIf side = 2 Then
        Set south = c
        Set edges(2).neighbor = c
    ElseIf side = 3 Then
        Set west = c
        Set edges(3).neighbor = c
    End If
     
    
End Function
