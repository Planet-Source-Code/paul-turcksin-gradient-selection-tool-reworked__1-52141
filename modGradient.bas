Attribute VB_Name = "modGradient"
Option Explicit

Private Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, ByRef pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Integer
Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Const GRADIENT_FILL_RECT_H = 0
Private Const GRADIENT_FILL_RECT_V = 1

Dim arVert(1) As TRIVERTEX
Dim gRect As GRADIENT_RECT

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Sub subShowGradient(sPic As PictureBox, sOrientation As Boolean, sDirection As Boolean, sInitialColor As Long, sFinalColor As Long)
'   sPic : destination picturebox
'   sOrientation : Horizontal (True) or Vertical (False)
'   sDirection : Left to right(True) or Right to left (False) (horizontal)
'                Top to bottom (True) or bottom to top (False) (vertical)
'   sInitialColor , sFinalColor

   Dim arByteClr(3) As Byte   ' used to convert (long) color to its components
   Dim arByteVert(7) As Byte   ' used to init color part of vertices array
   Dim iOrientation As Long
   
    On Local Error Resume Next
    
' init vertices : position, size and direction
   If sDirection Then
      arVert(0).X = 0: arVert(1).X = sPic.ScaleWidth
      arVert(0).Y = 0: arVert(1).Y = sPic.ScaleHeight
   Else
      arVert(0).X = sPic.ScaleWidth:  arVert(1).X = 0
      arVert(0).Y = sPic.ScaleHeight: arVert(1).Y = 0
      End If
   
' init vertices :colors, initial
   CopyMemory arByteClr(0), sInitialColor, 4
   arByteVert(1) = arByteClr(0)   ' red
   arByteVert(3) = arByteClr(1)   ' green
   arByteVert(5) = arByteClr(2)   ' blue
   CopyMemory arVert(0).Red, arByteVert(0), 8

' init vertices :colors, final
   CopyMemory arByteClr(0), sFinalColor, 4
   arByteVert(1) = arByteClr(0)   ' red
   arByteVert(3) = arByteClr(1)   ' green
   arByteVert(5) = arByteClr(2)   ' blue
   CopyMemory arVert(1).Red, arByteVert(0), 8

' init gradient rect
   gRect.UpperLeft = 0
   gRect.LowerRight = 1
    
   iOrientation = IIf(sOrientation, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)
   
   GradientFill sPic.hDC, arVert(0), 2, gRect, 1, iOrientation
   sPic.Refresh
    On Error GoTo 0
End Sub
