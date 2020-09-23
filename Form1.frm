VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gradient Tool"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frDirection 
      Caption         =   "Direction"
      Height          =   975
      Left            =   2880
      TabIndex        =   21
      Top             =   1920
      Width           =   1575
      Begin VB.OptionButton optDirection 
         Caption         =   "Right to left"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1400
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Left to right"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Value           =   -1  'True
         Width           =   1400
      End
   End
   Begin VB.Frame frOrientation 
      Caption         =   "Orientation"
      Height          =   975
      Left            =   360
      TabIndex        =   18
      Top             =   1920
      Width           =   1455
      Begin VB.OptionButton optOrientation 
         Caption         =   "Vertical"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1000
      End
      Begin VB.OptionButton optOrientation 
         Caption         =   "Horizontal"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Value           =   -1  'True
         Width           =   1000
      End
   End
   Begin VB.HScrollBar scrollBlue 
      Height          =   195
      Index           =   1
      LargeChange     =   10
      Left            =   3000
      Max             =   255
      TabIndex        =   12
      Top             =   1260
      Width           =   1455
   End
   Begin VB.HScrollBar scrollGreen 
      Height          =   195
      Index           =   1
      LargeChange     =   10
      Left            =   3000
      Max             =   255
      TabIndex        =   11
      Top             =   1020
      Width           =   1455
   End
   Begin VB.HScrollBar scrollRed 
      Height          =   195
      Index           =   1
      LargeChange     =   10
      Left            =   3000
      Max             =   255
      TabIndex        =   10
      Top             =   780
      Width           =   1455
   End
   Begin VB.HScrollBar scrollBlue 
      Height          =   195
      Index           =   0
      LargeChange     =   10
      Left            =   300
      Max             =   255
      TabIndex        =   9
      Top             =   1260
      Width           =   1455
   End
   Begin VB.HScrollBar scrollGreen 
      Height          =   195
      Index           =   0
      LargeChange     =   10
      Left            =   300
      Max             =   255
      TabIndex        =   8
      Top             =   1020
      Width           =   1455
   End
   Begin VB.HScrollBar scrollRed 
      Height          =   195
      Index           =   0
      LargeChange     =   10
      Left            =   300
      Max             =   255
      TabIndex        =   7
      Top             =   780
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   315
      Left            =   3480
      TabIndex        =   0
      Top             =   6000
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog DialogColor 
      Left            =   240
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2715
      Left            =   360
      ScaleHeight     =   2655
      ScaleWidth      =   4095
      TabIndex        =   5
      Top             =   3120
      Width           =   4155
   End
   Begin VB.Label lblColorText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   17
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblColorText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1800
      TabIndex        =   15
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1875
      TabIndex        =   14
      Top             =   960
      Width           =   1020
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1815
      TabIndex        =   13
      Top             =   780
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(click)"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1860
      TabIndex        =   6
      Top             =   420
      Width           =   1140
   End
   Begin VB.Label fSelectColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Final Color:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label fSelectColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   300
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Initial Color:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type MY_RGB
   Red As Byte
   Green As Byte
   Blue As Byte
   End Type
Dim arClr(1) As MY_RGB        ' (0) : initial color; (1) : final color
Private Const cInitial = 0
Private Const cFinal = 1

' used to skip the change event code of the scrollbars when the values of these
' scrollbars are update by code
Dim swSkipScroll As Boolean

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()

' precaution

   Picture1.ScaleMode = vbPixels
   Picture1.AutoRedraw = True
   
' init all color stuff to white
   fSelectColor(cInitial).BackColor = vbWhite:   subConvertClr cInitial
   fSelectColor(cFinal).BackColor = vbWhite:     subConvertClr cFinal
   
' init color common dialog
   DialogColor.CancelError = True
   DialogColor.Flags = cdlCCFullOpen Or cdlCCRGBInit
End Sub

Private Sub fSelectColor_Click(Index As Integer)
   On Error GoTo ErrColor
   
' call color common dialog
   With DialogColor
      .Color = fSelectColor(Index).BackColor
      .ShowColor
      fSelectColor(Index).BackColor = .Color
      End With
      
' extract color components
   subConvertClr Index
      
   subShowGradient Picture1, optOrientation(0).Value, optDirection(0).Value, _
                   fSelectColor(cInitial).BackColor, fSelectColor(cFinal).BackColor
ErrColor:
   On Error GoTo 0
End Sub

Private Sub optDirection_Click(Index As Integer)
   subShowGradient Picture1, optOrientation(0).Value, optDirection(0).Value, _
                   fSelectColor(cInitial).BackColor, fSelectColor(cFinal).BackColor
End Sub

Private Sub optOrientation_Click(Index As Integer)
   If optOrientation(0).Value Then
      optDirection(0).Caption = "Left to right"
      optDirection(1).Caption = "Right to left"
   Else
      optDirection(0).Caption = "Top to bottom"
      optDirection(1).Caption = "Bottom to top"
      End If
      
   subShowGradient Picture1, optOrientation(0).Value, optDirection(0).Value, _
                   fSelectColor(cInitial).BackColor, fSelectColor(cFinal).BackColor
End Sub

Private Sub scrollBlue_Change(Index As Integer)
' update stored colors and color composition on form. Show color on the form

' skip if updated by code
   If swSkipScroll Then Exit Sub
   
   With arClr(Index)
      .Blue = scrollBlue(Index).Value
      
      lblColorText(Index) = Format(.Red) & "," & Format(.Green) & "," & Format(.Blue)
      fSelectColor(Index).BackColor = RGB(.Red, .Green, .Blue)
      End With
      
   subShowGradient Picture1, optOrientation(0).Value, optDirection(0).Value, _
                   fSelectColor(cInitial).BackColor, fSelectColor(cFinal).BackColor
End Sub

Private Sub scrollGreen_Change(Index As Integer)
' update stored colors and color composition on form. Show color on the form

' skip if updated by code
   If swSkipScroll Then Exit Sub
   
   With arClr(Index)
      .Green = scrollGreen(Index).Value
      
      lblColorText(Index) = Format(.Red) & "," & Format(.Green) & "," & Format(.Blue)
      fSelectColor(Index).BackColor = RGB(.Red, .Green, .Blue)
      End With
      
   subShowGradient Picture1, optOrientation(0).Value, optDirection(0).Value, _
                   fSelectColor(cInitial).BackColor, fSelectColor(cFinal).BackColor
End Sub

Private Sub scrollRed_Change(Index As Integer)
' update stored colors and color composition on form. Show color on the form

' skip if updated by code
   If swSkipScroll Then Exit Sub
   
   With arClr(Index)
      .Red = scrollRed(Index).Value
      
      lblColorText(Index) = Format(.Red) & "," & Format(.Green) & "," & Format(.Blue)
      fSelectColor(Index).BackColor = RGB(.Red, .Green, .Blue)
      End With
      
   subShowGradient Picture1, optOrientation(0).Value, optDirection(0).Value, _
                   fSelectColor(cInitial).BackColor, fSelectColor(cFinal).BackColor
End Sub



'======================================================================================
'                                     LOCAL PROCEDURES
'____________________________________________________________________________________

Private Sub subConvertClr(sType As Integer)
'   sType : Initial or final
'  Converts the color selected in the label box to RGB values. These values are stored,
'  shown on the form and used to update the values of the scrollbars

   Dim arByte(3) As Byte
   
   CopyMemory arByte(0), fSelectColor(sType).BackColor, 4
   With arClr(sType)
      .Red = arByte(0)
      .Green = arByte(1)
      .Blue = arByte(2)
      
      lblColorText(sType) = Format(.Red) & "," & Format(.Green) & "," & Format(.Blue)
 
      ' avoid executing code in change event
      swSkipScroll = True
     
      scrollRed(sType).Value = .Red
      scrollGreen(sType).Value = .Green
      scrollBlue(sType).Value = .Blue
      swSkipScroll = False
      End With
      
End Sub

