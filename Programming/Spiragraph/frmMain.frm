VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Spiragraph"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar scrWidth 
      Height          =   255
      Left            =   240
      Max             =   5
      Min             =   1
      TabIndex        =   9
      Top             =   3360
      Value           =   1
      Width           =   1455
   End
   Begin VB.HScrollBar scrY 
      Height          =   255
      Left            =   240
      Max             =   5
      Min             =   1
      TabIndex        =   7
      Top             =   2760
      Value           =   1
      Width           =   1455
   End
   Begin VB.HScrollBar scrX 
      Height          =   255
      Left            =   240
      Max             =   5
      Min             =   1
      TabIndex        =   5
      Top             =   2160
      Value           =   1
      Width           =   1455
   End
   Begin VB.CheckBox chkColour 
      Caption         =   "Colour Cycle"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.HScrollBar scrIterations 
      Height          =   255
      LargeChange     =   2
      Left            =   240
      Max             =   10
      Min             =   1
      TabIndex        =   2
      Top             =   960
      Value           =   1
      Width           =   1455
   End
   Begin VB.HScrollBar scrStep 
      Height          =   255
      LargeChange     =   2
      Left            =   240
      Max             =   20
      Min             =   1
      TabIndex        =   0
      Top             =   360
      Value           =   1
      Width           =   1455
   End
   Begin VB.Label lblWidth 
      Caption         =   "Line Width: 1"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblY 
      Caption         =   "Y Multiplier: 1"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblX 
      Caption         =   "X Multiplier: 1"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblIterations 
      Caption         =   "Iterations: 1"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblStep 
      Caption         =   "Step: 1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DrawCircle(offsetX As Single, offsetY As Single, colour As Long)

Dim count As Integer, X As Long, Y As Long, deg2rad As Double
deg2rad = (3.14159265358979 / 180)

frmMain.CurrentX = (Sin(0) * 800) + offsetX
frmMain.CurrentY = (Cos(0) * 800) + offsetY

If chkColour.Value = 0 Then colour = 0

For count = 0 To 360
    X = (Sin(count * deg2rad) * 800) + offsetX
    Y = (Cos(count * deg2rad) * 800) + offsetY
    frmMain.Line -(X, Y), colour
    DoEvents
Next count

End Sub

Private Sub Form_Click()

Dim count As Integer, X As Single, Y As Single, deg2rad As Double, col As Integer
deg2rad = (3.14159265358979 / 180)

frmMain.Cls

For count = 0 To (360 * scrIterations.Value) Step scrStep.Value
    X = (Sin(count * scrX.Value * deg2rad) * 1800) + (frmMain.Width / 2)
    Y = (Cos(count * scrY.Value * deg2rad) * 1800) + (frmMain.Height / 2)
    col = (count / scrIterations.Value) * (255 / 360)
    DrawCircle X, Y, RGB(col / 4, col / 2, col / 3)
Next count

End Sub

Private Sub scrIterations_Change()

lblIterations.Caption = "Iterations:" + Str$(scrIterations.Value)

End Sub

Private Sub scrStep_Change()

lblStep.Caption = "Step:" + Str$(scrStep.Value)

End Sub

Private Sub scrWidth_Change()

lblWidth.Caption = "Line Width:" + Str$(scrWidth.Value)
frmMain.DrawWidth = scrWidth.Value

End Sub

Private Sub scrX_Change()

lblX.Caption = "X Multiplier:" + Str$(scrX.Value)

End Sub

Private Sub scrY_Change()

lblY.Caption = "Y Multiplier:" + Str$(scrY.Value)

End Sub
