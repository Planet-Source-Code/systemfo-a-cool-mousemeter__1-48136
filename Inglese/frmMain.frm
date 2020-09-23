VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E9C570&
   Caption         =   "Mouse meter"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   ScaleHeight     =   196
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E9C570&
      Caption         =   "On top"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Timer timVel 
      Interval        =   1000
      Left            =   0
      Top             =   360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E9C570&
      Caption         =   "Data"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label5 
         BackColor       =   &H00E9C570&
         Caption         =   "Speed (m/sec): "
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E9C570&
         Caption         =   "Speed (pixel/sec): "
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E9C570&
         Caption         =   "Total distance (m):"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E9C570&
         Caption         =   "Position:"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E9C570&
         Caption         =   "Total distance: (pixel):"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
   End
   Begin VB.Timer tmr1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PositionBefore As POINTAPI, PositionNow As POINTAPI
Private TDist As Long
Private BeforeDistance As Long, NowDistance As Long
Private SpeedPixel_Sec As Single, Speedml_Sec As Single

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Call api.StayOnTop(Me)
    Debug.Print Me.Top
Else
    Call api.NotOnTop(Me)
End If
End Sub

Private Sub Form_Load()
Call api.SetCursorPos(0, 0)
End Sub

Private Sub timVel_Timer()
BeforeDistance = NowDistance
NowDistance = TDist
'Velocità in Px\sec.
SpeedPixel_Sec = (NowDistance - BeforeDistance)
'Velocità in m/sec.
Speedml_Sec = SpeedPixel_Sec * Screen.TwipsPerPixelX / api.Twip_m
Label4 = "Speed (pixel/sec): " & SpeedPixel_Sec
Label5 = "Speed (m/sec): " & Format(Speedml_Sec, "#0.0#")
End Sub

Private Sub tmr1_Timer()
PositionBefore = PositionNow
Call api.GetCursorPos(PositionNow)
Label1 = "Position: " & PositionNow.x & ", " & PositionNow.y
TDist = TDist + Distance(PositionBefore, PositionNow)
Label2 = "Total distance (pixel): " & TDist
Label3 = "Total distance (m): " & Format(TDist * Screen.TwipsPerPixelX / api.Twip_m, "###0.000")
End Sub

Private Function Distance(p1 As POINTAPI, p2 As POINTAPI)
Dim v As Integer, h As Integer

v = Abs(p1.y - p2.y)
h = Abs(p1.x - p2.x)
'Pitagora...
Distance = Int(Sqr(v ^ 2 + h ^ 2))
End Function
