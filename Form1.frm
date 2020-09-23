VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FireWorks"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5160
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   6135
      Left            =   120
      ScaleHeight     =   6075
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.Timer Timer2 
         Interval        =   200
         Left            =   3960
         Top             =   1440
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************'
'* FireWorks V1.0           - By Jongmin Baek *'
'**********************************************'

'**********************************************'
'*                  Caution                   *'
'*                                            *'
'* You should put two timers, and a picturebox*'
'*                                            *'
'*  Their names should be Timer1, Timer2,     *'
'*              and Picture1                  *'
'*                                            *'
'* And form1 must not have Minimize button.   *'
'* If you use it, the program will make an    *'
'* error.                                     *'
'**********************************************'

Dim NowBeingUsed(100) As Boolean    'Do not touch it!
Dim Direction(100) As Integer       'the direction of the each flame
Dim Speed(100) As Integer           'the speed of the each flame
Dim Y(100) As Integer               'the location of the each flame
Dim X(100) As Integer               '               "
Dim ColorRGB(100, 2) As Long        'the color of the each flame(RGB)
                                        'ColorRGB(n, 0) -> Red value
                                        'ColorRGB(n, 1) -> Green value
                                        'ColorRGB(n, 2) -> Blue value

Dim Scatter As Integer              'It decides how far the flames are
                                    'scattered when they launched
Dim ScatterFalling As Integer       'It decides how far the flames are
                                    'scattered when they start falling.
Dim SpeedDif As Integer             'The larger it is, the distance between
                                    'the higher flame and the lower flame
                                    'become bigger.
Dim Size As Integer                 'Each flame's size
Private Sub Form_Load()
Timer1.Interval = 10
Timer2.Interval = 200   'Read Timer2_Timer and if you want, change it!
Form1.Caption = "FireWorks"
Scatter = 5
ScatterFalling = 60
SpeedDif = 30
Size = 30                   'These value are good for the view.
                            'If you want, test many values.
Form_Resize
End Sub
Private Sub Form_Resize()
Picture1.Top = 120
Picture1.Left = 120
On Error GoTo 100
Picture1.Width = Form1.Width - 345
Picture1.Height = Form1.Height - 545
            'These code change the picturebox's size when you resize
            'the form.
100 End Sub
Private Sub Timer1_Timer()
        'Whenever this timer work, all of the flames move as their
        'direction & speed
For i = 1 To 100
If NowBeingUsed(i) = False Then GoTo 20
        'If the flame isn't on the screen, we don't need to process.
Picture1.Line (X(i) - Size, Y(i) - Size)-(X(i) + Size, Y(i) + Size), 0, BF
        'erase the original flame.
X(i) = X(i) - Cos(Direction(i) * 3.141592654 / 180) * Speed(i)
Y(i) = Y(i) - Sin(Direction(i) * 3.141592654 / 180) * Speed(i)
        'move the location as its direction & speed
If Direction(i) >= 0 And Direction(i) <= 180 Then
        Speed(i) = Speed(i) - 7
        'The higher it fly, the slower its speed become.
        '(because of the gravity) so I make it slower as it goes higher.
        If Speed(i) < 0 Then
            Speed(i) = 0
            Direction(i) = Int(Rnd(1) * ScatterFalling * 2) + 270 - ScatterFalling
            'if its speed became zero, it should be fell down.
        End If
Else
        Speed(i) = Speed(i) + 7
            'I make the speed fast as it fall down
            'because of the gravity
        If Speed(i) > 80 Then
            'if the flame should be removed,
            NowBeingUsed(i) = False: GoTo 20
        Else
            GoTo 25
        End If
End If
'Don't touch - Picture1.Line (X(i) - Size, Y(i) - Size)-(X(i) + Size, Y(i) + Size), RGB(ColorRGB(i, 0), ColorRGB(i, 1), ColorRGB(i, 2)), BF: GoTo 20
A1 = Picture1.ScaleHeight / 40 + 150 + SpeedDif
A2 = ((A1 - Speed(i)) / 2 + A1 / 2) / A1
Picture1.Line (X(i) - Size * A2, Y(i) - Size * A2)-(X(i) + Size * A2, Y(i) + Size * A2), RGB(ColorRGB(i, 0), ColorRGB(i, 1), ColorRGB(i, 2)), BF: GoTo 20
'I draw a new flame for it (it is going higher)
25
R = ColorRGB(i, 0) * (80 - Speed(i)) / 80
G = ColorRGB(i, 1) * (80 - Speed(i)) / 80
B = ColorRGB(i, 2) * (80 - Speed(i)) / 80
Picture1.Line (X(i) - Size, Y(i) - Size)-(X(i) + Size, Y(i) + Size), RGB(R, G, B), BF
'I draw a new flame for it (it is falling down)
'As it fall down, its color become dark.
'Because it will be removed soon, so it should be ready for being removed.
20 Next i
End Sub
Private Sub Timer2_Timer()
        'Whenever this timer work, a new set of flames are launched
Randomize Timer
10 R = Int(Rnd(1) * 128) + 128
   G = Int(Rnd(1) * 128) + 128
   B = Int(Rnd(1) * 128) + 128  'Decide the color of the fire
If Abs(R - G) < 30 And Abs(G - B) < 30 And Abs(R - B) < 30 Then GoTo 10
    'if the color is like gray, it makes a new color
    '(because I don't like gray color).
LocationX = Int(Rnd(1) * (Picture1.ScaleWidth - 2000)) + 1000
MainDirection = Int(Rnd(1) * 40) + 70   'Main Direction
MainSpeed = Picture1.ScaleHeight / 40 + 50 + Int(Rnd(1) * 101) 'Main Speed
For i = 1 To 100
If NowBeingUsed(i) = True Then GoTo 30 'If it is being used, I can't use it
NowBeingUsed(i) = True

X(i) = LocationX
Y(i) = Picture1.ScaleHeight
ColorRGB(i, 0) = R
ColorRGB(i, 1) = G
ColorRGB(i, 2) = B      'Save the values.

Direction(i) = MainDirection + Int(Rnd(1) * Scatter * 2) - Scatter
Speed(i) = MainSpeed + Int(Rnd(1) * SpeedDif * 2) - SpeedDif
30 Next i
End Sub
