VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   404
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   119
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim w, h As Integer
Dim col(2000) As Integer        'our colours for each star
Dim star(2000) As p3d           'our 3d coords
Dim star2d(2000) As p2d            'our 2d coords
Dim ex, wy As Boolean               'camera flags
Dim cam As p2d                  'camera coords


Private Type p3d    'custom type for 3d coords we calculate from
X As Long
Y As Long
z As Long
End Type


Private Type p2d    'custom type for 2d coords we will put on screen
X As Long
Y As Long

End Type

Dim halt As Boolean
Const num = 1000        'amount of stars (compile if your using a high number for extra speed
Const zoff = 510        ' max distance for stars
Const speed = 2         'speed the stars and camera move at



Private Sub init()
Dim i As Integer        'apply to all INDIVIDUAL star co-ord
p.Width = Screen.Width / Screen.TwipsPerPixelX      'make the picture box the same size as the screen
p.Height = Screen.Height / Screen.TwipsPerPixelY    'make the picture box the same size as the screen

w = p.ScaleWidth        'get the scalewidth
h = p.ScaleHeight       'get the scaleheight

For i = 1 To num        'apply to all INDIVIDUAL star co-ord
star(i).X = Rnd * w     'Give each its random co-ords (3D)
star(i).Y = Rnd * h
star(i).z = Rnd * zoff
Next i

cam.X = Rnd * w            'random camera coords
cam.Y = Rnd * h
ex = True               'camera movement
wy = True
End Sub

Private Sub mve()
Dim i As Integer

For i = 1 To num                 'apply to all INDIVIDUAL star co-ord

star(i).z = star(i).z - speed       'decreasing the z co-ord brings the stars closer to us

If star(i).z < 75 Then               'when its close...
star(i).z = Rnd * zoff              'create a new one :D
star(i).X = Rnd * w
star(i).Y = Rnd * h

End If

Next i

End Sub

Private Sub convert()
Dim i As Integer

'3d to 2d conversion...
'We need to plot these 3d co-ords onto a 2d screen so:

For i = 1 To num            'apply to all INDIVIDUAL star co-ord

    col(i) = 255 - (star(i).z / 2)  'get the colour from its distance
    
    If star(i).z > 0 Then              'make sure there is no division by zero
        'formula to convert 3d to 2d
        star2d(i).X = cam.X - (256 * ((w / 2) - star(i).X)) / star(i).z
        star2d(i).Y = cam.Y - (256 * ((h / 2) - star(i).Y)) / star(i).z
    End If
    
Next i


End Sub
Private Sub Form_activate()
init
Do
DoEvents

camera  'move the direction of view
mve     'move the stars inwards
convert 'Change 3d coords to 2d (that we are able to plot)
p.Cls   'clear the screen to remove "trails"
penit   'draw the stars
DoEvents

Loop Until halt = True
End Sub

Private Sub p_Click()

halt = True 'clicking the screen ends the loop
End Sub

Private Sub penit()
Dim i As Integer

For i = 1 To num        'apply to all INDIVIDUAL star co-ord

p.PSet (star2d(i).X, star2d(i).Y), RGB(col(i), col(i), col(i)) 'draw stars (pset is perfect for this, setpixel too fast)
Next
End Sub

Private Sub p_KeyPress(KeyAscii As Integer)
End ' kill it ;)
End Sub
Private Sub camera()
'The bouncy camera

If cam.X < 0 Then ex = True        'decide which way the camera moves
If cam.X > w Then ex = False

If cam.Y < 0 Then wy = True
If cam.Y > h Then wy = False

                                'move it on x-axis
Select Case ex

Case True
cam.X = cam.X + speed
Case False
cam.X = cam.X - speed
End Select

Select Case wy                  'move it on y-axis
    
Case True
cam.Y = cam.Y + speed
Case False
cam.Y = cam.Y - speed
End Select

End Sub
Private Sub p_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'un-comment these and remove the call to camera to allow movement with the mouse
'cam.x = X
'cam.y = Y
End Sub
