VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3480
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      DrawMode        =   1  'Blackness
      DrawWidth       =   2
      Height          =   1935
      Left            =   600
      ScaleHeight     =   1875
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------+
'|                          DOT SCREEN SAVER - Kamaron Peterson                  |
'|  This is kind of a funny screen saver actually - I origionally wrote it on my |
'|graphing calculator, and not for the computer. But I decided I would share it  |
'|with the world, so if in an extremely wierd scenario where this is needed, it's|
'|there. I seriously doubt that would ever happen, but I still put it up. This   |
'|isn't the best way to write a program that does exactly this, but it's good for|
'|providing examples on how to use the keybd_event, picture boxes, rnd command,  |
'|and a few odd methoods of doing this. Have fun, I don't care if you use it     |
'|somewhere else, as long as you say somewhere where one can see that I made it, |
'|or at least give some sort of credit. Email me at KamJPetey@hotmail.com for    |
'|the graphing calculator code (for TI-83 and TI-84, all editions)               |
'|                                        -Kamaron Peterson.                     |
'+-------------------------------------------------------------------------------+

Private Xcoor, Ycoor As Double  'coordinates of dot being drawn
Private Direction As Double     'direction of the next dot to be drawn
Private ScreenWidth As Integer  'holds screen width information
Private ScreenHeight As Integer 'holds screen height information
Private Step As Integer         'how many dots have been drawn
Private xScale, yScale As Double 'how far away from the previous dot
                                    'the next dot will be, x and y

Private Sub Form_Click()
    End
End Sub

Private Sub Form_Load()
    Me.Left = 0                 'Sets form against the right of the screen.
    Me.Top = 0                  'sets form against the top of the screen.
    Me.Width = Screen.Width     'sets form to be as wide as the screen.
    Me.Height = Screen.Height   'sets form to be as tall as the screen.
    Picture1.Left = 0           'sets picture box to be against the right of the screen.
    Picture1.Top = 0            '... yeah. Basically the same thing as the form did earlier.
    Picture1.Width = Me.Width   '...
    Picture1.Height = Me.Height '...
    Xcoor = Rnd * (ScreenWidth) 'Sets the first dot as any random point in the screen (x axis)
    Ycoor = Rnd * (ScreenHeight) 'Sets the first dot as any random point in the screen (y axis)
    ScreenWidth = Screen.Width  'Sets ScreenWidth information
    ScreenHeight = Screen.Height 'Sets ScreenHeight information
    Step = 0                    'No dots have been drawn.
End Sub

Private Sub Picture1_Click()
    End     'The program ends when you click the screen.
End Sub

Private Sub Timer1_Timer()
    Direction = Int(Rnd * 8) + 1    'Sets Direction as any number, 1-8
    xScale = Rnd * 500              'Sets xScale as any random number from 1-500
    yScale = Rnd * 500              'Same for y
    Select Case Direction           'Loop - for a visual, imagine this:
    Case 1                          '------500 pixels from last dot drawn-----
    Xcoor = Xcoor - xScale          ' 1 -  up  |    2 - just    | 3 - up and |
    Ycoor = Ycoor + yScale          '   and    |       up       |    right   |
    Case 2                          '   left   |                |            |
    Ycoor = Ycoor + yScale          '----------------------------------------|
    Case 3                          ' 4 - just |                | 5 - just   |  500 pixels
    Xcoor = Xcoor + xScale          '   left   | Last Dot Drawn |   right    |  from last
    Ycoor = Ycoor + yScale          '          |    Position    |            |  dot drawn
    Case 4                          '----------------------------------------|
    Xcoor = Xcoor - xScale          ' 6 - down |  7 - just down |  8 - down  |
    Case 5                          '   and    |                |   and right|
    Xcoor = Xcoor + xScale          '    left  |                |            |
    Case 6                          '-----500 pixels from last dot drawn-----
    Xcoor = Xcoor - xScale          'The math operations basically move the next point
    Ycoor = Ycoor - yScale          'to any random point within 500 pixels of the last
    Case 7                          'point, and in it's appropriate section.
    Ycoor = Ycoor - yScale
    Case 8
    Xcoor = Xcoor + xScale
    Ycoor = Ycoor - yScale
    End Select
    If Xcoor <= 0 Then                      'If the point goes off the screen, then
        Xcoor = ScreenWidth - Abs(Xcoor)    'these operations simply put it back on,
    ElseIf Xcoor >= ScreenWidth Then        'but on the oposite side of the screen.
        Xcoor = Xcoor - ScreenWidth         'i.e. - if a point goes off the screen to
    End If                                  'the right by 80 pixels, then this math puts
    If Ycoor <= 0 Then                      'it back on, 80 pixels from the left of the
        Ycoor = ScreenHeight - Abs(Ycoor)   'screen, as well as vice versa and for the
    ElseIf Ycoor >= ScreenHeight Then       'y-coordinate.
        Ycoor = Ycoor - ScreenHeight
    End If
    Picture1.PSet (Xcoor, Ycoor), ForeColor 'Draws the point. Simple as that.
    If Step = 10000 Then              'If 10,000 points have been drawn, then
        Step = 0                      'Reset the count, and
        Picture1.Refresh              'Clear the picture box
        Form_Load                     'and restart the program.
    End If
    Step = Step + 1                   'Count the point that has just been drawn.
End Sub
