Attribute VB_Name = "Module1"
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

Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
  ByVal bScan As Byte, ByVal dwFlags As Long, _
  ByVal dwExtraInfo As Long)
'The above declaration declares keybd_event as a keystroke sent
'to the computer from the keyboard. Compare to "sendkeys".
'I use this because it's just the way I was taught, and it's better
'to use in terms of managing the picture easily, and sending it
'straight to the picture box.


Sub Main()
    Const VK_SNAPSHOT As Byte = &H2C       'Sets VK_SNAPSHOT as the Print_Scrn key.
    Call keybd_event(VK_SNAPSHOT, 0, 0, 0) 'Sends a fake keystroke to the computer
                                           '(Printscreen key)
    DoEvents                               'Waits for the computer to finish.
    Form1.Picture1.Picture = Clipboard.GetData(vbCFBitmap) 'Sends picture to pic. box
    Form1.Show                             'Shows the form, and starts the program.
End Sub

