Attribute VB_Name = "General"
Option Explicit

Sub Gradient(TheObject As Object, Redval&, Greenval&, Blueval&)

'TheObject can be any object that supports the Line method (like forms and pictures).
'Redval, Greenval, and Blueval are the Red, Green, and Blue starting values from 0 to 255.

Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$

'This will create 63 steps in the gradient. This looks smooth on 16-bit and 24-bit color.
'You can change this, but be careful. You can do some strange-looking stuff with it...
Step = (TheObject.Height / 63)
    
FillLeft = 0
FillRight = TheObject.Width
FillBottom = FillTop + Step

'If you changed the number of steps, change the number of reps to match it.
'If you don't, the gradient will look all funny.
For Reps = 1 To 63
    TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), BF
    'This decreases the RGB values to darken the color.
    'Lower the value for "squished" gradients. Raise it for incomplete gradients.
    'Also, if you change the number of steps, you will need to change this number.
    Redval = Redval - 4
    Greenval = Greenval - 4
    Blueval = Blueval - 4
    'This prevents the RGB values from becoming negative, which causes a runtime error.
    If Redval <= 0 Then Redval = 0
    If Greenval <= 0 Then Greenval = 0
    If Blueval <= 0 Then Blueval = 0
    FillTop = FillBottom
    FillBottom = FillTop + Step
Next

End Sub


