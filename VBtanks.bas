Attribute VB_Name = "VBTANKS1"
Public Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Sub blue_shot()
Randomize Timer
Form1.DrawWidth = 1
' WindRes = Wind resistance
' Gravity = Gravity (set to on earth)
' Velocity = Power entered by player
' initYvol = Initial Y velocity
' initXvol = Initial X velocity
' iTime = Time increment

WindRes = 0      'Wind Resistence
Gravity = -8.13  'Gravity
Velocity = Val(Form1.BluePower) / 100 '(Form1.ScaleWidth / 100) 'Get Entered Speed
'Calculate X & Y Initial Velocity
initYvol = round(Sin(Val(Form1.BlueAngle) * (3.14 / 180)), 2) * Velocity
initXvol = round(Cos(Val(Form1.BlueAngle) * (3.14 / 180)), 2) * Velocity
Form1.AutoRedraw = False
StartX = Form1.tank2.Left + Form1.tank2.Width / 2
StartY = Form1.tank2.Top + Form1.tank2.Height / 2
' START: Find tip of gun
Ang = -CInt(Form1.BlueAngle)
Ang = Ang * 3.1416 / 180
Length = 35
endx = -Length * Cos(Ang) + StartX
endy = Length * Sin(Ang) + StartY
' END: Found the tip
Do
    iTime = iTime + 0.01 'Increment Time
    disty = disty + (initYvol * iTime) + ((Gravity * (iTime ^ 2)) / 2) 'Distance along the Y Axis
    distx = distx - (initXvol * iTime) - ((WindRes * (iTime ^ 2)) / 2) 'Distance along the X Axis
    ' Check the color of the background where the bullet currently is
    landclr = Form1.Point(endx + distx, endy - disty)
    
    ' START: If enemy tank has been hit then...
    If endx + distx > Form1.tank1.Left And endx + distx < Form1.tank1.Left + Form1.tank1.Width And endy - disty > Form1.tank1.Top And endy - disty < Form1.tank1.Top + Form1.tank1.Height Then
        For explode = 0 To 100
        Form1.FillColor = QBColor(15 * Rnd)
        Form1.Circle (endx + distx, endy - disty), (100)
        Next explode
        MsgBox "Red dies! Nice shot blue!"
        ans = MsgBox("Do you want to play again?", 4)
        If ans = 6 Then
            landscape2 Form1
        Else
        Unload Form1
        End If
        Exit Sub
    End If
    ' END: A new game has started or current game ended
    If distx > Form1.ScaleWidth - endx Then distx = distx - Form1.ScaleWidth 'If shot goes out screen then return it
    If distx < (0 + Form1.ScaleWidth - endx) - Form1.ScaleWidth Then distx = distx + Form1.ScaleWidth 'If shot goes out screen then return it
    Form1.PSet (endx + distx, endy - disty), QBColor(15)
    DoEvents
    ' If the color of the background matches the land color then stop drawing lines and draw and explosion
    If landclr = 40960 Then GoTo 3
Loop While -disty < Form1.ScaleHeight 'And distx < Form1.ScaleWidth
3
Form1.FillColor = Form1.BackColor 'QBColor(9)
Form1.ForeColor = Form1.BackColor 'QBColor(9)
Form1.Circle (endx + distx, endy - disty), (25)
Form1.RedAngle.Enabled = True
Form1.RedPower.Enabled = True
Form1.tank2.Refresh
draw_gun
End Sub

Sub landscape(frm As Form)
' USAGE: landscape form#; example landscape frm will draw a landscape on frm
Randomize Timer
frm.Scale (0, 500)-(1000, -500)
frm.tank1.Visible = False 'HIDE TANK1
frm.tank2.Visible = False 'HIDE TANK2
frm.DrawWidth = 1

' START: Choose random sky color
r = Int(25 * Rnd) + 1
g = Int(25 * Rnd) + 1
b = Int(175 * Rnd) + 1
frm.BackColor = RGB(r, g, b)
' END: Sky color has been choosen

Static y
Static tank1pos
Static bump
'Static texture
frm.Cls
y = 0

' START: This draws landscape
For land = 0 To 1000
If y > 150 Then y = -150
    bump = Int(5 * Rnd) / 100
    y = y + bump
    bump = Sin(y)
    'texture = Int(frm.ScaleHeight * Rnd) + 1
    stars = Int(frm.ScaleHeight * Rnd) + 1
    'If texture > 0 Then texture = -texture
    If stars < 0 Then stars = -stars
    frm.Line (land, -500)-(land, 125 * bump), RGB(0, 175, 0) ' DRAW LAND
    'frm.PSet (land, 125 * bump + texture), RGB(150, 50, 50) ' ADD TEXTURE IN THE LAND
    frm.PSet (land, 125 * bump + stars), QBColor(15) ' ADD STARS ONLY ABOVE LAND
Next land
' END: Landscape and stars have been drawn

' START: Put tanks into position
frm.tank1.Move Int(500 * Rnd) + 1, 500
frm.tank2.Move Int(300 * Rnd) + 600, 500
frm.tank1.Visible = True
frm.tank2.Visible = True
frm.AutoRedraw = False
Do
frm.tank1.Move frm.tank1.Left, frm.tank1.Top - 10
If frm.Point(frm.tank1.Left, frm.tank1.Top - frm.tank1.Height) = 40960 And frm.Point(frm.tank1.Left + frm.tank1.Width, frm.tank1.Top - frm.tank1.Height) = 40960 Then Exit Do
Loop
Do
frm.tank2.Move frm.tank2.Left, frm.tank2.Top - 10
If frm.Point(frm.tank2.Left, frm.tank2.Top - frm.tank2.Height) = 40960 And frm.Point(frm.tank2.Left + frm.tank2.Width, frm.tank2.Top - frm.tank2.Height) = 40960 Then Exit Do
Loop
' END: Tanks are in position

' START: Draw moon
frm.FillColor = QBColor(14) ' SET FILLCOLOR TO YELLOW
frm.FillStyle = 0 ' SET FILLSTYLE TO SOLID
frm.Circle (1000 * Rnd, 400), (30), QBColor(14)
' END: Moon has been drawn

' START: Put boxes, scroll bars, and labels into postition
frm.RedPower.Move 20, -450
frm.Label2.Move 20, frm.RedPower.Top + frm.RedPower.Height
frm.RedAngle.Move 20, frm.Label2.Top + frm.Label2.Height + 5
frm.Label1.Move 20, frm.RedAngle.Top + frm.RedAngle.Height
frm.BluePower.Move 980 - frm.BluePower.Width, -450
frm.Label3.Move 980 - frm.Label3.Width, frm.BluePower.Top + frm.BluePower.Height
frm.BlueAngle.Move 980 - frm.BlueAngle.Width, frm.Label3.Top + frm.Label3.Height + 5
frm.Label4.Move 980 - frm.Label4.Width, frm.BlueAngle.Top + frm.BlueAngle.Height
frm.Fire.Move frm.ScaleWidth / 2, -450
' END: Everything is in place

' START: Show all boxes
frm.RedAngle.Visible = True
frm.BlueAngle.Visible = True
frm.RedPower.Visible = True
frm.BluePower.Visible = True
frm.Label1.Visible = True
frm.Label2.Visible = True
frm.Label3.Visible = True
frm.Label4.Visible = True
frm.Fire.Visible = True
' END: All boxes have been shown

frm.Label1.BackColor = RGB(0, 175, 0)
frm.Label2.BackColor = RGB(0, 175, 0)
frm.Label3.BackColor = RGB(0, 175, 0)
frm.Label4.BackColor = RGB(0, 175, 0)

frm.Label1.Refresh
frm.Label2.Refresh
frm.Label3.Refresh
frm.Label4.Refresh

' Rescale the form
frm.Scale (0, 0)-(1000, 1000)

' Make it red's turn
frm.Tag = "red"

End Sub

Sub red_shot()
Randomize Timer
Form1.DrawWidth = 1
' WindRes = Wind resistance
' Gravity = Gravity (set to on earth)
' Velocity = Power entered by player
' initYvol = Initial Y velocity
' initXvol = Initial X velocity
' iTime = Time increment

WindRes = 0      'Wind Resistence
Gravity = -8.13  'Gravity
Velocity = Val(Form1.RedPower) / 100 '(Form1.ScaleWidth / 100) 'Get Entered Speed
'Calculate X & Y Initial Velocity
initYvol = round(Sin(Val(Form1.RedAngle) * (3.14 / 180)), 2) * Velocity
initXvol = round(Cos(Val(Form1.RedAngle) * (3.14 / 180)), 2) * Velocity
Form1.AutoRedraw = False
StartX = Form1.tank1.Left + Form1.tank1.Width / 2
StartY = Form1.tank1.Top + Form1.tank1.Height / 2
' START: Find tip of gun
Ang = -CInt(Form1.RedAngle)
Ang = Ang * 3.1416 / 180
Length = 35
endx = Length * Cos(Ang) + StartX
endy = Length * Sin(Ang) + StartY
' END: Found the tip
Do
    iTime = iTime + 0.01 'Increment Time
    disty = disty + (initYvol * iTime) + ((Gravity * (iTime ^ 2)) / 2) 'Distance along the Y Axis
    distx = distx + (initXvol * iTime) + ((WindRes * (iTime ^ 2)) / 2) 'Distance along the X Axis
    landclr = Form1.Point(endx + distx, endy - disty)
    ' START: If the enemy tank is hit then...
    If endx + distx > Form1.tank2.Left And endx + distx < Form1.tank2.Left + Form1.tank2.Width And endy - disty > Form1.tank2.Top And endy - disty < Form1.tank2.Top + Form1.tank2.Height Then
        For explode = 0 To 100
        Form1.FillColor = QBColor(15 * Rnd)
        Form1.Circle (endx + distx, endy - disty), (100)
        Next explode
        MsgBox "Blue dies! Nice shot red!"
        ans = MsgBox("Do you want to play again?", 4)
        If ans = 6 Then
            landscape2 Form1
        Else
        Unload Form1
        End If
        Exit Sub
    End If
    ' END: A new game has started or current game ended
    If distx > Form1.ScaleWidth - endx Then distx = distx - Form1.ScaleWidth  'If shot goes out screen then return it
    If distx < (0 + Form1.ScaleWidth - endx) - Form1.ScaleWidth Then distx = distx + Form1.ScaleWidth 'If shot goes out screen then return it
    Form1.PSet (endx + distx, endy - disty), QBColor(12)
    DoEvents
    If landclr = 40960 Then GoTo 1
Loop While -disty < Form1.ScaleHeight
1
Form1.FillColor = Form1.BackColor 'QBColor(12)
Form1.ForeColor = Form1.BackColor 'QBColor(12)
Form1.Circle (endx + distx, endy - disty), (25)
Form1.RedAngle.Enabled = True
Form1.RedPower.Enabled = True
Form1.tank1.Refresh
draw_gun
End Sub

Function round(nValue As Double, nDigits As Integer) As Double
' USAGE: round(number,decimalpoint); example round(5.378,2) will come out to be 5.38
round = Int(nValue * (10 ^ nDigits) + 0.5) / (10 ^ nDigits)
End Function

Sub wait(howlong)
' USAGE: wait #ofseconds; example wait 3 will wait 3 seconds
temptime = Timer
Do
DoEvents
Loop While Timer < temptime + howlong
End Sub


Public Sub draw_gun()
Form1.DrawWidth = 3
If Form1.Tag = "red" Then
StartX = (Form1.tank1.Left + Form1.tank1.Width / 2)
StartY = (Form1.tank1.Top + Form1.tank1.Height / 2) + 8
Ang = -CInt(Form1.RedAngle)
Ang = Ang * 3.1416 / 180
Length = 18
endx = Length * Cos(Ang) + StartX
endy = Length * Sin(Ang) + StartY
Form1.tank1.Refresh
Form1.Line (StartX, StartY)-(endx, endy), QBColor(12)
End If

If Form1.Tag = "blue" Then
StartX = (Form1.tank2.Left + Form1.tank2.Width / 2)
StartY = (Form1.tank2.Top + Form1.tank2.Height / 2) + 8
Ang = -CInt(Form1.BlueAngle)
Ang = Ang * 3.1416 / 180
Length = 18
endx = -Length * Cos(Ang) + StartX
endy = Length * Sin(Ang) + StartY
Form1.tank2.Refresh
Form1.Line (StartX, StartY)-(endx, endy), QBColor(9)
End If

End Sub

Public Sub landscape2(frm As Form)
' USAGE: landscape2 form#; example: landscape2 frm will draw a landscape on frm

' START: Setting up the form for landscape
Randomize
frm.Scale (0, 0)-(1000, 1000)
frm.tank1.Visible = False 'HIDE TANK1
frm.tank2.Visible = False 'HIDE TANK2
frm.RedAngle = 45
frm.RedPower = 500
frm.BlueAngle = 45
frm.BluePower = 500
frm.AutoRedraw = False
frm.DrawWidth = 1
frm.Cls
' END: The form is ready for landscape

' START: Choose random sky color
r = Int(25 * Rnd) + 1
g = Int(25 * Rnd) + 1
b = Int(175 * Rnd) + 1
frm.BackColor = RGB(r, g, b)
' END: Sky color has been choosen

' START: This draws landscape
X2 = Int(100 * Rnd)
Y1 = Int(500 * Rnd) + 500
Y2 = Int(400 * Rnd) + 400
frm.Line (0, Y1)-(X2, Y2), RGB(0, 160, 0)
For land = 0 To 9999
If X2 > 1000 Then Exit For
X1 = X2
Y1 = Y2
X2 = X2 + Int(100 * Rnd)
Y2 = Y2 - Int(400 * Rnd) + 400
If Y2 > 900 Then Y2 = Y2 - Int(500 * Rnd) - 200
frm.Line (X1, Y1)-(X2, Y2), RGB(0, 160, 0)
'Line (land, 1000)-(X2, Y2) ' Makes a 3d land
Next land
frm.FillColor = RGB(0, 160, 0)
frm.FillStyle = 0
FloodFill frm.hdc, 1, 730, RGB(0, 160, 0)
FloodFill frm.hdc, 1000, 730, RGB(0, 160, 0)
' END: Landscape has been drawn

' START: Put boxes, scroll bars, and labels into postition
frm.RedPower.Move 20, 950
frm.Label2.Move 20, frm.RedPower.Top - frm.RedPower.Height
frm.RedAngle.Move 20, frm.Label2.Top - frm.Label2.Height - 5
frm.Label1.Move 20, frm.RedAngle.Top - frm.RedAngle.Height
frm.BluePower.Move 980 - frm.BluePower.Width, 950
frm.Label3.Move 980 - frm.Label3.Width, frm.BluePower.Top - frm.BluePower.Height
frm.BlueAngle.Move 980 - frm.BlueAngle.Width, frm.Label3.Top - frm.Label3.Height - 5
frm.Label4.Move 980 - frm.Label4.Width, frm.BlueAngle.Top - frm.BlueAngle.Height
frm.Fire.Move frm.ScaleWidth / 2, 950
' END: Everything is in place

' START: Show all boxes
frm.RedAngle.Visible = True
frm.BlueAngle.Visible = True
frm.RedPower.Visible = True
frm.BluePower.Visible = True
frm.Label1.Visible = True
frm.Label2.Visible = True
frm.Label3.Visible = True
frm.Label4.Visible = True
frm.Fire.Visible = True
' END: All boxes have been shown

' START: Make labels landscape color and refresh them
frm.Label1.BackColor = RGB(0, 160, 0)
frm.Label2.BackColor = RGB(0, 160, 0)
frm.Label3.BackColor = RGB(0, 160, 0)
frm.Label4.BackColor = RGB(0, 160, 0)
frm.Label1.Refresh
frm.Label2.Refresh
frm.Label3.Refresh
frm.Label4.Refresh
' END: The labels are ready

' START: Put tanks into position
frm.tank1.Move Int(500 * Rnd) + 1, 0
frm.tank2.Move Int(300 * Rnd) + 600, 0
frm.tank1.Visible = True
frm.tank2.Visible = True
' TANK1
Do
    If frm.Point(frm.tank1.Left, frm.tank1.Top + frm.tank1.Height) = 40960 And frm.Point(frm.tank1.Left + frm.tank1.Width, frm.tank1.Top + frm.tank1.Height) = 40960 Then Exit Do
    If frm.tank1.Top + frm.tank1.Height > frm.ScaleWidth Then
    frm.tank1.Top = frm.ScaleHeight - frm.tank1.Height
    Exit Do
    End If
    frm.tank1.Move frm.tank1.Left, frm.tank1.Top + 10
    DoEvents
Loop
' TANK2
Do
    If frm.Point(frm.tank2.Left, frm.tank2.Top + frm.tank2.Height) = 40960 And frm.Point(frm.tank2.Left + frm.tank2.Width, frm.tank2.Top + frm.tank2.Height) = 40960 Then Exit Do
    If frm.tank2.Top + frm.tank2.Height > frm.ScaleWidth Then
    frm.tank2.Top = frm.ScaleHeight - frm.tank2.Height
    Exit Do
    End If
    frm.tank2.Move frm.tank2.Left, frm.tank2.Top + 10
    DoEvents
Loop
' END: Tanks are in position

' START: Draw moon
frm.FillColor = QBColor(14) ' SET FILLCOLOR TO YELLOW
frm.FillStyle = 0 ' SET FILLSTYLE TO SOLID
frm.Circle (1000 * Rnd, 100), (30), QBColor(14)
' END: Moon has been drawn

' START: Draw stars
For stars = 0 To 300
    starx = Int(frm.ScaleWidth * Rnd)
    stary = Int(frm.ScaleHeight * Rnd)
    check = frm.Point(starx, stary)
    If check <> 40960 Then frm.PSet (starx, stary), QBColor(15)
Next stars
' END: Stars drawn

' Draw guns and make it Red's turn
frm.Tag = "blue"
draw_gun
frm.Tag = "red"
draw_gun

End Sub
