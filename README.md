<div align="center">

## 3D\-Stars in a Form


</div>

### Description

Draws a nice 3D-Starfield in a Form (uses X,Y,Z positions)

width a very short code

Shades each star depending on the distance.
 
### More Info
 
Create an Timer for the form called Timer1

None (maybe some flickering)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mikael Rickan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mikael-rickan.md)
**Level**          |Unknown
**User Rating**    |3.9 (59 globes from 15 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mikael-rickan-3d-stars-in-a-form__1-920/archive/master.zip)





### Source Code

```
'
Dim X(100), Y(100), Z(100) As Integer
Dim tmpX(100), tmpY(100), tmpZ(100) As Integer
Dim K As Integer
Dim Zoom As Integer
Dim Speed As Integer
Private Sub Form_Activate()
Speed = -1
K = 2038
Zoom = 256
Timer1.Interval = 1
For i = 0 To 100
  X(i) = Int(Rnd * 1024) - 512
  Y(i) = Int(Rnd * 1024) - 512
  Z(i) = Int(Rnd * 512) - 256
Next i
End Sub
Private Sub Timer1_Timer()
For i = 0 To 100
  Circle (tmpX(i), tmpY(i)), 5, BackColor
  Z(i) = Z(i) + Speed
  If Z(i) > 255 Then Z(i) = -255
  If Z(i) < -255 Then Z(i) = 255
  tmpZ(i) = Z(i) + Zoom
  tmpX(i) = (X(i) * K / tmpZ(i)) + (Form1.Width / 2)
  tmpY(i) = (Y(i) * K / tmpZ(i)) + (Form1.Height / 2)
  Radius = 1
  StarColor = 256 - Z(i)
  Circle (tmpX(i), tmpY(i)), 5, RGB(StarColor, StarColor, StarColor)
Next i
End Sub
```

