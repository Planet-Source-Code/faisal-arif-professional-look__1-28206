<div align="center">

## Professional look


</div>

### Description

This code will give a great effect to any control making the user interface much more professional
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Faisal  Arif](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/faisal-arif.md)
**Level**          |Beginner
**User Rating**    |3.5 (21 globes from 6 users)
**Compatibility**  |VB 5\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/faisal-arif-professional-look__1-28206/archive/master.zip)





### Source Code

```
Option Explicit
Private Sub Command1_Click()
'Create a shadow to the right and below of Text1 (TextBox)
Shadow Me, Text1
End Sub
Private Sub Shadow(fIn As Form, ctrlIn As Control)
Const SHADOW_COLOR = &H40C0& 'Shadow Color
Const SHADOW_WIDTH = 3 'Shadow Border Width
Dim iOldWidth As Integer
Dim iOldScale As Integer
'Save the current DrawWidth and ScaleMode
iOldWidth = fIn.DrawWidth
iOldScale = fIn.ScaleMode
fIn.ScaleMode = 3
fIn.DrawWidth = 1
'Draws the shadow around the control by drawing a gray
'box behind the control that's offset right and down.
fIn.Line (ctrlIn.Left + SHADOW_WIDTH, ctrlIn.Top + _
      SHADOW_WIDTH)-Step(ctrlIn.Width - 1, _
      ctrlIn.Height - 1), SHADOW_COLOR, BF
'Restore Old Setting
fIn.DrawWidth = iOldWidth
fIn.ScaleMode = iOldScale
End Sub
```

