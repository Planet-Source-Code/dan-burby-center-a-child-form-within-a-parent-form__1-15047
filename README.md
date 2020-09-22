<div align="center">

## Center a child form within a parent form


</div>

### Description

This function centers a child form within a parent form (not MDI), and can be used for custom messageboxes etc... Please try and leave a comment - Dan
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dan Burby](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dan-burby.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dan-burby-center-a-child-form-within-a-parent-form__1-15047/archive/master.zip)





### Source Code

```
Public Function CenterChild(Parent As Form, Child As Form)
On Local Error Resume Next
If Parent.WindowState = 1 Then
Exit Function
Else
Child.Left = (Parent.Left + (Parent.Width / 2)) - (Child.Width / 2)
Child.Top = (Parent.Top + (Parent.Height / 2)) - (Child.Height / 2)
End If
End Function
```

