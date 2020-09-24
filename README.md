<div align="center">

## Create multi\-Colored forms without using the API


</div>

### Description

Although using the API is a nice way to create multi-colored forms, there might be a reason why you would wish to create one without using the API.
 
### More Info
 
The routine requires several parameters to be passed to it. They are:

FormName - Used to indicate which form is to be colored Orientation% - Top to bottom or right to left painting effect RStart% - (0-255) value for Red GStart% - (0-255) value for Green BStart% - (0-255) value for Blue RInc% - Amount to increment or decrement for Red GInc% - Amount to increment or decrement for Green BInc% - Amount to increment or decrement for Blue


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Tips and Source Code](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-tips-and-source-code.md)
**Level**          |Unknown
**User Rating**    |4.0 (4 globes from 1 user)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-tips-and-source-code-create-multi-colored-forms-without-using-the-api__1-157/archive/master.zip)





### Source Code

```
Sub PaintForm (FormName As Form, Orientation%, RStart%, GStart%, BStart%, RInc%, GInc%, BInc%)
'  This routine does NOT use API calls
  On Error Resume Next
  Dim x As Integer, y As Integer, z As Integer, Cycles As Integer
  Dim R%, G%, B%
  R% = RStart%: G% = GStart%: B% = BStart%
  ' Dividing the form into 100 equal parts
  If Orientation% = 0 Then
    Cycles = FormName.ScaleHeight \ 100
  Else
    Cycles = FormName.ScaleWidth \ 100
  End If
  For z = 1 To 100
    x = x + 1
    Select Case Orientation
      Case 0: 'Top to Bottom
        If x > FormName.ScaleHeight Then Exit For
        FormName.Line (0, x)-(FormName.Width, x + Cycles - 1), RGB(R%, G%, B%), BF
      Case 1: 'Left to Right
        If x > FormName.ScaleWidth Then Exit For
        FormName.Line (x, 0)-(x + Cycles - 1, FormName.Height), RGB(R%, G%, B%), BF
    End Select
    x = x + Cycles
    R% = R% + RInc%: G% = G% + GInc%: B% = B% + BInc%
    If R% > 255 Then R% = 255
    If R% < 0 Then R% = 0
    If G% > 255 Then G% = 255
    If G% < 0 Then G% = 0
    If B% > 255 Then B% = 255
    If B% < 0 Then B% = 0
  Next z
End Sub
To paint a form call the PaintForm procedure as follows:
PaintForm Me, 1, 100, 0, 255, 1, 0, -1
Experiment with the parameters and see what you can come up with. Keep the values for the incrementing low so as to create a smooth transition, whether they are negative or positive numbers.
```

