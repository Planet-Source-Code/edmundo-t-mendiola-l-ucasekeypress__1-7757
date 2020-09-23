<div align="center">

## L/UCaseKeyPress


</div>

### Description

Use these functions in KeyPress events to format user entries as entry is typed. UCaseKeyPress() converts key pressed to uppercase while LCaseKeyPress() converts key pressed to lowercase.
 
### More Info
 
KeyAscii As Integer

The ASCII value of key pressed; parameter from KeyPress event.

These two functions are easily re-usable in any projects where form controls with KeyPress events exists. You may also use this function in situations where the equivalent upper or lower case ASCII value of a character is to be evaluated.

Returns the new value of KeyASCII to be returned by the KeyPress event.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Edmundo T\. Mendiola](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/edmundo-t-mendiola.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/edmundo-t-mendiola-l-ucasekeypress__1-7757/archive/master.zip)





### Source Code

```
Public Function LCaseKeyPress(ByRef KeyAscii As Integer) As Integer
  ' Useful in the KeyPress event to convert entry to LCase()
  LCaseKeyPress = Asc(LCase(Chr(KeyAscii)))
End Function
Public Function UCaseKeyPress(ByRef KeyAscii As Integer) As Integer
  ' Useful in the KeyPress event to convert entry to UCase()
  UCaseKeyPress = Asc(UCase(Chr(KeyAscii)))
End Function
```

