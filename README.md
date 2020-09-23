<div align="center">

## Perfect Pause


</div>

### Description

Pauses an operation while allowing other operations to run. This pause is date and time based. The sleep function freezes your computer. The timer function and timer controls stop at midnight because they return to a 0 value. The perfect pause continues where these stop. It's highly configurable.
 
### More Info
 
'Seconds are inputed as seconds ( 1 = 1sec). A boolean is inputed for a quick 'exit, and an integer of 0 or 1 are inputted to determine the type of return.

'it is assumed the user knows a small bit about programming and function calls 'within modules. It is assumed the users computer is keeping good time

'Returns a boolean


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[rbennett](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rbennett.md)
**Level**          |Unknown
**User Rating**    |3.7 (22 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rbennett-perfect-pause__1-3552/archive/master.zip)





### Source Code

```
'=========================
'Paste in a BAS module
'=========================
Option Explicit
Public exitPause As Boolean
Public Function timedPause(secs As Long)
 Dim secStart As Variant
 Dim secNow As Variant
 Dim secDiff As Variant
 Dim Temp%
 exitPause = False 'this is our early way out out of the pause
 secStart = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'get the starting seconds
 Do While secDiff < secs
 If exitPause = True Then Exit Do
 secNow = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'this is the current time and date at any itteration of the loop
 secDiff = DateDiff("s", secStart, secNow) 'this compares the start time with the current time
 Temp% = DoEvents
 Loop
End Function
'=============================
'Paste in a form with 1 command button
'=============================
Option Explicit
Private Sub Command1_Click()
 timedPause 25
 MsgBox "Time is up buddy!"
End Sub
```

