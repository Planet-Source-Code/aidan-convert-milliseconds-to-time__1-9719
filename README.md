<div align="center">

## Convert Milliseconds to Time


</div>

### Description

Converts a number of milliseconds to a time of form HH:MM:SS:hh
 
### More Info
 
No. of Milliseconds

Time


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aidan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aidan.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aidan-convert-milliseconds-to-time__1-9719/archive/master.zip)





### Source Code

```
' Enumerations:
Private Enum BeforeOrAfter
  Before
  After
End Enum
' ********** Procedure: Convert Milliseconds To Time **********
Public Function ConvertMillisecondsToTime(Milliseconds As Long, Optional IncludeHours As Boolean) As String
  ' Converts a number of Milliseconds to a time (HH:MM:SS:HH)
  Dim CurrentHSecs As Double, HSecs As Long, Mins As Long, Secs As Long, Hours As Double
  CurrentHSecs = Int((Milliseconds / 10) + 0.5)
  If IncludeHours Then
    Hours = Int(CurrentHSecs / 360000)
    CurrentHSecs = CurrentHSecs - (Hours * 360000)
  End If
  Mins = Int(CurrentHSecs / 6000)
  CurrentHSecs = CurrentHSecs - (Mins * 6000)
  Secs = Int((CurrentHSecs) / 100)
  CurrentHSecs = CurrentHSecs - (Secs * 100)
  HSecs = CurrentHSecs
  ConvertMillisecondsToTime = FixLength(Mins, 2) & ":" & FixLength(Secs, 2) & ":" & FixLength(HSecs, 2)
  If IncludeHours Then
    ConvertMillisecondsToTime = FixLength(Hours, 2) & ":" & ConvertMillisecondsToTime
  End If
End Function
' ********** Additional Subs/Functions Required **********
Private Function FixLength(Number As Variant, Length As Integer, Optional CharacterPosition As BeforeOrAfter = Before, Optional Character As String = "0") As String
  ' Inserts "0"'s before a number to make it a certain length
  Dim i As Integer, StrNum As String
  StrNum = CStr(Number)
  FixLength = StrNum
  For i = Len(StrNum) To Length - 1
    If CharacterPosition = Before Then
      FixLength = Character & FixLength
    Else
      FixLength = FixLength & Character
    End If
  Next i
End Function
```

