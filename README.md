<div align="center">

## HiByte,HiWord,LoByte,LoWord, MakeInt and MakeLong


</div>

### Description

Often especially when dealing with the API, byte and integer data types will be packed into LONG INTEGER (32 bit) values.

Thes snippets allow you to decode/encode these 32 bit longs:
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Duncan Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/duncan-jones.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/duncan-jones-hibyte-hiword-lobyte-loword-makeint-and-makelong__1-22740/archive/master.zip)





### Source Code

```
Public Function hiByte(ByVal w As Integer) As Byte
  If w And &H8000 Then
   hiByte = &H80 Or ((w And &H7FFF) \ &HFF)
  Else
   hiByte = w \ 256
  End If
End Function
Public Function HiWord(dw As Long) As Integer
 If dw And &H80000000 Then
   HiWord = (dw \ 65535) - 1
 Else
  HiWord = dw \ 65535
 End If
End Function
Public Function LoByte(w As Integer) As Byte
 LoByte = w And &HFF
End Function
Public Function LoWord(dw As Long) As Integer
 If dw And &H8000& Then
   LoWord = &H8000 Or (dw And &H7FFF&)
  Else
   LoWord = dw And &HFFFF&
  End If
End Function
Public Function MakeInt(ByVal LoByte As Byte, ByVal hiByte As Byte) As Integer
MakeInt = ((hiByte * &H100) + LoByte)
End Function
Public Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
MakeLong = ((HiWord * &H10000) + LoWord)
End Function
```

