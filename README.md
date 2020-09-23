<div align="center">

## AutoComplete Made EASY\!


</div>

### Description

This small script simply and easly creates an AutoComplete affect for your ComboBox. VERY EFFECTIVE AND VERY EASY TO UNDERSTAND.
 
### More Info
 
This code can be further inhanced by using

If Ucase(.text) = Ucase(......

I will leave the rest for you to discover and learn, but I think you will agree that this code is very promising, flawless and alot less complicated then all the other AutoCompleters!

The code returns the text that the use inputs into the combo box followed by the AutoCompleted text which is highlighted.

EG)

The user types "F"

The ComboBox then Displays a normal "F" which is then followed by "rog" which is highlighted to give "Frog" where "rog" is the .seltext. Understand?

NOTE: This is not hard to explain, sorry if you do not get what I am saying :)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Justin Stevens](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/justin-stevens.md)
**Level**          |Beginner
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/justin-stevens-autocomplete-made-easy__1-11833/archive/master.zip)





### Source Code

```
Private Sub ComboBox_KeyPress(KeyAscii As Integer)
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
On Error GoTo Oops
With ComboBox
Kounter = 0
For Kounter = 0 To .ListCount
If .Text = Left(.List(Kounter), Len(.Text)) Then
OldLength = Len(.Text)
.Text = .List(Kounter)
.SelStart = OldLength
.SelLength = Len(.Text) - OldLength
Timer1.Enabled = False
GoTo Oops
End If
Next Kounter
End With
Oops:
Timer1.Enabled = False
End Sub
```

