<div align="center">

## Sorting numeric/dates a listview


</div>

### Description

This code is an adapt form MSDN.

It will sort numeric - or on dates - any column with the apropriate tag (be aware that the column values must be all numeric or dates - according with the tag of the column )

It uses only one module - and on the form with the listview 1 line of code !

IF
 
### More Info
 
See How To on MSDN

Crash the application (it uses subclassing !) if the column values does not agree with with the tag of the column (if the tag is not date or numeric - then it uses standard sort order )


<span>             |<span>
---                |---
**Submitted On**   |2002-03-22 16:34:50
**By**             |[Ignat Andrei](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ignat-andrei.md)
**Level**          |Intermediate
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Sorting\_nu644613222002\.zip](https://github.com/Planet-Source-Code/ignat-andrei-sorting-numeric-dates-a-listview__1-32950/archive/master.zip)

### API Declarations

```
' the only line used on the form to sort
Private Sub lvwTest_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  SetListViewOrder lvwTest, ColumnHeader
End Sub
```





