<div align="center">

## Sort\_TwoDimensionBubble


</div>

### Description

Sorts a 2-dimensional array
 
### More Info
 
TempArray        Variant

iElement        Integer

iDimension       Integer

bAscOrder        Boolean

Best used for smaller arrays, since the bubblesort algorithm is not suited to very large arrays

Boolean if the sort was successful


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Gordon Fuller](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gordon-fuller.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gordon-fuller-sort-twodimensionbubble__1-6658/archive/master.zip)





### Source Code

```
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Name:     Sort_TwoDimensionBubble
' VB Version:  6.00
' Called by:  Procedures     Events
'        ----------     ------
'
' Author:    Gordon McI. Fuller
' Copyright:  ©2000 Force 10 Automation
' Created:   Friday, March 17, 2000
' Modified:   [Friday, March 17, 2000]
' Purpose:
' Inputs:  Param    Name          Type    Meaning
'      -----    ----          ----    -------
'            TempArray        Variant
'      Optional  iElement        Integer
'      Optional  iDimension       Integer = 1
'      Optional  bAscOrder        Boolean = True
' Returns:   True/False for success of the sort
' Global Used:
' Module used:
'------------------------------------------------------------
' Notes:    This is a bubble sort
'        For large arrays it may not be the most efficient
'          option, but I haven't found anything in a
'          multi-dimension sort using another algorithm.
'
'  Sample array  array(0,0) = Apples
'          array(0,1) = 5
'          array(0,2) = Tree
'          array(1,0) = Grapes
'          ...
'      Apples     5    Tree
'      Grapes     2    Vine
'      Pears      3    Tree
'  The iDimension is 1 because it am sorting by the "rows" of the
'    first dimension rather than the "columns" of the 2nd
'  Since we would want to sort by the numeric value,
'    the iElement variable is 1
'  bAscOrder indicates whether the sort order is ascending or descending
'
'  If the array were structured as
'         array(0,0) = "Apples"
'         array(1,0) = 5
'         array(2,0) = Tree
'         ...
'      Apples     Grapes   Tree
'      5        2      3
'      Tree      Vine    Tree
'  iDimension will be 2 since we are sorting on the "columns"
'  iElement will still be 1 since we are sorting by that numeric value
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Function Sort_TwoDimensionBubble(TempArray As Variant, _
            Optional iElement As Integer = 1, _
            Optional iDimension As Integer = 1, _
            Optional bAscOrder As Boolean = True) As Boolean
  Dim arrTemp As Variant
  Dim i%, j%
  Dim NoExchanges As Integer
  On Error GoTo Error_BubbleSort
  ' Loop until no more "exchanges" are made.
  If iDimension% = 1 Then
    ReDim arrTemp(1, UBound(TempArray, 2))
  Else
    ReDim arrTemp(UBound(TempArray, 1), 1)
  End If
  Do
    NoExchanges = True
    ' Loop through each element in the array.
    If iDimension% = 1 Then
      For i% = LBound(TempArray, iDimension%) To UBound(TempArray, iDimension%) - 1
        ' If the element is greater than the element
        ' following it, exchange the two elements.
        If (bAscOrder And (TempArray(i%, iElement%) > TempArray(i% + 1, iElement%))) _
            Or (Not bAscOrder And (TempArray(i%, iElement%) < TempArray(i% + 1, iElement%))) _
          Then
            NoExchanges = False
            For j% = LBound(TempArray, 2) To UBound(TempArray, 2)
              arrTemp(1, j%) = TempArray(i%, j%)
            Next j%
            For j% = LBound(TempArray, 2) To UBound(TempArray, 2)
              TempArray(i%, j%) = TempArray(i% + 1, j%)
            Next j%
            For j% = LBound(TempArray, 2) To UBound(TempArray, 2)
              TempArray(i% + 1, j%) = arrTemp(1, j%)
            Next j%
        End If
      Next i%
    Else
      For i% = LBound(TempArray, iDimension%) To UBound(TempArray, iDimension%) - 1
        ' If the element is greater than the element
        ' following it, exchange the two elements.
        If (bAscOrder And (TempArray(iElement%, i%) > TempArray(iElement%, i% + 1))) _
            Or (Not bAscOrder And (TempArray(iElement%, i%) < TempArray(iElement%, i% + 1))) _
          Then
            NoExchanges = False
            For j% = LBound(TempArray, 1) To UBound(TempArray, 1)
              arrTemp(j%, 1) = TempArray(j%, i%)
            Next j%
            For j% = LBound(TempArray, 1) To UBound(TempArray, 1)
              TempArray(j%, i%) = TempArray(j%, i% + 1)
            Next j%
            For j% = LBound(TempArray, 1) To UBound(TempArray, 1)
              TempArray(j%, i% + 1) = arrTemp(j%, 1)
            Next j%
        End If
      Next i%
    End If
  Loop While Not (NoExchanges)
  Sort_TwoDimensionBubble = True
  On Error GoTo 0
  Exit Function
Error_BubbleSort:
  On Error GoTo 0
  Sort_TwoDimensionBubble = False
End Function
```

