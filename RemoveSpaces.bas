Attribute VB_Name = "RemoveSpaces"
Sub NoSpaces()
 Dim c As Range
 For Each c In Selection.Cells
 c = Trim(c)
 Next
 End Sub
