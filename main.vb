Sub 清除内容()
ActiveSheet.Cells.Delete
'ActiveSheet.Cells.Clear'有bug
'ActiveSheet.Rows.Clear
'Range("A1:V99").Clear
'ActiveSheet.UsedRange.ClearContents '和楼下的搭配使用有奇效
'ActiveSheet.UsedRange.ClearFormats
Rows.RowHeight = 14.5
Columns.ColumnWidth = 20
End Sub

