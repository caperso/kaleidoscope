Public Sub elfin2()
Dim carrot As String
Dim str0, str1 As String
Dim str, h, i, j, k, dn As String
Dim a, b, c, d, e, f, g, bmax, elserow As Long

b = 1 '列
a = 2 '行
c = 2
e = 1 '行数

'自动删除空白列（为了更懒地配合EDS使用.....↓
bmax = ActiveWorkbook.ActiveSheet.UsedRange.Columns.Count
Do While b <= bmax
If Len(Cells(a, b)) = 0 Then
Columns(b).Delete
bmax = bmax - 1
Else
b = b + 1
End If
Loop
'..........................................↑


'出现PN的情况..............................↓
b = 1 '列
Do While b < bmax
If InStr(UCase(Cells(1, b)), "PN") <> 0 Then
dn = InputBox("请输入图号（比如HAA304GA）：", "(*?▽?*)", "") 'get drawing number
dn = UCase(dn)
End If
b = b + 1
Loop
'..........................................↑

a = 2 '行
b = 2 '列
g = 2 'ENDIF行数
Do While Len(Cells(a, b)) <> 0
If UCase(Cells(a, b - 1)) <> "ELSE" Then
'If InStr(UCase(Cells(a, b - 1)), "ELSE") = 0 Then

Do While Len(Cells(a - 1, b)) <> 0
str0 = Cells(1, b - 1)
str = Cells(a, b - 1)

If UCase(str) = "ALL" Then
Else
carrot = carrot + "IF "

If InStr(str, "(") = 0 And InStr(str, ")") = 0 And InStr(str, "[") = 0 And InStr(str, "]") = 0 Then
h = 0
Do While h <= Len(str) - Len(Replace(str, ",", ""))
If UCase(str0) = "PN" Then
If InStr(str, "<>") <> 0 Then
'carrot = carrot + "REL_MODEL_NAME<>""" + CStr(dn) + Replace(UCase(Split(str, ",")(h)), "<>", "") + """&&"
carrot = carrot + "string_starts(REL_MODEL_NAME,""" + CStr(dn) + Replace(UCase(Split(str, ",")(h)), "<>", "") + """)==NO&&"
Else
'carrot = carrot + "REL_MODEL_NAME==""" + CStr(dn) + UCase(Split(str, ",")(h)) + """||"
carrot = carrot + "string_starts(REL_MODEL_NAME,""" + CStr(dn) + UCase(Split(str, ",")(h)) + """)==YES||"
End If
Else
If IsNumeric(Replace(Replace(Replace(str, ",", ""), "<>", ""), ".", "")) = True Then '
If InStr(str, "<>") <> 0 Then
carrot = carrot + str0 + "<>" + Replace(Split(str, ",")(h), "<>", "") + "&&"
Else
carrot = carrot + str0 + "==" + Split(str, ",")(h) + "||"
End If
Else
If InStr(str, "<>") <> 0 Then
carrot = carrot + str0 + "<>""" + Replace(Split(str, ",")(h), "<>", "") + """&&"
Else
carrot = carrot + str0 + "==""" + Split(str, ",")(h) + """||"
End If
End If
End If
h = h + 1
Loop
carrot = Left(carrot, Len(carrot) - 2)
Else
If InStr(str, "$") > 0 Then
i = Split(str, ",")(0)
i = Right(i, Len(i) - 1)
k = str0 + Left(str, 1) + i
k = Replace(k, "(", ">")
k = Replace(k, ")", "<")
k = Replace(k, "[", ">=")
k = Replace(k, "]", "<=")
carrot = carrot + k
Else
i = Split(str, ",")(0)
i = Right(i, Len(i) - 1)
j = Split(str, ",")(1)
j = Left(j, Len(j) - 1)
k = str0 + Left(str, 1) + i + "&&" + str0 + Right(str, 1) + j
k = Replace(k, "(", ">")
k = Replace(k, ")", "<")
k = Replace(k, "[", ">=")
k = Replace(k, "]", "<=")
carrot = carrot + k
End If
End If
End If
d = b
b = b + 1
If UCase(str) <> "ALL" Then
carrot = carrot + Chr(10)
g = g + 1
e = e + 1
End If
Loop

str1 = CStr(Cells(a, b - 1))
str1 = Replace(str1, "INT", "FLOOR")
str1 = Replace(str1, "ROUNDUP", "CEIL")
str1 = Replace(str1, "ROUNDDOWN", "FLOOR")

If InStr(Cells(1, b - 1), "PN") > 0 Then
carrot = carrot + Cells(1, b - 1) + "=""" + str1 + """" + Chr(10)
Else
carrot = carrot + Cells(1, b - 1) + "=" + str1 + Chr(10)    '12123132314rdegvtrtrbrnb
End If
e = e + 1

Else
elserow = a
End If
Do While c < g
carrot = carrot + "ENDIF" + Chr(10)
e = e + 1
c = c + 1
Loop
c = 2
carrot = carrot
'carrot = carrot + Chr(10)
e = e + 1
a = a + 1
If b > bmax Then
bmax = b
End If
b = 2
g = 2
Loop



If elserow > 0 Then
'carrot = CStr(Cells(1, bmax - 1)) + "=" + CStr(Cells(elserow, bmax - 1)) + Chr(10) + carrot
str1 = CStr(Cells(elserow, bmax - 1))
str1 = Replace(str1, "INT", "FLOOR")
str1 = Replace(str1, "ROUNDUP", "CEIL")
str1 = Replace(str1, "ROUNDDOWN", "FLOOR")

If InStr(Cells(1, bmax - 1), "PN") > 0 Then
carrot = CStr(Cells(1, bmax - 1)) + "=""" + str1 + """" + Chr(10) + carrot
Else
carrot = CStr(Cells(1, bmax - 1)) + "=" + str1 + Chr(10) + carrot
End If
e = e + 1
End If
carrot = "/*----------------" + CStr(Cells(1, bmax - 1)) + "↓----------------*/" + Chr(10) + carrot + "/*----------------" + CStr(Cells(1, bmax - 1)) + "↑----------------*/" + Chr(10)

Range(Cells(a + 1, 1), Cells(a + 1, d)).Merge
Range(Cells(a + 2, 1), Cells(a + 2, d)).Merge
Cells(a + 1, 1) = "关系："
Cells(a + 2, 1) = carrot
e = e * 14
If e > 409 Then
e = 409
End If

Rows(a + 2).RowHeight = e

Rows(a + 1).RowHeight = 25

'Cells(a + 2, 1).CenterVertically = True
Cells(a + 1, 1).Font.Size = 18
'Cells(a + 2, 1).Interior.Color = RGB(229, 245, 255)
Cells(a + 2, 1).Interior.color = RGB(229, 245, 255)


'On Error GoTo ErrHandle
'ErrHandle:
   'If Err.Number Then MsgBox Err.Number
 
'Cells(a + 2, 1).Copy '这个也行，但是下面的更优秀
With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
.SetText carrot
.PutInClipboard
End With

End Sub

