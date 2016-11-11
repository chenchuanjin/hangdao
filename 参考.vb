 Private Sub CommandButton1_Click()
c = Val(TextBox1.Text)
d = Val(TextBox2.Text)
x = Val(TextBox3.Text)
dj1 = Val(TextBox4.Text)
TextBox5.Text = d - x - dj1

If OptionButton1.Value = True Then

If OptionButton3.Value = True Then
For i = 0 To Int(Val(TextBox7.Text) / 2) Step 1
i = i
fk1 = c * Val(TextBox6.Text)
fk2 = (d + x) * i
Dim 封口 As AcadLine
Dim 封口起点(2) As Double
Dim 封口端点(2) As Double
封口起点(0) = fk2
封口起点(1) = 0 - Val(TextBox10.Text)
封口端点(0) = fk2
封口端点(1) = fk1 + Val(TextBox8.Text)
Set 封口 = ThisDrawing.ModelSpace.AddLine(封口起点, 封口端点)
Next
Else
For i = 0 To Int(Val(TextBox7.Text) / 2) Step 1
i = i
fk1 = c * Val(TextBox6.Text)
fk2 = (d + d) * i
Dim 封口1 As AcadLine
Dim 封口起点1(2) As Double
Dim 封口端点1(2) As Double
封口起点1(0) = fk2
封口起点1(1) = 0 - Val(TextBox10.Text)
封口端点1(0) = fk2
封口端点1(1) = fk1 + Val(TextBox8.Text)
Set 封口1 = ThisDrawing.ModelSpace.AddLine(封口起点1, 封口端点1)
Next
End If

If OptionButton3.Value = True Then
For n = 0 To Int(Val(TextBox7.Text) / 2 + (Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2 - 1) Step 1
n = n
fk12 = c * Val(TextBox6.Text)
fk22 = d + (d + x) * n
Dim 封口2 As AcadLine
Dim 封口起点2(2) As Double
Dim 封口端点2(2) As Double
封口起点2(0) = fk22
封口起点2(1) = 0 - Val(TextBox10.Text)
封口端点2(0) = fk22
封口端点2(1) = fk12 + Val(TextBox8.Text)
Set 封口2 = ThisDrawing.ModelSpace.AddLine(封口起点2, 封口端点2)
Next
Else
For n = 0 To Int(Val(TextBox7.Text) / 2 + (Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2 - 1) Step 1
n = n
fk12 = c * Val(TextBox6.Text)
fk22 = d + (d + d) * n
Dim 封口21 As AcadLine
Dim 封口起点21(2) As Double
Dim 封口端点21(2) As Double
封口起点21(0) = fk22
封口起点21(1) = 0 - Val(TextBox10.Text)
封口端点21(0) = fk22
封口端点21(1) = fk12 + Val(TextBox8.Text)
Set 封口21 = ThisDrawing.ModelSpace.AddLine(封口起点21, 封口端点21)
Next
End If

If OptionButton3.Value = True Then
For m = 0 To (Val(TextBox6.Text)) Step 1
m = m
fk13 = (d + x) * Int(Val(TextBox7.Text) / 2)
fk23 = c * m
Dim 封口3 As AcadLine
Dim 封口起点3(2) As Double
Dim 封口端点3(2) As Double
封口起点3(0) = 0 - Val(TextBox9.Text)
封口起点3(1) = fk23
封口端点3(0) = fk13 + (Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2 * d + Val(TextBox11.Text)
封口端点3(1) = fk23
Set 封口3 = ThisDrawing.ModelSpace.AddLine(封口起点3, 封口端点3)
Next
Else
For m = 0 To (Val(TextBox6.Text)) Step 1
m = m
fk13 = (d + d) * Int(Val(TextBox7.Text) / 2)
fk23 = c * m
Dim 封口31 As AcadLine
Dim 封口起点31(2) As Double
Dim 封口端点31(2) As Double
封口起点31(0) = 0 - Val(TextBox9.Text)
封口起点31(1) = fk23
封口端点31(0) = fk13 + Val(TextBox11.Text) + d * ((Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2)
封口端点31(1) = fk23
Set 封口31 = ThisDrawing.ModelSpace.AddLine(封口起点31, 封口端点31)
Next
End If

Else
If OptionButton2.Value = True Then

If OptionButton3.Value = True Then
For i = 0 To Int(Val(TextBox7.Text) / 2) Step 1
i = i
fk1 = c * Val(TextBox6.Text)
fk2 = dj1 + (d + x) * i
Dim 封口12 As AcadLine
Dim 封口起点12(2) As Double
Dim 封口端点12(2) As Double
封口起点12(0) = fk2
封口起点12(1) = 0 - Val(TextBox10.Text)
封口端点12(0) = fk2
封口端点12(1) = fk1 + Val(TextBox8.Text)
Set 封口12 = ThisDrawing.ModelSpace.AddLine(封口起点12, 封口端点12)
Next
Else
For i = 0 To Int(Val(TextBox7.Text) - 1) Step 1
i = i
fk1 = c * Val(TextBox6.Text)
fk2 = dj1 + d * i
Dim 封口121 As AcadLine
Dim 封口起点121(2) As Double
Dim 封口端点121(2) As Double
封口起点121(0) = fk2
封口起点121(1) = 0 - Val(TextBox10.Text)
封口端点121(0) = fk2
封口端点121(1) = fk1 + Val(TextBox8.Text)
Set 封口121 = ThisDrawing.ModelSpace.AddLine(封口起点121, 封口端点121)
Next
End If

If OptionButton3.Value = True Then
For n = 0 To Int(Val(TextBox7.Text) / 2 + (Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2 - 1) Step 1
n = n
fk12 = c * Val(TextBox6.Text)
fk22 = dj1 + x + (d + x) * n
Dim 封口122 As AcadLine
Dim 封口起点122(2) As Double
Dim 封口端点122(2) As Double
封口起点122(0) = fk22
封口起点122(1) = 0 - Val(TextBox10.Text)
封口端点122(0) = fk22
封口端点122(1) = fk12 + Val(TextBox8.Text)
Set 封口122 = ThisDrawing.ModelSpace.AddLine(封口起点122, 封口端点122)
Next
Else
For n = 0 To Int(Val(TextBox7.Text) - 1) Step 1
n = n
fk12 = c * Val(TextBox6.Text)
fk22 = (x + dj1) + d * n
Dim 封口222 As AcadLine
Dim 封口起点222(2) As Double
Dim 封口端点222(2) As Double
封口起点222(0) = fk22
封口起点222(1) = 0 - Val(TextBox10.Text)
封口端点222(0) = fk22
封口端点222(1) = fk12 + Val(TextBox8.Text)
Set 封口222 = ThisDrawing.ModelSpace.AddLine(封口起点222, 封口端点222)
Next
End If

If OptionButton3.Value = True Then
For m = 0 To (Val(TextBox6.Text)) Step 1
m = m
fk13 = (d + x) * Int(Val(TextBox7.Text) / 2)
fk23 = c * m
Dim 封口311 As AcadLine
Dim 封口起点311(2) As Double
Dim 封口端点311(2) As Double
封口起点311(0) = 0 - Val(TextBox9.Text)
封口起点311(1) = fk23
封口端点311(0) = fk13 + (Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2 * d + Val(TextBox11.Text) + dj1 + dj2
封口端点311(1) = fk23
Set 封口311 = ThisDrawing.ModelSpace.AddLine(封口起点311, 封口端点311)
Next
Else
For m = 0 To (Val(TextBox6.Text)) Step 1
m = m
fk13 = d * (Int(Val(TextBox7.Text) - 1))
fk23 = c * m
Dim 封口312 As AcadLine
Dim 封口起点312(2) As Double
Dim 封口端点312(2) As Double
封口起点312(0) = dj1 - Val(TextBox9.Text) - dj1
封口起点312(1) = fk23
封口端点312(0) = fk13 + d + Val(TextBox11.Text)
封口端点312(1) = fk23
Set 封口312 = ThisDrawing.ModelSpace.AddLine(封口起点312, 封口端点312)
Next
End If

Else
MsgBox "请选择一个正面或是反面进行绘图！", vbOKOnly, "提示："
End If
End If

End Sub