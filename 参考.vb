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
Dim ��� As AcadLine
Dim ������(2) As Double
Dim ��ڶ˵�(2) As Double
������(0) = fk2
������(1) = 0 - Val(TextBox10.Text)
��ڶ˵�(0) = fk2
��ڶ˵�(1) = fk1 + Val(TextBox8.Text)
Set ��� = ThisDrawing.ModelSpace.AddLine(������, ��ڶ˵�)
Next
Else
For i = 0 To Int(Val(TextBox7.Text) / 2) Step 1
i = i
fk1 = c * Val(TextBox6.Text)
fk2 = (d + d) * i
Dim ���1 As AcadLine
Dim ������1(2) As Double
Dim ��ڶ˵�1(2) As Double
������1(0) = fk2
������1(1) = 0 - Val(TextBox10.Text)
��ڶ˵�1(0) = fk2
��ڶ˵�1(1) = fk1 + Val(TextBox8.Text)
Set ���1 = ThisDrawing.ModelSpace.AddLine(������1, ��ڶ˵�1)
Next
End If

If OptionButton3.Value = True Then
For n = 0 To Int(Val(TextBox7.Text) / 2 + (Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2 - 1) Step 1
n = n
fk12 = c * Val(TextBox6.Text)
fk22 = d + (d + x) * n
Dim ���2 As AcadLine
Dim ������2(2) As Double
Dim ��ڶ˵�2(2) As Double
������2(0) = fk22
������2(1) = 0 - Val(TextBox10.Text)
��ڶ˵�2(0) = fk22
��ڶ˵�2(1) = fk12 + Val(TextBox8.Text)
Set ���2 = ThisDrawing.ModelSpace.AddLine(������2, ��ڶ˵�2)
Next
Else
For n = 0 To Int(Val(TextBox7.Text) / 2 + (Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2 - 1) Step 1
n = n
fk12 = c * Val(TextBox6.Text)
fk22 = d + (d + d) * n
Dim ���21 As AcadLine
Dim ������21(2) As Double
Dim ��ڶ˵�21(2) As Double
������21(0) = fk22
������21(1) = 0 - Val(TextBox10.Text)
��ڶ˵�21(0) = fk22
��ڶ˵�21(1) = fk12 + Val(TextBox8.Text)
Set ���21 = ThisDrawing.ModelSpace.AddLine(������21, ��ڶ˵�21)
Next
End If

If OptionButton3.Value = True Then
For m = 0 To (Val(TextBox6.Text)) Step 1
m = m
fk13 = (d + x) * Int(Val(TextBox7.Text) / 2)
fk23 = c * m
Dim ���3 As AcadLine
Dim ������3(2) As Double
Dim ��ڶ˵�3(2) As Double
������3(0) = 0 - Val(TextBox9.Text)
������3(1) = fk23
��ڶ˵�3(0) = fk13 + (Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2 * d + Val(TextBox11.Text)
��ڶ˵�3(1) = fk23
Set ���3 = ThisDrawing.ModelSpace.AddLine(������3, ��ڶ˵�3)
Next
Else
For m = 0 To (Val(TextBox6.Text)) Step 1
m = m
fk13 = (d + d) * Int(Val(TextBox7.Text) / 2)
fk23 = c * m
Dim ���31 As AcadLine
Dim ������31(2) As Double
Dim ��ڶ˵�31(2) As Double
������31(0) = 0 - Val(TextBox9.Text)
������31(1) = fk23
��ڶ˵�31(0) = fk13 + Val(TextBox11.Text) + d * ((Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2)
��ڶ˵�31(1) = fk23
Set ���31 = ThisDrawing.ModelSpace.AddLine(������31, ��ڶ˵�31)
Next
End If

Else
If OptionButton2.Value = True Then

If OptionButton3.Value = True Then
For i = 0 To Int(Val(TextBox7.Text) / 2) Step 1
i = i
fk1 = c * Val(TextBox6.Text)
fk2 = dj1 + (d + x) * i
Dim ���12 As AcadLine
Dim ������12(2) As Double
Dim ��ڶ˵�12(2) As Double
������12(0) = fk2
������12(1) = 0 - Val(TextBox10.Text)
��ڶ˵�12(0) = fk2
��ڶ˵�12(1) = fk1 + Val(TextBox8.Text)
Set ���12 = ThisDrawing.ModelSpace.AddLine(������12, ��ڶ˵�12)
Next
Else
For i = 0 To Int(Val(TextBox7.Text) - 1) Step 1
i = i
fk1 = c * Val(TextBox6.Text)
fk2 = dj1 + d * i
Dim ���121 As AcadLine
Dim ������121(2) As Double
Dim ��ڶ˵�121(2) As Double
������121(0) = fk2
������121(1) = 0 - Val(TextBox10.Text)
��ڶ˵�121(0) = fk2
��ڶ˵�121(1) = fk1 + Val(TextBox8.Text)
Set ���121 = ThisDrawing.ModelSpace.AddLine(������121, ��ڶ˵�121)
Next
End If

If OptionButton3.Value = True Then
For n = 0 To Int(Val(TextBox7.Text) / 2 + (Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2 - 1) Step 1
n = n
fk12 = c * Val(TextBox6.Text)
fk22 = dj1 + x + (d + x) * n
Dim ���122 As AcadLine
Dim ������122(2) As Double
Dim ��ڶ˵�122(2) As Double
������122(0) = fk22
������122(1) = 0 - Val(TextBox10.Text)
��ڶ˵�122(0) = fk22
��ڶ˵�122(1) = fk12 + Val(TextBox8.Text)
Set ���122 = ThisDrawing.ModelSpace.AddLine(������122, ��ڶ˵�122)
Next
Else
For n = 0 To Int(Val(TextBox7.Text) - 1) Step 1
n = n
fk12 = c * Val(TextBox6.Text)
fk22 = (x + dj1) + d * n
Dim ���222 As AcadLine
Dim ������222(2) As Double
Dim ��ڶ˵�222(2) As Double
������222(0) = fk22
������222(1) = 0 - Val(TextBox10.Text)
��ڶ˵�222(0) = fk22
��ڶ˵�222(1) = fk12 + Val(TextBox8.Text)
Set ���222 = ThisDrawing.ModelSpace.AddLine(������222, ��ڶ˵�222)
Next
End If

If OptionButton3.Value = True Then
For m = 0 To (Val(TextBox6.Text)) Step 1
m = m
fk13 = (d + x) * Int(Val(TextBox7.Text) / 2)
fk23 = c * m
Dim ���311 As AcadLine
Dim ������311(2) As Double
Dim ��ڶ˵�311(2) As Double
������311(0) = 0 - Val(TextBox9.Text)
������311(1) = fk23
��ڶ˵�311(0) = fk13 + (Val(TextBox7.Text) / 2 - Int(Val(TextBox7.Text) / 2)) * 2 * d + Val(TextBox11.Text) + dj1 + dj2
��ڶ˵�311(1) = fk23
Set ���311 = ThisDrawing.ModelSpace.AddLine(������311, ��ڶ˵�311)
Next
Else
For m = 0 To (Val(TextBox6.Text)) Step 1
m = m
fk13 = d * (Int(Val(TextBox7.Text) - 1))
fk23 = c * m
Dim ���312 As AcadLine
Dim ������312(2) As Double
Dim ��ڶ˵�312(2) As Double
������312(0) = dj1 - Val(TextBox9.Text) - dj1
������312(1) = fk23
��ڶ˵�312(0) = fk13 + d + Val(TextBox11.Text)
��ڶ˵�312(1) = fk23
Set ���312 = ThisDrawing.ModelSpace.AddLine(������312, ��ڶ˵�312)
Next
End If

Else
MsgBox "��ѡ��һ��������Ƿ�����л�ͼ��", vbOKOnly, "��ʾ��"
End If
End If

End Sub