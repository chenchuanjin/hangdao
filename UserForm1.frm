VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "����Զ�����"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12630
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox1_Change()

End Sub

'�������
Private Sub CommandButton1_Click()

a1 = Val(TextBox1.Text)
a2 = Val(TextBox2.Text)
a3 = Val(TextBox3.Text)

k = a1 / 2
g = a2
h = a3
'������ƫ��
z = k - 300 - 1250 - 150


If OptionButton1.Value = True Then

'��1
Dim waikuang1 As AcadLWPolyline
Dim points(0 To 9) As Double
points(0) = 0: points(1) = 0
points(2) = k: points(3) = 0
points(4) = k: points(5) = -300
points(6) = k + h: points(7) = -300
points(8) = k + h: points(9) = g

Set waikuang1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
'��2
Dim waikuang2 As AcadLWPolyline
Dim waikuang2points(0 To 3) As Double
waikuang2points(0) = k: waikuang2points(1) = 0
waikuang2points(2) = k: waikuang2points(3) = g

Set waikuang2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(waikuang2points)

'��Բ1��2
Dim yuan1 As AcadArc
Dim yuan2 As AcadArc

Dim yuan1xin(0 To 2) As Double
Dim yuan1banjing As Double
Dim yuan1qsjd  As Double
Dim yuan1zzjd As Double
yuan1xin(0) = 0: yuan1xin(1) = g: yuan1xin(2) = 0
yuan1banjing = k
yuan1qsjd = 0
yuan1zzjd = 3.15

Set yuan1 = ThisDrawing.ModelSpace.AddArc(yuan1xin, yuan1banjing, yuan1qsjd, yuan1zzjd)

Dim yuan2xin(0 To 2) As Double
Dim yuan2banjing As Double
Dim yuan2qsjd  As Double
Dim yuan2zzjd As Double
yuan2xin(0) = 0: yuan2xin(1) = g: yuan2xin(2) = 0
yuan2banjing = k + h
yuan2qsjd = 0
yuan2zzjd = 3.15

Set yuan2 = ThisDrawing.ModelSpace.AddArc(yuan2xin, yuan2banjing, yuan2qsjd, yuan2zzjd)

'�����ᣨ��1��2��
Dim zhoupoint1(0 To 2) As Double
Dim zhoupoint2(0 To 2) As Double
zhoupoint1(0) = 0: zhoupoint1(1) = 0: zhoupoint1(2) = 0
zhoupoint2(0) = 0: zhoupoint2(1) = 10: zhoupoint2(2) = 0
' ������1��2��
Dim mirrorwaikuang1 As AcadLWPolyline
Set mirrorwaikuang = waikuang1.Mirror(zhoupoint1, zhoupoint2)
'mirrorwaikuang.color = acRed


Dim mirrorwaikuang2 As AcadLWPolyline
Set mirrorwaikuang2 = waikuang2.Mirror(zhoupoint1, zhoupoint2)
'mirrorwaikuang2.color = acRed

' �в�ֱ��
Dim zhongline As AcadLWPolyline
Dim zlpoints(0 To 3) As Double

zlpoints(0) = k: zlpoints(1) = g

zlpoints(2) = -k: zlpoints(3) = g

Set zhongline = ThisDrawing.ModelSpace.AddLightWeightPolyline(zlpoints)


'����
Dim tianxian As AcadLWPolyline
Dim txpoints(0 To 7) As Double
txpoints(0) = -z + 150 + 625 - 500 - 50: txpoints(1) = g + 300 - 50
txpoints(2) = -z + 150 + 625 - 500: txpoints(3) = g + 300
txpoints(4) = -z + 150 + 625 - 500 + 1000: txpoints(5) = g + 300
txpoints(6) = -z + 150 + 625 - 500 + 1000 + 50: txpoints(7) = 300 + g - 50


Set tianxian = ThisDrawing.ModelSpace.AddLightWeightPolyline(txpoints)


'��

Dim kuangche  As AcadLWPolyline
Dim kcpoints(0 To 7) As Double
kcpoints(0) = -(z - 150): kcpoints(1) = 300 + 1300 - 150
kcpoints(2) = -(z - 150): kcpoints(3) = 300 + 1300
kcpoints(4) = -(z - 150) + 1250: kcpoints(5) = 300 + 1300
kcpoints(6) = -(z - 150) + 1250: kcpoints(7) = 300 + 1300 - 150

Set kuangche = ThisDrawing.ModelSpace.AddLightWeightPolyline(kcpoints)


'���

Dim guidao  As AcadLWPolyline
Dim gdpoints(0 To 11) As Double
gdpoints(0) = -z + 150 + 625: gdpoints(1) = 150 + 75
gdpoints(2) = -z + 150 + 625 + 625: gdpoints(3) = 150 + 75
gdpoints(4) = -z + 150 + 625 + 625: gdpoints(5) = 150 + 75 - 150
gdpoints(6) = -z + 150 + 625 + 625 - 1250: gdpoints(7) = 150 + 75 - 150
gdpoints(8) = -z + 150 + 625 + 625 - 1250: gdpoints(9) = 150 + 75
gdpoints(10) = -z + 150 + 625: gdpoints(11) = 150 + 75
Set guidao = ThisDrawing.ModelSpace.AddLightWeightPolyline(gdpoints)


Dim tiegui As AcadLWPolyline
Dim tgpoint(0 To 7) As Double
tgpoint(0) = -z + 150 + 625 + 300: tgpoint(1) = 150 + 75
tgpoint(2) = -z + 150 + 625 + 300: tgpoint(3) = 150 + 75 + 75
tgpoint(4) = -z + 150 + 625 + 300 - 50: tgpoint(5) = 150 + 75 + 75
tgpoint(6) = -z + 150 + 625 + 300 + 50: tgpoint(7) = 150 + 75 + 75
Set tiegui = ThisDrawing.ModelSpace.AddLightWeightPolyline(tgpoint)





'������(���� �� �����
Dim zhoupoint3(0 To 2) As Double
Dim zhoupoint4(0 To 2) As Double
zhoupoint3(0) = -z: zhoupoint3(1) = 0: zhoupoint3(2) = 0
zhoupoint4(0) = -z: zhoupoint4(1) = 10: zhoupoint4(2) = 0

'������(���죩
Dim zhoupoint5(0 To 2) As Double
Dim zhoupoint6(0 To 2) As Double
zhoupoint5(0) = -z + 150 + 625: zhoupoint5(1) = 0: zhoupoint5(2) = 0
zhoupoint6(0) = -z + 150 + 625: zhoupoint6(1) = 10: zhoupoint6(2) = 0


' �������죩
Dim mirrortianxian As AcadLWPolyline
Set mirrortianxian = tianxian.Mirror(zhoupoint3, zhoupoint4)


Dim mirrorkuangche As AcadLWPolyline
Set mirrorkuangche = kuangche.Mirror(zhoupoint3, zhoupoint4)

Dim mirrorguidao As AcadLWPolyline
Set mirrorguidao = guidao.Mirror(zhoupoint3, zhoupoint4)


Dim mirrortiegui As AcadLWPolyline
Set mirrortiegui = tiegui.Mirror(zhoupoint5, zhoupoint6)


Dim mirror2tiegui As AcadLWPolyline
Set mirror2tiegui = tiegui.Mirror(zhoupoint3, zhoupoint4)

Dim mirrormirrortiegui As AcadLWPolyline
Set mirrormirrortiegui = mirrortiegui.Mirror(zhoupoint3, zhoupoint4)

'mirrorwaikuang.color = acRed




'����

Dim dz1  As AcadLWPolyline
Dim dz1points(0 To 3) As Double
dz1points(0) = -z - 1250 - 300 - 150: dz1points(1) = 150
dz1points(2) = -z - 1250 - 150: dz1points(3) = 150
Set dz1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(dz1points)

Dim dz2  As AcadLWPolyline
Dim dz2points(0 To 3) As Double
dz2points(0) = -z - 150: dz2points(1) = 150
dz2points(2) = -z - 150 + 300: dz2points(3) = 150
Set dz2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(dz2points)

Dim dz3  As AcadLWPolyline
Dim dz3points(0 To 3) As Double
dz3points(0) = -z + 150 + 1250: dz3points(1) = 150
dz3points(2) = -z + 150 + 1250 + (2 * k - 600 - 2500): dz3points(3) = 150
Set dz3 = ThisDrawing.ModelSpace.AddLightWeightPolyline(dz3points)



'��ע






ElseIf OptionButton2.Value = True Then

'��1
Dim swaikuang1 As AcadLWPolyline
Dim spoints(0 To 9) As Double
spoints(0) = 0: spoints(1) = 0
spoints(2) = k: spoints(3) = 0
spoints(4) = k: spoints(5) = -300
spoints(6) = k + h: spoints(7) = -300
spoints(8) = k + h: spoints(9) = g

Set swaikuang1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(spoints)
'��2
Dim swaikuang2 As AcadLWPolyline
Dim swaikuang2points(0 To 3) As Double
swaikuang2points(0) = k: swaikuang2points(1) = 0
swaikuang2points(2) = k: swaikuang2points(3) = g

Set swaikuang2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(swaikuang2points)

'��Բ1��2
Dim syuan1 As AcadArc
Dim syuan2 As AcadArc

Dim syuan1xin(0 To 2) As Double
Dim syuan1banjing As Double
Dim syuan1qsjd  As Double
Dim syuan1zzjd As Double
syuan1xin(0) = 0: syuan1xin(1) = g: syuan1xin(2) = 0
syuan1banjing = k
syuan1qsjd = 0
syuan1zzjd = 3.15

Set syuan1 = ThisDrawing.ModelSpace.AddArc(syuan1xin, syuan1banjing, syuan1qsjd, syuan1zzjd)

Dim syuan2xin(0 To 2) As Double
Dim syuan2banjing As Double
Dim syuan2qsjd  As Double
Dim syuan2zzjd As Double
syuan2xin(0) = 0: syuan2xin(1) = g: syuan2xin(2) = 0
syuan2banjing = k + h
syuan2qsjd = 0
syuan2zzjd = 3.15

Set syuan2 = ThisDrawing.ModelSpace.AddArc(syuan2xin, syuan2banjing, syuan2qsjd, syuan2zzjd)

'�����ᣨ��1��2��
Dim szhoupoint1(0 To 2) As Double
Dim szhoupoint2(0 To 2) As Double
szhoupoint1(0) = 0: szhoupoint1(1) = 0: szhoupoint1(2) = 0
zhoupoint2(0) = 0: szhoupoint2(1) = 10: szhoupoint2(2) = 0
' ������1��2��
Dim smirrorwaikuang1 As AcadLWPolyline
Set smirrorwaikuang = swaikuang1.Mirror(szhoupoint1, szhoupoint2)
'mirrorwaikuang.color = acRed


Dim smirrorwaikuang2 As AcadLWPolyline
Set smirrorwaikuang2 = swaikuang2.Mirror(szhoupoint1, szhoupoint2)
'mirrorwaikuang2.color = acRed

' �в�ֱ��
Dim szhongline As AcadLWPolyline
Dim szlpoints(0 To 3) As Double

szlpoints(0) = k: szlpoints(1) = g

szlpoints(2) = -k: szlpoints(3) = g

Set szhongline = ThisDrawing.ModelSpace.AddLightWeightPolyline(szlpoints)


'����
Dim stianxian As AcadLWPolyline
Dim stxpoints(0 To 7) As Double
stxpoints(0) = -z + 150 + 625 - 500 - 50: stxpoints(1) = g + 300 - 50
stxpoints(2) = -z + 150 + 625 - 500: stxpoints(3) = g + 300
stxpoints(4) = -z + 150 + 625 - 500 + 1000: stxpoints(5) = g + 300
stxpoints(6) = -z + 150 + 625 - 500 + 1000 + 50: stxpoints(7) = 300 + g - 50


Set stianxian = ThisDrawing.ModelSpace.AddLightWeightPolyline(stxpoints)


'��

Dim skuangche  As AcadLWPolyline
Dim skcpoints(0 To 7) As Double
skcpoints(0) = -(z - 150): skcpoints(1) = 300 + 1300 - 150
skcpoints(2) = -(z - 150): skcpoints(3) = 300 + 1300
skcpoints(4) = -(z - 150) + 1250: skcpoints(5) = 300 + 1300
skcpoints(6) = -(z - 150) + 1250: skcpoints(7) = 300 + 1300 - 150

Set skuangche = ThisDrawing.ModelSpace.AddLightWeightPolyline(skcpoints)


'���

Dim sguidao  As AcadLWPolyline
Dim sgdpoints(0 To 11) As Double
sgdpoints(0) = -z + 150 + 625: sgdpoints(1) = 150 + 75
sgdpoints(2) = -z + 150 + 625 + 625: sgdpoints(3) = 150 + 75
sgdpoints(4) = -z + 150 + 625 + 625: sgdpoints(5) = 150 + 75 - 150
sgdpoints(6) = -z + 150 + 625 + 625 - 1250: sgdpoints(7) = 150 + 75 - 150
sgdpoints(8) = -z + 150 + 625 + 625 - 1250: sgdpoints(9) = 150 + 75
sgdpoints(10) = -z + 150 + 625: sgdpoints(11) = 150 + 75
Set sguidao = ThisDrawing.ModelSpace.AddLightWeightPolyline(sgdpoints)


Dim stiegui As AcadLWPolyline
Dim stgpoint(0 To 7) As Double
stgpoint(0) = -z + 150 + 625 + 300: stgpoint(1) = 150 + 75
stgpoint(2) = -z + 150 + 625 + 300: stgpoint(3) = 150 + 75 + 75
stgpoint(4) = -z + 150 + 625 + 300 - 50: stgpoint(5) = 150 + 75 + 75
stgpoint(6) = -z + 150 + 625 + 300 + 50: stgpoint(7) = 150 + 75 + 75
Set stiegui = ThisDrawing.ModelSpace.AddLightWeightPolyline(stgpoint)





'������(���� �� �����
Dim szhoupoint3(0 To 2) As Double
Dim szhoupoint4(0 To 2) As Double
szhoupoint3(0) = -z: szhoupoint3(1) = 0: szhoupoint3(2) = 0
szhoupoint4(0) = -z: szhoupoint4(1) = 10: szhoupoint4(2) = 0

'������(���죩
Dim szhoupoint5(0 To 2) As Double
Dim szhoupoint6(0 To 2) As Double
szhoupoint5(0) = -z + 150 + 625: szhoupoint5(1) = 0: szhoupoint5(2) = 0
szhoupoint6(0) = -z + 150 + 625: szhoupoint6(1) = 10: szhoupoint6(2) = 0


' �������죩
Dim smirrortianxian As AcadLWPolyline
Set smirrortianxian = stianxian.Mirror(szhoupoint3, szhoupoint4)


Dim smirrorkuangche As AcadLWPolyline
Set smirrorkuangche = skuangche.Mirror(szhoupoint3, szhoupoint4)

Dim smirrorguidao As AcadLWPolyline
Set smirrorguidao = sguidao.Mirror(szhoupoint3, szhoupoint4)


Dim smirrortiegui As AcadLWPolyline
Set smirrortiegui = stiegui.Mirror(szhoupoint5, szhoupoint6)


Dim smirror2tiegui As AcadLWPolyline
Set smirror2tiegui = stiegui.Mirror(szhoupoint3, szhoupoint4)

Dim smirrormirrortiegui As AcadLWPolyline
Set smirrormirrortiegui = smirrortiegui.Mirror(szhoupoint3, szhoupoint4)

'mirrorwaikuang.color = acRed




'����

Dim sdz1  As AcadLWPolyline
Dim sdz1points(0 To 3) As Double
sdz1points(0) = -z - 1250 - 300 - 150: sdz1points(1) = 150
sdz1points(2) = -z - 1250 - 150: sdz1points(3) = 150
Set sdz1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(sdz1points)

Dim sdz2  As AcadLWPolyline
Dim sdz2points(0 To 3) As Double
sdz2points(0) = -z - 150: sdz2points(1) = 150
sdz2points(2) = -z - 150 + 300: sdz2points(3) = 150
Set sdz2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(sdz2points)

Dim sdz3  As AcadLWPolyline
Dim sdz3points(0 To 3) As Double
sdz3points(0) = -z + 150 + 1250: sdz3points(1) = 150
sdz3points(2) = -z + 150 + 1250 + (2 * k - 600 - 2500): sdz3points(3) = 150
Set sdz3 = ThisDrawing.ModelSpace.AddLightWeightPolyline(sdz3points)


'��ע


Dim sbz As AcadDimAligned
Dim sbzpointq(0 To 2) As Double
Dim sbzPointz(0 To 2) As Double
Dim sbzwzpoint(0 To 2) As Double


sbzpointq(0) = 125
sbzpointq(1) = 300
sbzpointq(2) = 0
sbzPointz(0) = 725
sbzPointz(1) = 300
sbzPointz(2) = 0
sbzwzpoint(0) = 0
sbzwzpoint(1) = 400
sbzwzpoint(2) = 0
Set sbz = ThisDrawing.ModelSpace.AddDimAligned(sbzpointq, sbzPointz, sbzwzpoint)




Else: OptionButton3.Value = True


Dim bz2 As AcadDimAligned
Dim bz2pointq(0 To 2) As Double
Dim bz2Pointz(0 To 2) As Double
Dim bz2wzpoint(0 To 2) As Double


bz2pointq(0) = 125
bz2pointq(1) = 300
bz2pointq(2) = 0
bz2Pointz(0) = 725
bz2Pointz(1) = 10000
bz2Pointz(2) = 0
bz2wzpoint(0) = 0
bz2wzpoint(1) = 400
bz2wzpoint(2) = 0
Set bz2 = ThisDrawing.ModelSpace.AddDimAligned(bz2pointq, bz2Pointz, bz2wzpoint)

End If



End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Activate()
ComboBox1.List = Array("����", "˫��")


'�Զ����ע��ʽ

Dim dimStyle As AcadDimStyle
    Set dimStyle = ThisDrawing.DimStyles.Add("dimStyle1")
    ThisDrawing.ActiveDimStyle = dimStyle '����ñ�ע��ʽ
   
   With ThisDrawing
       '��һ�鶨��ȫ�ֺ����Ա�������
         .SetVariable "DimScale", 1     '����ȫ�ֱ�������
         .SetVariable "DimLFac", 1   '���Ա�������. '1'=1:1, '2'=2:1,'.5'=1:2��
        '������͵ı�ע����
        .SetVariable "DimADec", 0      '���ƽǶȱ�ע����ʾ��ȷλ��
        .SetVariable "DimAssoc", 2     '���Ʊ�ע����Ĺ�����
                                       'ʵ���ϸ�ϵͳ������ͼ�ο���
        .SetVariable "DimASz", 10        '���Ƴߴ��ߡ����߼�ͷ�Ĵ�С�������ƹ��ߵĴ�С
        .SetVariable "DimAtFit", 3    '���ߴ���ߵĿռ䲻����ͬʱ���±�ע���ֺͼ�ͷʱ,ȷ�������ߵ����з�ʽ
                                        '0 �����ֺͼ�ͷ�������ڳߴ����֮��
                                        '1  ���ƶ���ͷ��Ȼ���ƶ�����
                                        '2  ���ƶ����֣�Ȼ���ƶ���ͷ
                                        '3  �ƶ����ֺͼ�ͷ�нϺ��ʵ�һ��
        .SetVariable "DimAUnit", 0     '���ýǶȱ�ע�ĵ�λ��ʽ
                                       '0 ʮ���ƶ���
        .SetVariable "DimAZin", 0      '�ԽǶȱ�ע�����㴦��
                                       '0 ��ʾ����ǰ����ͺ�����
        .SetVariable "DimBlk", ""      '���óߴ��߻�����ĩ����ʾ�ļ�ͷ��
                                       '"" ʵ�ıպ�
        .SetVariable "DimBlk1", ""     '�� DIMSAH ϵͳ������ʱ�����óߴ��ߵ�һ���˵�ļ�ͷ
        .SetVariable "DimBlk2", ""     '�� DIMSAH ϵͳ������ʱ�����óߴ��ߵڶ����˵�ļ�ͷ
        .SetVariable "DimClrD", 256     'Ϊ�ߴ��ߡ���ͷ�ͱ�ע����ָ����ɫ
        .SetVariable "DimClrE", 256    'Ϊ�ߴ����ָ����ɫ������ɫ������������Ч����ɫ���
        .SetVariable "DimClrT", 256     'Ϊ��ע����ָ����ɫ
         .SetVariable "DimDec", 0       '���ñ�ע����λ��ʾ��С��λλ��
        .SetVariable "DimExe", 0        'ָ���ߴ���߳����ߴ��ߵľ���
        .SetVariable "DimExO", 0       'ָ���ߴ����ƫ��ԭ��ľ���
        .SetVariable "DimFrac", 0      '�� DIMLUNIT ϵͳ��������Ϊ 4���������� 5��������ʱ���÷�����ʽ
        .SetVariable "DimGap", 0.5     '���ߴ��߷ֳɶ���������֮����ñ�ע����ʱ�����ñ�ע������Χ�ľ���
        .SetVariable "DimJust", 0      '���Ʊ�ע���ֵ�ˮƽλ��
                                        '0  ���������ڳߴ���֮�ϣ����ڳߴ����֮�����ж���
                                        '1  ���ڵ�һ���ߴ���߷��ñ�ע����
                                        '2  ���ڵڶ����ߴ���߷��ñ�ע����
                                        '3  ����ע���ַ��ڵ�һ���ߴ�������ϣ�����֮����
                                        '4  ����ע���ַ��ڵڶ����ߴ�������ϣ�����֮����
        .SetVariable "DimLwd", acLnWtByLayer 'ָ���ߴ��ߵ��߿�
        .SetVariable "DimLwe", acLnWtByLayer 'ָ���ߴ���ߵ��߿�
        .SetVariable "DimPost", ""     'ָ����ע����ֵ������ǰ׺���׺���������߶�ָ����
        .SetVariable "DimRnd", 0       '�����б�ע�������뵽ָ��ֵ
        .SetVariable "DimSAh", 0       '���Ƴߴ��߼�ͷ�����ʾ
        .SetVariable "DimSD1", 0       '�����Ƿ��ֹ��ʾ��һ���ߴ���
        .SetVariable "DimSD2", 0       '�����Ƿ��ֹ��ʾ�ڶ����ߴ���
        .SetVariable "DimSE1", 0       '�����Ƿ��ֹ��ʾ��һ���ߴ����
        .SetVariable "DimSE2", 0       '�����Ƿ��ֹ��ʾ�ڶ����ߴ����
        .SetVariable "DimSOXD", 0      '�����Ƿ�����ߴ��߻��Ƶ��ߴ����֮��
        .SetVariable "DimTAD", 0       '����������Գߴ��ߵĴ�ֱλ��
                                       '0 ��ע�����ڳߴ����֮����з���
                                        '1  ���ǳߴ��߲���ˮƽ���õĻ��߳ߴ�����ڵ����ֱ�ǿ��Ϊˮƽ����
                                        '(DIMTIH = 1)������ͽ���ע���ַ����ڳߴ��ߵ��Ϸ�����ע������ײ�
                                        '���ߵ��ߴ��ߵľ���ֵ����ϵͳ����DIMGAP �ĵ�ǰֵ��
        .SetVariable "DimTIH", 0       '�������б�ע���ͣ������ע���⣩�ı�ע�����ڳߴ�����ڵ�λ��
                                        '0 ��� ��������ߴ��߶���
                                        '1 �� ������ˮƽ����
        .SetVariable "DimTIX", 1      '�ڳߴ����֮���������
                                        '0 ��� ������ע���͵Ĳ�ͬ����ͬ���������ԺͽǶȱ�ע��AutoCAD
                                        '�����ַ��õ��ߴ����֮�䣨������㹻�Ŀռ䣩�����ڲ����ڷ���Բ
                                        '��Բ���еİ뾶��ע��ֱ����ע��DIMTIX ��Ч������ǿ�ƽ����ַŵ�Բ��Բ��֮��
                                        '1 �� ����ע���ֻ����ڳߴ����֮�䣬��ʹ AutoCAD ͨ������Щ���ַ����ڳߴ����֮�⡣
        .SetVariable "DimTMOVE", 2      '���ñ�ע���ֵ��ƶ�����
                                        '0  �ߴ��ߺͱ�ע����һ���ƶ�
                                        '1  ���ƶ���ע����ʱ���һ������
                                        '2  �����ע���������ƶ��������������
        .SetVariable "DimTOFL", 0      '�����Ƿ񽫳ߴ��߻����ڳߴ����֮�䣨��ʹ���ַ����ڳߴ����֮�⣩
        .SetVariable "DimTOH", 0       '���Ʊ�ע�����ڳߴ�������λ��
        .SetVariable "DimTSz", 20      'ָ�����Ա�ע���뾶��ע�Լ�ֱ����ע�������ͷ��Сб�߳ߴ�
        .SetVariable "DimTVP", 1        '���Ƴߴ����Ϸ����·���ע���ֵĴ�ֱλ��
        .SetVariable "DimTxSty", "STANDARD"     'ָ����ע��������ʽ
        .SetVariable "DimTxt", 50         'ָ����ע���ֵĸ߶ȣ����ǵ�ǰ������ʽ���й̶��ĸ߶�
        .SetVariable "DimUPT", 0        '�����û���λ���ֵ�ѡ��
        .SetVariable "DimZIn", 0        '�����Ƿ������λֵ�����㴦��
'
       '���廻�㵥λ������
        .SetVariable "DimAlt", 0        '���Ʊ�ע�л��㵥λ����ʾ
        .SetVariable "DimAltD", 4       '���ƻ��㵥λ��С��λ��λ��
        .SetVariable "DimAltF", 25.4    '���ƻ��㵥λ����
        .SetVariable "DimAltRnd", 0     '���뻻���ע��λ
        .SetVariable "DimAltTD", 4      '���ñ�ע���㵥λ����ֵС��λ��λ��
        .SetVariable "DimAltTZ", 0      '�����Ƿ�Թ���ֵ�����㴦��
        .SetVariable "DimAltU", 2       'Ϊ���б�ע��ʽ�壨�Ƕȱ�ע���⣩���㵥λ���õ�λ��ʽ
        .SetVariable "DimAltZ", 0       '�����Ƿ�Ի��㵥λ��עֵ�����㴦��
        .SetVariable "DimAPost", ""     'Ϊ���б�ע���ͣ��Ƕȱ�ע���⣩�Ļ����ע����ֵָ������ǰ׺���׺�������߶�ָ����
   End With
    '��ע��ʽ�����Դ�ͼ��������ʽ�л��
   dimStyle.CopyFrom ThisDrawing
End Sub



