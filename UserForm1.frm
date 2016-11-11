VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "巷道自动生成"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12630
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox1_Change()

End Sub

'生成巷道
Private Sub CommandButton1_Click()

a1 = Val(TextBox1.Text)
a2 = Val(TextBox2.Text)
a3 = Val(TextBox3.Text)

k = a1 / 2
g = a2
h = a3
'与中线偏移
z = k - 300 - 1250 - 150


If OptionButton1.Value = True Then

'外1
Dim waikuang1 As AcadLWPolyline
Dim points(0 To 9) As Double
points(0) = 0: points(1) = 0
points(2) = k: points(3) = 0
points(4) = k: points(5) = -300
points(6) = k + h: points(7) = -300
points(8) = k + h: points(9) = g

Set waikuang1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
'外2
Dim waikuang2 As AcadLWPolyline
Dim waikuang2points(0 To 3) As Double
waikuang2points(0) = k: waikuang2points(1) = 0
waikuang2points(2) = k: waikuang2points(3) = g

Set waikuang2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(waikuang2points)

'顶圆1，2
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

'镜像轴（外1外2）
Dim zhoupoint1(0 To 2) As Double
Dim zhoupoint2(0 To 2) As Double
zhoupoint1(0) = 0: zhoupoint1(1) = 0: zhoupoint1(2) = 0
zhoupoint2(0) = 0: zhoupoint2(1) = 10: zhoupoint2(2) = 0
' 镜像（外1外2）
Dim mirrorwaikuang1 As AcadLWPolyline
Set mirrorwaikuang = waikuang1.Mirror(zhoupoint1, zhoupoint2)
'mirrorwaikuang.color = acRed


Dim mirrorwaikuang2 As AcadLWPolyline
Set mirrorwaikuang2 = waikuang2.Mirror(zhoupoint1, zhoupoint2)
'mirrorwaikuang2.color = acRed

' 中部直线
Dim zhongline As AcadLWPolyline
Dim zlpoints(0 To 3) As Double

zlpoints(0) = k: zlpoints(1) = g

zlpoints(2) = -k: zlpoints(3) = g

Set zhongline = ThisDrawing.ModelSpace.AddLightWeightPolyline(zlpoints)


'天线
Dim tianxian As AcadLWPolyline
Dim txpoints(0 To 7) As Double
txpoints(0) = -z + 150 + 625 - 500 - 50: txpoints(1) = g + 300 - 50
txpoints(2) = -z + 150 + 625 - 500: txpoints(3) = g + 300
txpoints(4) = -z + 150 + 625 - 500 + 1000: txpoints(5) = g + 300
txpoints(6) = -z + 150 + 625 - 500 + 1000 + 50: txpoints(7) = 300 + g - 50


Set tianxian = ThisDrawing.ModelSpace.AddLightWeightPolyline(txpoints)


'矿车

Dim kuangche  As AcadLWPolyline
Dim kcpoints(0 To 7) As Double
kcpoints(0) = -(z - 150): kcpoints(1) = 300 + 1300 - 150
kcpoints(2) = -(z - 150): kcpoints(3) = 300 + 1300
kcpoints(4) = -(z - 150) + 1250: kcpoints(5) = 300 + 1300
kcpoints(6) = -(z - 150) + 1250: kcpoints(7) = 300 + 1300 - 150

Set kuangche = ThisDrawing.ModelSpace.AddLightWeightPolyline(kcpoints)


'轨道

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





'镜像轴(天线 矿车 轨道）
Dim zhoupoint3(0 To 2) As Double
Dim zhoupoint4(0 To 2) As Double
zhoupoint3(0) = -z: zhoupoint3(1) = 0: zhoupoint3(2) = 0
zhoupoint4(0) = -z: zhoupoint4(1) = 10: zhoupoint4(2) = 0

'镜像轴(铁轨）
Dim zhoupoint5(0 To 2) As Double
Dim zhoupoint6(0 To 2) As Double
zhoupoint5(0) = -z + 150 + 625: zhoupoint5(1) = 0: zhoupoint5(2) = 0
zhoupoint6(0) = -z + 150 + 625: zhoupoint6(1) = 10: zhoupoint6(2) = 0


' 镜像（铁轨）
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




'道砟

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



'标注






ElseIf OptionButton2.Value = True Then

'外1
Dim swaikuang1 As AcadLWPolyline
Dim spoints(0 To 9) As Double
spoints(0) = 0: spoints(1) = 0
spoints(2) = k: spoints(3) = 0
spoints(4) = k: spoints(5) = -300
spoints(6) = k + h: spoints(7) = -300
spoints(8) = k + h: spoints(9) = g

Set swaikuang1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(spoints)
'外2
Dim swaikuang2 As AcadLWPolyline
Dim swaikuang2points(0 To 3) As Double
swaikuang2points(0) = k: swaikuang2points(1) = 0
swaikuang2points(2) = k: swaikuang2points(3) = g

Set swaikuang2 = ThisDrawing.ModelSpace.AddLightWeightPolyline(swaikuang2points)

'顶圆1，2
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

'镜像轴（外1外2）
Dim szhoupoint1(0 To 2) As Double
Dim szhoupoint2(0 To 2) As Double
szhoupoint1(0) = 0: szhoupoint1(1) = 0: szhoupoint1(2) = 0
zhoupoint2(0) = 0: szhoupoint2(1) = 10: szhoupoint2(2) = 0
' 镜像（外1外2）
Dim smirrorwaikuang1 As AcadLWPolyline
Set smirrorwaikuang = swaikuang1.Mirror(szhoupoint1, szhoupoint2)
'mirrorwaikuang.color = acRed


Dim smirrorwaikuang2 As AcadLWPolyline
Set smirrorwaikuang2 = swaikuang2.Mirror(szhoupoint1, szhoupoint2)
'mirrorwaikuang2.color = acRed

' 中部直线
Dim szhongline As AcadLWPolyline
Dim szlpoints(0 To 3) As Double

szlpoints(0) = k: szlpoints(1) = g

szlpoints(2) = -k: szlpoints(3) = g

Set szhongline = ThisDrawing.ModelSpace.AddLightWeightPolyline(szlpoints)


'天线
Dim stianxian As AcadLWPolyline
Dim stxpoints(0 To 7) As Double
stxpoints(0) = -z + 150 + 625 - 500 - 50: stxpoints(1) = g + 300 - 50
stxpoints(2) = -z + 150 + 625 - 500: stxpoints(3) = g + 300
stxpoints(4) = -z + 150 + 625 - 500 + 1000: stxpoints(5) = g + 300
stxpoints(6) = -z + 150 + 625 - 500 + 1000 + 50: stxpoints(7) = 300 + g - 50


Set stianxian = ThisDrawing.ModelSpace.AddLightWeightPolyline(stxpoints)


'矿车

Dim skuangche  As AcadLWPolyline
Dim skcpoints(0 To 7) As Double
skcpoints(0) = -(z - 150): skcpoints(1) = 300 + 1300 - 150
skcpoints(2) = -(z - 150): skcpoints(3) = 300 + 1300
skcpoints(4) = -(z - 150) + 1250: skcpoints(5) = 300 + 1300
skcpoints(6) = -(z - 150) + 1250: skcpoints(7) = 300 + 1300 - 150

Set skuangche = ThisDrawing.ModelSpace.AddLightWeightPolyline(skcpoints)


'轨道

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





'镜像轴(天线 矿车 轨道）
Dim szhoupoint3(0 To 2) As Double
Dim szhoupoint4(0 To 2) As Double
szhoupoint3(0) = -z: szhoupoint3(1) = 0: szhoupoint3(2) = 0
szhoupoint4(0) = -z: szhoupoint4(1) = 10: szhoupoint4(2) = 0

'镜像轴(铁轨）
Dim szhoupoint5(0 To 2) As Double
Dim szhoupoint6(0 To 2) As Double
szhoupoint5(0) = -z + 150 + 625: szhoupoint5(1) = 0: szhoupoint5(2) = 0
szhoupoint6(0) = -z + 150 + 625: szhoupoint6(1) = 10: szhoupoint6(2) = 0


' 镜像（铁轨）
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




'道砟

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


'标注


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
ComboBox1.List = Array("单轨", "双轨")


'自定义标注格式

Dim dimStyle As AcadDimStyle
    Set dimStyle = ThisDrawing.DimStyles.Add("dimStyle1")
    ThisDrawing.ActiveDimStyle = dimStyle '激活该标注样式
   
   With ThisDrawing
       '第一组定义全局和线性比例因子
         .SetVariable "DimScale", 1     '设置全局比例因子
         .SetVariable "DimLFac", 1   '线性比例因子. '1'=1:1, '2'=2:1,'.5'=1:2等
        '定义典型的标注特性
        .SetVariable "DimADec", 0      '控制角度标注的显示精确位数
        .SetVariable "DimAssoc", 2     '控制标注对象的关联性
                                       '实际上该系统变量由图形控制
        .SetVariable "DimASz", 10        '控制尺寸线、引线箭头的大小。并控制钩线的大小
        .SetVariable "DimAtFit", 3    '当尺寸界线的空间不足以同时放下标注文字和箭头时,确定这两者的排列方式
                                        '0 将文字和箭头均放置于尺寸界线之外
                                        '1  先移动箭头，然后移动文字
                                        '2  先移动文字，然后移动箭头
                                        '3  移动文字和箭头中较合适的一个
        .SetVariable "DimAUnit", 0     '设置角度标注的单位格式
                                       '0 十进制度数
        .SetVariable "DimAZin", 0      '对角度标注作消零处理
                                       '0 显示所有前导零和后续零
        .SetVariable "DimBlk", ""      '设置尺寸线或引线末端显示的箭头块
                                       '"" 实心闭合
        .SetVariable "DimBlk1", ""     '当 DIMSAH 系统变量打开时，设置尺寸线第一个端点的箭头
        .SetVariable "DimBlk2", ""     '当 DIMSAH 系统变量打开时，设置尺寸线第二个端点的箭头
        .SetVariable "DimClrD", 256     '为尺寸线、箭头和标注引线指定颜色
        .SetVariable "DimClrE", 256    '为尺寸界线指定颜色。此颜色可以是任意有效的颜色编号
        .SetVariable "DimClrT", 256     '为标注文字指定颜色
         .SetVariable "DimDec", 0       '设置标注主单位显示的小数位位数
        .SetVariable "DimExe", 0        '指定尺寸界线超出尺寸线的距离
        .SetVariable "DimExO", 0       '指定尺寸界线偏移原点的距离
        .SetVariable "DimFrac", 0      '在 DIMLUNIT 系统变量设置为 4（建筑）或 5（分数）时设置分数格式
        .SetVariable "DimGap", 0.5     '当尺寸线分成段以在两段之间放置标注文字时，设置标注文字周围的距离
        .SetVariable "DimJust", 0      '控制标注文字的水平位置
                                        '0  将文字置于尺寸线之上，并在尺寸界线之间置中对正
                                        '1  紧邻第一条尺寸界线放置标注文字
                                        '2  紧邻第二条尺寸界线放置标注文字
                                        '3  将标注文字放在第一条尺寸界线以上，并与之对齐
                                        '4  将标注文字放在第二条尺寸界线以上，并与之对齐
        .SetVariable "DimLwd", acLnWtByLayer '指定尺寸线的线宽
        .SetVariable "DimLwe", acLnWtByLayer '指定尺寸界线的线宽
        .SetVariable "DimPost", ""     '指定标注测量值的文字前缀或后缀（或者两者都指定）
        .SetVariable "DimRnd", 0       '将所有标注距离舍入到指定值
        .SetVariable "DimSAh", 0       '控制尺寸线箭头块的显示
        .SetVariable "DimSD1", 0       '控制是否禁止显示第一条尺寸线
        .SetVariable "DimSD2", 0       '控制是否禁止显示第二条尺寸线
        .SetVariable "DimSE1", 0       '控制是否禁止显示第一条尺寸界线
        .SetVariable "DimSE2", 0       '控制是否禁止显示第二条尺寸界线
        .SetVariable "DimSOXD", 0      '控制是否允许尺寸线绘制到尺寸界线之外
        .SetVariable "DimTAD", 0       '控制文字相对尺寸线的垂直位置
                                       '0 标注文字在尺寸界线之间居中放置
                                        '1  除非尺寸线不是水平放置的或者尺寸界线内的文字被强制为水平放置
                                        '(DIMTIH = 1)，否则就将标注文字放置在尺寸线的上方。标注文字最底部
                                        '基线到尺寸线的距离值就是系统变量DIMGAP 的当前值。
        .SetVariable "DimTIH", 0       '控制所有标注类型（坐标标注除外）的标注文字在尺寸界线内的位置
                                        '0 或关 将文字与尺寸线对齐
                                        '1 或开 将文字水平放置
        .SetVariable "DimTIX", 1      '在尺寸界线之间绘制文字
                                        '0 或关 结果随标注类型的不同而不同。对于线性和角度标注，AutoCAD
                                        '将文字放置到尺寸界线之间（如果有足够的空间）。对于不适于放入圆
                                        '或圆弧中的半径标注和直径标注，DIMTIX 无效并总是强制将文字放到圆或圆弧之外
                                        '1 或开 将标注文字绘制在尺寸界线之间，即使 AutoCAD 通常将这些文字放置于尺寸界线之外。
        .SetVariable "DimTMOVE", 2      '设置标注文字的移动规则
                                        '0  尺寸线和标注文字一起移动
                                        '1  在移动标注文字时添加一条引线
                                        '2  允许标注文字自由移动而不用添加引线
        .SetVariable "DimTOFL", 0      '控制是否将尺寸线绘制在尺寸界线之间（即使文字放置在尺寸界线之外）
        .SetVariable "DimTOH", 0       '控制标注文字在尺寸界线外的位置
        .SetVariable "DimTSz", 20      '指定线性标注、半径标注以及直径标注中替代箭头的小斜线尺寸
        .SetVariable "DimTVP", 1        '控制尺寸线上方或下方标注文字的垂直位置
        .SetVariable "DimTxSty", "STANDARD"     '指定标注的文字样式
        .SetVariable "DimTxt", 50         '指定标注文字的高度，除非当前文字样式具有固定的高度
        .SetVariable "DimUPT", 0        '控制用户定位文字的选项
        .SetVariable "DimZIn", 0        '控制是否对主单位值作消零处理
'
       '定义换算单位的特性
        .SetVariable "DimAlt", 0        '控制标注中换算单位的显示
        .SetVariable "DimAltD", 4       '控制换算单位中小数位的位数
        .SetVariable "DimAltF", 25.4    '控制换算单位乘数
        .SetVariable "DimAltRnd", 0     '舍入换算标注单位
        .SetVariable "DimAltTD", 4      '设置标注换算单位公差值小数位的位数
        .SetVariable "DimAltTZ", 0      '控制是否对公差值作消零处理
        .SetVariable "DimAltU", 2       '为所有标注样式族（角度标注除外）换算单位设置单位格式
        .SetVariable "DimAltZ", 0       '控制是否对换算单位标注值作消零处理
        .SetVariable "DimAPost", ""     '为所有标注类型（角度标注除外）的换算标注测量值指定文字前缀或后缀（或两者都指定）
   End With
    '标注样式的特性从图形已有样式中获得
   dimStyle.CopyFrom ThisDrawing
End Sub



