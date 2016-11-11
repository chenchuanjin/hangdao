'自创建标注样式1
Public Function AddDimStyle1()
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
        .SetVariable "DimASz", 1.5        '控制尺寸线、引线箭头的大小。并控制钩线的大小
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
        .SetVariable "DimExe", 1        '指定尺寸界线超出尺寸线的距离
        .SetVariable "DimExO", 6       '指定尺寸界线偏移原点的距离
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
        .SetVariable "DimTAD", 1       '控制文字相对尺寸线的垂直位置
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
        .SetVariable "DimTSz", 0      '指定线性标注、半径标注以及直径标注中替代箭头的小斜线尺寸
        .SetVariable "DimTVP", 0        '控制尺寸线上方或下方标注文字的垂直位置
        .SetVariable "DimTxSty", "STANDARD"     '指定标注的文字样式
        .SetVariable "DimTxt", 1.8         '指定标注文字的高度，除非当前文字样式具有固定的高度
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
End Function