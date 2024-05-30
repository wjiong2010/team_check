"""
File Name: excel_format.py
Author: 小蜗牛慢慢爬
Date: 20190818
Description: excel单元格的格式
"""

from xlwt import Alignment
from xlwt import Borders
from xlwt import Font
from xlwt import Pattern
from xlwt import XFStyle

"""设置标题的格式类型"""
# 居中设置: 水平居中， 上下居中
alig = Alignment()
alig.horz = Alignment.HORZ_CENTER
alig.vert = Alignment.VERT_CENTER

# 边界设置: 边框颜色黑色， 宽度为1
bd = Borders()
bd.top = 1
bd.bottom = 1
bd.left = 1
bd.right = 1
bd.diag_colour = 0x0

# 单元格底色， 绿色
pt = Pattern()
pt.pattern = Pattern.SOLID_PATTERN
'''
pattern.pattern_fore_colour = 5 # May be: 8 through 63.
0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue,
5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green,
18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta,
21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
'''
# pt.pattern_back_colour = 0xff
pt.pattern_fore_colour = 3

# 字体，宋体加粗
fn = Font()
fn.bold = True
fn.name = u'宋体'
fn.height = 210

# title format
title_format = XFStyle()
title_format.pattern = pt
title_format.alignment = alig
title_format.borders = bd
title_format.font = fn

"""设置文本的格式类型"""
# 居中设置: 水平左对齐， 上下居中
alig = Alignment()
alig.horz = Alignment.HORZ_LEFT
alig.vert = Alignment.VERT_CENTER
alig.wrap = 1  # 自动换行

# 边界设置: 边框颜色黑色， 宽度为1
bd = Borders()
bd.top = 1
bd.bottom = 1
bd.left = 1
bd.right = 1
bd.diag_colour = 0x0  # 黑色

# 单元格底色，白色
pt = Pattern()
pt.pattern = Pattern.SOLID_PATTERN
# pt.pattern_back_colour = 0xff
pt.pattern_fore_colour = 1

# 字体，宋体不加粗，大小210
fn = Font()
fn.bold = False
fn.name = u'宋体'
fn.height = 210

# text_format
text_format = XFStyle()
text_format.pattern = pt
text_format.alignment = alig
text_format.borders = bd
text_format.font = fn
