# -*- coding: utf-8 -*- #  Copyright 2012 MIZUNO Hiroki # 
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.
import re
import base64
import win32com.client
import os
import os.path
from blockdiag.utils import XY,Box,Size
from types import StringType
from webcolors import name_to_rgb, hex_to_rgb
from blockdiag.imagedraw import base, textfolder
class ExcelShapeImageDraw(base.ImageDraw):
    MSO_SHAPE_RECTANGE = 1
    MSO_SHAPE_OVAL     = 9
    MSO_SHAPE_ARC      = 25
    MSO_LINE_DASH      = 4
    XL_CENTER = -4108
    XL_RIGHT  = -4152

    def __init__(self, filename, **kwargs):
        self.excel         = win32com.client.Dispatch("Excel.Application") 
        self.excel.Visible = True
        self.DisplayAlerts = False
        self.book = self.excel.Workbooks.Add()

    def set_canvas_size(self, size):
        pass

    def path(self, pd, **kwargs):
        pass

    def rgb(self, rgb):
        if type(rgb) is StringType and rgb[0] == '#':
            rgb = hex_to_rgb(rgb)
        elif type(rgb) is StringType:
            rgb = name_to_rgb(rgb)
        r,g,b = rgb
        return (b << 16) + (g << 8) + r

    def width(self, box):
        return box[2] - box[0]

    def height(self, box):
        return box[3] - box[1]

    def set_style(self, shape, kwargs):
        fill    = kwargs.get('fill')
        outline = kwargs.get('outline')
        tick    = kwargs.get('tick')

        if tick != None:
            shape.Line.Weight = 1

        if outline != None:
            shape.Line.ForeColor.RGB = self.rgb(outline)
        else:
            shape.Line.Visible = False

        if fill != None:
            shape.Fill.ForeColor.RGB = self.rgb(fill)
        else:
            shape.Fill.Visible = False

    def rectangle(self, box, **kwargs):
        left   = box[0]
        top    = box[1]
        width  = self.width(box)
        height = self.height(box)

        shape = self.book.ActiveSheet.Shapes.AddShape(self.MSO_SHAPE_RECTANGE, left, top, width, height);
        self.set_style(shape, kwargs)

    def text(self, xy, string, font, **kwargs):
        pass

    def textarea(self, box, string, font, **kwargs):
        left   = box[0]
        top    = box[1]
        width  = self.width(box)
        height = self.height(box)

        fill = kwargs.get('fill', 'none')
        shape = self.book.ActiveSheet.Shapes.AddShape(self.MSO_SHAPE_RECTANGE, left, top, width, height)
        shape.Line.Visible = 0

        shape.Fill.ForeColor.RGB = 0xFFFFFF
        shape.Fill.Visible = False
        chars = shape.TextFrame.Characters()
        chars.Text = string
        chars.Font.Color = self.rgb(fill)
        chars.Font.Size = font.size
        if kwargs.get('halign') == 'center':
            shape.TextFrame.HorizontalAlignment = self.XL_CENTER
        elif kwargs.get('halign') == 'right':
            shape.TextFrame.HorizontalAlignment = self.XL_RIGHT
        shape.TextFrame.VerticalAlignment = self.XL_CENTER

    def line(self, xy, **kwargs):
        fill    = kwargs.get('fill')
        style   = kwargs.get('style')
        x1,y1 = xy[0]
        x2,y2 = xy[1]

        line = self.book.ActiveSheet.Shapes.AddLine(x1, y1, x2, y2)

        if fill != None:
            line.Line.ForeColor.RGB = self.rgb(kwargs['fill'])
        if style == 'dashed':
            line.Line.DashStyle = self.MSO_LINE_DASH

    def arc(self, xy, start, end, **kwargs):
        w = self.width(xy)
        h = self.height(xy)
        pt = XY(xy[0], xy[1])
        shape = self.book.ActiveSheet.Shapes.AddShape(self.MSO_SHAPE_ARC, pt.x, pt.y, w, h)
        shape.Adjustments.SetItem(1, start - end)
        shape.Rotation = start + 90
        shape.Left     = xy[0]
        shape.Top      = xy[1]

    def ellipse(self, xy, **kwargs):
        w = self.width(xy)
        h = self.height(xy)
        pt = XY(xy[0], xy[1])

        shape = self.book.ActiveSheet.Shapes.AddShape(self.MSO_SHAPE_OVAL, pt.x, pt.y, w, h)
        self.set_style(shape, kwargs)

    def polygon(self, xy, **kwargs):
        form = self.book.ActiveSheet.Shapes.BuildFreeform(0, xy[0][0], xy[0][1])

        for (x,y) in xy[1:]:
            form.AddNodes(0, 0, x, y) 
        shape = form.ConvertToShape()
        self.set_style(shape, kwargs)

    def save(self, filename, size, format):
        path = os.path.abspath(filename)
        if os.path.exists(path):
            os.remove(path)
        self.book.SaveAs(path, 56)
        self.excel.Quit()

    def textlinesize(self, string, font, **kwargs):
        return Size(len(string)*15,10)

    def textsize(self, string, font, maxwidth=None, **kwargs):
        if maxwidth is None:
            maxwidth = 65535

        box = Box(0, 0, maxwidth, 65535)
        textbox = textfolder.get(self, box, string, font=None, **kwargs)
        return textbox.outlinebox.size

def setup(self):
    from blockdiag.imagedraw import install_imagedrawer
    install_imagedrawer('shape.xls', ExcelShapeImageDraw)
