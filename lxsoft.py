#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Last Update:

import win32com.client
import pywintypes
import pythoncom
import os

def GoodString(value):
    """
    >>> GoodString(2)
    '2'
    """
    try:
        return str(value)
    except UnicodeEncodeError:
        return value

class Lxsoft():
    def __init__(self, log, path=''):
        self.log = log
        self.lx = self.__registerLxsoft()
        self.InitExcel()
        if path:
            self.Open(path)
            self.SheetIndex()
            self.SheetCount()

    def __registerLxsoft(self):
        '''注册插件
        '''
        pythoncom.CoInitialize()
        command = "regsvr32 /s atl.dll"
        os.system(command)
        try:
            lx = win32com.client.Dispatch('Lazy.LxjExcel')
        except pywintypes.com_error:
            command = 'regsvr32 /s LazyOffice.dll'
            if os.system(command) == 0:
                lx = win32com.client.Dispatch('Lazy.LxjExcel')
                # assert(dm.Ver == '2.1138')
                return lx
            else:
                self.log.error('regsver32 error')
        else:
            return lx
     
    def InitExcel(self):
        self.path = ''
        self.maxsheet = 1
        self.index = 1
        self.sindex = 1

    def __del__(self):
        self.Close()

    def __Trueindex(self, index):
        if index == '':
            return self.index
        else:
            return index

    def Open(self, path, visual=0):
        # visual:1可见,0不可见 opencd:打开密码(如果存在)
        # writecd:写入密码 flag:只读方式打开
        self.path = path
        self.index =  self.lx.ExcelOpen(path, visual)
        return self.index

    def Close(self, index=''):
        index = self.__Trueindex(index)
        return self.lx.ExcelClose(index)

    def SheetIndex(self, index=''):
        # 获取当前标签页序号
        index = self.__Trueindex(index)
        self.sindex = self.lx.SheetIndex(index)
        return self.sindex

    def SheetName(self, Sindex, index=''):
        # 获取标签页序号(或名称)为Sindex的标签名称(或索引)
        index = self.__Trueindex(index)
        return self.lx.SheetGetName(Sindex, index)

    def SheetCount(self, index=''):
        # 获取标签页总数
        index = self.__Trueindex(index)
        self.maxsheet =  self.lx.SheetCount(index)
        return self.maxsheet

    def SheetAdd(self, Sindex, index=''):
        # 在第Sindex个标签页之前新建一个标签
        # 如果Sindex是字符串,则在最后新建一个名叫Sindex的标签
        index = self.__Trueindex(index)
        self.log.debug(u'新建标签: '+GoodString(Sindex))
        return self.lx.SheetAdd(Sindex, index)

    def SheetRename(self, Sindex, name, index=''):
        # Sindex可为序号或者名称
        index = self.__Trueindex(index)
        return self.lx.SheetRename(Sindex, name, index)

    def SheetDel(self, Sindex, index=''):
        index = self.__Trueindex(index)
        self.log.debug(u'删除标签: '+GoodString(Sindex))
        return self.lx.SheetDel(Sindex, index)

    def Write(self, Sindex, x, y, string, index=''):
        # Write(1,3,2,"内容",Index)向单元格(3,2)即'B3'写入内容
        index = self.__Trueindex(index)
        self.log.debug(u'向单元格(%s,%s)写入内容: %s' % (GoodString(x),GoodString(y),GoodString(string)))
        return self.lx.ExcelWrite(Sindex, x, y, string, index)

    def Read(self, Sindex, x, y, index=''):
        index = self.__Trueindex(index)
        string = self.lx.ExcelRead(Sindex, x, y, index)[0]
        self.log.debug(u'读取单元格(%s,%s): %s' % (GoodString(x),GoodString(y),GoodString(string)))
        return string

    def differ(self, Sindex, x1, y1, x2, y2, index1='', index2=''):
        index1 = self.__Trueindex(index1)
        index2 = self.__Trueindex(index2)
        return int(self.Read(Sindex, x1, y1, index1)) - int(self.Read(Sindex, x2, y2, index2))

    def Cells(self, tab1, action, tab2, index=''):
        # 如设置tab2为负数，将只进行复制
        # 如设置tab2为0，将在所有标签页后面新建一个被复制的表
        # 将工作表第1个标签页整表复制到第3个标签页:self.lx.ExcelCells(1, u"复制", 3, Index)
        index = self.__Trueindex(index)
        self.log.debug(u'表%s %s -> %s' % (GoodString(tab1),action,GoodString(tab2)))
        return self.lx.ExcelCells(tab1, action, tab2, index)

    def Rows(self, tab, row1, action, row2, index=''):
        # 如设置row2为负数，将只进行复制
        # 复制第1个标签页第2行到第8行:self.lx.ExcelRows(1, 2, u"复制", 8, Index)
        # 第2行模糊查找1:self.lx.ExcelRows(1, 2, u"模糊查找", 1, Index)
        index = self.__Trueindex(index)
        self.log.debug(u'行%s %s -> %s' % (GoodString(row1),action,GoodString(row2)))
        return self.lx.ExcelRows(tab, row1, action, row2, index)

    def Columns(self, tab, col1, action, col2, index=''):
        # 如设置col2为负数，将只进行复制
        # 复制第1个标签页第3列到第6列:self.lx.ExcelColumns(1, 3, u"复制", 6, Index)
        # 第3列模糊查找1:self.lx.ExcelColumns(1,3, u"模糊查找", 1, Index)
        index = self.__Trueindex(index)
        self.log.debug(u'列%s %s -> %s' % (GoodString(col1),action,GoodString(col2)))
        return self.lx.ExcelColumns(tab, col1, action, col2, index)

    def Range(self, tab, ranges, action, destination, index=''):
        # 如设置destination为负数，将只进行复制
        # 复制第1个标签页"B2:C5"区域:self.lx.ExcelRange(1, "B2:C5", "复制", -1, Index)
        # 区域C2:E5模糊查找1:self.lx.ExcelRange(1, "C2:E5",u"模糊查找", 1, Index)
        index = self.__Trueindex(index)
        self.log.debug(u'区域%s %s -> %s' % (GoodString(tab),action,GoodString(destination)))
        return self.lx.ExcelRange(tab, ranges, action, destination, index)


if __name__ == '__main__':
    import log
    import time
    path = 'F:\\test.xls'
    mylog = log.Log()
    lx = Lxsoft(mylog, path)
    # print int(lx.Read(lx.SheetCount(), 5, 7))==int(time.strftime("%Y%m%d"))
    # a = lx.Read(lx.maxsheet, 5, 7)
    a = lx.SheetCount()
    print(lx.Columns(a, 7, u"模糊查找", "20150722")[0][0])
    # lx.Cells(lx.SheetCount(), 5, u"字体颜色", 3)
    # print lx.differ(lx.maxsheet, 17, 7, 17, 3)
    # print a
    # lx.Write(lx.maxsheet, 5, 7, "20150709")
    # lx.Cells(80, u"复制", 0, index)
    lx.Close()
    # lx = Lxsoft(config.log)
    # index = lx.Open(path, 0)
    # lx.Close(index)


