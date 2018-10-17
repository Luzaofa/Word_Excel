from win32com import client as wc


class ExcelHelper(object):

    def __init__(self, filepath, Debug=False):
        """
        :param filepath:
        :param Debug: 控制过程是否可视化
        """
        self.excelApp = wc.Dispatch('Excel.Application')
        self.excelApp.Visible = Debug
        self.excelApp.DisplayAlerts = 0
        self.myexcel = self.excelApp.Workbooks.Open(filepath)

    def find_sheet(self, number):
        """查找操作的sheet"""
        sheet_name = 'Sheet' + str(number)
        try:
            self.sheet = self.myexcel.Sheets(sheet_name).Select()
        except Exception as msg:
            print(msg)

    def select_position_by_range(self, sheet_num, range):
        """选择Excel第sheet_num个Sheet中的区域"""
        mySheet = self.find_sheet(sheet_num)
        mySheet.Range(range).Select()

    def get_cell_value(self, row, col):
        """获取某个单元格的值（传入具体行列位置）"""
        return self.excelApp.Cells(row, col).Text

    def get_value_by_range(self, range):
        """通过range获取单元格的值（传入具体range位置）"""
        return self.excelApp.Range(range).Text

    def set_cell_value(self, row, col, value):
        """单元格赋值（传入具体行列位置）"""
        self.excelApp.Cells(row, col).value = value

    def set_value_by_range(self, range, value):
        """通过range为单元格赋值"""
        self.excelApp.Range(range).Value = value

    def set_y_minvalue(self, pic_name, y_minvalue):
        """更改条形图Pic_name的Y轴最小值为Y_value"""
        self.myexcel.ActiveSheet.ChartObjects(pic_name).Activate()
        self.myexcel.ActiveChart.Axes(2).Select()
        self.myexcel.ActiveChart.Axes(2).MinimumScale = y_minvalue
        self.myexcel.Application.CommandBars("Format Object").Visible = False

    def set_x_valuerange_spacing(self, pic_name, x_value_range_start, x_value_range_end):
        """设置图片pic_name的X轴坐标数据范围（需提供开始单元格和结束单元格位置）以及X轴数据间隔（最大数量限制为9）"""
        self.myexcel.ActiveSheet.ChartObjects(pic_name).Activate()
        self.myexcel.ActiveChart.PlotArea.Select()
        self.myexcel.ActiveChart.FullSeriesCollection(1).Values = "=Sheet2!$K${0}:$K${1}".format(str(x_value_range_start),
                                                                                            str(x_value_range_end))
        self.myexcel.ActiveChart.FullSeriesCollection(2).Values = "=Sheet2!$L${0}:$L${1}".format(str(x_value_range_start),
                                                                                            str(x_value_range_end))
        self.myexcel.ActiveChart.FullSeriesCollection(2).XValues = "=Sheet2!$J${0}:$J${1}".format(str(x_value_range_start),
                                                                                             str(x_value_range_end))
        # X 轴间隔
        Space_num = int((x_value_range_end - x_value_range_start) / 9)
        self.myexcel.ActiveChart.Axes(1).Select()
        self.myexcel.ActiveChart.Axes(1).Select()
        self.myexcel.ActiveChart.Axes(1).TickMarkSpacing = Space_num

    def del_lenged(self, pic_name, lenged_num):
        """删除图片Pic_name的第lenged_num个图例"""
        self.myexcel.ActiveSheet.ChartObjects(pic_name).Activate()
        self.myexcel.ActiveChart.Legend.Select()
        self.myexcel.ActiveChart.Legend.LegendEntries(lenged_num).Select()
        self.excelApp.Selection.Delete()

    def set_lenged_position(self, pic_name, lenged_width, lenged_height, lenged_left, lenged_top):
        """更改图片pic_name的图例位置（长、宽、距左、顶部位置）"""
        self.myexcel.ActiveSheet.ChartObjects(pic_name).Activate()
        self.myexcel.ActiveChart.PlotArea.Select()
        self.myexcel.ActiveChart.Legend.Select()
        self.excelApp.Selection.Width = lenged_width
        self.excelApp.Selection.Height = lenged_height
        self.excelApp.Selection.Left = lenged_left
        self.excelApp.Selection.Top = lenged_top

    def set_lable_position(self, pic_name, lenged_num, lenged_left, lenged_top):
        """更改图片pic_name的数据标签lenged_num的位置（距左、顶部位置）"""
        self.myexcel.ActiveSheet.ChartObjects(pic_name).Activate()
        self.myexcel.ActiveChart.FullSeriesCollection(lenged_num).Points(1).DataLabel.Select()
        self.excelApp.Selection.Left = lenged_left
        self.excelApp.Selection.Top = lenged_top

    def claer_rows(self, sheet_num, start_line):
        """清除excel中第sheet_num个Sheet多余整行数据（开始行到1000行所有整行数据）"""
        clear_rows = str(start_line) + ':1000'
        mySheet = self.find_sheet(sheet_num)
        mySheet.Rows(clear_rows).ClearContents()

    def del_rows(self, sheet_num, start_line):
        """删除excel中第sheet_num个Sheet多余数据行（开始行到1000行所有整行数据）"""
        del_rows = str(start_line) + ':1000'
        mySheet = self.find_sheet(sheet_num)
        mySheet.Rows(del_rows).Delete()

    def claer_rows_by_range(self, sheet_num, start_range):
        """清除excel中第sheet_num个Sheet指定单元格数据（clear_range）"""
        clear_range = "L{0}:L1000".format(start_range)
        self.select_position_by_range(sheet_num, clear_range)
        self.excelApp.Selection.ClearContents()

    def copy_img(self, img_name):
        """复制Excel中图片"""
        try:
            self.excelApp.ActiveSheet.ChartObjects(img_name).Activate()
            self.excelApp.ActiveChart.ChartArea.Copy()
        except Exception as msg:
            print("复制图片出错:{0}".format(msg))

    def copy_img_to_file(self, sheet_num, pic_range, pic_name, pic_savePath):
        """复制Excel指定区域为图片将其保存到指定文件夹"""
        from PIL import ImageGrab
        mySheet = self.find_sheet(sheet_num)
        mySheet.Range(pic_range).CopyPicture()
        # mySheet.Paste(mySheet.Range('K1'))
        self.excelApp.Selection.ShapeRange.Name = pic_name
        try:
            mySheet.Shapes(pic_name).Copy()
        except Exception as msg:
            print("复制图片出错:{0}".format(msg))
        img = ImageGrab.grabclipboard()  # 获取图片数据
        img.save(pic_savePath)

    def copy_table(self, range):
        """复制Excel中的表格"""
        try:
            self.excelApp.Range(range).Select()
            self.excelApp.Selection.Copy()
        except Exception as msg:
            print("复制表格出错:{0}".format(msg))

    def save(self):
        """保存文件"""
        self.myexcel.Save()

    def saveAs(self, filepath):
        """文件另存为"""
        self.myexcel.SaveAs(filepath)

    def close(self):
        """关闭文件"""
        self.myexcel.Close()
        self.excelApp.Quit()
