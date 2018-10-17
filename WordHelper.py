import win32com.client


class WordHelper(object):

    def __init__(self, filepath, Debug=False):
        """
        :param filepath:
        :param Debug: 控制过程是否可视化
        """
        self.wordApp = win32com.client.Dispatch('word.Application')
        self.wordApp.Visible = Debug
        self.myDoc = self.wordApp.Documents.Open(filepath)

    def replace_bookmark_text(self, key, value):
        """
        置换单个书签(文本类型)
        :param key: 书签名称
        :param value: 值
        :return:
        """
        try:
            self.myDoc.Bookmarks[str(key)].Range.Text = value
            return 1
        except:
            print('找不到书签{key}'.format(key=key))
            return 0

    def replace_bookmarks_text(self, paramdict={}):
        """
        置换多个书签(文本类型)
        :param paramdict: 书签名-植入值 键值对
        :return: 置换数量
        """
        effected_num = 0
        for key, value in paramdict.items():
            effected_num += self.replace_bookmark_text(key, value)
        return effected_num

    def move_to_bookmark(self, bookmark_key):
        """
        移动到书签
        :param bookmark_key:
        :return:
        """
        if bookmark_key is None or bookmark_key == '':
            raise Exception('书签名不能为空')
        try:
            self.point = self.wordApp.ActiveDocument.Bookmarks(bookmark_key).Select()
        except Exception as msg:
            print("找不到书签{bookmark_key} {errormsg}".format(bookmark_key=bookmark_key, errormsg=msg))

    def paste_img(self, bookmark_key):
        """
        图片粘贴（需要先复制）
        :return:
        """
        try:
            self.wordApp.ActiveDocument.Bookmarks(bookmark_key).Select()
        except Exception as msg:
            print('找不到书签 {}'.format(msg))
            return 0
        # 对粘贴失败的图片进行处理
        try:
            self.wordApp.Selection.Paste()
        except Exception as msg:
            # print('粘贴图片失败 {}'.format(msg))
            return -1
        return 1

    def paste_table(self, bookmark_key):
        """粘贴表格"""
        try:
            self.wordApp.ActiveDocument.Bookmarks(bookmark_key).Select()
            self.wordApp.Selection.Paste()
            self.wordApp.Selection.Tables(1).AutoFitBehavior(1)
            self.wordApp.Selection.Tables(1).AutoFitBehavior(2)
        except Exception as msg:
            print("粘贴表格失败 {}".format(msg))

    def export_pdf(self, output_file_path):
        """Word转PDF"""
        self.myDoc.ExportAsFixedFormat(output_file_path, 17, Item=7, CreateBookmarks=0)

    def save(self):
        """保存文件"""
        self.myDoc.Save()

    def saveAs(self, filepath):
        """文件另存为"""
        self.myDoc.SaveAs(filepath)

    def close(self):
        """关闭应用"""
        self.myDoc.Close()
        self.wordApp.Quit()


