import os
import sys

import ExcelHelper
import WordHelper
import Config


class MainService(object):

    def __init__(self, template_path):

        self.excelhelper = ExcelHelper.ExcelHelper(f'{template_path}Template.xlsx')
        self.wordhelper = WordHelper.WordHelper(f'{template_path}Template.doc')

    def main(self, target_path, save_name):
        """
            选择Sheet2，将Word模板图片替换为Excel中的图片，同时替换页眉作者姓名
            最终生成Excel、Word、PDF目标文件
        """
        self.excelhelper.find_sheet(2)

        for key, value in Config.REPLACE_PIC_PASTE.items():
            self.excelhelper.copy_img(value)
            self.wordhelper.paste_img(key)

        self.wordhelper.replace_bookmarks_text(Config.REPLACE_STR_DICT)

        saveName = target_path + save_name
        self.excelhelper.saveAs(f'{saveName}.xlsx')
        self.wordhelper.export_pdf(f'{saveName}.pdf')
        self.wordhelper.saveAs(f'{saveName}.doc')

def kill_process():
    """关闭后台Excel、Word所占进程"""
    command = 'taskkill /F /IM {process_name}'
    for i in ['EXCEL.EXE', 'WINWORD.EXE']:
        os.system(command.format(process_name=i))


if __name__ == '__main__':

    root_path = os.getcwd()

    save_name = 'New_template'

    if len(sys.argv) == 2:
        save_name = sys.argv[1]

    template_path = f'{root_path}\\Files\\Template\\'
    target_path = f'{root_path}\\Files\\target\\'

    kill_process()

    MainService = MainService(template_path)

    MainService.main(target_path, save_name)

    kill_process()

