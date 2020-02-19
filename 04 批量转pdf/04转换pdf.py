import win32com.client
import pythoncom
import os


 
 
class Word_2_PDF(object):
 
    def __init__(self, filepath, Debug=False):
        """
        :param filepath:
        :param Debug: 控制过程是否可视化
        """
        self.wordApp = win32com.client.Dispatch('word.Application')
        self.wordApp.Visible = Debug
        self.myDoc = self.wordApp.Documents.Open(filepath)
 
    def export_pdf(self, output_file_path):
        """
        将Word文档转化为PDF文件
        :param output_file_path:
        :return:
        """
        self.myDoc.ExportAsFixedFormat(output_file_path, 17, Item=7, CreateBookmarks=0)

    def close(self):
        self.wordApp.Quit()
         
if __name__ == '__main__':

    rootpath = os.getcwd()  # 文件夹路径
    save_path = os.getcwd()   # PDF储存位置
 
    #rootpath = 'C:\\python\\07 批量转pdf\\'       # 文件夹根目录
    
    pythoncom.CoInitialize()

    os_dict = {root:[dirs, files] for root, dirs, files in os.walk(rootpath)}
    for parent, dirnames, filenames in os.walk(rootpath):
        for filename in filenames:
            if u'.doc' in filename and u'~$' not in filename:
                  # 直接保存为PDF文件
                #print(rootpath+filename)
                a = Word_2_PDF(rootpath +'\\'+ filename, True)
                title = filename.split('.')[0]  # 删除.docx
                a.export_pdf(rootpath  +'\\'+ title+'.pdf')
                


    print('转化完成')
  
   
 
    
