import os
import win32com.client
import docx
import re


def doc2docx(ori_path_abs, conv_path_abs):
    word = win32com.client.Dispatch('Word.application')
    doc = word.Documents.Open(ori_path_abs)
    doc.SaveAs2(conv_path_abs, FileFormat=16)
    doc.Close()


class DocProc:
    def __init__(self, filename):
        self.doc = docx.Document(filename)
    
    def get_text(self):
        for para in self.doc.paragraphs:
            num = self.get_report_num(para.text)
            print(num)
        #print(len(self.doc.paragraphs))
        #print(self.doc.paragraphs[21].text)
        #print(self.doc.paragraphs[21]._element.xml)
        #print(self.doc.paragraphs[19].text)
        #print(self.doc.paragraphs[19]._element.xml)
    
    def get_report_num(self, text):
        num = re.search('鲁海会济审字.*$', text)
        if num:
            return num.group()
        else:
            return None
    
    def get_report_time(self, para):
        print(para._element.xml)
    
    def get_company_name(self, text):
        print(para._element.xml)

        


def conv_doc2docx(filename):
    """

    :param filename: doc文档的文件名
    :return: None
    """
    # 用os.path.splitext把文件名和扩展名拆开来，分别存为filename_base和filename_ext
    filename_base,filename_ext = os.path.splitext(filename)
    # 文件名加上.docx扩展名就是转换以后的文件名了。当然实际的转换在后面完成
    filename_conv = filename_base+'.docx'
    # 下面两行是检验前面提到的两个问题
    assert '~$' not in filename_base, f"~$ in filename:{filename}"
    assert filename_ext=='.doc', f'{filename} should be .doc file'

    # 下面一行用basedir目录和文件名形成word文档的相对链接
    path = os.path.join(basedir,filename)
    # 下面一行用basedir目录和文件名形成转换后word文档（docx）的相对链接
    path_conv = os.path.join(basedir,filename_conv)
    # 下面两行是把相对链接转换成绝对链接
    ori_path_abs = os.path.abspath(path)
    conv_path_abs = os.path.abspath(path_conv)
    word = win32com.client.Dispatch('Word.application')
    try:
        doc = word.Documents.Open(ori_path_abs)
        doc.SaveAs2(conv_path_abs, FileFormat=16)
        doc.Close()
    except Exception as e:
        print(f'Fail to convert:{filename}')
        print(e)


if __name__ == '__main__':
    basedir = 'F:/海天会计事务所/2017年度审计报告/济宁华都房地产开发有限公司/'
    #filename = basedir + '审计报告.doc'
    #conv_doc2docx(filename)
    filename2 = basedir + '审计报告.docx'
    doc_proc = DocProc(filename2)
    doc_proc.get_text()

