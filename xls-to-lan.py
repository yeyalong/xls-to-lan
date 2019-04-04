import chardet
# import binascii
import xlrd
import xlwt
import os
import shutil

class FileProcess():
    def __init__(self):
        self.language_col_number_ = 0
        self.list_language = []
        self.list_first = []
        self.list_second = []

    def ReadExcel(self, language, excel_language_name):
        for i in range(0, ncols):
            row_first = table.cell(1, i).value  # 取第二行的值
            if row_first == language:
                self.language_col_number_ = i

        for nrow in range(0, nrows):  # 遍历每一行
            language_value = table.cell(nrow, self.language_col_number_).value  # 取language列的值
            col_first = table.cell(nrow, 0).value  #取第一列的值
            col_second = table.cell(nrow, 1).value  #取第二列的值
            self.list_language.append(language_value)
            self.list_first.append(col_first)
            self.list_second.append(col_second)

        book_write = xlwt.Workbook(encoding=encode_name)
        sheet = book_write.add_sheet('test')
        for i in range(len(self.list_first)):
            sheet.write(i, 0, self.list_first[i])
        for i in range(len(self.list_second)):
            sheet.write(i, 1, self.list_second[i])
        for i in range(len(self.list_language)):
            sheet.write(i, 2, self.list_language[i])

        if os.path.exists(excel_language_name):
            os.remove(excel_language_name)
        book_write.save(excel_language_name)
        self.list_language.clear()
        self.list_first.clear()
        self.list_second.clear()

    def strs(self, row):
        try:
            values = ""
            for i in range(len(row)):
                values = values + str(row[i])
            return values
        except:
            raise

    def xls_txt(self, xls_name, txt_name):
        try:
            data = xlrd.open_workbook(xls_name)
            if os.path.exists(txt_name):
                os.remove(txt_name)
            sqlfile = open(txt_name, "a", encoding=encode_name)
            table = data.sheets()[0]
            nrows = table.nrows
            for ronum in range(0, nrows):
                row = table.row_values(ronum)
                values = process.strs(row) + "\n"  # 调用函数，将行数据拼接成字符串
                sqlfile.writelines(values)
            sqlfile.close()
        except:
            pass

    # 获取文件编码类型
    def GetEncode(self, file):
        # 二进制方式读取，获取字节数据，检测类型
        with open(file, 'rb') as f:
            data = f.read()
            # print(binascii.hexlify(data), end=' ')
            return chardet.detect(data)['encoding']

    def SetEncode(self):
        set_encode = encoding
        if encoding == 'GB2312':
            set_encode = 'GB18030'
        return set_encode

    def Replace(self, file):
        with open(file, 'r', encoding=self.SetEncode()) as infopen:
            lines = infopen.readlines()
        with open(file, 'w', encoding=self.SetEncode()) as outfopen:
            for line in lines:
                if (line[:1] != '=') and (line[-2:-1] != '=') and (('UCCode' in line) or ('#menu' in line)):
                        if '\t' in line:
                            line = line.replace('\t', '')
                        if '\'' in line:
                            line = line.replace('\'', '')
                        if '\"' in line:
                            line = line.replace('\"', '')
                        if line == '\n':
                            line = line.strip("\n")
                        outfopen.writelines(line)

    def DeleteLf(self, file):
        with open(file, "rb") as f1, open("%s.bak" % file, "wb") as f2:
            for line in f1:
                if b'\\n' in line:
                    line = line.replace(b'\\n', b'\n')
                f2.write(line)
        os.remove(file)
        os.rename("%s.bak" % file, file)

    def Convert(self, file):
        with open(file, "rb") as f1, open("%s.ba" % file, "wb") as f2:
            for line in f1:
                line_str = str(line, encoding='utf8')
                for i in line_str:
                    byte_i = bytes(i, encoding='utf8')
                    if len(byte_i) != 1:
                        byte_many = bytes(str(byte_i)[2:-1], encoding='utf8')
                        f2.write(byte_many)
                    else:
                        f2.write(bytes(i, encoding='utf8'))
        os.remove(file)
        os.rename("%s.ba" % file, file)

if __name__ == '__main__':
    process = FileProcess()
    if os.path.exists('./xls/') == False:
        os.mkdir('./xls/')
    if os.path.exists('./txt/') == False:
        os.mkdir('./txt/')
    if os.path.exists('./lan/') == False:
        os.mkdir('./lan/')
    excel_all_name = "./F2000语言文件(完整版).xls"
    book = xlrd.open_workbook(excel_all_name)
    table = book.sheet_by_index(0)
    nrows = table.nrows  # 获取行总数
    ncols = table.ncols  # 获取列总数

    for i in range(3, ncols):
        language_name = table.cell(1, i).value  # 取第二行的值
        language_name = os.path.splitext(language_name)[0]
        encode_name = table.cell(2, i).value  # 取第三行的值
        if (language_name != '') and (encode_name != ''):
            excel_language_name = './xls/' + language_name + ".xls"
            txt_language_name = './txt/' + language_name + ".txt"
            lan_language_name = './lan/' + language_name + ".lan"
            process.ReadExcel(language_name, excel_language_name)
            process.xls_txt(excel_language_name, txt_language_name)
            encoding = process.GetEncode(txt_language_name)
            process.SetEncode()
            process.Replace(txt_language_name)
            process.DeleteLf(txt_language_name)
            if encoding == 'utf-8':
                process.Convert(txt_language_name)
            shutil.copy(txt_language_name, lan_language_name)