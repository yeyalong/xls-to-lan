# -*- coding: utf-8 -*-
import chardet
import xlrd
import xlwt
import os
import shutil
import configparser

class FileProcess():
    def Ini(self):
        config = configparser.ConfigParser()
        config.read('F2000_lan_tool.ini')
        self.language_add1 = config.getint("section_col_number", "language_add_col_1")
        self.language_add2 = config.getint("section_col_number", "language_add_col_2")
        self.excel_all_name = config.get("set_path", "xls_all_path")
        self.xls_path = config.get("set_path", "xls_path")
        self.txt_path = config.get("set_path", "txt_path")
        self.lan_path = config.get("set_path", "lan_path")

    def ReadExcel(self, language, excel_language_name, excel_language_name_3col):
        for i in range(0, self.ncols):
            row_first = self.table.cell(1, i).value  # 取第二行的值
            if row_first == language:
                self.language_col_number_ = i

        book_write = xlwt.Workbook(encoding=self.encode_name)
        sheet = book_write.add_sheet('test')
        book_write_3col = xlwt.Workbook(encoding=self.encode_name)
        sheet_3col = book_write_3col.add_sheet('test')

        for nrow in range(0, self.nrows):  # 遍历每一行
            language_value = self.table.cell(nrow, self.language_col_number_).value  # 取language列的值
            add1_value = self.table.cell(nrow, self.language_add1).value  # 取add1列的值
            add2_value = self.table.cell(nrow, self.language_add2).value  # 取add2列的值
            col_first = self.table.cell(nrow, 0).value  #取第一列的值
            col_second = self.table.cell(nrow, 1).value  #取第二列的值
            col_third = self.table.cell(nrow, 2).value  #取第三列的值

            sheet.write(nrow, 0, col_first)
            sheet.write(nrow, 1, col_second)
            sheet.write(nrow, 2, col_third)
            sheet.write(nrow, 3, add1_value)
            sheet.write(nrow, 4, add2_value)
            if (self.language_col_number_ != self.language_add1) and (
                    self.language_col_number_ != self.language_add2):
                sheet.write(nrow, 5, language_value)

            sheet_3col.write(nrow, 0, col_first)
            sheet_3col.write(nrow, 1, col_second)
            sheet_3col.write(nrow, 2, language_value)

        if os.path.exists(excel_language_name):
            os.remove(excel_language_name)
        book_write.save(excel_language_name)
        if os.path.exists(excel_language_name_3col):
            os.remove(excel_language_name_3col)
        book_write_3col.save(excel_language_name_3col)

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
            sqlfile = open(txt_name, "a", encoding=self.encode_name)
            table = data.sheets()[0]
            nrows = table.nrows
            for ronum in range(0, nrows):
                row = table.row_values(ronum)
                values = self.strs(row) + "\n"  # 调用函数，将行数据拼接成字符串
                sqlfile.writelines(values)
            sqlfile.close()
        except:
            pass

    # 获取文件编码类型
    def GetEncode(self, file):
        # 二进制方式读取，获取字节数据，检测类型
        with open(file, 'rb') as f:
            data = f.read()
            return chardet.detect(data)['encoding']

    def SetEncode(self):
        set_encode = self.encoding
        if self.encoding == 'GB2312':
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

    def Convert(self, file, file_convert):
        with open(file, "rb") as f1, open(file_convert, "wb") as f2:
            for line in f1:
                line_str = str(line, encoding='utf8')
                for i in line_str:
                    byte_i = bytes(i, encoding='utf8')
                    if len(byte_i) != 1:
                        byte_many = bytes(str(byte_i)[2:-1], encoding='utf8')
                        f2.write(byte_many)
                    else:
                        f2.write(bytes(i, encoding='utf8'))
    def MakeDir(self):
        if os.path.exists(self.xls_path) == False:
            os.mkdir(self.xls_path)
        if os.path.exists(self.txt_path) == False:
            os.mkdir(self.txt_path)
        if os.path.exists(self.lan_path) == False:
            os.mkdir(self.lan_path)

    def Main(self):
        book = xlrd.open_workbook(self.excel_all_name)
        self.table = book.sheet_by_index(0)
        self.nrows = self.table.nrows  # 获取行总数
        self.ncols = self.table.ncols  # 获取列总数

        for i in range(3, self.ncols):
            language_name = self.table.cell(1, i).value  # 取第二行的值
            language_name = os.path.splitext(language_name)[0]
            self.encode_name = self.table.cell(2, i).value  # 取第三行的值
            if (language_name != '') and (self.encode_name != ''):
                excel_language_name = self.xls_path + language_name + ".xls"
                txt_language_name = self.txt_path + language_name + ".txt"
                lan_language_name = self.lan_path + language_name + ".lan"
                excel_language_name_3col = self.xls_path + language_name + "_3col.xls"
                txt_language_name_convert = self.txt_path + language_name + "_convert.txt"
                self.ReadExcel(language_name, excel_language_name, excel_language_name_3col)
                self.xls_txt(excel_language_name_3col, txt_language_name)
                os.remove(excel_language_name_3col)
                self.encoding = self.GetEncode(txt_language_name)
                self.SetEncode()
                self.Replace(txt_language_name)
                self.DeleteLf(txt_language_name)
                if self.encoding == 'utf-8':
                    self.Convert(txt_language_name, txt_language_name_convert)
                    shutil.copy(txt_language_name_convert, lan_language_name)
                    os.remove(txt_language_name_convert)
                else:
                    shutil.copy(txt_language_name, lan_language_name)

if __name__ == '__main__':
    process = FileProcess()
    process.Ini()
    process.MakeDir()
    process.Main()