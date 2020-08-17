import os
import glob
import docx2txt
from collections import Counter


class File:
    def __init__(self, filepath):
        self.filepath = filepath

    def read_text(self):
        text = ''
        if str(self.filepath).endswith('.docx'):
            text = docx2txt.process(str(self.filepath))
            return text
        elif str(self.filepath).endswith('.hwp'):
            pass


def get_file_path(path):
    files = glob.glob(path + '/*')
    if files:
        files = [file for file in files if file.endswith('.docx') or file.endswith('.hwp')]
        return files
    else:
        print('분석할 파일이 폴더내에 존재 하지 않습니다.')
        # print('기본 디렉토리 위치 : ', DEFAULT_FILEPATH)


def main():
    DEFAULT_FILEPATH = '/Users/limkaram/PycharmProjects/resume_keyword_analysis'
    file_list = get_file_path(DEFAULT_FILEPATH)
    print(file_list)

    templ = docx2txt.process(file_list[1])
    print(templ)


if __name__ == '__main__':
    main()