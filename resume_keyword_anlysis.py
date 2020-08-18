import os
import glob
import docx2txt
import subprocess
from collections import Counter



class File:
    def __init__(self, filepath):
        self.filepath = filepath

    def read_text(self):
        if str(self.filepath).endswith('.docx'):
            text = docx2txt.process(str(self.filepath))
            return text
        elif str(self.filepath).endswith('.hwp'):
            text = subprocess.check_output('hwp5txt test_hwp.hwp', shell=True, encoding='UTF-8')
            return text


def analyzing_keyword():  # 작성 필요
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
    file_list = get_file_path(os.getcwd())
    print(file_list)

    # templ = docx2txt.process(file_list[1])
    # print(templ)
    file = File(file_list[1])
    file.read_text()


if __name__ == '__main__':
    main()