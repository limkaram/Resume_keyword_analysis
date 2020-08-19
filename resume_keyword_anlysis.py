import os
import glob
import docx2txt
import subprocess
from konlpy.tag import Okt
from collections import Counter


class File:
    def __init__(self, filepath):
        self.filepath = filepath

    def read_text(self):
        if str(self.filepath).endswith('.docx'):
            text = docx2txt.process(str(self.filepath))
            return text
        elif str(self.filepath).endswith('.hwp'):
            # pyhwp 다운로드 : pip install --user --pre pyhwp
            # shell=Ture시 셀기반이기때문에, args를 리스트로 넘겨줄 필요 없음(문자열로 넘겨주면 됨)
            text = subprocess.check_output('hwp5txt test_hwp.hwp', shell=True, encoding='UTF-8')
            return text

    def analyze_text(self, raw_text):
        text = Okt()
        # text.nouns(raw_text) : 명사 추출
        # text.morphs(raw_text) : 형태소 추출
        # text.pos(raw_text) : 품사 부착하여 추출
        # print(okt.tagset) : 각 품사 태그의 기호와 의미 확인 가능
        return text.nouns(raw_text)


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
    file_text = file.read_text()
    print(file.analyze_text(file_text))


if __name__ == '__main__':
    main()