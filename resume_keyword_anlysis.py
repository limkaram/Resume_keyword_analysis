import os
import glob
import docx2txt
import subprocess
from konlpy.tag import Okt
from collections import Counter

DEFALUT_DICECTORY_PATH = os.getcwd()  # 분석 예정 파일 저장되어 있는 폴더 PATH


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


def get_file_list(path):
    files = glob.glob(path + '/*')
    if files:
        files = [file for file in files if file.endswith('.docx') or file.endswith('.hwp')]
        return files


def main():
    file_list = get_file_list(DEFALUT_DICECTORY_PATH)
    if not file_list:
        print('키워드 분석 가능한 파일이 폴더내 존재하지 않습니다.')
    while file_list:
        print('==분석 가능한 파일 목록==')
        for index, path_of_file in enumerate(file_list):
            print('{0}. {1}'.format(index+1, os.path.basename(path_of_file)))
        print('\n')
        print('키워드 분석하고자 하는 파일 번호 입력')
        print('※모든 파일을 일괄 분석코자 하는 경우 "0" 입력')

        try:
            input_num = int(input('입력 >>'))
            if input_num > len(file_list)-1:  # 입력이 file_list 범위를 넘는 경우
                print('올바르지 않은 입력입니다.')
            elif input_num == 0:  # 모든 파일 일괄 분석 경우
                pass
            elif 1 <= input_num <= len(file_list):  # 원하는 하나의 파일을 분석하는 경우
                pass
        except ValueError:
            print('Warning : 올바르지 않은 입력, 정수 입력 필요.\n')


    # templ = docx2txt.process(file_list[1])
    # print(templ)
    # file = File(file_list[1])
    # file_text = file.read_text()
    # print(file.analyze_text(file_text))


if __name__ == '__main__':
    main()