import os
import glob
import docx2txt
import subprocess
import pandas as pd
from matplotlib import pyplot as plt
import seaborn as sns
from konlpy.tag import Okt
from collections import Counter

# 시각화 한글 폰트 깨짐 해결
import matplotlib.font_manager as fm
fname = os.path.join('C:\\Windows\\Fonts\\malgun.ttf')  # ttf 확장자 Path 입력
font = fm.FontProperties(fname=fname).get_name()
plt.rcParams["font.family"] = font

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

    # def show_tagset(self):
    #     show = Okt()
    #     print(show.tagset)

    # noinspection PyMethodMayBeStatic
    def extract_noun(self, raw_text):
        # 단어(명사)+빈도수 dataframe 반환 함수
        text = Okt()
        nouns = text.nouns(raw_text)  # 명사 추출, 형태 : ('단어', '빈도수')

        counter = Counter(nouns)
        word_with_freq_list = counter.most_common(15)  # 추출된 명사+빈도수 중 Top 15 추출

        # 출력할 테이블 생성
        words = [word_with_freq[0] for word_with_freq in word_with_freq_list]  # 명사만 추출
        freq = [word_with_freq[1] for word_with_freq in word_with_freq_list]  # 빈도수만 추출
        result_table = pd.DataFrame(columns=['단어', '빈도수'])
        result_table['단어'] = words
        result_table['빈도수'] = freq

        return result_table

    # noinspection PyMethodMayBeStatic
    def extract_adjective(self, raw_text):
        # 단어(형용사)+빈도수 dataframe 반환 함수
        text = Okt()
        pos_list = text.pos(raw_text)
        adjectives = [pos[0] for pos in pos_list if pos[1] == 'Adjective']

        counter = Counter(adjectives)
        word_with_freq_list = counter.most_common(15)  # 추출된 형용사+빈도수 중 Top 15 추출

        # 출력할 테이블 생성
        words = [word_with_freq[0] for word_with_freq in word_with_freq_list]  # 명사만 추출
        freq = [word_with_freq[1] for word_with_freq in word_with_freq_list]  # 빈도수만 추출
        result_table = pd.DataFrame(columns=['단어', '빈도수'])
        result_table['단어'] = words
        result_table['빈도수'] = freq

        return result_table

    # noinspection PyMethodMayBeStatic
    def extract_adverb(raw_text):
        # 단어(부사)+빈도수 dataframe 반환 함수
        text = Okt()
        pos_list = text.pos(raw_text)
        adverbs = [pos[0] for pos in pos_list if pos[1] == 'Adverb']

        counter = Counter(adverbs)
        word_with_freq_list = counter.most_common(15)  # 추출된 형용사+빈도수 중 Top 15 추출

        # 출력할 테이블 생성
        words = [word_with_freq[0] for word_with_freq in word_with_freq_list]  # 명사만 추출
        freq = [word_with_freq[1] for word_with_freq in word_with_freq_list]  # 빈도수만 추출
        result_table = pd.DataFrame(columns=['단어', '빈도수'])
        result_table['단어'] = words
        result_table['빈도수'] = freq

        return result_table

    # noinspection PyMethodMayBeStatic
    def show_barplot(self, noun_df, adjective_df, adverb_df):
        # 막대그래프 시각화
        plt.figure(figsize=(10, 6))

        # 명사 빈도수 시각화
        plt.subplot(311)
        sns.barplot(data=noun_df, x='단어', y='빈도수')
        plt.xlabel('단어')
        plt.ylabel('빈도수')
        plt.title('<명사 빈도수 그래프>')
        plt.tight_layout()

        # 형용사 빈도수 시각화
        plt.subplot(312)
        sns.barplot(data=adjective_df, x='단어', y='빈도수')
        plt.xlabel('단어')
        plt.ylabel('빈도수')
        plt.title('<형용사 빈도수 그래프>')
        plt.tight_layout()

        # 부사 빈도수 시각화
        plt.subplot(313)
        sns.barplot(data=adverb_df, x='단어', y='빈도수')
        plt.xlabel('단어')
        plt.ylabel('빈도수')
        plt.title('<부사 빈도수 그래프>')
        plt.tight_layout()

        plt.show()

    # noinspection PyMethodMayBeStatic
    def show_wordcloud(self, df):
        pass


def get_file_list(path):
    files = glob.glob(path + '/*')
    if files:
        files = [file for file in files if file.endswith('.docx') or file.endswith('.hwp')]
        return files


def main():
    file_list = get_file_list(DEFALUT_DICECTORY_PATH)

    if len(file_list) == 0:
        print('키워드 분석 가능한 파일이 폴더내 존재하지 않습니다.')

    while len(file_list) > 0:
        print('==분석 가능한 파일 목록==')
        for index, path_of_file in enumerate(file_list):
            print('{0}. {1}'.format(index+1, os.path.basename(path_of_file)))
        print('\n')

        try:
            input_num = int(input('키워드 분석하고자 하는 파일 번호 입력 >> '))
            print('키워드 분석 중...')
            print('\n')
            if input_num > len(file_list):  # 입력이 file_list 범위를 넘는 경우
                print('Warning : 파일 번호 범주내 정수 입력 필요.')
            elif 1 <= input_num <= len(file_list):  # 원하는 하나의 파일을 분석하는 경우
                analysis_target_file = File(file_list[input_num-1])  # 입력한 숫자는 리스트 요소의 인덱스보다 +1이므로 -1 해줌
                result = analysis_target_file.read_text()
                # print(result)

                nouns_df = analysis_target_file.extract_noun(result)
                adjective_df = analysis_target_file.extract_adjective(result)
                adverb_df = analysis_target_file.extract_adverb(result)
                print('====명사 분석 결과====')
                print(nouns_df)
                print('\n')

                print('====형용사 분석 결과====')
                print(adjective_df)
                print('\n')

                print('====부사 분석 결과====')
                print(adverb_df)

                analysis_target_file.show_barplot(nouns_df, adjective_df, adverb_df)
                print('\n')
                print('키워드 분석 완료!!!')
                print('\n')
        except ValueError:
            print('Warning : 정수 입력 필요.\n')


if __name__ == '__main__':
    main()