import matplotlib.pyplot as plt
from imageio import imread
from wordcloud import WordCloud,STOPWORDS
import xlrd
import pandas 
import jieba
import jieba.analyse
import numpy as np
 
def create_word_cloud(word_dict, destination_file):

    wc = WordCloud(
        font_path=r'C:\\Windows\\Fonts\\simsun.ttc',
        collocations=False, # 去重
        stopwords=STOPWORDS,
        max_words=100,
        width=2000,
        height=1200,
        background_color='white',
        #mask=back_coloring
    )
    
    wc.generate_from_frequencies(word_dict)
    wc.to_file(destination_file)

    plt.imshow(wc)
    plt.axis("off")
    plt.show()

def generate_word_dict_from_excel(filepath, index_sheet):

    data_tmp = {}
    with open('D:\\keywords1.txt', 'r', encoding='utf-8') as f:
        line = f.readline().strip()

        while line:
            values = str.split(line, ',')
            count = int(values[1])
            if data_tmp.__contains__(values[0]):
                data_tmp[values[0]] =  data_tmp[values[0]] + count
            else:
                data_tmp[values[0]] = count
            line = f.readline().strip()

    return data_tmp

def generate_word_frequency_from_excel(dict_file, data_source_file, output_file):
    jieba.load_userdict(dict_file)
    data = pandas.read_excel(data_source_file, usecols=[2])
    data_desc = data[["APPLY_DESC"]].copy()

    segments = []
    for index, row in data_desc.iterrows():
        content = row.APPLY_DESC
        try:
            #TextRank 关键词抽取，只获取固定词性
            words = jieba.analyse.textrank(content, topK=50,withWeight=False,allowPOS=('ns', 'n', 'vn', 'v'))
            splitedStr = ''
            for word in words:
            # 记录全局分词
                segments.append({'word':word, 'count':1})
                splitedStr += word + ' '
        except AttributeError:
            pass
    dfSg = pandas.DataFrame(segments)

    # 词频统计
    dfWord = dfSg.groupby('word')['count'].sum()

    #导出csv
    dfWord.to_csv(output_file,encoding='gbk')

if __name__ == '__main__': 
    output_file = "E:\\keywords.csv"
    generate_word_frequency_from_excel("D:\\dict.txt", "D:\\file_20200804_152851.xls", output_file)
    word_dict = generate_word_dict_from_excel(output_file, 0)
    create_word_cloud(word_dict, "E:\\wordcloud_test.png")