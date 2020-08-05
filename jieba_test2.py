import matplotlib.pyplot as plt
from imageio import imread
from wordcloud import WordCloud,STOPWORDS
import xlrd
import pandas 
import jieba
import jieba.analyse
import numpy as np
from sklearn.feature_extraction.text import  TfidfVectorizer
from sklearn.cluster import KMeans


def jieba_tokenize(text):
    return jieba.lcut(text)
    
text_list = []
tfidf_vectorizer = TfidfVectorizer(tokenizer=jieba_tokenize, lowercase=False)
 
def load_data():

    jieba.load_userdict("D:\\dict.txt")
    data = pandas.read_excel("D:\\file_20200804_152851.xls", usecols=[2])
    data_desc = data[["APPLY_DESC"]].copy()

    for index, row in data_desc.iterrows():
        content = row.APPLY_DESC
        try:
        #TextRank 关键词抽取，只获取固定词性
            words = jieba.analyse.textrank(content, topK=50, withWeight=False, allowPOS=('ns', 'n', 'vn', 'v'))
            text_list.append(' '.join(words))
        except AttributeError:
            pass

    file=open('E:\\raw_data.txt','w')
    file.write(str(text_list))
    file.close()

def k_mean():
    tfidf_matrix = tfidf_vectorizer.fit_transform(text_list)
 
    num_clusters = 100
    km_cluster = KMeans(n_clusters=num_clusters, max_iter=300, n_init=1, \
                    init='k-means++',n_jobs=1)

    result = km_cluster.fit_predict(tfidf_matrix)
 
    print("Predicting result: ", result)

if __name__ == '__main__': 
    load_data()
    # k_mean()
