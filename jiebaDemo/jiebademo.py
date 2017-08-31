#!/usr/bin/env python3
#encoding=utf-8

import docx
import jieba
import jieba.posseg as pseg
import matplotlib.pyplot as plt
from wordcloud import WordCloud
import PIL
import numpy as np
from operator import itemgetter

class DocxFreqWord(object):

	def __init__(self, docurl):
		self.docurl = docurl

	##
	# 解析docx文档并过滤掉无用的词
	# @return 词数组 e.g. ['文档', '书房', '文档', '案例', '梅花', '文档']
	#
	def wordsFromDocx(self):
		document = docx.Document(self.docurl)
		docText = ''
		for paragraph in document.paragraphs:
			docText += paragraph.text

		# 语义化解析文字
		words = list(pseg.cut(docText))
		
		filted_arr = []
		filter_key = ['x', 'uj', 'c', 'a', 'p', 'ul', 'c', 'f', 'p', 'm', 'r', 'd', 'u']
		for word in words:
			# word: pair('Hadoop', 'eng') --> 'hadoop': word.word ; 'eng' : word.flag
			if word.flag in filter_key or len(word.word) < 2:
				pass
			else:
				filted_arr.append(word.word)
		return filted_arr

	##
	# 桶排序得到词频向量
	# @params: words_arr 通过wordsFromDocx得到的词数组
	# @return: 词频向量 {'文档': 3, '书房': 1, '案例': 1, '梅花': 1}
	#
	def getWordFreq(self, words_arr):
		d = dict()
		for word in words_arr:
			if d.get(word) is None:
				d[word] = 1
			else:
				d[word] = d[word] + 1
		return d

	
if __name__ == '__main__':
	docxFreqWord = DocxFreqWord('research_report.docx')
	wordarr = docxFreqWord.wordsFromDocx()
	wordfreq = docxFreqWord.getWordFreq(wordarr)

	fontpath = r'wqy-zenhei.ttc'
	alice_mask = np.array(PIL.Image.open('bigdata.jpg'))
	my_wordcloud = WordCloud(font_path = fontpath, max_words = 2000, mask = alice_mask, 
		max_font_size = 40, width = 1200, height = 800, background_color = "white").generate_from_frequencies(wordfreq)
	my_wordcloud.to_file("re.jpg")
	plt.imshow(my_wordcloud)
	plt.axis("off")
	plt.show()