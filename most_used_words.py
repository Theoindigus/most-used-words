# Copyrigth © 2023 Егоров Антон (Egorov Anton) (Theoindigus) <indigus@aegor.ru>
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation 
# files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, 
# modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Soft-
# ware is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRAN-
# TIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT 
# HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

import docx
import sys
import re
import os.path
from progress.bar import IncrementalBar

# Подготовка программы
file_not_exists = True
file_name = ''
file_path = ''
while(file_not_exists):
	print('Enter the file name. The file should be in the same directory with the script')
	file_name = input()
	file_path = sys.path[0] + '\\' + file_name + '.docx'
	if not os.path.exists(file_path):
		print("There is no such file")
	else:
		file_not_exists = False

# Создание паттерна по которому из текста будут выделяться слова вместе со знаками пунктуации
base_pattern = r'\s*([-,.!?:;\'"/\\0-9A-Za-zА-Яа-яЁё]+)\s*'
# Создание паттерна по которому из текста будут выделяться слова без знаков пунктуации
only_liters_pattern = r'\s*([-\'"A-Za-zА-Яа-яЁё]+)\s*'

# Для облегчения обработки файла стоит исключить из него базовый набор
# общеупотребимых слов, а также отдельные буквы алфавита 
exclude_words_set = set()
if not os.path.exists(sys.path[0] + '\\exclude_words.docx'):
	print('The file with the list of excluded words is missing in the script directory')
else:
	for p in docx.Document(sys.path[0] + '\\exclude_words.docx').paragraphs:
		exclude_words_set.add(p.text)


# Получение исходного документа
# Извлечение из него параграфов текста
# Создание прогресс-бара
origin_doc = docx.Document(file_path)
paragraphs = origin_doc.paragraphs
bar = IncrementalBar('Paragraphs processing:         ', max = len(paragraphs))

# Создание словаря для хранения всех слов без повторов в качестве ключа
# и количества их повторов в качестве значения
# Создание документа, в который будет записан отделённый от всего документа
# текст без форматирования, то есть чистый текст
clear_set = dict()
clear_doc = docx.Document()

for p in paragraphs:
	match = re.findall(base_pattern, p.text)
	if match:
		clear_text = '\t'
		for m in match:
			if m:
				clear_text += m + ' '
				tmp = re.match(only_liters_pattern, m)
				if tmp:
					tmp = tmp.group(0)
					if tmp not in exclude_words_set:
						if tmp not in clear_set:
							clear_set[tmp.lower()] = 1
						else:
							clear_set[tmp.lower()] = clear_set[tmp.lower()] + 1
		clear_doc.add_paragraph(clear_text)
	bar.next()
bar.finish()

max_search_words = int(input('Enter the maximum number of most used words to be found:\n'))
if max_search_words > len(clear_set):
	max_search_words = len(clear_set)
bar = IncrementalBar('Search for the most used words:', max = max_search_words)
most_used_words = list()
for i in range(0, max_search_words):
	max_used_count = 0
	max_used_word = ''
	for e in clear_set:
		if clear_set[e] > max_used_count:
			max_used_word = e
			max_used_count = clear_set[e]
	most_used_words.append(max_used_word)
	del clear_set[max_used_word]
	bar.next()
most_used_words = sorted(most_used_words)
bar.finish()

bar = IncrementalBar('Files saving:                  ', max = len(most_used_words) + 1)
clear_doc.save(sys.path[0] + '\\' + file_name + '_clear_text.docx')
bar.next()
most_used_words_doc = docx.Document()
for w in most_used_words:
	most_used_words_doc.add_paragraph(w)
	bar.next()
most_used_words_doc.save(sys.path[0] + '\\' + file_name + '_most_used_words.docx')
bar.finish()