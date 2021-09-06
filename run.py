from docx import Document
from collections import Counter
import os

path = 'xxx'
file_list = os.listdir(path)

all_content = ""
for file in file_list:
    if file != '.DS_Store':
        document = Document(path + '/' + file)
        for paragraph in document.paragraphs:
            all_content = all_content + paragraph.text
            
all_content = all_content.lower()
all_content = all_content.replace(".", " ")
all_content = all_content.replace(",", " ")
all_content = all_content.replace("(", " ")
all_content = all_content.replace(")", " ")

all_word_list = all_content.split()

all_bigram_list = []
for i in range(0, len(all_word_list)-1):
    all_bigram_list.append(all_word_list[i] + ' ' + all_word_list[i+1])

obj = Counter(all_word_list)
temp = obj.most_common(1000)
#word
for i in range(0, len(temp)):
    print(temp[i][0])
#count
for i in range(0, len(temp)):
    print(temp[i][1])

obj = Counter(all_bigram_list)
temp = obj.most_common(1000)
#word
for i in range(0, len(temp)):
    print(temp[i][0])
#count
for i in range(0, len(temp)):
    print(temp[i][1])
