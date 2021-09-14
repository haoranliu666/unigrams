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

punctuations = ["\n", "\t", ".", ",", ";", "(", ")", ":", "—", "[", "]", "/", "–", "-"]
numbers = [" 0 ", " 1 ", " 2 ", " 3 ", " 4 ", " 5 ", " 6 ", " 7 ", " 8 ", " 9 "]
letters = [" a ", " b ", " c ", " d ", " e ", " f ", " g ", " h ", " i ", " j ", " k ", " l ", " m ", " n ", " o ", " p ", " q ", " r ", " s ", " t ", " u ", " v ", " w ", " x ", " y ", " z "]
articles = [" the ", " a ", " an "]
#https://www.talkenglish.com/vocabulary/top-50-prepositions.aspx
prepositions = [" of ", " with ", " at ", " from ", " into ", " during ", " including ", " until ", " against ", " among ", " throughout ", " despite ", " towards ", " upon ", " concerning ", " to ", " in ", " for ", " on ", " by ", " about ", " like ", " through ", " over ", " before ", " between ", " after ", " since ", " without ", " under ", " within ", " along ", " following ", " across ", " behind ", " beyond ", " plus ", " except ", " but ", " up ", " out ", " around ", " down ", " off ", " above ", " near "]
#https://www.myenglishpages.com/english/grammar-lesson-auxiliary-verbs.php
auxiliaryverbs = [" be ", " am ", " are ", " is ", " was ", " were ", " being ", " can ", " could ", " do ", " did ", " does ", " doing ", " have ", " had ", " has ", " having ", " may ", " might ", " must ", " shall ", " should ", " will ", " would "]
conjunctions = [" for ", " and ", " nor ", " but ", " or ", " yet ", " so "]
#https://www.thefreedictionary.com/List-of-pronouns.htm
pronouns = [" all ", " another ", " any ", " anybody ", " anyone ", " anything ", " as ", " aught ", " both ", " each ", " each other ", " either ", " enough ", " everybody ", " everyone ", " everything ", " few ", " he ", " her ", " hers ", " herself ", " him ", " himself ", " his ", " i ", " idem ", " it ", " its ", " itself ", " many ", " me ", " mine ", " most ", " my ", " myself ", " naught ", " neither ", " no one ", " nobody ", " none ", " nothing ", " nought ", " one ", " one another ", " other ", " others ", " ought ", " our ", " ours ", " ourself ", " ourselves ", " several ", " she ", " some ", " somebody ", " someone ", " something ", " somewhat ", " such ", " suchlike ", " that ", " thee ", " their ", " theirs ", " theirself ", " theirselves ", " them ", " themself ", " themselves ", " there ", " these ", " they ", " thine ", " this ", " those ", " thou ", " thy ", " thyself ", " us ", " we ", " what ", " whatever ", " whatnot ", " whatsoever ", " whence ", " where ", " whereby ", " wherefrom ", " wherein ", " whereinto ", " whereof ", " whereon ", " wherever ", " wheresoever ", " whereto ", " whereunto ", " wherewith ", " wherewithal ", " whether ", " which ", " whichever ", " whichsoever ", " who ", " whoever ", " whom ", " whomever ", " whomso ", " whomsoever ", " whose ", " whosever ", " whosesoever ", " whoso ", " whosoever ", " ye ", " yon ", " yonder ", " you ", " your ", " yours ", " yourself ", " yourselves "]
years = ["1990", "1991", "1992", "1993", "1994", "1995", "1996", "1997", "1998", "1999", "2000", "2001", "2002", "2003", "2004", "2005", "2006", "2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020"]
months = [" jan ", " feb ", " mar ", " apr ", " may ", " june ", " july ", " aug ", " sep ", " oct ", " nov ", " dec "]

all_content = all_content.lower()
all_delete = punctuations + numbers + letters + articles + prepositions + auxiliaryverbs + conjunctions + pronouns + years + months
for temp in all_delete:
    all_content = all_content.replace(temp, " ")

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
temp = obj.most_common(500)
#word
for i in range(0, len(temp)):
    print(temp[i][0])
#count
for i in range(0, len(temp)):
    print(temp[i][1])

