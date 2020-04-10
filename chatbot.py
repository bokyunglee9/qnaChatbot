import openpyxl, nltk, string, math
from nltk import word_tokenize, pos_tag
from nltk.corpus import stopwords, wordnet
from nltk.stem.wordnet import WordNetLemmatizer
from textblob import TextBlob
from collections import Counter
from functions import *


endProgram = False
while not endProgram:
    category = opening()
    if category == 'Q' :
        endProgram = True
    else:
        main(category)
