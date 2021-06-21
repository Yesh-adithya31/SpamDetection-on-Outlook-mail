# -*- coding: utf-8 -*-
"""
Created on Mon Jun 21 18:08:38 2021

@author: Yesh Adithya
"""


import win32com.client as client

outlook = client.Dispatch('Outlook.Application')

namespace = outlook.GetNameSpace('MAPI')

account = namespace.Folders['codse182f-071@student.nibm.lk']

inbox = account.Folders['Inbox']

messages = inbox.Items

message = messages.GetLast()

body_content = message.Body

print(body_content)



import matplotlib.pyplot as plt
import nltk 
from nltk.corpus import stopwords
import numpy as np
import pandas as pd 
import seaborn as sns
import string

from sklearn.feature_extraction.text import CountVectorizer
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import MultinomialNB
from sklearn.utils.multiclass import unique_labels
from sklearn.metrics import accuracy_score


nltk.download('stopwords')

def process_text(body_content):
    
    nopunc = [char for char in body_content if char not in string.punctuation]
    nopunc = ''.join(nopunc)

    clean_words = [word for word in nopunc.split() if word.lower() not in stopwords.words('english')]

    return clean_words






########################### 4th ##############################

df4 = pd.read_csv('spam.csv')

df4 = df4 [["Label","EmailText"]]

print(df4)

df4['Label'] = np.where(df4['Label']=='spam','spam', 'ham')

df4['EmailText'].head().apply(process_text)

X_train4, X_test4, Y_train4, Y_test4 = train_test_split(df4['EmailText'], df4['Label'], random_state=0)

vectorizer4 = CountVectorizer(ngram_range=(1, 2)).fit(X_train4)
X_train_vectorized4 = vectorizer4.transform(X_train4)
X_train_vectorized4.toarray().shape


modelf = MultinomialNB(alpha=0.1)
modelf.fit(X_train_vectorized4, Y_train4)

###########################  ##############################

pred = modelf.predict(vectorizer4.transform(
    [
       body_content
    ])
            ) 
   

print(pred)

# AFTER Prediction Go  to this path
if pred == 'ham':
    print("Not a Spam Mail")
else :
    print("It is a Spam Mail")