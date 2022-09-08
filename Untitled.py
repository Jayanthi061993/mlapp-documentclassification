#!/usr/bin/env python
# coding: utf-8

# In[1]:


from subprocess import Popen, PIPE
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams 
from pdfminer.pdfpage import PDFPage
import io
from io import StringIO
import os
import glob
import comtypes.client
import sys
import string
from nltk.corpus import stopwords
import matplotlib.pyplot as plt
from nltk.stem import WordNetLemmatizer
import warnings
warnings.filterwarnings('ignore')

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import json
import nltk
import re
import csv


# In[5]:


wdFormatPDF = 17 # selecting the PDF format 

filelocation = 'D:/JAAM_PROJECTS/Document Classification Project POC Deployment/Raw Data/Taxes'

filelist=os.listdir(filelocation)
doccollection=[]
for files in filelist: #Traversing through all the files in the location to find the doc files
    files=os.path.join(filelocation,files)
    doccollection.append(files)
for x in doccollection:
    if x.endswith('.doc'):
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(os.path.abspath(x))
        doc.SaveAs(os.path.abspath(x+'.pdf'), FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()


# In[3]:


import pandas as pd
def convert_pdf_to_txt(path):
    #alltexts = []
    filelist=os.listdir(path)
    documentcollection=[]
    for files in filelist:
        files=os.path.join(path,files)
        documentcollection.append(files)
    for ifiles in documentcollection:
        if ifiles.endswith('.pdf') or ifiles.endswith('.PDF'): #different extensions on the raw data
            with open(ifiles, 'rb') as fh:
                for page in PDFPage.get_pages(fh, 
                                              caching=True,
                                              check_extractable=True):
                    resource_manager = PDFResourceManager()
                    fake_file_handle = io.StringIO()
                    converter = TextConverter(resource_manager, fake_file_handle)
                    page_interpreter = PDFPageInterpreter(resource_manager, converter)
                    page_interpreter.process_page(page)
 
                    text = fake_file_handle.getvalue() # extraction of the text data
                    yield text
 
                    # closing open handles 
                    converter.close()
                    fake_file_handle.close()
        
    #return alltexts


# In[6]:


filepath='D:/JAAM_PROJECTS/Document Classification Project POC Deployment/Raw Data/Taxes'
textcontents = convert_pdf_to_txt(filepath)
dftaxes = pd.DataFrame(textcontents, columns = ['Text_Data']) 
dftaxes['Category'] = 'Taxes' # Adding the  label


# In[7]:


print(dftaxes)


# In[8]:


filepath='D:/JAAM_PROJECTS/Document Classification Project POC Deployment/Raw Data/Human Resources'
hrcontents = convert_pdf_to_txt(filepath)
dfhr = pd.DataFrame(hrcontents, columns = ['Text_Data']) 
dfhr['Category'] = 'Human Resouces'


# In[9]:


print(dfhr)


# In[10]:


filepath='D:/JAAM_PROJECTS/Document Classification Project POC Deployment/Raw Data/Agreements'
agreementcontents = convert_pdf_to_txt(filepath)
dfagreement = pd.DataFrame(agreementcontents, columns = ['Text_Data']) 
dfagreement['Category'] = 'Agreement'


# In[11]:


print(dfagreement)


# In[12]:


frames = [dftaxes, dfhr, dfagreement]
finalframe = pd.concat(frames,sort=False)
finalframe = finalframe[['Text_Data','Category']]
finalframe = finalframe.reset_index(drop=True)
finalframe[:]


# In[13]:


# Create a new column 'category_id' with encoded categories 
finalframe['Category_Id'] = finalframe['Category'].factorize()[0]
category_id_df = finalframe[['Category', 'Category_Id']]


# In[14]:


finalframe


# In[15]:


# Dictionaries for future use
category_to_id = dict(category_id_df.values)
id_to_category = dict(category_id_df[['Category', 'Category_Id']].values)
# New dataframe
finalframe.head()


# In[16]:


fig = plt.figure(figsize=(8,6))
colors = ['grey','grey','grey','grey','grey','grey','grey','grey','grey', 'grey','darkblue','darkblue','darkblue']
finalframe.groupby('Category').Text_Data.count().sort_values().plot.barh(ylim=0, color=colors, title= 'GRAPH')
plt.xlabel('Number of ocurrences', fontsize = 10);


# In[17]:


from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.feature_extraction.text import TfidfTransformer
tfidf = TfidfVectorizer(sublinear_tf=True, min_df=5, ngram_range=(1, 2), stop_words='english')
# We transform each complaint into a vector
features = tfidf.fit_transform(finalframe.Text_Data).toarray()
labels = finalframe.Category_Id
print("Each of the %d Text Data is represented by %d features (TF-IDF score of unigrams and bigrams)" %(features.shape))


# In[18]:


print(features)
print(features.shape)


# In[19]:


print(labels)


# In[20]:


from sklearn.feature_selection import chi2
import numpy as np
# Finding the three most correlated terms with each of the product categories
N = 3
for Category, Category_Id in sorted(category_to_id.items()):
  features_chi2 = chi2(features, labels == Category_Id)
  indices = np.argsort(features_chi2[0])
  feature_names = np.array(tfidf.get_feature_names())[indices]
  unigrams = [v for v in feature_names if len(v.split(' ')) == 1]
  bigrams = [v for v in feature_names if len(v.split(' ')) == 2]
  print("n==> %s:" %(Category))
  print("  * Most Correlated Unigrams are: %s" %(', '.join(unigrams[-N:])))
  print("  * Most Correlated Bigrams are: %s" %(', '.join(bigrams[-N:])))


# In[21]:


from sklearn.model_selection import train_test_split
X = finalframe['Text_Data'] # Collection of documents
y = finalframe['Category'] # Target or the labels we want to predict (i.e., the 3 different complaints of products)
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.25, random_state = 0)


# In[22]:


from sklearn.naive_bayes import MultinomialNB
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier
from sklearn.svm import LinearSVC
from sklearn.model_selection import cross_val_score
from sklearn.metrics import confusion_matrix
from sklearn import metrics

models = [
    RandomForestClassifier(n_estimators=100, max_depth=5, random_state=0),
    LinearSVC(),
    MultinomialNB(),
    LogisticRegression(random_state=0),
]


# In[23]:


# 5 Cross-validation
CV = 5
cv_df = pd.DataFrame(index=range(CV * len(models)))


# In[24]:


entries = []
for model in models:
  model_name = model.__class__.__name__
  accuracies = cross_val_score(model, features, labels, scoring='accuracy', cv=CV)
  for fold_idx, accuracy in enumerate(accuracies):
    entries.append((model_name, fold_idx, accuracy))
cv_df = pd.DataFrame(entries, columns=['model_name', 'fold_idx', 'accuracy'])


# In[25]:


mean_accuracy = cv_df.groupby('model_name').accuracy.mean()
std_accuracy = cv_df.groupby('model_name').accuracy.std()

acc = pd.concat([mean_accuracy, std_accuracy], axis= 1, 
          ignore_index=True)
acc.columns = ['Mean Accuracy', 'Standard deviation']
acc


# In[26]:


plt.figure(figsize=(8,5))
sns.boxplot(x='model_name', y='accuracy', 
            data=cv_df, 
            color='lightblue', 
            showmeans=True)
plt.title("MEAN ACCURACY (cv = 5)n", size=14);


# In[27]:


X_train, X_test, y_train, y_test,indices_train,indices_test = train_test_split(features, labels, finalframe.index, test_size=0.25, random_state=1)
model = LinearSVC()
model.fit(X_train, y_train)
y_pred = model.predict(X_test)


# In[28]:


# Classification report
print('ttttCLASSIFICATIION METRICSn')
print(metrics.classification_report(y_test, y_pred, target_names= finalframe['Category'].unique()))


# In[29]:


X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.25, random_state = 0)


# In[30]:


tfidf = TfidfVectorizer(sublinear_tf=True, min_df=5, ngram_range=(1, 2), stop_words='english')


# In[31]:


fitted_vectorizer = tfidf.fit(X_train)
tfidf_vectorizer_vectors = fitted_vectorizer.transform(X_train)


# In[32]:


model = LinearSVC().fit(tfidf_vectorizer_vectors, y_train)


# In[33]:


demofile = 'D:/JAAM_PROJECTS/Document Classification Project POC Deployment/f1040.pdf'


# In[34]:


# Parsing through the sample document and extracting the textual data
def convert2txt():
    alltexts = []
    with open(demofile, 'rb') as fh:
        rsrcmgr = PDFResourceManager()
        retstr = StringIO()
        codec = 'utf-8'
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        fp = open(demofile, 'rb')
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        password = ""
        maxpages = 0
        caching = True
        pagenos=set()

        for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
            interpreter.process_page(page)

        text = retstr.getvalue()
        alltexts.append(text)
        fp.close()
        device.close()
        retstr.close()
        
    return alltexts 


# In[35]:


textdata = convert2txt()


# In[36]:


textdata


# In[38]:


print(model.predict(fitted_vectorizer.transform(textdata)))


# In[39]:


import pickle

pickle.dump(fitted_vectorizer,open('fitted_vectorizer.pkl', 'wb'))
pickle.dump(model, open('model.pkl', 'wb'))


# In[ ]:




