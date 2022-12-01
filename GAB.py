import pandas as pd
import tensorflow as tf
from tensorflow.keras.preprocessing.sequence import pad_sequences
from tensorflow.keras.preprocessing.text import Tokenizer
from tensorflow.keras.layers import Embedding, LSTM, Dense, Dropout, Bidirectional
from tensorflow.keras.models import Sequential
from tensorflow.keras.optimizers import Adam
from livelossplot import PlotLossesKeras
from sklearn.metrics import accuracy_score
import numpy as np
import seaborn as sns
from matplotlib import pyplot as plt
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from sklearn.model_selection import train_test_split
from zipfile import ZipFile
from spellchecker import SpellChecker
import json
import os
import re
import io

Reviews = pd.read_excel('C:\\Users\\George.Burns\\Documents\\Projects\\WWT Customer Reviews Data.xlsx')
Reviews = Reviews[['Rating','Review']]
print(Reviews)

rating_counts = Reviews['Rating'].value_counts()
df_counts = pd.DataFrame({'rating': rating_counts.index, 'count': rating_counts.values})

sns.set_theme(style="darkgrid")
ax = sns.barplot(x="rating", y="count", data=df_counts)
plt.show()

Reviews['Review'] = Reviews['Review'].apply(lambda x: re.sub('\d+', '', x))

nltk.download("stopwords")
nltk.download('punkt')
stop_words = set(stopwords.words('english'))

def remove_stop_words(sentence):
  tokens = word_tokenize(sentence)
  tokens = [tok for tok in tokens if not tok in stop_words]
  return ' '.join(tokens)

Reviews['Review'] = Reviews['Review'].apply(lambda x: remove_stop_words(x))

texts = Reviews['Review'].values
ratings = Reviews['Rating'].astype(int).values

x_train, x_test, y_train, y_test = train_test_split(texts, ratings, test_size=0.05, shuffle=True)

tokenizer = Tokenizer()
tokenizer.fit_on_texts(texts)
total_words = len(tokenizer.word_index) + 1

train_sequences = tokenizer.texts_to_sequences(x_train)
test_sequences = tokenizer.texts_to_sequences(x_test)

word_index = tokenizer.word_index
max_sequence_len = max([len(x) for x in train_sequences])

train_sequences = np.array(pad_sequences(train_sequences, maxlen=max_sequence_len, padding='post'))
test_sequences = np.array(pad_sequences(test_sequences, maxlen=max_sequence_len, padding='post'))

embedding_dim = 100

base_path = 'C:\\Users\\George.Burns\\Documents\\Projects\\'
embeddings_index = {};
with ZipFile(os.path.join(base_path, 'glovetwitter100d.zip'), 'r') as zip:
  with io.TextIOWrapper(zip.open(f"glove.twitter.27B.{embedding_dim}d.txt"), encoding="utf-8") as f:
    for line in f:
        values = line.split();
        word = values[0];
        coefs = np.asarray(values[1:], dtype='float32');
        embeddings_index[word] = coefs;

oov_words = []

embeddings_matrix = np.zeros((total_words, embedding_dim));
for word, i in word_index.items():
  embedding_vector = embeddings_index.get(word);
  if embedding_vector is not None:
    embeddings_matrix[i] = embedding_vector;
  else :
    oov_words.append(word)

oov_words = list(set(oov_words))

def get_autocorrections(oov_words):
  spell = SpellChecker()
  auto_corrections = {}

  for word in oov_words:
      # Get the one `most likely` answer
      auto_corrections[word] = spell.correction(word)
      if len(auto_corrections) % 100 == 0:
        print(f'corrected: {len(auto_corrections)} words from a total of {len(oov_words)}')
  
  with open(os.path.join(base_path, 'oov_words.json'), 'w') as json_file:
    json.dump(auto_corrections, json_file)

def load_autocorrections_from_json():
  with open(os.path.join(base_path, 'oov_words.json')) as json_file:
    return json.load(json_file)

#auto_corrections = get_autocorrections(oov_words)

# for oov_word, correction in auto_corrections.items():
#   embedding_vector = embeddings_index.get(correction);
# if embedding_vector is not None:
#     embeddings_matrix[word_index[oov_word]] = embedding_vector
    
def build_model(total_words, embedding_dim, input_length, embeddings_matrix):
  model = Sequential([
  tf.keras.layers.Embedding(total_words, embedding_dim, input_length=input_length, weights=[embeddings_matrix], trainable=False),
  tf.keras.layers.Bidirectional(tf.keras.layers.LSTM(64)),
  tf.keras.layers.Dropout(0.5),
  tf.keras.layers.Dense(128, activation='relu'),
  tf.keras.layers.Dropout(0.5),
  tf.keras.layers.Dense(64, activation='relu'),
  tf.keras.layers.Dropout(0.5),
  tf.keras.layers.Dense(1),
])
  model.compile(
    loss='mse', 
    optimizer=tf.keras.optimizers.Adam(), 
    metrics=['mae'])
  return model

model = build_model(total_words, embedding_dim, max_sequence_len, embeddings_matrix)

print(model.summary())

history = model.fit(
    x=train_sequences, 
    y=y_train, 
    validation_split=0.5, 
    epochs=50, 
    #callbacks=[PlotLossesKeras()],
    shuffle=True)

yhat = model(test_sequences)
preds = np.round(yhat).astype(int)

df_preds = pd.DataFrame({
    'y_true': y_test.astype(int),
    'y_pred': preds[:,0], 
    'y_hat': yhat[:,0],
    'mae': np.abs(y_test - yhat[:,0]),
    'review': x_test
    })
print('total predictions', len(df_preds))
df_errors = df_preds[df_preds['y_true'] != df_preds['y_pred']]
print('total errors', len(df_errors))
print('accuracy score', accuracy_score(df_preds['y_true'], df_preds['y_pred']))