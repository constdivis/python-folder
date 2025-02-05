import os

from nltk.tokenize import RegexpTokenizer
from nltk.tokenize import sent_tokenize
import pymorphy3

import pandas as pd

tokenizer = RegexpTokenizer(r'\w+')
morph = pymorphy3.MorphAnalyzer()

def get_pos_tags(txt, language='russian'):
    sents = sent_tokenize(txt, language=language)
    morph_rslts = []
    lemms_rslts = []
    tags_result = []
    word_tokens = []
    for s in sents:
        words = tokenizer.tokenize(s)
        morph_rslt = [morph.parse(w) for w in words]
        lemms = [w[0].normal_form for w in morph_rslt]
        w_tokens = [w[0].word for w in morph_rslt]
        tags = [[i.tag._grammemes_tuple for i in w] for w in morph_rslt]
        morph_rslts.append(morph_rslt)
        lemms_rslts.append(lemms)
        tags_result.append(tags)
        word_tokens.append(w_tokens)
    dct = {
        'sents' : sents,
        'tokens' : word_tokens,
        'lemms' : lemms_rslts,
        'tags' : tags_result,
    }
    return dct

def get_file_names(dir_name):
    f_data = []
    for r, dirs, files in os.walk(dir_name):
        for f in files:
            f_data.append((os.path.splitext(f)[0], os.path.join(r, f)))
    return f_data

def get_freq_table(txt_dir='txt',
                  pos = ['VERB', 'PRTF', 'NOUN', 'ADJF']):
    file_data = get_file_names(txt_dir)

    pos_dct_lst = []
    for d in file_data:
        with open(d[1], 'r') as f:
            txt = f.readlines()
        txt = ' '.join(txt)
        pos_dct_lst.append(get_pos_tags(txt))

    pos_data = []
    for i in pos_dct_lst:
        tags = i['tags']
        lemms = i['lemms']
        tokens = i['tokens']
        tok_lem_pos = []
        for n, s in enumerate(tags):
            for m, w in enumerate(s):
                for p in pos:
                    if p in w[0]:
                        tok_lem_pos.append((tokens[n][m], lemms[n][m], p))
        pos_data.append(tok_lem_pos)

    dfs = []
    for d in pos_data:
        df = pd.DataFrame([[i[1], i[2]] for i in d])
        df.columns = ['lemma', 'pos']
        df = df.groupby(by=['lemma', 'pos']).size().reset_index()
        df.columns.values[2] = 'freq'
        df.sort_values(['freq', 'lemma'], ascending=[False, True], inplace=True)
        dfs.append(df)

    with pd.ExcelWriter('pos_freq.xlsx') as writer:
        for n, i in enumerate(dfs):
            i.to_excel(writer, sheet_name=f'{file_data[n][0]}', index=False)
    print('done')
