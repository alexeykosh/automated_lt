#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# author: Alexey Koshevoy

import pandas as pd
from lingcorpora import Corpus
from pymorphy2 import MorphAnalyzer
import warnings

warnings.filterwarnings('error')

parser = MorphAnalyzer()
rus_corp = Corpus('rus')
# rus_results = rus_corp.search('узкое окно')


class Completer:

    def __init__(self, path, sheet):
        self.path = path
        self.sheet = sheet
        self.words = []
        self.data, self.columns, self.index = self._opener()

    def _opener(self):
        data = pd.read_excel(self.path, sheet_name=self.sheet).fillna(0)
        return data, data.columns, data.index

    def _counter(self):
        data, columns, indexs = self._opener()
        for column in columns:
            for index in indexs:
                word = ''.join(c for c in index if c not in '(){}<>')
                gender, number = parser.parse(word.split(' ', 1)[0])[0].tag.gender, \
                                  parser.parse(word.split(' ', 1)[0])[0].tag.number
                if number == 'plur':
                    try:
                        self.words.append(
                            [str(parser.parse(column)[0].inflect({number}).
                                 word) + ' ' + str(word), [column, index]])
                    except AttributeError:
                        pass
                try:
                    self.words.append([str(parser.parse(column)[0].inflect({gender,
                                                                       number}).
                                      word) + ' ' + str(word), [column, index]])
                except AttributeError:
                    pass

    def _searcher(self):
        for element in self.words:
            try:
                rus_corp.search(element[0])
                element.append(1)
            except UserWarning:
                element.append(0)

    def _writer(self):
        for element in self.words:
            self.data.at[element[1][1], element[1][0]] = element[2]
        print(self.data.head(10))

        self.data.to_csv('try_table.csv')


if __name__ == '__main__':
    a = Completer(path='/Users/alexey/Documents/Python/automated_lt/questionnai'
                       're_size.xlsx', sheet='русский_стандртный_вид')
    a._counter()
    a._searcher()
    print(a.words)
    a._writer()
