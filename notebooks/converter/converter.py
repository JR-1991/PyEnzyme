'''
Created on 21.03.2021

@author: HD
'''

import pandas as pd
import numpy as np
import json
import datetime

class ConverterExcelJSON():

    def toJSON(self, path, out=None):
        '''
        TODO DOKU
        '''
        self.path = path
        self.outJSON : dict[str, object] = {}
        self.__parseXLSX()
        print(json.dumps(self.outJSON, indent=4))
        if not(out == None):
            with open(out+'json.json','w', encoding='utf-8') as f:
                json.dump(self.outJSON, f, ensure_ascii=False, indent=4)


    def __parseXLSX(self):
        '''
        TODO DOKU
        '''
        doc = pd.read_excel(self.path, sheet_name=['General Information'])
        gi = doc['General Information']

        self.__getGI(gi)

    def __getGI(self, gi):
        '''
        TODO DOKU
        Open question: How should we handle NaN?
        '''
        values = self.__filter(gi, ['Title', 'Date of creation', 'DOI (optional)', 'PubMedID (optional)', 'URL (optional)'], x=1)
        self.outJSON['name'] = values['Title']
        ######## key 'date' guessed
        self.outJSON['date'] = values['Date of creation'].strftime('%d.%m.%Y')
        ######## key 'doi' guessed
        self.outJSON['doi'] = values['DOI (optional)']
        ######## key 'pubmed' guessed
        self.outJSON['pubmed'] = values['PubMedID (optional)']
        ######## key 'url' guessed
        self.outJSON['url'] = values['URL (optional)']
        

    def __filter(self, df, conds: list[str], y: int =0, x: int =0, isSingleValue: bool=True):
        '''
        TODO DOKU
        '''
        values = dict()
        for cond in conds:
            coords = self.__getCoords(df, cond)
            values[cond] = df.iloc[coords[0]+y, coords[1]+x]
        return values

    def __getCoords(self, df, cond):
        '''
        TODO DOKU
        '''
        return[(x, y) for x, y in zip(*np.where(df.values == cond))][0]

if __name__ == '__main__':
    path = '../datasets/EnzymeML_Template_Example.xlsx'
    ConverterExcelJSON().toJSON(path, out = '../datasets/')