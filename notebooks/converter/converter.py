'''
Created on 21.03.2021

@author: HD
'''

import pandas as pd
import numpy as np
import json
import os


class ConverterExcelJSON():

    def toJSON(self, path, outpath=None, outname=None):
        '''
        TODO DOKU
        '''
        self.path = path
        self.outJSON: dict[str, object] = {}
        self.__parseXLSX()
        print(json.dumps(self.outJSON, indent=4))
        if not(outpath is None):
            assert(os.path.exists(outpath)), "Given outpath does not exist."
            if not(outname is None):
                if not(outname[-5:] == '.json'):
                    outname = outname + '.json'
                self.__writeJSON(os.path.join(outpath, outname))
            else:
                self.__writeJSON(os.path.join(outpath, self.outJSON['name'] + '.json'))
        else:
            outpath = os.path.split(path)[0]
            if not(outname is None):
                if not(outname[-5:] == '.json'):
                    outname = outname + '.json'
                self.__writeJSON(os.path.join(outpath, outname))
            else:
                self.__writeJSON(os.path.join(outpath, self.outJSON['name'] + '.json'))

    def __writeJSON(self, path):
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self.outJSON, f, ensure_ascii=False, indent=4)

    def __parseXLSX(self):
        '''
        TODO DOKU
        '''
        doc = pd.read_excel(self.path, sheet_name=['General Information'])
        gi = doc['General Information']
        ### DEBUG
        print(gi)
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
        values = self.__filter(gi, ['Family Name', 'Given Name'], y=1, isSingleValue=False)
        print(values)

    def __filter(self, df, conds: list[str], y: int = 0, x: int = 0, isSingleValue: bool = True):
        '''
        TODO DOKU
        '''
        values = dict()
        isFirst = True
        iy = y
        ix = x
        for cond in conds:
            coords = self.__getCoords(df, cond)
            if isSingleValue:
                values[cond] = df.iloc[coords[0] + iy, coords[1] + ix]
            # we use the first condition to determine how many cells are filled before we reach a nan
            elif isFirst:
                values[cond] = []
                isFirst = False

                while True:
                    value = df.iloc[coords[0] + iy, coords[1] + ix]
                    # if cell contains nan we have found all filled cellvalues and can leave to while loop
                    if value is np.nan:
                        break
                    values[cond].append(value)
                    iy += y
                    ix += x
            # we take for all other conditions as many cellvalues as for the first, if they contain nan we save an empty string
            else:
                values[cond] = []
                if y == 0:
                    for i in range(coords[1] + x, coords[1] + ix):
                        value = df.iloc[coords[0] + iy, i]
                        if value is np.nan:
                            values[cond].append('')
                        else:
                            values[cond].append(value)
                elif x == 0:
                    for i in range(coords[0] + y, coords[0] + iy):
                        value = df.iloc[i, coords[1] + ix]
                        if value is np.nan:
                            values[cond].append('')
                        else:
                            values[cond].append(value)
        return values

    def __getCoords(self, df, cond):
        '''
        TODO DOKU
        '''
        return[(x, y) for x, y in zip(*np.where(df.values == cond))][0]


if __name__ == '__main__':
    path = '../datasets/EnzymeML_Template_Example.xlsm'
    ConverterExcelJSON().toJSON(path)
