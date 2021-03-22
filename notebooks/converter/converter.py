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
        self.indexFinder = {}
        self.outJSON: dict[str, object] = {}
        self.__parseXLSX()
        # DEBUG
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
            # we take for all other conditions as many cellvalues as for the first
            else:
                values[cond] = []
                if y == 0:
                    for i in range(coords[1] + x, coords[1] + ix):
                        value = df.iloc[coords[0] + iy, i]
                        values[cond].append(value)
                elif x == 0:
                    for i in range(coords[0] + y, coords[0] + iy):
                        value = df.iloc[i, coords[1] + ix]
                        values[cond].append(value)
        return values

    def __getCoords(self, df, cond):
        '''
        TODO DOKU
        '''
        return[(x, y) for x, y in zip(*np.where(df.values == cond))][0]

    def __getGI(self, gi):
        '''
        TODO DOKU
        Open question: How should we handle NaN?
        '''
        values = self.__filter(gi, ['Title', 'Date of creation', 'DOI (optional)', 'PubMedID (optional)', 'URL (optional)'], x=1)
        self.outJSON['name'] = values['Title']
        ######## not implemented yet?
        # self.outJSON['date'] = values['Date of creation'].strftime('%d.%m.%Y')
        ######## key 'doi' guessed
        self.outJSON['doi'] = values['DOI (optional)']
        self.outJSON['pubmedID'] = values['PubMedID (optional)']
        ######## key 'url' guessed
        self.outJSON['url'] = values['URL (optional)']
        # TODO add mail and institute
        values = self.__filter(gi, ['Family Name', 'Given Name'], y=1, isSingleValue=False)
        # TODO add author information to outJSON

    def __getProteins(self, proteins):
        '''
        TODO DOKU
        '''
        self.outJSON['protein'] = []
        values = self.__filter(proteins, ['Name', 'ID', 'Sequence', 'Vessel', 'Constant', 'Source organism', 'EC Number', 'UniProt ID'], y=1, isSingleValue=False)
        self.indexFinder['proteins'] = values['Name']
        for i in range(len(values['Name'])):
            proteinDic = {}
            # id not wanted
            # proteinDic['id_'] = values['ID'][i]
            proteinDic['name'] = values['Name'][i]
            proteinDic['sequence'] = values['Sequence'][i]
            # TODO may need to access VesselDic
            proteinDic['compartment'] = values['Vessel'][i]
            proteinDic['constant'] = (values['Constant'][i] == 'Constant')
            # Source organism is optional is only added if an Source organism is given
            if not(values['Source organism'][i] is np.nan):
                proteinDic['organism'] = values['Source organism'][i]
            # EC Number is optional
            if not(values['EC Number'][i] is np.nan):
                proteinDic['ecnumber'] = values['EC Number'][i]
            # UniProt ID is optional
            if not(values['UniProt ID'][i] is np.nan):
                proteinDic['uniprotid'] = values['UniProt ID'][i]
            # Protein Dic needs substanceunits, these are not given in the excel sheet so I cheat here!!!
            proteinDic['substanceunits'] = 'umole / l'
            # Protein Dic needs init_conc, which is only given in the data sheet and differs for differnet experiments!!!
            # So I cheat here again!!!
            proteinDic['init_conc'] = 1.0
            self.outJSON['protein'].append(proteinDic)

    def __getReactants(self, reactants):
        '''
        TODO DOKU
        '''
        self.outJSON['reactant'] = []
        values = self.__filter(reactants, ['Name', 'ID', 'Vessel', 'Constant', 'SMILES', 'InCHI'], y=1, isSingleValue=False)
        self.indexFinder['reactants'] = values['Name']
        for i in range(len(values['Name'])):
            reactantDic = {}
            # id not wanted
            # reactantDic['id_'] = values['ID'][i]
            reactantDic['name'] = values['Name'][i]
            # TODO may need to access VesselDic
            reactantDic['compartment'] = values['Vessel'][i]
            reactantDic['constant'] = (values['Constant'][i] == 'Constant')
            # InCHI is optional is only added if an InCHI is given
            if not(values['InCHI'][i] is np.nan):
                reactantDic['inchi'] = values['InCHI'][i]
            # SMILES is optional
            if not(values['SMILES'][i] is np.nan):
                reactantDic['smiles'] = values['SMILES'][i]
            self.outJSON['reactant'].append(reactantDic)

    def __getReactions(self, reactions):
        '''
        TODO DOKU
        '''
        self.outJSON['reaction'] = []
        values = self.__filter(reactions, ['Name', 'ID', 'Temperature value', 'Temperature unit', 'pH value', 'Educts', 'Products', 'Proteins', 'Modifiers'], y=1, isSingleValue=False)
        for i in range(len(values['Name'])):
            reacDic = {}
            # id not wanted
            # reacDic['id_'] = values['ID'][i]
            reacDic['name'] = values['Name'][i]
            reacDic['temperature'] = values['Temperature value'][i]
            reacDic['tempunit'] = values['Temperature unit'][i]
            reacDic['ph'] = values['pH value'][i]
            # Educts, Products and Proteins can be more than one -> need lists
            # All need stoich key, but this is not given in excel so I cheat here!!!
            reacDic['educts'] = []
            educts = values['Educts'][i].split(', ')
            for educt in educts:
                eductDic = {}
                eductDic['species'] = educt
                eductDic['stoich'] = 1.0
                ind = self.indexFinder['reactants'].index(educt)
                eductDic['constant'] = self.outJSON['reactant'][ind]['constant']
                eductDic['replicates'] = []
                reacDic['educts'].append(eductDic)
            reacDic['products'] = []
            products = values['Products'][i].split(', ')
            for product in products:
                prodDic = {}
                prodDic['species'] = product
                prodDic['stoich'] = 1.0
                ind = self.indexFinder['reactants'].index(product)
                prodDic['constant'] = self.outJSON['reactant'][ind]['constant']
                prodDic['replicates'] = []
                reacDic['products'].append(prodDic)
            reacDic['proteins'] = []
            prots = values['Proteins'][i].split(', ')
            for prot in prots:
                protDic = {}
                protDic['species'] = prot
                protDic['stoich'] = 1.0
                ind = self.indexFinder['proteins'].index(prot)
                protDic['constant'] = self.outJSON['protein'][ind]['constant']
                protDic['replicates'] = []
                reacDic['proteins'].append(protDic)
            # Modifiers not always given
            if not(values['Modifiers'][i] is np.nan):
                # TODO wants a list
                reacDic['modifiers'] = []
                modifiers = values['Proteins'][i].split(', ')
                for modifier in modifiers:
                    modDic = {}
                    modDic['species'] = modifier
                    modDic['stoich'] = 1.0
                    try:
                        ind = self.indexFinder['reactants'].index(modifier)
                        modDic['constant'] = self.outJSON['reactant'][ind]['constant']
                    # TODO except...
                    except ValueError:
                        pass
                    try:
                        ind = self.indexFinder['proteins'].index(modifier)
                        modDic['constant'] = self.outJSON['protein'][ind]['constant']
                    except ValueError:
                        # Should not happen...
                        modDic['constant'] = True
                    reacDic['modifiers'].append(modDic)
            else:
                reacDic['modifiers'] = []
            # reaction needs reversible key, this is not given in the excel sheet so I cheat here!!!
            reacDic['reversible'] = True
            self.outJSON['reaction'].append(reacDic)

    def __getVessel(self, vessels):
        '''
        TODO DOKU
        This function is a placeholder till I understand mapping, for now you can only get ONE Vessel not more!!!
        '''
        values = self.__filter(vessels, ['Name', 'ID', 'Volume value', 'Volume unit'], y=1)
        vesselDic = {}
        vesselDic['id_'] = values['ID']
        vesselDic['name'] = values['Name']
        vesselDic['size'] = values['Volume value']
        vesselDic['unit'] = values['Volume unit']
        self.outJSON['vessel'] = vesselDic

    def __getVessels(self, vessels):
        '''
        TODO DOKU
        This function creates a list of all vessels in the excel-sheet, but the API expects a mapping not a list,
        maybe this function will be useful later on
        '''
        self.outJSON['vessel'] = []
        values = self.__filter(vessels, ['Name', 'ID', 'Volume value', 'Volume unit'], y=1, isSingleValue=False)
        for i in range(len(values['Name'])):
            vesselDic = {}
            vesselDic['id_'] = values['ID'][i]
            vesselDic['name'] = values['Name'][i]
            vesselDic['size'] = values['Volume value'][i]
            vesselDic['unit'] = values['Volume unit'][i]
            self.outJSON['vessel'].append(vesselDic)

    def __parseXLSX(self):
        '''
        TODO DOKU
        '''
        doc = pd.read_excel(self.path, sheet_name=['General Information', 'Vessels', 'Reactants', 'Proteins', 'Reactions'])
        self.__getGI(doc['General Information'])
        self.__getVessel(doc['Vessels'])
        self.__getReactants(doc['Reactants'])
        self.__getProteins(doc['Proteins'])
        self.__getReactions(doc['Reactions'])

    def __writeJSON(self, path):
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self.outJSON, f, ensure_ascii=False, indent=4)


if __name__ == '__main__':
    path = '../datasets/EnzymeML_Template_Example.xlsm'
    ConverterExcelJSON().toJSON(path)
