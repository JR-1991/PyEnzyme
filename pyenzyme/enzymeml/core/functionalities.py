'''
Created on 09.06.2020

@author: JR
'''
from builtins import isinstance


def TypeChecker( value, obj ):
    
    if isinstance( value, obj ):
        return value
    else:
        raise TypeError(f"Expected {str(obj)} got {str(type(value))}")