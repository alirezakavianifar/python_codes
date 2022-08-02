# -*- coding: utf-8 -*-
"""
Created on Thu Jul  7 16:40:48 2022

@author: Msi
"""

n_retries = 0
n_downloads = 0

def retry(func):
    def try_it(*args, **kwargs):
        global n_retries
        try:
            result = func(*args, **kwargs)
            return result
        
        except Exception as e:
            n_retries += 1
            print(e)
            if n_retries < 3:
                try_it(*args, **kwargs)
                
    return try_it


def retry_with_arguments(func):
    def try_it(*args, **kwargs):
        global n_retries
        try:
            result = func(*args, **kwargs)
            return result
        
        except Exception as e:
            n_retries += 1
            print(e)
            if n_retries < 3:
                try_it(*args, **kwargs)
                
    return try_it
            
@retry        
def test():
    global n_downloads
    x = 0
    x = [1,2,3,4,5]
    c = 0
    for index, i in enumerate(x[c:]):
        n_downloads += 1
        print(n_downloads)
    raise Exception('error')
        
        
if __name__=='__main__':
    
    test()