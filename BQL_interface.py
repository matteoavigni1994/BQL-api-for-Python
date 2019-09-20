# -*- coding: utf-8 -*-
"""
Created on Thu Sep 19 14:48:10 2019

@author: matteoavigni
"""




import time
import win32com.client
import pandas as pd
import os
from os import listdir
from os.path import isfile, join
import math


waiting_time = 5





def BQL(let = '', get='', universe='', settings = ''):
    """
    DESCRIPTION: equivalente della funzione BQL in python
    INPUT:
        let = (optional)(local variable) can contain one or more variable name
            and value pairs. Each name and value pair begins with # and ends
            with ; (#lastPrice = px_last(dates=range(-2M, 0D)); )
        get = (the Expression) can contain one or more comma-separated data items.
            Optional parameters that control the format or calculation of a
            specific data item are enclosed within parentheses adjacent to the
            data item and separated by commas (e.g., px_last(dates=-1M) )
        universe = (the Universe) can contain one or more comma separated
            tickers within square brackets, each surrounded by a set of single
            quotation marks (eg, ['IBM US Equity','AAPL US Equity'] ). it can
            also contain the members() function and an index ticker, which
            creates a universe that includes all members of the specified index
            (e.g. members(['SPX Index'] )
        settings = (optional)(Global Parameters) can contain parameters that
            control the format or calculation of all data items in your BQL.Query
            formula and are comma-separated (eg, dates=-1M,per=w )
    OUTPUT:
        df = output della query [pandas dataframe]
    """

    if get == '' or universe == '':
        print('FORMULA ERROR: you have to specify get and universe')
        import sys
        sys.exit()
    ##-----------------------------------------------------------------------------
    # Write and Update
    ##-----------------------------------------------------------------------------
    # Start an instance of Excel
    xlapp = win32com.client.DispatchEx("Excel.Application")
    try:
        xlapp.Workbooks.Open('C:/blp/API/Office Tools/BloombergUI.xla') #
    except:
        print("\nAPI ERROR: manca l'add-in di Bloomberg per Excel\n")
        import sys
        sys.exit()

    support_file = get_support_file()
    wb = xlapp.workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = 'query'
    ws2 = wb.Worksheets.Add(After=ws)
    ws2.Name = 'data'

    ##-----------------------------------------------------------------------------
    # Building queries
    ##-----------------------------------------------------------------------------
    query_final = ''
    if let != '':
        query_final = query_final + 'let('+let+ ') '
    query_final = query_final + 'get('+get+ ') '
    query_final = query_final + 'for('+universe+ ') '
    if settings != '':
        query_final = query_final + 'with('+settings+ ') '

    ncells = math.floor(len(query_final)/250) + 1*((len(query_final)%250)!=0)
    relations = ''
    for i in range(ncells):
        ws.Cells(i+1,1).Value = query_final[i*250:min((i+1)*250, len(query_final)-1)]
        relations = relations +' & query!A'+str(i+1) if relations != '' else relations +'query!A'+str(i+1)

        
    ws2.Cells(1,1).Value = 'f=BQL.Query('+ relations + ')'
    ws2.Cells(1,1).Replace('f=', '=')


    ##-----------------------------------------------------------------------------
    # Refreshing excel book
    ##-----------------------------------------------------------------------------
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    xlapp.DisplayAlerts = False
    time.sleep(waiting_time)
    wb.SaveAs(support_file)
    
    xlapp.DisplayAlerts = True
    xlapp.Quit()
    
    ##-----------------------------------------------------------------------------
    # Read and Remove Excel
    ##-----------------------------------------------------------------------------    
    df = pd.read_excel(support_file, 'data')
    time.sleep(2)
    os.remove(support_file)
    
    return df
    
    




def get_support_file():
    """
    DESCRIPTION: crea un file di supporto per ogni nuova richiesta e lo elimina quando ha finito
    INPUT:
        
    OUTPUT:
        support_file = location del file di supporto [str]
    """
    
    location = 'R:\\root\\BQL\\support_files'
    onlyfiles = [f for f in listdir(location) if isfile(join(location, f)) and f.split('.')[-1] == 'xlsx']
    
    i = 0
    actual_file = 'support.xlsx'
    file_present = True
    while file_present:
        actual_file = 'support' + str(i) + '.xlsx'
        i += 1
        if actual_file not in onlyfiles:
            file_present = False

    return location + '\\' + actual_file



if __name__ == '__main__':
    
    test = BQL( let = '',
                get = "IS_EPS(FPR='2015A').VALUE", 
                universe = "members(['SPX Index'])",
                settings = '')







