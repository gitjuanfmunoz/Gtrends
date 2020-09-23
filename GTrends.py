#!/usr/bin/env python
import pandas as pd
import xlsxwriter
from pytrends.request import TrendReq
from datetime import datetime

fecha_inicial, ahora = '2020-06-01', datetime.now()
fecha_hoy = ahora.strftime("%Y-%m-%d")
timeframe = (fecha_inicial + ' ' + fecha_hoy)

kw_sv = ['Tigo', 'Claro', 'Digicel', 'Movistar']
kw_hn = ['Tigo', 'Claro', 'Cablecolor', 'Hondutel']
kw_cr = ['Tigo', 'Cabletica', 'Kolbi', 'Movistar', 'Claro']
kw_bo = ['Tigo', 'Entel', 'Costas', 'AXS', 'Viva']
kw_co = ['Tigo', 'Claro', 'Movistar','Avantel','ETB']
kw_py,kw_gt = ['Tigo', 'Claro', 'Personal'], ['Tigo', 'Claro', 'Movistar']
geo_sv,geo_hn,geo_cr,geo_bo,geo_py,geo_co,geo_gt = ['SV'],['HN'],['CR'],['BO'],['PY'],['CO'],['GT']
interests = pd.DataFrame()

def keyword_interest(kw_list, timeframe, geo):
    pytrends = TrendReq()
    pytrends.build_payload(kw_list,cat=0, timeframe=timeframe, geo=geo, gprop='')
    interests = pytrends.interest_over_time().reset_index()
    return interests

for geo in geo_sv:
    interest = keyword_interest(kw_sv, timeframe, geo)
    globals()['' + str(geo)] = interests.append(interest)
    
for geo in geo_hn:
    interest = keyword_interest(kw_hn, timeframe, geo)
    globals()['' + str(geo)] = interests.append(interest)
    
for geo in geo_bo:
    interest = keyword_interest(kw_bo, timeframe, geo)
    globals()['' + str(geo)] = interests.append(interest)
    
for geo in geo_py:
    interest = keyword_interest(kw_py, timeframe, geo)
    globals()['' + str(geo)] = interests.append(interest)
    
for geo in geo_co:
    interest = keyword_interest(kw_co, timeframe, geo)
    globals()['' + str(geo)] = interests.append(interest)
    
for geo in geo_gt:
    interest = keyword_interest(kw_gt, timeframe, geo)
    globals()['' + str(geo)] = interests.append(interest)
    
for geo in geo_cr:
    interest = keyword_interest(kw_cr, timeframe, geo)
    globals()['' + str(geo)] = interests.append(interest)
    
writer = pd.ExcelWriter('GTrends_Tigo.xlsx', engine='xlsxwriter')
SV.to_excel(writer, sheet_name='SV')
HN.to_excel(writer, sheet_name='HN')
CR.to_excel(writer, sheet_name='CR')
BO.to_excel(writer, sheet_name='BO')
PY.to_excel(writer, sheet_name='PY')
CO.to_excel(writer, sheet_name='CO')
GT.to_excel(writer, sheet_name='GT')
writer.save()