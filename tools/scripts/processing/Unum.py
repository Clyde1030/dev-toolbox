import numpy as np
import pandas as pd
from calendar import month_name
import openpyxl as xl # report sales
import xlwings as xw
from xlwings.utils import rgb_to_int
import logging
import shutil
import os
import re

LOGGER = logging.getLogger(__name__)
# LOGGER = logging.getLogger()
# formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
# handler = logging.StreamHandler()
# handler.setFormatter(formatter)
# LOGGER.addHandler(handler)
# LOGGER.setLevel(logging.DEBUG)

class unum:

    def __init__(self, year, month, *csvdst):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        self.csvdst = csvdst or f'C:\dev\Production\data\{year}\{self.monthApplied}'
        self.carrierId = 1267
        self.datadir = r'J:\Acctng\Production\{}\Data\{}'.format(year,'UNUM')
        self.goanywhere =  r'J:\Acctng\Revenue\goanywhere\UnumProvident'
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        self.csvname = f"P-1267-DIS-{year}-{self.monthApplied}-1.txt"
        self.exportCols = ["SourceFileName","CarrierID","MemFirmID","CarrierContracteeID","CarrierContracteeName",
                            "ProductID","CarrierProductID","YearApplied","MonthApplied","ProducerID",
                            "CarrierProducerID","CarrierProducerName","PolicyNumber","IssueDate","InsuredName",
                            "YTDAnnualizedPrem" ,"YTDAnnualizedLowNon","YTDFace","RiskNumber","RiskName",
                            "Exchange1035","SplitPercentage","Replacement","CarrierProductName","PolicyOwner",
                            "Exchange","NLG","Modco","YRT","CSO2001"]

    def getrawmonth(self, year=None, month=None):
        '''Unum sends YTD production report through GoAnywhere every month. This fuction fetches the corresponding month's
        raw data from Goanywhere. Parameters year and month is default to self.year and self.month if not specified'''

        isFileFound = True

        if not year:
            year = self.year
        if not month:
            month = self.month
        regex = re.compile(f'{month_name[month]} {year-2000} Sales Data - M.Financial.xlsx')
        
        LOGGER.info(f'Searching {month_name[month]} file for carrierid {self.carrierId}...')

        f = []
        # If processing month is December, raw data will be delivered in next year January
        if month == 12:
            for dirs, _, filenames in os.walk(os.path.join(self.goanywhere,f'{year+1}','01')):
                f.extend([os.path.join(dirs,f) for f in filenames if re.search(regex,os.path.join(dirs,f))])            
        # Otherwise, raw data will be delivered in the next month.
        else:
            for dirs, _, filenames in os.walk(os.path.join(self.goanywhere,f'{year}',str(month+1).zfill(2))):
                f.extend([os.path.join(dirs,f) for f in filenames if re.search(regex,os.path.join(dirs,f))])        
        
        if len(f) == 0:
            LOGGER.info(f'File not available for {month_name[self.month]} yet. Go back to check Goanywhere folder.')
            path = None
            isFileFound = False
        elif len(f) == 1:
            path = f[0]            
        else: 
            path = f[len(f)-1] # To find the latest file

        if path != None:
            if os.path.exists(os.path.join(self.datadir,os.path.basename(path)))==False:
                shutil.copy2(path,self.datadir)
                LOGGER.info(f'{os.path.basename(path)} is copied to {month_name[self.month]} data directory')
            else:
                LOGGER.info(f'{os.path.basename(path)} has already existed in data directory.')            
        
        return isFileFound

    def getcsv(self, datayear=None, datamonth=None):
        '''This function generates csv files for individual life business. datayear and datamonth are provided if a different month's data is reported as default month'''

        if not datayear:
            datayear = self.year
        if not datamonth:
            datamonth = self.month

        regex = re.compile(f'{month_name[datamonth]} {datayear-2000} Sales Data - M.Financial.xlsx')

        # filter the listdit based on the regex to find us file name
        file = list(filter(regex.match,os.listdir(self.datadir)))[0]

        LOGGER.info(f'Processing UNUM disability with {datayear} {month_name[datamonth]} data...')

        df = pd.read_excel(os.path.join(self.datadir,file),sheet_name="Data Set 1", converters={'Policy Number':str,'Policy Number Long':str,'Producer ID':str}, engine='openpyxl')

        # Special Formatting
        df = df[df["CY Sale Premium"]!=0]
        df["Policy Number"] = df["Policy Number"].str.lstrip(' ')
        df['Policy Number Long'] = df['Policy Number Long'].str.rstrip().str.lstrip().str[-7:]
        df['PM'] = df.apply(lambda x: x['Product Market'] if x['Product Line2'] == "IDI" else (x['Product Market']+'-'+x['New or NBOC'] if x['Product Line2'] == "VWB" else x['Product Line2']),axis=1)

        # ETL 
        df["SourceFileName"] = f"P-1267-DIS-{self.year}-{self.monthApplied}-1-IDI" 
        df["CarrierID"] = 1267
        df["MemFirmID"] = 0
        df["CarrierContracteeID"] = ""
        df["CarrierContracteeName"] = ""
        df["ProductID"] = 0
        df["CarrierProductID"] = df.apply(lambda x: 'MONOGRAM 500' if x['PM'][0:2] == '50' 
                                            else ( 'Monogram II '+ x['PM'] if (( x['Product Line']=='INCOME SERIES' and x['Special Group']=='MONOGRAM') or x['Product Line']=='IDI-850')
                                                    else 'NON-MONOGRAM'+x['PM']) , axis=1)
        df["YearApplied"] = self.year
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df['Producer ID']
        df["CarrierProducerName"] = df['Producer Name'].str[0:50]
        df["PolicyNumber"] = df.apply(lambda x: x['Policy Number'] if (x['Policy Number Long'] == '' or x['Policy Number Long'] == ' ' or pd.isnull(x['Policy Number Long']) == True) else x['Policy Number Long'], axis=1)
        df["IssueDate"] = df['Sale Cr Date']
        df["InsuredName"] = ""
        df["YTDAnnualizedPrem"] = df['CY Sale Premium'] #.astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = 0
        df["YTDFace"] = 0
        df["RiskNumber"] = df.apply(lambda x: x['Policy Number Long'] if (x['Policy Number'] == '' or x['Policy Number'] == ' ' or pd.isnull(x['Policy Number']) == True) else x['Policy Number'], axis=1)
        df["RiskName"] = df["Policy Holder Name"].str[0:50]
        df["Exchange1035"] = ""
        df["SplitPercentage"] = ""
        df["Replacement"] = ""
        df["CarrierProductName"] = df.apply(lambda x: x['Product Line']+"-"+x['PM'] if x['Product Line'] == 'VWB' else x['Product Line']+" "+x['PM'], axis = 1)
        df["PolicyOwner"] = ""
        df["Exchange"] = ""
        df["NLG"] = ""
        df["Modco"] = ""
        df["YRT"] = ""
        df["CSO2001"] = ""   
        
        # Report
        GroupTarget = df.loc[df["Product Group"].str.contains('Group',case = False)==True,"YTDAnnualizedPrem"].sum()
        GroupExcess = df.loc[df["Product Group"].str.contains('Group',case = False)==True,"YTDAnnualizedLowNon"].sum()
        VBTarget = df.loc[df["Product Group"].str.contains('VB',case = False)==True,"YTDAnnualizedPrem"].sum()
        VBExcess = df.loc[df["Product Group"].str.contains('VB',case = False)==True,"YTDAnnualizedLowNon"].sum()
        IDITarget = df.loc[df["Product Group"].str.contains('IDI',case = False)==True,"YTDAnnualizedPrem"].sum()
        IDIExcess = df.loc[df["Product Group"].str.contains('IDI',case = False)==True,"YTDAnnualizedLowNon"].sum()

        LOGGER.info(f'Target Premium {month_name[self.month]} YTD Group = {GroupTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD Group = {GroupExcess}')
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD VB = {VBTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD VB = {VBExcess}')
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD IDI = {IDITarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD IDI = {IDIExcess}')
        
        LOGGER.info(f'Accessing flash workbook and updating values...')
        with xw.App(visible=False) as app:
            flash = xw.Book(self.flash)
            notesbreakdown = flash.sheets['Notes-Breakdown']
            notesbreakdown.range('Unum_Group_target').value = GroupTarget
            notesbreakdown.range('Unum_Group_excess').value = GroupExcess
            notesbreakdown.range('Unum_VWB_target').value = VBTarget
            notesbreakdown.range('Unum_VWB_excess').value = VBExcess
            notesbreakdown.range('Unum_ID_target').value = IDITarget
            notesbreakdown.range('Unum_ID_excess').value = IDIExcess
            notesbreakdown.range('Unum_Group_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('Unum_Group_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('Unum_VWB_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('Unum_VWB_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('Unum_ID_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('Unum_ID_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        # Export
        df["YTDAnnualizedPrem"] = df['YTDAnnualizedPrem'].astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df['YTDAnnualizedLowNon'].astype(float).round(2).map("{:,.2f}".format)
        df = df[self.exportCols]
        df.to_csv(os.path.join(self.csvdst,self.csvname), index=False, sep = '|')
        # df.to_excel(os.path.join(self.csvdst,'test.xlsx'), index=False)
        LOGGER.info(f'{self.csvname} is saved at {self.csvdst}.')


def processUnum(year, month):
    '''To generate a YTD report, if report month's data is not available, previous month's data would be used temporarily until that month's data is arrived.'''
    un = unum(year,month)
    isFileFound = un.getrawmonth(year, month)
    if isFileFound == True:
        un.getcsv(year, month)
    else:
        try:
            if month == 1:
                un.getcsv(year-1, 12)
            else:
                un.getcsv(year,month -1)
        except IndexError:
            LOGGER.warning(f'Unum {month_name[month-1]} file not available.')

# if __name__ == '__main__':
#     processUnum(2022, 11)
    
