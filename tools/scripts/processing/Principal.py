import pandas as pd
import numpy as np
import openpyxl as xl
import xlwings as xw
from xlwings.utils import rgb_to_int
from calendar import month_name
import logging
import os
import re

LOGGER = logging.getLogger(__name__)
# LOGGER = logging.getLogger()
# formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
# handler = logging.StreamHandler()
# handler.setFormatter(formatter)
# LOGGER.addHandler(handler)
# LOGGER.setLevel(logging.DEBUG)

class principal:
    
    def __init__(self, year, month, *csvdst):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        # self.data = data
        self.csvdst = csvdst or f'C:\dev\Production\data\{year}\{self.monthApplied}' # r'J:\Systems\Production & Override\Production\Premium Data\Current Month\Formatted'
        self.carrierId = 1710
        self.regexstring = f'M Financial - {month_name[month]} Sales Report.xlsx' + '|' + f'M Holdings AUM {self.monthApplied}{self.year-2000}.xlsx'
        self.regex = re.compile(self.regexstring)
        self.datadir =  f'J:\Acctng\Production\{year}\Data\Principal' 
        self.csvname = f"P-1710-401-{year}-{self.monthApplied}-1.txt" 
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        self.exportCols = ["SourceFileName","CarrierID","MemFirmID","CarrierContracteeID","CarrierContracteeName",
                            "ProductID","CarrierProductID","YearApplied","MonthApplied","ProducerID",
                            "CarrierProducerID","CarrierProducerName","PolicyNumber","IssueDate","InsuredName",
                            "YTDAnnualizedPrem" ,"YTDAnnualizedLowNon","YTDFace","RiskNumber","RiskName",
                            "Exchange1035","SplitPercentage","Replacement","CarrierProductName","PolicyOwner",
                            "Exchange","NLG","Modco","YRT","CSO2001"]
    
    def getrawmonth(self, year=None, month=None):

        if not year: 
            year = self.year
        if not month:
            month = self.month

        regex = f'M Financial - {month_name[month]} Sales Report.xlsx' + '|' + f'M Holdings AUM {str(month).zfill(2)}{year-2000}.xlsx'
        regex = re.compile(regex)

        isFileFound = True

        LOGGER.info(f'Searching {month_name[month]} file for carrierid {self.carrierId}...')
        
        file = [os.path.join(self.datadir,f) for f in os.listdir(self.datadir) if re.search(regex,f)]        
                    
        if len(file) == 1:
            LOGGER.info(f'Principal {month_name[month]} is available')        
        elif len(file) == 0:
            isFileFound = False
            LOGGER.info(f'Principal {month_name[month]} data not available')
        else:
            isFileFound = False
            LOGGER.info(f'There are {len(file)} Principal files found for {month_name[month]}')

        return isFileFound

    def getcsv(self, year=None, month=None): # version starting from 202209

        if not year:
            year = self.year
        if not month:
            month = self.month

        regex = f'M Financial - {month_name[month]} Sales Report.xlsx' + '|' + f'M Holdings AUM {str(month).zfill(2)}{year-2000}.xlsx'
        regex = re.compile(regex)
        file = list(filter(regex.match,os.listdir(self.datadir)))[0]

        LOGGER.debug(f'Processing Principal 401 with {year} {month_name[month]} data...')
        df = pd.read_excel(os.path.join(self.datadir,file), engine='openpyxl')

        # Filter for only this year's sales
        df = df[pd.DatetimeIndex(df["AssignmentEffDt"]).year == self.year ]
        # df = df[df["M Financial Effect Dt"] > pd.Timestamp(self.year, 1, 1)] # Previous Filter

        ## ETL Logic ##
        df["SourceFileName"] = f"P-1710-401-{self.year}-{self.monthApplied}-1"
        df["CarrierID"] = self.carrierId
        df["MemFirmID"] = 0
        df["CarrierContracteeID"] = df["HierarchyName"].str.strip()
        df["CarrierContracteeName"] = df["HierarchyName"].str.strip()
        df["ProductID"] = 0
        df["CarrierProductID"] = "Principal 401k"
        df["YearApplied"] = self.year
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df["MktrRole"].str.strip()
        df["CarrierProducerName"] = df["MktrRole"].str.strip()
        df["PolicyNumber"] = df["Acct"]
        df["IssueDate"] = pd.to_datetime(df["AssignmentEffDt"], format='%m/%d/%Y').dt.normalize()
        df["InsuredName"] = df["AcctName"].str.strip()
        df["YTDAnnualizedPrem"] = df["AUM"]
        df["YTDAnnualizedLowNon"] = 0
        df["YTDFace"] = ""
        df["RiskNumber"] = ""
        df["RiskName"] = df["AcctName"].str.strip()
        df["Exchange1035"] = ""
        df["SplitPercentage"] = df["Rate"]
        df["Replacement"] = ""
        df["CarrierProductName"] = df["Product"]
        df["PolicyOwner"] = ""
        df["Exchange"] = ""
        df["NLG"] = ""
        df["Modco"] = ""
        df["YRT"] = ""
        df["CSO2001"] = ""

        # Report
        target = df["YTDAnnualizedPrem"].sum()
        excess = df["YTDAnnualizedLowNon"].sum()
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD = {target}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD = {excess}')

        LOGGER.info(f'Accessing flash workbook and updating values...')        
        with xw.App(visible=False) as app:
            flash = xw.Book(self.flash)
            notesbreakdown = flash.sheets['Notes-Breakdown']
            notesbreakdown.range('Principal_target').value = target
            notesbreakdown.range('Principal_excess').value = excess
            notesbreakdown.range('Principal_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('Principal_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        ## Export ##
        df["YTDAnnualizedPrem"] = df['YTDAnnualizedPrem'].astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df['YTDAnnualizedLowNon'].astype(float).round(2).map("{:,.2f}".format)
        df = df[self.exportCols]
        df.to_csv(os.path.join(self.csvdst,self.csvname), index=False, sep='|')
        # df.to_excel(os.path.join(self.csvdst,'test.xlsx'), index=False)
        LOGGER.info(f'{self.csvname} is saved at {self.csvdst}.')


def processPrin(year, month):
    '''To generate a YTD report, if report month's data is not available, previous month's data would be used 
    temporarily until that month's data is arrived.'''
    prin = principal(year, month)    
    isFileFound = prin.getrawmonth(year, month)
    if isFileFound == True:
        prin.getcsv(year, month)
    else:
        try:
            if month == 1:
                prin.getcsv(year-1, 12)
            else:
                prin.getcsv(year,month -1)
        except IndexError:
            LOGGER.warning(f'Principal {month_name[month-1]} file not available.')


# if __name__ == '__main__':
#     processPrin(2022,9)


