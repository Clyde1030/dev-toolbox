from calendar import month_name, monthrange
import logging
import shutil
import os
import re

import pandas as pd
import xlwings as xw
from xlwings.utils import rgb_to_int

LOGGER = logging.getLogger(__name__)
# LOGGER = logging.getLogger()
# formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
# handler = logging.StreamHandler()
# handler.setFormatter(formatter)
# LOGGER.addHandler(handler)
# LOGGER.setLevel(logging.DEBUG)

class mhs:

    def __init__(self, year, month, *csvdst):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        self.csvdst = csvdst or f'C:\dev\Production\data\{year}\{self.monthApplied}'
        self.csvname = 'P-1999-LIF-1.txt'
        self.carrierId = 1999
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        self.exportCols = ["SourceFileName","CarrierID","MemFirmID","CarrierContracteeID","CarrierContracteeName",
                            "ProductID","CarrierProductID","YearApplied","MonthApplied","ProducerID",
                            "CarrierProducerID","CarrierProducerName","PolicyNumber","IssueDate","InsuredName",
                            "YTDAnnualizedPrem" ,"YTDAnnualizedLowNon","YTDFace","RiskNumber","RiskName",
                            "Exchange1035","SplitPercentage","Replacement","CarrierProductName","PolicyOwner",
                            "Exchange","NLG","Modco","YRT","CSO2001"]


    def getrawmonth(self, year, month):

        if not year:
            year = self.year
        if not month:
            month = self.month

        LOGGER.info(f'Searching {year} {month_name[self.month]} MHS data...')

        if month == 12:
            nextmonth = '01'
            regexString = r"mfgovr_cron.{}0101.".format(year+1)+"\d{6}.txt"
        else:
            nextmonth = month +1
            regexString = r"mfgovr_cron.{}{}01.".format(year, str(nextmonth).zfill(2))+"\d{6}.txt"

        goanywhere =  r'\\mfh.local\data\CorpData\FinancialTeam\Production and Sabbatical Reports\Production'
        regex = re.compile(regexString)  
        
        file = []

        for dirs, _, filenames in os.walk(os.path.join(goanywhere, str(year))):
            file.extend([os.path.join(dirs,f) for f in filenames if re.search(regex,f)])            
        for dirs, _, filenames in os.walk(os.path.join(goanywhere, str(year+1))):
            file.extend([os.path.join(dirs,f) for f in filenames if re.search(regex,f)])            

        if len(file)==0:
            isFileFound = False
            LOGGER.info(f'{year} {month_name[month]} MHS Data not available yet.')

        else:
            isFileFound = True
            for i in file:
                if os.path.exists(os.path.join(self.csvdst, self.csvname))==False:
                    shutil.copy2(i,os.path.join(self.csvdst, self.csvname))
                    # shutil.copy2(i,os.path.join(r'J:\Systems\Production & Override\Production\Premium Data\Current Month\Formatted', self.csvname))
                    LOGGER.info(f'{os.path.basename(i)} is copied to {self.csvdst} as {self.csvname}')
                else:
                    LOGGER.info(f'{os.path.basename(i)} is already in data directory.')      
        return isFileFound      


def processMHS(year, month):
    m = mhs(year,month)
    m.getrawmonth(year,month)
        

if __name__ == '__main__':
    processMHS(2022, 11)
