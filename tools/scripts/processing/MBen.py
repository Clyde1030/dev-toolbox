import pandas as pd
from calendar import month_name, monthrange
from datetime import datetime
# import openpyxl as xl # report sales
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

class mben:
    def __init__(self, year, month, *csvdst):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        # self.data = data
        self.csvdst = csvdst or f'C:\dev\Production\data\{year}\{self.monthApplied}'
        self.carrierId = 'MBen'  
        self.datadir = f'J:\Acctng\Production\{year}\Data\MBen'
        self.goanywhere =  None     
        self.regex = re.compile(f'{year} {self.monthApplied} Year To Date DataXfer.xlsx')
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        self.csvname = f"P-MBen-LIF-{year}-{self.monthApplied}-1.txt"
        self.exportCols = ["SourceFileName","CarrierID","MemFirmID","CarrierContracteeID","CarrierContracteeName",
                            "ProductID","CarrierProductID","YearApplied","MonthApplied","ProducerID",
                            "CarrierProducerID","CarrierProducerName","PolicyNumber","IssueDate","InsuredName",
                            "YTDAnnualizedPrem" ,"YTDAnnualizedLowNon","YTDFace","RiskNumber","RiskName",
                            "Exchange1035","SplitPercentage","Replacement","CarrierProductName","PolicyOwner",
                            "Exchange","NLG","Modco","YRT","CSO2001"]

    def getcsv(self, datayear=None, datamonth=None):

        if not datayear:
            datayear = self.year
        if not datamonth:
            datamonth = self.month

        file = os.path.join(self.datadir,f'{datayear} {str(datamonth).zfill(2)} Year To Date DataXfer.xlsx')
        LOGGER.info(f'Processing Mben with {datayear} {month_name[datamonth]} data...')
        df = pd.read_excel(file, keep_default_na=False, engine='openpyxl')


        ##Filter Out Data Not Needed#
        remove_lst = ['JHV','JH','LINC','METLF','LINCV','NATW','NATWV','PMD','PL','PLV','PRUD','PRUC','PRUDV','TCREF','TCRFV','TIAA','UNUM','PENNM','JHLVI','SYM','PROT','JHVLI','DLICV','TRANS'] # Check the last three carrier code 
        df = df[~df['CarrierCode'].isin(remove_lst)]

        ## ETL Logic ##
        df["SourceFileName"] = f"P-MBen-LIF-{self.year}-{self.monthApplied}-1"
        df["CarrierID"] = df["CarrierCode"].apply(lambda x: '1802' if x== 'GENAM' else
                                                            '1807' if x== 'NLIC' else 
                                                            '1829' if x=='GW' else 
                                                            '1829' if x=='GWV' else 
                                                            '1804' if x=='MASS' else 
                                                            '1883' if x=='MIDL' else 
                                                            '1871' if x=='NYLV' else 
                                                            '1871' if x=='NYL' else 
                                                            '1889' if x=='NA' else 
                                                            '1868' if x=='OHNAT' else 
                                                            '1885' if x=='SBLI' else 
                                                            '1806' if x=='MINNL' else 
                                                            '1892' if x=='NWMUT' else '' )
        df["MemFirmID"] = '2058' 
        df["CarrierContracteeID"] = 'M Benefit Solutions'
        df["CarrierContracteeName"] = 'M Benefit Solutions'
        df["ProductID"] = 0
        df["CarrierProductID"] = df["ProductName"]
        df["YearApplied"] = self.year
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = '3194'
        df["CarrierProducerID"] = 'Don Friedman'
        df["CarrierProducerName"] = 'Don Friedman'
        df["PolicyNumber"] = df["PolicyNumber"]
        df["IssueDate"] = df["IssueDate"]
        df["InsuredName"] = df["InsuredName"]
        df["YTDAnnualizedPrem"] = 0.1*df['Premium']
        df["YTDAnnualizedPrem"] = df['YTDAnnualizedPrem'] #.astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = 0.9*df['Premium']
        df["YTDAnnualizedLowNon"] = df['YTDAnnualizedLowNon'] #.astype(float).round(2).map("{:,.2f}".format)
        df["YTDFace"] = df["FaceAmount"]
        df["RiskNumber"] = df["TransactionNumber"]
        df["RiskName"] = df["PolicyOwner"].str[:50]
        df["Exchange1035"] = ""
        df["SplitPercentage"] = ""
        df["Replacement"] = ""
        df["CarrierProductName"] = ""
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
            notesbreakdown = flash.sheets['M Securities']
            notesbreakdown.range('Mben_target').value = target
            notesbreakdown.range('Mben_excess').value = excess
            notesbreakdown.range('Mben_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('Mben_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        # Export
        df["YTDAnnualizedPrem"] = df['YTDAnnualizedPrem'].astype(float).round(4).map("{:,.4f}".format)
        df["YTDAnnualizedLowNon"] = df['YTDAnnualizedLowNon'].astype(float).round(4).map("{:,.4f}".format)
        df = df[self.exportCols]
        df.to_csv(os.path.join(self.csvdst,self.csvname), sep='|',index=False)
        # df.to_excel(os.path.join(self.csvdst,'test.xlsx'), index=False)
        LOGGER.info(f'{self.csvname} is saved at {self.csvdst}.')


def processMben(year, month):
    ben = mben(year,month)
    try:
        ben.getcsv(year, month)
    except FileNotFoundError:
        LOGGER.warning(f'Mben {month_name[month]} data not available.')
        try: 
            if month == 1:
                ben.getcsv(year-1, 12)
            else:
                ben.getcsv(year,month-1)
        except IndexError:
            LOGGER.warning(f'Mben {month_name[month-1]} data not available.')
            

# if __name__ == '__main__':
#     processMben(2022, 11)




