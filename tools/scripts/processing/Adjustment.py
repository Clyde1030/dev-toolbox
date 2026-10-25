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

class adjustment:

    def __init__(self, year, month, *csvdst):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        self.csvdst = csvdst or f'C:\dev\Production\data\{year}\{self.monthApplied}'
        self.file = r'J:\Systems\Production & Override\Production\Premium Data\Current Month\Raw\P-M-Adj-YYYY-MM-1.xlsm' 
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm') 
        self.csvname = f"P-M-Adj-{year}-{self.monthApplied}-1.txt"
        self.exportCols = ["SourceFileName","CarrierID","MemFirmID","CarrierContracteeID","CarrierContracteeName",
                            "ProductID","CarrierProductID","YearApplied","MonthApplied","ProducerID",
                            "CarrierProducerID","CarrierProducerName","PolicyNumber","IssueDate","InsuredName",
                            "YTDAnnualizedPrem" ,"YTDAnnualizedLowNon","YTDFace","RiskNumber","RiskName",
                            "Exchange1035","SplitPercentage","Replacement","CarrierProductName","PolicyOwner",
                            "Exchange","NLG","Modco","YRT","CSO2001"]


    def getcsv(self, year=None, month=None):

        if not year:
            year = self.year
        if not month:
            month = self.month

        LOGGER.info(f'Processing adjustment data...')

        df = pd.read_excel(self.file,sheet_name="Export", converters = {'CarrierID':str, 'YearApplied':int, 'MonthApplied':str,'PolicyNumber':str}, engine='openpyxl')
        df = df[pd.isnull(df['CarrierID'])==False]        
        
        # # Report
        VoyaTarget = df.loc[df["CarrierID"]=='1127',"YTDAnnualizedPrem"].sum()
        VoyaExcess = df.loc[df["CarrierID"]=='1127',"YTDAnnualizedLowNon"].sum()
        MasterCareTarget = df.loc[df["CarrierID"]=='1380',"YTDAnnualizedPrem"].sum()
        MasterCareExcess = df.loc[df["CarrierID"]=='1380',"YTDAnnualizedLowNon"].sum()
        OneAmericaTarget = df.loc[df["CarrierID"]=='1383',"YTDAnnualizedPrem"].sum()
        OneAmericaExcess = df.loc[df["CarrierID"]=='1383',"YTDAnnualizedLowNon"].sum()
        MutualOmahaTarget = df.loc[df["CarrierID"]=='1384',"YTDAnnualizedPrem"].sum()
        MutualOmahaExcess = df.loc[df["CarrierID"]=='1384',"YTDAnnualizedLowNon"].sum()

        LOGGER.info(f'Target Premium {month_name[self.month]} YTD Voya = {VoyaTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD Voya = {VoyaExcess}')
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD MasterCare = {MasterCareTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD MasterCare = {MasterCareExcess}')
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD OneAmerica = {OneAmericaTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD OneAmerica = {OneAmericaExcess}')
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD Mutual Of Omaha = {MutualOmahaTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD Mutual Of Omaha = {MutualOmahaExcess}')
        
        LOGGER.info(f'Accessing flash workbook and updating values...')
        with xw.App(visible=False) as app:
            flash = xw.Book(self.flash)
            notesbreakdown = flash.sheets['Notes-Breakdown']
            
            notesbreakdown.range('Voya_401_adj').value = VoyaTarget # notesbreakdown['E26'].value = VoyaTarget
            notesbreakdown.range('MasterCare').value = MasterCareTarget # notesbreakdown['B120'].value = MasterCareTarget
            notesbreakdown.range('OneAmerica').value = OneAmericaTarget # notesbreakdown['B120'].value = MasterCareTarget
            notesbreakdown.range('MutualOm').value = MutualOmahaTarget # notesbreakdown['B120'].value = MasterCareTarget

            notesbreakdown.range('Voya_401_adj').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('MasterCare').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('OneAmerica').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('MutualOm').api.Font.Color = rgb_to_int((0, 102, 204))

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


def processAdj(year, month):
    adj = adjustment(year, month)
    adj.getcsv()

# if __name__ == '__main__':
#     processAdj(2022,11)
    


        
