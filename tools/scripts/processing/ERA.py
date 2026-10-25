import pandas as pd
from calendar import month_name
import xlwings as xw
from xlwings.utils import rgb_to_int
import logging
import os

LOGGER = logging.getLogger(__name__)
# LOGGER = logging.getLogger()
# formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
# handler = logging.StreamHandler()
# handler.setFormatter(formatter)
# LOGGER.addHandler(handler)
# LOGGER.setLevel(logging.DEBUG)

class era:

    def __init__(self, year, month, csvdst=None):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        self.datadir = f'J:\Acctng\Production\{year}\Data\ERA'
        self.carrierid = 1340
        self.goanywhere = f'J:\Acctng\Revenue\goanywhere\Exceptional Risk'
        if not csvdst:
            self.csvdst = f'C:\dev\Production\data\{year}\{self.monthApplied}' #r'J:\Systems\Production & Override\Production\Premium Data\Current Month\Formatted'
        else:
            self.csvdst = csvdst
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm') # os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}t.xlsm')        
        self.csvname = f'P-{self.carrierid}-DIS-{year}-{self.monthApplied}-1.txt'
        self.exportCols = ["SourceFileName","CarrierID","MemFirmID","CarrierContracteeID","CarrierContracteeName",
                            "ProductID","CarrierProductID","YearApplied","MonthApplied","ProducerID",
                            "CarrierProducerID","CarrierProducerName","PolicyNumber","IssueDate","InsuredName",
                            "YTDAnnualizedPrem" ,"YTDAnnualizedLowNon","YTDFace","RiskNumber","RiskName",
                            "Exchange1035","SplitPercentage","Replacement","CarrierProductName","PolicyOwner",
                            "Exchange","NLG","Modco","YRT","CSO2001"]

    def isFileFound(self, year=None, month=None):        

        isFileFound = True
        if not year:
            year = self.year
        if not month:
            month = self.month
        monthfilename = f'M Financial Report {year*100+month}.xlsx'

        if os.path.exists(os.path.join(self.datadir,monthfilename)):
            LOGGER.info(f'{monthfilename} exists')
        else:
            LOGGER.info(f'{monthfilename} does not exists')
            isFileFound = False
        return isFileFound

    def getcsv(self, year=None, month=None):
        
        if not year:
            year = self.year
        if not month:
            month = self.month

        file = r'J:\Systems\Production & Override\Production\Premium Data\Current Month\Raw\P-1340-DIS-YYYY-MM-1.xlsm'

        LOGGER.info(f'Processing ERA data...')

        df = pd.read_excel(file,sheet_name="Export", converters = {'CarrierID':str, 'YearApplied':int, 'MonthApplied':str,'PolicyNumber':str}, engine='openpyxl')
        df = df[pd.isnull(df['PolicyNumber'])==False]        

        # file = f'M Financial Report {year*100+month}.xlsx'

        # trandetail = pd.read_excel(os.path.join(self.datadir,file), sheet_name=1, engine='openpyxl')
        # poldetail = pd.read_excel(os.path.join(self.datadir,file), sheet_name=0, engine='openpyxl')
   
        df["SourceFileName"] = f"P-1064-LIF-{self.year}-{self.monthApplied}-1"
        # df["CarrierID"] = self.carrierid
        # df["MemFirmID"] = 0
        # df["CarrierContracteeID"] = df['Agency_Number']
        # df["CarrierContracteeName"] = df['Agency_Name']
        # df["ProductID"] = 0
        # df["CarrierProductID"] = df['Product_Code'].apply(lambda x: '{0:0>10}'.format(x))
        df["YearApplied"] = str(self.year)
        df["MonthApplied"] = self.monthApplied
        # df["ProducerID"] = 0
        # df["CarrierProducerID"] = df['Agent_Number']
        # df["CarrierProducerName"] = df['Agent_Name']
        # df["PolicyNumber"] = df['Policy_Number']
        # df["IssueDate"] = df['Issue_Date']
        # df["InsuredName"] = df['Insured_Name']
        # df["YTDAnnualizedPrem"] = df['Target_Premium'] #.astype(float).round(2).map("{:,.2f}".format)
        # df["YTDAnnualizedLowNon"] = df['Excess_Premium'] #.astype(float).round(2).map("{:,.2f}".format)
        # df["YTDFace"] = df['Face_Amount'] #.astype(float).round(2).map("{:,.2f}".format)
        # df["RiskNumber"] = ""
        # df["RiskName"] = ""
        # df["Exchange1035"] = df['1035']
        # df["SplitPercentage"] = df['Policy_Count']
        # df["Replacement"] = ""
        # df["CarrierProductName"] = df['Product_Name']
        # df["PolicyOwner"] = ""
        # df["Exchange"] = ""
        # df["NLG"] = ""
        # df["Modco"] = ""
        # df["YRT"] = ""
        # df["CSO2001"] = ""    

        # Report
        target = df["YTDAnnualizedPrem"].sum()
        excess = df["YTDAnnualizedLowNon"].sum()
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD = {target}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD = {excess}')

        LOGGER.info(f'Accessing flash workbook and updating values...')
        with xw.App(visible=False) as app:
            flash = xw.Book(self.flash)
            notesbreakdown = flash.sheets['Notes-Breakdown']
            notesbreakdown.range('ERA').value = target
            notesbreakdown.range('ERA').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        # Export
        df["YTDAnnualizedPrem"] = df['YTDAnnualizedPrem'].astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df['YTDAnnualizedLowNon'].astype(float).round(2).map("{:,.2f}".format)
        df = df[self.exportCols]
        df.to_csv(os.path.join(self.csvdst,self.csvname), index=False, sep = '|')
        # df.to_excel(os.path.join(self.csvdst,'test.xlsx'), index=False)
        LOGGER.info(f'{self.csvname} is saved at: {self.csvdst}')


def processEra(year, month):
    '''To generate a YTD report, if report month's data is not available, previous month's data would be used 
    temporarily until that month's data is arrived.'''
    er = era(year, month)    
    er.getcsv(year, month)
    # isFileFound = er.isFileFound(year, month)
    # if isFileFound == True:
    #     er.getcsv(year, month)
    # else:
    #     try:
    #         if month == 1:
    #             er.getcsv(year-1, 12)
    #         else:
    #             er.getcsv(year,month -1)
    #     except FileNotFoundError:
    #         LOGGER.warning(f'ERA {month_name[month-1]} file not available.')


# if __name__ == '__main__':
#     processEra(2022,10)


















