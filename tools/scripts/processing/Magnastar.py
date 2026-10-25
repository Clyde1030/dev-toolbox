import pandas as pd
from calendar import month_name, monthrange
import openpyxl as xl # report sales
import logging
import shutil
import xlrd
import xlwings as xw
from xlwings.utils import rgb_to_int
import os
import re

LOGGER = logging.getLogger(__name__)
# LOGGER = logging.getLogger()
# formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
# handler = logging.StreamHandler()
# handler.setFormatter(formatter)
# LOGGER.addHandler(handler)
# LOGGER.setLevel(logging.DEBUG)

class magnastar:

    def __init__(self, year, month, *csvdst):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        # self.data = data
        self.csvdst = csvdst or f'C:\dev\Production\data\{year}\{self.monthApplied}'
        self.carrierId = 'Magn'
        self.regex = re.compile(f'{year}{self.monthApplied} PPVUL Magnastar Production Report.xls')  
        self.datadir = f'J:\Acctng\Production\{year}\Data\Mag'
        self.goanywhere =  f'J:\MLife\Magnastar\Data\{year}\{self.monthApplied}\Monthly Reports\{year}{self.monthApplied} PPVUL Magnastar Production Report.xls'     # Ben saves the file here every month:
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        self.csvname = f"P-Magn-LIF-{year}-{self.monthApplied}-1.txt"
        self.exportCols = ["SourceFileName","CarrierID","MemFirmID","CarrierContracteeID","CarrierContracteeName",
                            "ProductID","CarrierProductID","YearApplied","MonthApplied","ProducerID",
                            "CarrierProducerID","CarrierProducerName","PolicyNumber","IssueDate","InsuredName",
                            "YTDAnnualizedPrem" ,"YTDAnnualizedLowNon","YTDFace","RiskNumber","RiskName",
                            "Exchange1035","SplitPercentage","Replacement","CarrierProductName","PolicyOwner",
                            "Exchange","NLG","Modco","YRT","CSO2001"]

    def fetchrawmonth(self, year=None, month=None):

        if not year:
            year = self.year
        if not month:
            month = self.month
        regex = re.compile(f'{year*100+month} PPVUL Magnastar Production Report.xls')

        isFileFound = True

        LOGGER.info(f'Searching {month_name[month]} file for carrier {self.carrierId}...')
        if os.path.exists(self.goanywhere):
            file = os.path.basename(self.goanywhere)
            if os.path.exists(os.path.join(self.datadir,file))==True:
                LOGGER.info(f'{os.path.basename(file)} is already in data directory.')            
            else:
                shutil.copy2(self.goanywhere,self.datadir)
                LOGGER.info(f'{file} is now copied to {month_name[self.month]} data directory.')
        else:    
            isFileFound = False
            LOGGER.warning(f'Magnastar {month_name[self.month]} data not availablr yet.')
        return isFileFound

    def getcsv(self, datayear=None, datamonth=None):

        if not datayear:
            datayear = self.year
        if not datamonth:
            datamonth = self.month

        regex = re.compile(f'{datayear*100+datamonth} PPVUL Magnastar Production Report.xls')

        file = list(filter(regex.match,os.listdir(self.datadir)))[0]
        LOGGER.debug(f'Processing Magnastar PPVUL with {datayear} {month_name[datamonth]} data...')
        df = pd.read_excel(os.path.join(self.datadir,file), sheet_name=0, engine='xlrd')

        df["SourceFileName"] = f"P-Magn-LIF-{self.year}-{self.monthApplied}-1"
        df["CarrierID"] = df["Carrier"].apply(lambda x: '1064' if x == 'JHL' else '1064' if x=='JHN' else '1111' if x=='PLN' else '1116' if x=='PR1' else '1111' if x=='PLA' else '' )
        df["MemFirmID"] = 0
        df["CarrierContracteeID"] = df["Agency #"]
        df["CarrierContracteeName"] = df["Agency Name"].str[0:50]
        df["ProductID"] = 0
        df["CarrierProductID"] = df["Product"]+'-'+df["CarrierID"]+'-'+df["Single/Joint"]
        df["YearApplied"] = self.year
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df["Agent #"]
        df["CarrierProducerName"] = df["Agent Name"].str[0:50]
        df["PolicyNumber"] = df["CellID"]
        df["IssueDate"] = df["Issue Date"]
        df["InsuredName"] = ""
        df["YTDAnnualizedPrem"] = df["Target Premium"]
        df["YTDAnnualizedLowNon"] = df["Excess Premium"]
        df["YTDFace"] = df["Amount DB"]
        df["RiskNumber"] = ""
        df["RiskName"] = ""
        df["Exchange1035"] = ""
        df["SplitPercentage"] = df["Split%"]
        df["Replacement"] = ""
        df["CarrierProductName"] = "Magnastar"
        df["PolicyOwner"] = ""
        df["Exchange"] = ""
        df["NLG"] = ""
        df["Modco"] = ""
        df["YRT"] = ""
        df["CSO2001"] = ""

        
        # Report
        PruTarget = df.loc[df["CarrierID"]=='1116',"YTDAnnualizedPrem"].sum()
        PruExcess = df.loc[df["CarrierID"]=='1116',"YTDAnnualizedLowNon"].sum()
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD Prudential = {PruTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD Prudential = {PruExcess}')

        PLTarget = df.loc[df["CarrierID"]=='1111',"YTDAnnualizedPrem"].sum()
        PLExcess = df.loc[df["CarrierID"]=='1111',"YTDAnnualizedLowNon"].sum()
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD Pacific Life = {PLTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD Pacific Life = {PLExcess}')

        LOGGER.info(f'Accessing flash workbook and updating values...')
        with xw.App(visible=False) as app:
            flash = xw.Book(self.flash)
            notesbreakdown = flash.sheets['Notes-Breakdown']
            notesbreakdown.range('Pru_Mag_target').value = PruTarget
            notesbreakdown.range('Pru_Mag_excess').value = PruExcess
            notesbreakdown.range('Pru_Mag_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('Pru_Mag_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('PL_Mag_target').value = PLTarget
            notesbreakdown.range('PL_Mag_excess').value = PLExcess           
            notesbreakdown.range('PL_Mag_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('PL_Mag_excess').api.Font.Color = rgb_to_int((0, 102, 204))
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


def processMag(year, month):
    '''To generate a YTD report, if report month's data is not available, previous month's data would be used temporarily until that month's data is arrived.'''
    mag = magnastar(year, month)
    isFileFound = mag.fetchrawmonth()
    if isFileFound == True:
        mag.getcsv(year, month)
    else:
        try:
            if month == 1:
                mag.getcsv(year-1, 12)
            else:
                mag.getcsv(year,month -1)
        except IndexError:
            LOGGER.warning(f'Magnastar {month_name[month-1]} file not available.')


# if __name__ == '__main__':
#     processMag(2022, 10)

