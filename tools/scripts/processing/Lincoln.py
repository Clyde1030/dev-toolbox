import openpyxl as xl
import xlwings as xw
from xlwings.utils import rgb_to_int
from calendar import month_name
import shutil
import pandas as pd
import logging
import re
import os

LOGGER = logging.getLogger(__name__)
# LOGGER = logging.getLogger()
# formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
# handler = logging.StreamHandler()
# handler.setFormatter(formatter)
# LOGGER.addHandler(handler)
# LOGGER.setLevel(logging.DEBUG)

class lincoln:

    def __init__(self,year,month,*csvdst):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        # self.df = df
        # self.dst = dst
        self.carrierId = 1743
        self.regex = re.compile(f"Production Report {self.monthApplied} {year}.xlsx",re.IGNORECASE) 
        self.YTDregex = re.compile("Production Report") 
        self.datadir = f'J:\Acctng\Production\{year}\Data\LFG'
        self.goanywhere = f'J:\Acctng\Revenue\goanywhere\Lincoln'
        self.csvname = f"P-1743-LIF-{year}-{self.monthApplied}-1.txt"
        self.csvdst = csvdst or f'C:\dev\Production\data\{year}\{self.monthApplied}'
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        self.exportCols = ["SourceFileName","CarrierID","MemFirmID","CarrierContracteeID","CarrierContracteeName",
                            "ProductID","CarrierProductID","YearApplied","MonthApplied","ProducerID",
                            "CarrierProducerID","CarrierProducerName","PolicyNumber","IssueDate","InsuredName",
                            "YTDAnnualizedPrem" ,"YTDAnnualizedLowNon","YTDFace","RiskNumber","RiskName",
                            "Exchange1035","SplitPercentage","Replacement","CarrierProductName","PolicyOwner",
                            "Exchange","NLG","Modco","YRT","CSO2001"]

    def _unprotect_wb(self, filepath, password):
        '''save over the original workbook without the password'''
        app = xw.App(visible=False)
        wb = xw.Book(filepath,password=password)
        wb.save(filepath,password='')
        wb.close()

    def fetchRawMonth(self, datayear=None, datamonth=None):

        isFileFound = True

        if not datayear:
            datayear = self.year
        if not datamonth:
            datamonth = self.month

        regex = re.compile(f"Production Report {datamonth} {datayear}.xlsx",re.IGNORECASE)

        LOGGER.info(f'Searching {month_name[datamonth]} file for carrierid {self.carrierId}...')
        path = []
        f = []
        for dirs, _, filenames in os.walk(self.goanywhere):
            f = [os.path.join(dirs,f) for f in filenames if re.search(regex,os.path.join(dirs,f))]
            path.extend(f)
        if len(path) == 0:
            isFileFound = False
            LOGGER.warning(f'1743 Lincoln {month_name[self.month]} production report not available.')
        else:
            for p in path:
                shutil.copy2(p,self.datadir)
            LOGGER.info(f'1743 Lincoln: {len(path)} {month_name[self.month]} raw file is copied to {self.datadir}.')
        return isFileFound


    def getcsv(self, datayear=None, datamonth=None):
        
        isFileFound = True

        if not datayear:
            datayear = self.year
        if not datamonth:
            datamonth = self.month

        regex = re.compile(f"Production Report {datamonth} {datayear}.xlsx",re.IGNORECASE)

        file = list(filter(regex.match,os.listdir(self.datadir)))[0]
        wb = os.path.join(self.datadir,file)
        
        LOGGER.debug(f'decrypting Lincoln raw file...')
        self._unprotect_wb(wb, password='production') 

        # Read in the file and cleaning missing data for some columns
        LOGGER.debug(f'Processing Lincoln with {datayear} {month_name[datamonth]} data...')
        df = pd.read_excel(wb, sheet_name='Data',header=3, converters={'Agent SSN':str,'Product Type':str}, engine='openpyxl')
        df = df[pd.isnull(df['Policy Number'])==False]

        ## ETL Logic ##
        df["SourceFileName"] = f"P-1743-LIF-{self.year}-{self.monthApplied}-1"
        df["CarrierID"] = self.carrierId
        df["MemFirmID"] = ""
        df["CarrierContracteeID"] = df['Producer Name'].str.lstrip()
        df["CarrierContracteeName"] = df['Producer Name'].str.rstrip()
        df["ProductID"] = 0
        df["CarrierProductID"] = df['Product Name'].str[0:50]
        df["YearApplied"] = self.year
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df['Agent SSN']
        df["CarrierProducerName"] = df['Producer Name'].str.rstrip()
        df["PolicyNumber"] = df['Policy Number']
        df["IssueDate"] = df['Issue Date']
        df["InsuredName"] = df['Insured Name']
        df["isAnnuity"] = df["Product Type"].apply(lambda x: True if x=='Annuity' else False)
        df["YTDAnnualizedPrem"] = df.apply(lambda x: 0 if x['isAnnuity']==True else x['ACTP'], axis=1)
        df["YTDAnnualizedLowNon"] = df.apply(lambda x: x["ACTP"]+x["Dump In Premium"] if x['isAnnuity']==True else x["Dump In Premium"], axis=1)
        df["YTDFace"] = df['Base Face Value'] 
        df["RiskNumber"] = ""
        df["RiskName"] = ""
        df["Exchange1035"] = df['Internal Replace Indicator'].apply(lambda x: 'Y' if x=='RO' else 'N')
        df["SplitPercentage"] = df['Policy Count']
        df["Replacement"] = ""
        df["CarrierProductName"] = df['Product Name'].str[0:50]
        df["PolicyOwner"] = ""
        df["Exchange"] = ""
        df["NLG"] = df.apply(lambda x: 'Y' if x["Product Type"]=='No Lapse' else 'N', axis=1)
        df["Modco"] =df.apply(lambda x: 'N' if x["Modco"]=='No' else '', axis=1)
        df["YRT"] = ""
        df["CSO2001"] = ""

        # Report
        AnnuityTarget = 0
        AnnuityExcess = df.loc[df["Product Type"].str.contains('Annuity')==True,"ACTP"].sum()
        LifeTarget = df["YTDAnnualizedPrem"].sum()
        LifeExcess = df["YTDAnnualizedLowNon"].sum()-AnnuityExcess
        LTCTarget = df.loc[df["CarrierProductID"].str.contains('MoneyGuard|MG',case = False)==True,"YTDAnnualizedPrem"].sum()
        LTCExcess = df.loc[df["CarrierProductID"].str.contains('MoneyGuard|MG',case = False)==True,"YTDAnnualizedLowNon"].sum()

        LOGGER.info(f'Target Premium {month_name[self.month]} YTD Life = {LifeTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD Life = {LifeExcess}')
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD Annuity = {AnnuityTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD Annuity = {AnnuityExcess}')
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD LTC Hybrid = {LTCTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD LTC Hybrid = {LTCExcess}')
        
        LOGGER.info(f'Accessing flash workbook and updating values...')
        with xw.App(visible=False) as app:
            flash = xw.Book(self.flash)
            notesbreakdown = flash.sheets['Notes-Breakdown']
            notesbreakdown.range('LFG_lif_target').value = LifeTarget
            notesbreakdown.range('LFG_lif_excess').value = LifeExcess
            notesbreakdown.range('LFG_Ann_excess').value = AnnuityExcess
            notesbreakdown.range('LFG_LTC_target').value = LTCTarget
            notesbreakdown.range('LFG_LTC_excess').value = LTCExcess
            notesbreakdown.range('LFG_lif_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('LFG_lif_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('LFG_Ann_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('LFG_LTC_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('LFG_LTC_excess').api.Font.Color = rgb_to_int((0, 102, 204))
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


def processlfg(year, month):
    '''To generate a YTD report, if report month's data is not available, previous month's data would be used 
    temporarily until that month's data is arrived.'''
    lin = lincoln(year, month)    
    isFileFound = lin.fetchRawMonth(year, month)
    if isFileFound == True:
        lin.getcsv(year, month)
    else:
        try:
            if month == 1:
                lin.getcsv(year-1, 12)
            else:
                lin.getcsv(year,month -1)
        except IndexError:
            LOGGER.warning(f'Nationwide {month_name[month-1]} file not available.')


# if __name__ == '__main__':
#     processlfg(2022,11)










########################################################################################################################################
    # def _decrypt_wb(self, filepath, password):
    #     '''return a decrypted workbook that is readable for pandas'''
    #     import io
    #     import msoffcrypto as mso # Python tool and library for decrypting encrypted MS Office files with password, intermediate key, or private key which generated its escrow key.
    #     decrypted_wb = io.BytesIO()
    #     with open(filepath,'rb') as file:
    #         office_file = mso.OfficeFile(file)
    #         office_file.load_key(password=password)
    #         office_file.decrypt(decrypted_wb)
    #     return decrypted_wb
# wb = self._decrypt_wb(wb,'production')