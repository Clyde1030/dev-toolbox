import pandas as pd
from calendar import month_name, monthrange
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

class prudential:
    def __init__(self, year, month, *csvdst):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        self.csvdst = csvdst or f'C:\dev\Production\data\{year}\{self.monthApplied}'
        self.carrierId = 1116
        self.regexlif = re.compile(f'Prudential {year}.*(YTD Production.xlsx)$')  
        self.regexann = re.compile(f'M-M Financial Sales by Firm by IP {self.monthApplied}-{year}.xlsx')  
        self.datadir = f'J:\Acctng\Production\{year}\Data\Prudential'
        self.goanywhere =  f'J:\Acctng\Revenue\goanywhere\Prudential'
        self.csvnamelif = f"P-1116-LIF-{year}-{self.monthApplied}-1.txt"
        self.csvnameann = f"P-1116-ANN-{year}-{self.monthApplied}-1.txt"
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        self.exportCols = ["SourceFileName","CarrierID","MemFirmID","CarrierContracteeID","CarrierContracteeName",
                            "ProductID","CarrierProductID","YearApplied","MonthApplied","ProducerID",
                            "CarrierProducerID","CarrierProducerName","PolicyNumber","IssueDate","InsuredName",
                            "YTDAnnualizedPrem" ,"YTDAnnualizedLowNon","YTDFace","RiskNumber","RiskName",
                            "Exchange1035","SplitPercentage","Replacement","CarrierProductName","PolicyOwner",
                            "Exchange","NLG","Modco","YRT","CSO2001"]


    def getcsvlif(self, datayear=None, datamonth=None):

        if not datayear:
            datayear = self.year
        if not datamonth:
            datamonth = self.month

        file = list(filter(self.regexlif.match,os.listdir(os.path.join(self.datadir,str(datamonth).zfill(2)))))[0]

        LOGGER.info(f'Processing Prudential Life with {datayear} {month_name[datamonth]} data...')

        df = pd.read_excel(os.path.join(self.datadir, str(datamonth).zfill(2),file),converters={"AGENCY NUMBER":str}, engine='openpyxl')
       

        ## ETL Logic ##
        df["SourceFileName"] = f"P-1116-LIF-{self.year}-{self.monthApplied}-1"
        df["CarrierID"] = self.carrierId
        df["MemFirmID"] = 0
        df["CarrierContracteeID"] = df["AGENCY NUMBER"]
        df["CarrierContracteeName"] = df["AGENCY NAME"].str.rstrip()
        df["ProductID"] = 0
        df["CarrierProductID"] = df["Product Code"]
        df["YearApplied"] = self.year
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df["AGENT NUMBER"]
        df["CarrierProducerName"] = df["AGENT NAME"]
        df["PolicyNumber"] = df["POLICY_ID"].str.rstrip()
        df["IssueDate"] = df["ISSUE DATE"]
        df["InsuredName"] = df["INSURED NAME"]
        df["YTDAnnualizedPrem"] = df["TARGET PREMIUM"].fillna(0) #.astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df["EXCESS PREMIUM"].fillna(0) #.astype(float).round(2).map("{:,.2f}".format)
        df["YTDFace"] = df["FACE AMOUNT"].fillna(0).astype(float).round(2).map("{:,.2f}".format)
        df["RiskNumber"] = ""
        df["RiskName"] = ""
        df["Exchange1035"] = df["EXCHANGE 1035"].fillna('N')
        df["SplitPercentage"] = df["POLICY COUNT"]
        df["Replacement"] = ""
        df["CarrierProductName"] = df["Product Name"]+'-'+df["Issue State"].fillna('')
        df["PolicyOwner"] =df["Policy Owner"].fillna('')
        df["Exchange"] = df["INTERNAL EXCHANGE"].fillna('')
        df["NLG"] = df["NO LAPSE GUARANTEE"].fillna('')
        df["Modco"] = df["Modco Reinsurance Treaty"].apply(lambda x: 'Y' if x == 'Modco' else ('' if pd.isnull(x)==True else 'N')) 
        df["YRT"] = df["Modco Reinsurance Treaty"].apply(lambda x: 'Y' if x == 'YRT' else ('' if pd.isnull(x)==True else 'N')) 
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
            notesbreakdown.range('Pru_lif_target').value = target
            notesbreakdown.range('Pru_lif_excess').value = excess
            notesbreakdown.range('Pru_lif_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('Pru_lif_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        # Export
        df["YTDAnnualizedPrem"] = df['YTDAnnualizedPrem'].astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df['YTDAnnualizedLowNon'].astype(float).round(2).map("{:,.2f}".format)
        df = df[self.exportCols]
        df.to_csv(os.path.join(self.csvdst,self.csvnamelif), index=False, sep = '|')
        # df.to_excel(os.path.join(self.csvdst, 'test.xlsx'), index = False)
        LOGGER.info(f'{self.csvnamelif} is saved at {self.csvdst}.')



    def getcsvann(self, datayear=None, datamonth=None):

        if not datayear:
            datayear = self.year
        if not datamonth:
            datamonth = self.month

        regexann = re.compile(f'M-M Financial Sales by Firm by IP {str(datamonth).zfill(2)}-{datayear}.xlsx')

        file = list(filter(regexann.match,os.listdir(os.path.join(self.datadir,str(datamonth).zfill(2)))))[0]

        LOGGER.info(f'Processing Prudential Annuity with {datayear} {month_name[datamonth]} data...')

        df = pd.read_excel(os.path.join(self.datadir,str(datamonth).zfill(2),file), engine='openpyxl') #,converters={"Investment Professional key":str}

        # df = df.head(df.shape[0] -1) This is from the old version.
        df.iloc[:,8] = df.iloc[:,8].str.rstrip()
        df = df[~df.iloc[:,8].isin(['Prudential Defined Income','ASK','VIP'])]

        ## ETL Logic ##
        df["SourceFileName"] = f"P-1116-ANN-{self.year}-{self.monthApplied}-1"
        df["CarrierID"] = self.carrierId
        df["MemFirmID"] = 0
        df["CarrierContracteeID"] = df["FP Individual ID"] #df["Investment Professional key"]
        df["CarrierContracteeName"] = df["FP Full Name"] #IP Name
        df["ProductID"] = 0
        df["CarrierProductID"] = df["Parent Product"]
        df["YearApplied"] = self.year
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df["FP Individual ID"] #Investment Professional key
        df["CarrierProducerName"] = df["FP Full Name"] #IP Name
        df["PolicyNumber"] = df["Contract Number"] # Contract Num
        df["IssueDate"] = df["Contract Issue Date"]
        df["InsuredName"] = df["Contract Owner Full Name"].str.rstrip().str[0:50] #Owner Name
        df["YTDAnnualizedPrem"] = df["Transaction Amount"].str.replace(',','').astype(float).fillna(0) #.fillna(0) #.astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = 0
        df["YTDFace"] = 0
        df["RiskNumber"] = ""
        df["RiskName"] = ""
        df["Exchange1035"] = ""
        df["SplitPercentage"] = ""
        df["Replacement"] = ""
        df["CarrierProductName"] = df["Product Name"].str[0:50]
        df["PolicyOwner"] = ""
        df["Exchange"] = ""
        df["NLG"] = ""
        df["Modco"] = ""
        df["YRT"] = ""
        df["CSO2001"] = ""

        # Report
        target = df["YTDAnnualizedPrem"].sum()
        excess = df["YTDAnnualizedLowNon"].sum()
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD Annuity = {target}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD Annuity = {excess}')

        LOGGER.info(f'Accessing flash workbook and updating values...')        
        with xw.App(visible=False) as app:
            flash = xw.Book(self.flash)
            notesbreakdown = flash.sheets['Notes-Breakdown']
            notesbreakdown.range('Pru_Ann_target').value = target
            notesbreakdown.range('Pru_Ann_excess').value = excess
            notesbreakdown.range('Pru_Ann_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('Pru_Ann_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        # Export
        df["YTDAnnualizedPrem"] = df['YTDAnnualizedPrem'].astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df['YTDAnnualizedLowNon'].astype(float).round(2).map("{:,.2f}".format)
        df = df[self.exportCols]
        df.to_csv(os.path.join(self.csvdst,self.csvnameann), index=False, sep = '|')
        # df.to_excel(os.path.join(self.csvdst, 'test.xlsx'), index = False)

        LOGGER.info(f'{self.csvnameann} is saved at {self.csvdst}.')


def processPru(year, month):
    '''To generate a YTD report, if report month's data is not available, previous month's data would be used temporarily until that month's data is arrived.'''
    pru = prudential(year,month)
    
    try:
        pru.getcsvlif(year, month)
    except FileNotFoundError:
        LOGGER.warning(f'Prudential {month_name[month]} Life data not available.')
        try: 
            if month == 1:
                pru.getcsvlif(year-1, 12)
            else:
                pru.getcsvlif(year,month-1)
        except IndexError:
            LOGGER.warning(f'Prudential {month_name[month-1]} Life data not available.')
    
    try:
        pru.getcsvann(year, month)
    except FileNotFoundError:
        LOGGER.warning(f'Prudential {month_name[month]} Annuity data not available.')
        try: 
            if month == 1:
                pru.getcsvann(year-1, 12)
            else:
                pru.getcsvann(year,month-1)
        except IndexError:
            LOGGER.warning(f'Prudential {month_name[month-1]} Annuity data not available.')
            

# if __name__ == '__main__':
#     processPru(2022, 10)



    


        

    


        
