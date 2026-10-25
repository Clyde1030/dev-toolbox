import pandas as pd
from calendar import month_name, monthrange
from datetime import datetime
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

class pacificlife:
    def __init__(self, year, month, *csvdst):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        self.csvdst = csvdst or f'C:\dev\Production\data\{year}\{self.monthApplied}'
        self.carrierId = 1111
        self.regexann = re.compile('M_FINANCIAL_RPT1_\d{8}.txt')  
        self.regexlif = re.compile('MPAIDPEN_WEEKLY_\d{8}.txt')  
        # self.regexlif = f'MPAIDPEN_WEEKLY_{year}{str(month+1).zfill(2)}'+'\d{2}.txt'         
        self.datadir = f'J:\Acctng\Production\{year}\Data\PL'
        self.goanywhere =  f'J:\Acctng\Revenue\goanywhere\Pacific Life'
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        self.csvnamelif = f"P-1111-LIF-{self.year}-{self.monthApplied}-1.txt"
        self.csvnameann = f"P-1111-ANN-{self.year}-{self.monthApplied}-2.txt"
        self.exportCols = ["SourceFileName","CarrierID","MemFirmID","CarrierContracteeID","CarrierContracteeName",
                            "ProductID","CarrierProductID","YearApplied","MonthApplied","ProducerID",
                            "CarrierProducerID","CarrierProducerName","PolicyNumber","IssueDate","InsuredName",
                            "YTDAnnualizedPrem" ,"YTDAnnualizedLowNon","YTDFace","RiskNumber","RiskName",
                            "Exchange1035","SplitPercentage","Replacement","CarrierProductName","PolicyOwner",
                            "Exchange","NLG","Modco","YRT","CSO2001"]

    def _days_between(self, d1, d2):
        # return the day difference between two given dates
        d1 = datetime.strptime(d1, "%Y%m%d")
        d2 = datetime.strptime(d2, "%Y%m%d")
        return abs((d2 - d1).days)

    def _chooseYTDFile(self, year=None, month=None):
        # PL only sends YTD life production report through GoAnywhere every week. This method is designed to choose the most suitable 
        # YTD report based on the date that report is received. Although it doesn't really matter what date we actually
        # use since it is YTD sales, the report date ideally should be close to 30th or 31st as much as possible.
        # Here the first day after the end of month is used. The function return the file date choosen.
        if not year:
            year = self.year
        if not month:
            month = self.month

        monthend = f'{year}{str(month).zfill(2)}{monthrange(self.year,self.month)[1]}' # Determine the last day of the month in YYYYMMDD
        file = []
        
        # Create a collection of PL life raw data paths. If it is December, look over January next year
        if self.month == 12:
            # for dirs, _, filenames in os.walk(os.path.join(self.goanywhere,f'{self.year}','12')):
            #     file.extend([os.path.join(dirs,f) for f in filenames if re.search(self.regexlif,os.path.join(dirs,f))])
            for dirs, _, filenames in os.walk(os.path.join(self.goanywhere,f'{year+1}','01')):
                file.extend([os.path.join(dirs,f) for f in filenames if re.search(self.regexlif,os.path.join(dirs,f))])            
        else:
            # for dirs, _, filenames in os.walk(os.path.join(self.goanywhere,f'{self.year}',f'{self.monthApplied}')):
            #     file.extend([os.path.join(dirs,f) for f in filenames if re.search(self.regexlif,os.path.join(dirs,f))])
            for dirs, _, filenames in os.walk(os.path.join(self.goanywhere,f'{year}',str(month+1).zfill(2))):
                file.extend([os.path.join(dirs,f) for f in filenames if re.search(self.regexlif,os.path.join(dirs,f))])            
        
        # A dictionary that contain the date file is received, file path and the date difference from last day of the month
        Path = dict()        
        for p in file:
            date = os.path.basename(p).split('_')[2].replace('.txt','')
            Path[date]={}
            Path[date]['path'] = p
            Path[date]['diff'] = self._days_between(date,monthend)

        # A group for date that satisfy the criteria: the nearest day after end of month
        group = []
        for d in Path.keys():
            if int(d)>int(monthend):
                group.append(d)
        group.sort()
        lifpath = None 
        try:
            lifpath = Path[group[0]]['path']
            LOGGER.debug(r'The closest file date found is {}'.format(group[0]))
        except IndexError:
            LOGGER.warning(f'File not available for {month_name[self.month]} yet.')
        except Exception as e:
            LOGGER.warning(e)
        return lifpath

    def getrawmonthlif(self, year, month):
        isFileFound = True
        LOGGER.info(f'Searching {month_name[self.month]} file for carrierid {self.carrierId}...')
        lifpath = self._chooseYTDFile(year, month)
        if lifpath != None:
            # date = os.path.basename(lifpath).split('_')[2].replace('.txt','') #MPAIDPEN_WEEKLY_20220101.txt
            if os.path.exists(os.path.join(self.datadir,str(month).zfill(2),os.path.basename(lifpath)))==False:
                shutil.copy2(lifpath,os.path.join(self.datadir,str(month).zfill(2)))
                LOGGER.info(f'{os.path.basename(lifpath)} is copied to {month_name[self.month]} data directory')
            else:
                LOGGER.info(f'{os.path.basename(lifpath)} is already in data directory.')            
        else:
            isFileFound = False
        return isFileFound

    def getcsvlif(self, year=None, month=None):
        if not year:
            year = self.year
        if not month:
            month = self.month

        # filter the listdit based on the regex to find us file name
        file = list(filter(self.regexlif.match,os.listdir(os.path.join(self.datadir,str(month).zfill(2)))))[0]
        LOGGER.info(f'Processing PL life with {year} {month_name[month]} data...')
        df = pd.read_csv(os.path.join(self.datadir,str(month).zfill(2),file),sep=';', header=None)

        df = df[~(df.iloc[:,22].isin([0]) & df.iloc[:,23].isin([0]))]
        df = df[(df[15].str.contains('MAGNASTAR')==False)|pd.isnull(df[15])==True] # 

        df["SourceFileName"] = f"P-1111-LIF-{self.year}-{self.monthApplied}-1"
        df["CarrierID"] = self.carrierId
        df["MemFirmID"] = 0
        df["CarrierContracteeID"] = df[4].str.rstrip()
        df["CarrierContracteeName"] = df[5].str.rstrip()
        df["ProductID"] = 0
        df["CarrierProductID"] = df[11]
        df["YearApplied"] = self.year
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df[6]
        df["CarrierProducerName"] = df[7]
        df["PolicyNumber"] = df[10]
        df["IssueDate"] = pd.to_datetime(df[12].astype(str), format='%m%d%Y').dt.normalize()
        df["InsuredName"] = df[28].str.cat(df[29].str[:8], sep = " - ", na_rep = "")
        df["YTDAnnualizedPrem"] = df[22]#.astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df[23]#.astype(float).round(2).map("{:,.2f}".format)
        df["YTDFace"] = df[24]#.astype(float).round(2).map("{:,.2f}".format)
        df["RiskNumber"] = df[30]
        df["RiskName"] = df[31].str[0:50]
        df["Exchange1035"] = df[13]
        df["SplitPercentage"] = df[25]
        df["Replacement"] = ""
        df["CarrierProductName"] = df[15].str.cat(df[16], sep=" - ", na_rep = "")
        df["PolicyOwner"] = df[28]
        df["Exchange"] = df[14]
        df["NLG"] = df[26].str[:1]
        df["Modco"] = df[27].str[:1]
        df["YRT"] = ""
        df["CSO2001"] = df[32].str[:1]

        # Report
        MagTarget = df.loc[df["PolicyNumber"].str.contains('VM',case = False)==True,"YTDAnnualizedPrem"].sum()
        MagExcesss = df.loc[df["PolicyNumber"].str.contains('VM',case = False)==True,"YTDAnnualizedLowNon"].sum()
        LTCTarget = df.loc[df[15].str.contains('PremierCare Adv',case = False)==True,"YTDAnnualizedPrem"].sum()
        LTCExcesss = df.loc[df[15].str.contains('PremierCare Adv',case = False)==True,"YTDAnnualizedLowNon"].sum()
        target = df["YTDAnnualizedPrem"].sum()
        excess = df["YTDAnnualizedLowNon"].sum()
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD = {target}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD = {excess}')
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD MAG = {MagTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD MAG = {MagExcesss}')
        LOGGER.info(f'Target Premium {month_name[self.month]} YTD LTC = {LTCTarget}')
        LOGGER.info(f'Excess Premium {month_name[self.month]} YTD LTC = {LTCExcesss}')

        LOGGER.info(f'Accessing flash workbook and updating values...')
        with xw.App(visible=False) as app:
            flash = xw.Book(self.flash)
            notesbreakdown = flash.sheets['Notes-Breakdown']
            notesbreakdown.range('C67').value = LTCTarget
            notesbreakdown.range('D67').value = LTCExcesss
            notesbreakdown.range('PL_lif_target').value = target
            notesbreakdown.range('PL_lif_excess').value = excess
            notesbreakdown.range('C67').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('D67').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('PL_lif_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('PL_lif_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        # Export
        df["YTDAnnualizedPrem"] = df['YTDAnnualizedPrem'].astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df['YTDAnnualizedLowNon'].astype(float).round(2).map("{:,.2f}".format)
        df = df[self.exportCols]
        df.to_csv(os.path.join(self.csvdst,self.csvnamelif), index=False, sep = '|')
        # df.to_excel(os.path.join(self.csvdst,'test.xlsx'), index=False)
        LOGGER.info(f'{self.csvnamelif} is saved at {self.csvdst}.')



    def getcsvann(self, year=None, month=None):

        if not year:
            year = self.year
        if not month:
            month = self.month

        # filter the listdit based on the regex to find us file name
        file = list(filter(self.regexann.match,os.listdir(os.path.join(self.datadir,str(month).zfill(2)))))[0]
        LOGGER.info(f'Processing PL Annuity with {year} {month_name[month]} data...')
        df = pd.read_csv(os.path.join(self.datadir,str(month).zfill(2),file), delimiter=";", converters={"Agency Number":str,"Policy Number":str})

        ## ETL Logic ##
        df["SourceFileName"] = f"P-1111-ANN-{self.year}-{self.monthApplied}-1"
        df["CarrierID"] = self.carrierId
        df["MemFirmID"] = 0
        df["CarrierContracteeID"] = df["Agency Number"]
        df["CarrierContracteeName"] = df["Agency Name"]
        df["ProductID"] = 0
        df["CarrierProductID"] = df["Product Code"]
        df["YearApplied"] = self.year
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df["Agent Number"]
        df["CarrierProducerName"] = df["Agent Name"]
        df["PolicyNumber"] = df["Policy Number"]
        df["IssueDate"] = df["Issue Date"].str.rstrip()
        df["InsuredName"] = df["Annuitant Name"]
        df["YTDAnnualizedPrem"] = df["M Sales"].replace({'\$': '', ',': ''}, regex=True).astype(float)
        df["YTDAnnualizedLowNon"] = 0
        df["YTDFace"] = 0
        df["RiskNumber"] = ""
        df["RiskName"] = ""
        df["Exchange1035"] = ""
        df["SplitPercentage"] = df["Policy Count "]
        df["Replacement"] = ""
        df["CarrierProductName"] = df["Product Name"]
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
            notesbreakdown['PL_Ann_target'].value = target
            notesbreakdown['PL_Ann_excess'].value = excess
            notesbreakdown.range('PL_Ann_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('PL_Ann_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        # Export
        df["YTDAnnualizedPrem"] = df['YTDAnnualizedPrem'].astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df['YTDAnnualizedLowNon'].astype(float).round(2).map("{:,.2f}".format)
        df = df[self.exportCols]
        df.to_csv(os.path.join(self.csvdst, self.csvnameann), sep='|',index=False)
        # df.to_excel(os.path.join(self.csvdst,'test.xlsx'), index=False)
        LOGGER.info(f'{self.csvnameann} is saved at {self.csvdst}.')

def processPL(year, month):
    pl = pacificlife(year, month)
    # Life 
    isFileFound = pl.getrawmonthlif(year, month)
    if isFileFound == True:
        pl.getcsvlif(year,month)
    else: 
        try:
            if month == 1:
                pl.getcsvlif(year-1, 12)
            else:
                pl.getcsvlif(year,month -1)
        except IndexError:
            LOGGER.warning(f'Pacific Life {month_name[month-1]} Life file not available.')
        
    # Annuity
    try:
        pl.getcsvann(year, month)
    except FileNotFoundError:
        LOGGER.warning(f'Pacific Life {month_name[month]} Annuity data not available.')
        try: 
            if month == 1:
                pl.getcsvann(year-1, 12)
            else:
                pl.getcsvann(year,month-1)
        except FileNotFoundError:
            LOGGER.warning(f'Pacific Life {month_name[month-1]} Annuity data not available.')


# if __name__ == '__main__':
#     pl = pacificlife(2022, 12)
#     pl.getcsvlif()
#     pl.getcsvann()