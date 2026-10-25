import pandas as pd
from calendar import month_name, monthrange
from datetime import datetime
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

class johnhancock:

    def __init__(self, year, month, *csvdst):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        self.csvdst = csvdst or f'C:\dev\Production\data\{year}\{self.monthApplied}'
        self.carrierId = 1064
        self.regexNY = re.compile('ff_rd_mgroup_extract_ny_\d{4}-\d{2}-\d{2}.txt')  
        self.regexUS = re.compile('ff_rd_mgroup_extract_us_\d{4}-\d{2}-\d{2}(.txt)$')
        self.regex401 = re.compile(f'{month_name[month]} {year} sales report.xlsx')  
        self.datadir = f'J:\Acctng\Production\{year}\Data\JH'
        self.goanywhere =  f'J:\Acctng\Revenue\goanywhere\John Hancock'
        self.csvnameUS = f"P-1064-LIF-{self.year}-{self.monthApplied}-1.txt"
        self.csvnameNY = f"P-1064-LIF-{self.year}-{self.monthApplied}-2.txt"        
        self.csvname401e = f"P-1064-401-{self.year}-{self.monthApplied}-1.txt"
        self.csvname401s = f"P-1064-401-{self.year}-{self.monthApplied}-2.txt"
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm')        
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
        '''JH sends production report through GoAnywhere every week. This method is designed to choose the most suitable 
        YTD report based on the date that report is received. Although it doesn't really matter what date we actually
        use since it is YTD sales, the report date ideally should be close to 30th or 31st of the month as much as possible.
        Here the first day after the end of month is used. The function return the file date chosen.'''        
        if not year:
            year = self.year
        if not month:
            month = self.month

        monthend = f'{year}{str(month).zfill(2)}{monthrange(self.year,self.month)[1]}'
        Path = dict()
        file = []
        if month == 12:
            for dirs, _, filenames in os.walk(os.path.join(self.goanywhere,f'{year+1}','01')):
                file.extend([os.path.join(dirs,f) for f in filenames if re.search(self.regexUS,os.path.join(dirs,f))])            
        else:
            for dirs, _, filenames in os.walk(os.path.join(self.goanywhere,f'{year}',str(month+1).zfill(2))):
                file.extend([os.path.join(dirs,f) for f in filenames if re.search(self.regexUS,os.path.join(dirs,f))])            
        for p in file:
            date = os.path.basename(p).split('_')[5].replace('.txt','').replace('-','')
            Path[date]={}
            Path[date]['path'] = p
            Path[date]['diff'] = self._days_between(date,monthend)
        group = []
        for d in Path.keys():
            if int(d)>int(monthend):
                group.append(d)
        group.sort()
        uspath = None 
        try:
            uspath = Path[group[0]]['path']
            LOGGER.debug(r'The closest file date found is {}'.format(group[0]))
        except IndexError:
            LOGGER.warning(f'File not available for {month_name[self.month]} yet.')
        except Exception as e:
            LOGGER.warning(e)
        
        return uspath

    def getrawmonth(self, year, month):
        isFileFound = True
        LOGGER.info(f'Searching {month_name[self.month]} file for carrierid {self.carrierId}...')
        uspath = self._chooseYTDFile(year, month)
        if uspath != None:
            # date = os.path.basename(uspath).split('_')[5].replace('.txt','')
            # LOGGER.debug(r'The closest file date found is {}'.format(date))
            nypath = uspath.replace('us','ny')
            for file in [uspath,nypath]:
                if os.path.exists(os.path.join(self.datadir,str(month).zfill(2),os.path.basename(file)))==False:
                    shutil.copy2(file,os.path.join(self.datadir,str(month).zfill(2)))
                    LOGGER.info(f'{os.path.basename(file)} is copied to {month_name[self.month]} data directory')
                else:
                    LOGGER.info(f'{os.path.basename(file)} is already in data directory.')            
        else:
            isFileFound = False
        return isFileFound

    def getcsvus(self, year=None, month=None):
        # This function generates csv files for us life business, there should only be one us file in data directory after running getrawmonth

        if not year:
            year = self.year
        if not month:
            month = self.month
        
        # filter the listdit based on the regex to find US file name
        file = list(filter(self.regexUS.match,os.listdir(os.path.join(self.datadir,str(month).zfill(2)))))[0]        
        LOGGER.info(f'Processing JH life - US with {year} {month_name[month]} data...')
        
        df = pd.read_csv(os.path.join(self.datadir,str(month).zfill(2),file),delimiter=',')
        # df.to_excel(os.path.join(self.datadir,'test1.xlsx'),index=False)        
        df = df[df['DURATION'] < 3]
        df["SourceFileName"] = f"P-1064-LIF-{self.year}-{self.monthApplied}-1"
        df["CarrierID"] = self.carrierId
        df["MemFirmID"] = 0
        df["CarrierContracteeID"] = df['Agency_Number']
        df["CarrierContracteeName"] = df['Agency_Name'].str.rstrip()
        df["ProductID"] = 0
        df["CarrierProductID"] = df['Product_Code'].apply(lambda x: '{0:0>10}'.format(x))
        df["YearApplied"] = str(self.year)
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df['Agent_Number']
        df["CarrierProducerName"] = df['Agent_Name'].str.rstrip()
        df["PolicyNumber"] = df['Policy_Number']
        df["IssueDate"] = df['Issue_Date']
        df["InsuredName"] = df['Insured_Name'].str.rstrip()
        df["YTDAnnualizedPrem"] = df['Target_Premium']
        df["YTDAnnualizedLowNon"] = df['Excess_Premium']
        df["YTDFace"] = df['Face_Amount'].astype(float).round(2).map("{:,.2f}".format)
        df["RiskNumber"] = ""
        df["RiskName"] = ""
        df["Exchange1035"] = df['1035']
        df["SplitPercentage"] = df['Policy_Count']
        df["Replacement"] = ""
        df["CarrierProductName"] = df['Product_Name'].str.rstrip()
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
            notesbreakdown.range('JH_lif_target_us').value = target # notesbreakdown['C7'].value = target
            notesbreakdown.range('JH_lif_excess_us').value = excess # notesbreakdown['D7'].value = excess
            notesbreakdown.range('JH_lif_target_us').api.Font.Color = rgb_to_int((0, 102, 204)) # notesbreakdown.range('C7:D7').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('JH_lif_excess_us').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            # flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        # Export
        df["YTDAnnualizedPrem"] = df['Target_Premium'].astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df['Excess_Premium'].astype(float).round(2).map("{:,.2f}".format)
        df = df[self.exportCols]
        df.to_csv(os.path.join(self.csvdst,self.csvnameUS), index=False, sep = '|')
        LOGGER.info(f'{self.csvnameUS} is saved at {self.csvdst}.')


    def getcsvny(self, year=None, month=None):

        if not year:
            year = self.year
        if not month:
            month = self.month

        # filter the listdit based on the regex to find US file name
        file = list(filter(self.regexNY.match,os.listdir(os.path.join(self.datadir,str(month).zfill(2)))))[0]        

        LOGGER.info(f'Processing JH life - NY with {year} {month_name[month]} data...')
        df = pd.read_csv(os.path.join(self.datadir,str(month).zfill(2),file),delimiter=',')

        df = df[df['DURATION'] < 3]
        df["SourceFileName"] = f"P-1064-LIF-{self.year}-{self.monthApplied}-2"
        df["CarrierID"] = self.carrierId
        df["MemFirmID"] = 0
        df["CarrierContracteeID"] = df['Agency_Number']
        df["CarrierContracteeName"] = df['Agency_Name'].str.rstrip()
        df["ProductID"] = 0
        df["CarrierProductID"] = df['Product_Code'].apply(lambda x: '{0:0>10}'.format(x))
        df["YearApplied"] = str(self.year)
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df['Agent_Number']
        df["CarrierProducerName"] = df['Agent_Name'].str.rstrip()
        df["PolicyNumber"] = df['Policy_Number']
        df["IssueDate"] = df['Issue_Date']
        df["InsuredName"] = df['Insured_Name'].str.rstrip()
        df["YTDAnnualizedPrem"] = df['Target_Premium']
        df["YTDAnnualizedLowNon"] = df['Excess_Premium']
        df["YTDFace"] = df['Face_Amount'].astype(float).round(2).map("{:,.2f}".format)
        df["RiskNumber"] = ""
        df["RiskName"] = ""
        df["Exchange1035"] = df['1035']
        df["SplitPercentage"] = df['Policy_Count']
        df["Replacement"] = ""
        df["CarrierProductName"] = df['Product_Name'].str.rstrip()
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
            notesbreakdown.range('JH_lif_target_ny').value = target # notesbreakdown['C9'].value = target
            notesbreakdown.range('JH_lif_excess_ny').value = excess # notesbreakdown['D9'].value = excess
            notesbreakdown.range('JH_lif_target_ny').api.Font.Color = rgb_to_int((0, 102, 204)) # notesbreakdown.range('C9:D9').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('JH_lif_excess_ny').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            # flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        # Export
        df["YTDAnnualizedPrem"] = df['Target_Premium'].astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df['Excess_Premium'].astype(float).round(2).map("{:,.2f}".format)
        df = df[self.exportCols]
        df.to_csv(os.path.join(self.csvdst,self.csvnameNY), index=False, sep = '|')
        LOGGER.info(f'{self.csvnameNY} is saved at {self.csvdst}.')


    def getcsv401(self, year=None, month=None):

        if not year:
            year = self.year
        if not month:
            month = self.month

        # filter the listdit based on the regex to find 401K file name
        file = os.path.join(self.datadir,str(month).zfill(2),f'{month_name[month]} {year} sales report.xlsx')
        LOGGER.info(f"Processing JH 401K with {year} {month_name[month]} data...")

        df = pd.read_excel(os.path.join(self.datadir,str(month).zfill(2),file), sheet_name='Sales', engine='openpyxl')

        df["SourceFileName"] = df.apply(lambda x: f"P-1064-401-{self.year}-{self.monthApplied}-2" if x['Product Type']=='Signature' else f"P-1064-401-{self.year}-{self.monthApplied}-1",axis=1) 
        df["CarrierID"] = self.carrierId
        df["MemFirmID"] = 0
        df["CarrierContracteeID"] = 'M Holdings Securities, Inc.'
        df["CarrierContracteeName"] = 'M Holdings Securities, Inc.'
        df["ProductID"] = 0
        df["CarrierProductID"] = df['Product Type'].str.rstrip()
        df["YearApplied"] = self.year
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df['Financial Rep Name'].str.rstrip()
        df["CarrierProducerName"] = df['Financial Rep Name'].str.rstrip()
        df["PolicyNumber"] = df['Contract Number']
        df["IssueDate"] = ""
        df["InsuredName"] = df['Contract Name'].str.rstrip().str[0:50]
        df["YTDAnnualizedPrem"] = df['Total Production']
        df["YTDAnnualizedLowNon"] = 0
        df["YTDFace"] = ""
        df["RiskNumber"] = ""
        df["RiskName"] = ""
        df["Exchange1035"] = ""
        df["SplitPercentage"] = ""
        df["Replacement"] = ""
        df["CarrierProductName"] = df['Product Type'].str.rstrip() 
        df["PolicyOwner"] = df['Contract Name'].str.rstrip().str[0:50]
        df["Exchange"] = ""
        df["NLG"] = ""
        df["Modco"] = ""
        df["YRT"] = ""
        df["CSO2001"] = ""

        # Report 
        targete = df.loc[df['CarrierProductID']=='Enterprise',"YTDAnnualizedPrem"].sum()
        excessus = 0
        targets = df.loc[df['CarrierProductID']=='Signature',"YTDAnnualizedPrem"].sum()
        excessny = 0

        LOGGER.info(f'Target Premium {month_name[self.month]} Enterprise = {targete}')
        LOGGER.info(f'Target Premium {month_name[self.month]} Signature = {targets}')
        LOGGER.info(f'Accessing flash workbook and updating values...')

        with xw.App(visible=False) as app:
            flash = xw.Book(self.flash)
            notesbreakdown = flash.sheets['Notes-Breakdown']
            notesbreakdown.range('JH_401_Ent').value = targete
            notesbreakdown.range('JH_401_Sig').value = targets
            notesbreakdown.range('JH_401_Ent').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('JH_401_Sig').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        # # # Export
        df["YTDAnnualizedPrem"] = df['Total Production'].astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = 0

        df = df[self.exportCols]

        df[df["SourceFileName"]==f"P-1064-401-{self.year}-{self.monthApplied}-1"].to_csv(os.path.join(self.csvdst,self.csvname401e), index=False, sep = '|')
        df[df["SourceFileName"]==f"P-1064-401-{self.year}-{self.monthApplied}-2"].to_csv(os.path.join(self.csvdst,self.csvname401s), index=False, sep = '|')

        LOGGER.info(f'{self.csvname401e} is saved at {self.csvdst}.')
        LOGGER.info(f'{self.csvname401s} is saved at {self.csvdst}.')



def processJH(year, month):
    jh = johnhancock(year,month)
    
    # Life 
    isFileFound = jh.getrawmonth(year, month)
    if isFileFound == True:
        jh.getcsvus(year,month)
        jh.getcsvny(year,month)
    else: 
        try:
            if month == 1:
                jh.getcsvus(year-1, 12)
                jh.getcsvny(year-1, 12)
            else:
                jh.getcsvus(year,month -1)
                jh.getcsvny(year,month -1)
        except IndexError:
            LOGGER.warning(f'John Hancock Life {month_name[month-1]} file not available.')
        
    # 401K
    try:
        jh.getcsv401(year, month)
    except FileNotFoundError:
        LOGGER.warning(f'John Hancock {month_name[month]} 401K data not available.')
        try: 
            if month == 1:
                jh.getcsv401(year-1, 12)
            else:
                jh.getcsv401(year,month-1)
        except FileNotFoundError:
            LOGGER.warning(f'John Hancock {month_name[month-1]} 401K data not available.')

# if __name__ == '__main__':
#     jh = johnhancock(2022,12)
#     jh.getcsvus(2022,12)
#     jh.getcsvny(2022,12)
#     jh.getcsv401(2022,12)
