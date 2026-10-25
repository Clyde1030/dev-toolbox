import pandas as pd
from calendar import month_name
from xlwings.utils import rgb_to_int
import xlwings as xw
import shutil
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

class protective:

    def __init__(self, year, month, *csvdst):
        self.year = year
        self.month = month
        self.monthApplied = str(month).zfill(2)
        self.csvdst = csvdst or f'C:\dev\Production\data\{year}\{self.monthApplied}'
        self.csvname = f"P-1390-LIF-{year}-{self.monthApplied}-1.txt"
        self.carrierId = 1390
        self.regex = re.compile('MFinancial_D')  
        self.datadir = f'J:\Acctng\Production\{year}\Data\Protective\{self.monthApplied}'
        self.goanywhere =  f'J:\Acctng\Revenue\goanywhere\Protective Life\{year}'
        self.flash = os.path.join(f'C:\dev\Production',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        # self.flash = os.path.join(f'J:\Acctng\Production\{year}\Summary - Flash',f'Production Summary {year}-{self.monthApplied}.xlsm')        
        self.ytdfolder = f'J:\Acctng\Production\{self.year}\Data\Protective\YTD'
        self.exportCols = ["SourceFileName","CarrierID","MemFirmID","CarrierContracteeID","CarrierContracteeName",
                            "ProductID","CarrierProductID","YearApplied","MonthApplied","ProducerID",
                            "CarrierProducerID","CarrierProducerName","PolicyNumber","IssueDate","InsuredName",
                            "YTDAnnualizedPrem" ,"YTDAnnualizedLowNon","YTDFace","RiskNumber","RiskName",
                            "Exchange1035","SplitPercentage","Replacement","CarrierProductName","PolicyOwner",
                            "Exchange","NLG","Modco","YRT","CSO2001"]

    # This function automate the process of copying the raw data file for a single month to the working directory
    def fetchRawMonth(self):
        '''Only fetch the month's data from Goanywhere folder. Save them in the corresponding month folder in the data directory'''
        LOGGER.info(f'Searching {month_name[self.month]} file for carrierid {self.carrierId}...')
        path = []
        f = []
        for dirs, _, filenames in os.walk(os.path.join(self.goanywhere,self.monthApplied)):
            f = [os.path.join(dirs,f) for f in filenames if re.search(self.regex,os.path.join(dirs,f))]
            path.extend(f)
        if len(path) != 0:
            for p in path:        
                shutil.copy2(p,self.datadir)
            LOGGER.info(f'1390 Protective: {len(path)} {month_name[self.month]} raw data files are copied to {self.datadir}.')
        else:
            LOGGER.warning(f'1390 Protective {month_name[self.month]} production reports not available.')


    # This function automate the process of moving the raw data files year to date to the working directory
    def fetchRawYTD(self):
        '''Copy all text files received from GoAnywhere as of the month to the YTD data directory'''
        LOGGER.info(f'Searching {month_name[self.month]} YTD files for carrierid {self.carrierId}...')
        if os.path.exists(self.ytdfolder):
            shutil.rmtree(self.ytdfolder)
        os.mkdir(self.ytdfolder)

        ytdmonth = []
        path = []
        files = []

        for i in range(1,self.month+1):
            if i <= self.month:
                ytdmonth.append(str(i).zfill(2))
        for m in ytdmonth:
            for dirs, _, filenames in os.walk(os.path.join(self.goanywhere,m)):
                files = [os.path.join(dirs,f) for f in filenames if re.search(self.regex,os.path.join(dirs,f))]
                path.extend(files)
        for p in path:
            shutil.copy2(p,self.ytdfolder)
        LOGGER.info(f'1390 Protective: {len(path)} raw data files are copied to {self.ytdfolder}.')
        
    # Since the text files only includes one single month's data, to get a YTD collection, this function combines all entries from each text file year-to-date.
    def getcsvYTD(self):
        '''Aggreagte all files in YTD data folder oin the data directory into an excel file. Read in, process and generate the csv file with Pandas'''
        
        # f is a list of raw Protective data file that should be process
        f = []
        for dir, _, filenames in os.walk(self.ytdfolder):
            data_paths = [os.path.join(dir, f) for f in filenames if re.search(self.regex, f)]  
        f.extend(data_paths)

        logging.debug(f'Combining {len(f)} text files in YTD folder...')
        for file in f:
            with open(file,'r') as input:
                LOGGER.info(f'Now processing {file}...')
                with open(os.path.join(self.ytdfolder,'temp.txt'),'a') as output:
                    content = input.readlines()
                    for i in range(0,len(content)):
                        newrow = content[i]
                        if newrow.startswith('00'):
                            output.write(newrow.replace('\x00','').replace('\t','|'))
        logging.debug('YTD temp.txt generated')

        df = pd.read_csv(os.path.join(self.ytdfolder,'temp.txt'),delimiter='|'
                ,names=["BP#","FA #","PRODUCER'S NAME","PRODUCER'S NPN","TRANS DATE","LIST BILL","NAME OF BANK/COMPANY","POLICY #",
                        "INSURED'S NAME","PLAN","COVERAGE ID","TRANS TYPE","PREMIUM/ASSET VALUE","BILLING FREQUENCY","RATE","SHARE",
                        "EARNINGS","POLICY YEAR","PRODUCT ID","PRODUCT DESC","TARGET/EXCESS"]
                ,dtype={'BP#':str,"PRODUCER'S NPN":str,"TRANS DATE":str,'PLAN':str,'PRODUCT ID':str})
        df["PREMIUM/ASSET VALUE"] = df["PREMIUM/ASSET VALUE"].str.replace(',','').astype(float)
        df.to_excel(os.path.join(self.ytdfolder,f'Protective{month_name[self.month]}full.xlsx'),index=False)
        logging.info(f'ProtectiveYTD.xlsx generated.')
        
        if os.path.exists(os.path.join(self.ytdfolder,'temp.txt'))==True:
            os.remove(os.path.join(self.ytdfolder,'temp.txt')) 

        # ETL Logic ##
        df["SourceFileName"] = f"P-1390-LIF-{self.year}-{self.monthApplied}-1"
        df["CarrierID"] = self.carrierId
        df["MemFirmID"] = 0
        df["CarrierContracteeID"] = df["PRODUCER'S NPN"].str.rstrip()
        df["CarrierContracteeName"] = df["PRODUCER'S NAME"].str.rstrip().str.replace(':',',')
        df["ProductID"] = 0
        df["CarrierProductID"] = df["PRODUCT DESC"].str.rstrip()
        df["YearApplied"] = self.year
        df["MonthApplied"] = self.monthApplied
        df["ProducerID"] = 0
        df["CarrierProducerID"] = df["PRODUCER'S NPN"].str.rstrip()
        df["CarrierProducerName"] = df["PRODUCER'S NAME"].str.rstrip()
        df["PolicyNumber"] = df["POLICY #"].str.strip()
        df["IssueDate"] = pd.to_datetime(df["TRANS DATE"], format='%m/%d/%Y').dt.normalize()
        df["InsuredName"] = df["INSURED'S NAME"].str.rstrip()
        df["YTDAnnualizedPrem"] = df["PREMIUM/ASSET VALUE"]*df["EARNINGS"]/(df["EARNINGS"].abs()) # This is to get the sign of cash flow. Sometime there are redo and undo activities which will not be presented in the PREMIUM/ASSET VALUE column
        df["YTDAnnualizedLowNon"] = 0
        df["YTDFace"] = ""
        df["RiskNumber"] = ""
        df["RiskName"] = ""
        df["Exchange1035"] = ""
        df["SplitPercentage"] = df["SHARE"].str.replace('%','').astype(float)*0.01
        df["Replacement"] = ""
        df["CarrierProductName"] = df["COVERAGE ID"].str.rstrip()
        df["PolicyOwner"] = df["NAME OF BANK/COMPANY"].str.rstrip()
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
            notesbreakdown.range('Prot_target').value = target
            notesbreakdown.range('Prot_excess').value = excess
            notesbreakdown.range('Prot_target').api.Font.Color = rgb_to_int((0, 102, 204))
            notesbreakdown.range('Prot_excess').api.Font.Color = rgb_to_int((0, 102, 204))
            flash.save(self.flash)
            flash.close()
        LOGGER.info(f'{self.flash} is updated and saved')

        # Export
        df["YTDAnnualizedPrem"] = df["YTDAnnualizedPrem"].astype(float).round(2).map("{:,.2f}".format)
        df["YTDAnnualizedLowNon"] = df["YTDAnnualizedLowNon"].astype(float).round(2).map("{:,.2f}".format)
        df = df[self.exportCols]
        df.to_csv(os.path.join(self.csvdst,self.csvname), index=False, sep='|')
        # df.to_excel(os.path.join(self.csvdst,'test.xlsx'), index=False)
        LOGGER.info(f'{self.csvname} is saved at {self.csvdst}.')


# Combine the whole class into a function so it can be more easily used in the API
def processProt(year, month):
    pro = protective(year,month)
    pro.fetchRawMonth()
    pro.fetchRawYTD()
    pro.getcsvYTD()


# if __name__ == '__main__':
#     pro = protective(2022,11)
#     pro.getcsvYTD()


