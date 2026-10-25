import os
import win32com.client as win32
import calendar as c
import pyodbc
import pandas as pd
import openpyxl as xl

month = 9
year = 2022

####################################################################################################################################
## Tami Wroblewski

requestFolder = r'J:\Acctng\Production\{}\Requests'.format(year)
connection_string = (r'DRIVER=' + 'SQL Server' + r';'
                    r'SERVER=' + 'pdxvmdb11' + r';'
                    r'DATABASE=' + 'ProductionAndOverride' + r';'
                    r'Trusted_Connection=yes')
conn = pyodbc.connect(connection_string)

PL_MarketShare_qry = '''
declare @reportDate date set @reportDate = '''+"'"+str(year)+"-"+str(month).zfill(2)+"-"+str(c.monthrange(year, month)[1])+"'"+'''

;with a as (
	select
		pd.YearApplied,
		c.[Name] MemberFirm,
		case when c.name in ('M Benefit Solutions', 'Southern Wealth Management, LLP','MullinTBG','Groff Team Partners, LLC','Executive Compensation Group, LLC, The',
							'Jones Lowry','Meridian Financial Group, LLC','Peck Financial','Peter M. Williams & Company','Pfleger Financial Group, Inc.',
							'Premier Partners, LLC','Retirement & Insurance Resources, LLC','Thomas Financial Group','United Financial Consultants','Groff Team Partners, LLC',
							'Vie International Financial Services, Ltd.','Enza Financial, LLC','Garlikov & Associates, Inc.','GDK & Company','HEIRMARK, LTD',
							'Koptis Organization, LLC, The','Ownership Advisors, Inc.','Ressourcement, Inc.','HEIRMARK, LTD') then 'Central' 
			when c.name in ('Knight Planning Corporation') then 'East'  
			when c.name in ('Miscellaneous Firms','Mezrah Financial Group') then 'PNW' else  st.Territory end Territory,
		ct.[Type],
		case when c.ReportToCarrier = '1' then 0 else sum(pd.YTDAnnualizedPrem) end AllTarget,
		case when c.ReportToCarrier = '1' then 0 else sum(pd.YTDAnnualizedLowNon) end AllExcess,
		case when pd.CarrierID = '1111' then sum(pd.YTDAnnualizedPrem) else 0 end PLTarget,
		case when pd.CarrierID = '1111' then sum(pd.YTDAnnualizedLowNon) else 0 end PLExcess
	from PremiumDetail pd join Company c on pd.MemFirmID = c.COID
		join product p on pd.ProductID = p.ID
		join ProdType pt on p.ProdTypeID = pt.ID
		join LOB l on pt.LOBID = l.ID
		join CompanyType ct on c.CompanyTypeID = ct.ID
		join States st on c.State = st.StateID
	where 
	pd.MonthApplied = month(@reportDate)
	and (c.CompanyTypeID in ('1', '2') or (c.CompanyTypeID = '3' and year(@reportDate)=YearApplied)) 
	and pd.CarrierID < '1800' 
	and l.ID in ('2') 
	and pt.id not in ('10', '11', '16')
	and c.Name not in ('Cambridge Consulting Group, LLC', 'GFG Strategic Advisors, LLC', 'Paul L. MacCaskill')
	group by pd.YearApplied, st.Territory,c.[Name], c.COID, ct.[Type], c.ReportToCarrier, c.CompanyTypeID, pd.CarrierID, c.ReportToCarrier
) select YearApplied, MemberFirm, [Type],Territory, sum(AllTarget) AllTarget, sum(AllExcess) AllExcess, sum(PLTarget) PLTarget, sum(PLExcess) PLExcess
from a
group by YearApplied, MemberFirm, Territory,[Type]
order by YearApplied desc, MemberFirm, [Type]

'''
data = pd.read_sql(PL_MarketShare_qry,conn)
data.to_excel(os.path.join(requestFolder,'Tami Wroblewski','Market Share',f'PL Market Share YTD {c.month_name[month]} {year}.xlsx'))

PL_Production_by_Type_qry = '''
declare @reportDate date
set @reportDate = '''+"'"+str(year)+"-"+str(month).zfill(2)+"-"+str(c.monthrange(year, month)[1])+"'"+'''
select
	pd.YearApplied,
	pd.MonthApplied,
	c.[Name] Carrier,
	c2.[Name] MemberFirm,
	case when c2.name in ('M Benefit Solutions', 'Southern Wealth Management, LLP','MullinTBG','Groff Team Partners, LLC','Executive Compensation Group, LLC, The',
					      'Jones Lowry','Meridian Financial Group, LLC','Peck Financial','Peter M. Williams & Company','Pfleger Financial Group, Inc.',
						  'Premier Partners, LLC','Retirement & Insurance Resources, LLC','Thomas Financial Group','United Financial Consultants','Groff Team Partners, LLC',
						  'Vie International Financial Services, Ltd.','Enza Financial, LLC','Garlikov & Associates, Inc.','GDK & Company','HEIRMARK, LTD',
						  'Koptis Organization, LLC, The','Ownership Advisors, Inc.','Ressourcement, Inc.','HEIRMARK, LTD') then 'Central'  
		 when c2.name in ('Knight Planning Corporation') then 'East'  
		 when c2.name in ('Mezrah Financial Group','Miscellaneous Firms') then 'PNW' 
		 else  st.Territory end Territory,
	case when pt.[Type] like '%annuity%' then 'Annuity' else pt.[Type] end [Type],
	sum(pd.YTDAnnualizedPrem) TargetPremium,
	sum(pd.YTDAnnualizedLowNon) ExcessPremium
from premiumdetail pd
	join Company c on c.COID = pd.CarrierID
	join Company c2 on c2.COID = pd.MemFirmID
	join States st on st.StateID =c2.State
	join Product p on p.ID = pd.ProductID
	join ProdType pt on pt.ID = p.ProdTypeID
	join LOB l on l.ID = pt.LOBID
where pd.YearApplied = year(@reportDate) and pd.MonthApplied = month(@reportDate) 
	  and pd.CarrierID = '1111'
	  and c2.Name not in ('Cambridge Consulting Group, LLC', 'GFG Strategic Advisors, LLC', 'Paul L. MacCaskill')
group by pd.YearApplied, pd.MonthApplied, c.[Name], st.Territory,c2.[Name], case when pt.[Type] like '%annuity%' then 'Annuity' else pt.[Type] end
order by MemberFirm

--Per Tami Wroblewski's email on 5/18/2022, move Mezrah Financial Group from Central to PNW -YSL
'''
PL_Bytype = pd.read_sql(PL_Production_by_Type_qry,conn)
PL_Bytype.to_excel(os.path.join(requestFolder,'Tami Wroblewski','Production by Type',f'PL Production thru {c.month_name[month]} {year}.xlsx'),header=True, index=False)







# construct Outlook application instance
olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

# construct the email item object
mailItem = olApp.CreateItem(0)
mailItem.Display()
mailItem.Subject = 'M Production Breakdowns - Market Share, Product Type'
mailItem.BodyFormat = 1
# mailItem.To = '<TWroblewski@pacificlife.com>'
mailItem.To = '<bboy80345@gmail.com>'
mailItem.CC = '<Blake.Myer@mfin.com>; <Hans.Avery@mfin.com>'
mailItem.Attachments.Add(os.path.join(requestFolder,'Tami Wroblewski','Market Share',f'PL Market Share YTD {c.month_name[month]} {year}.xlsx'))
mailItem.Attachments.Add(os.path.join(requestFolder,'Tami Wroblewski','Production by Type',f'PL Production thru {c.month_name[month]} {year}.xlsx'))
# mailItem.Attachments.Add(os.path.join(requestFolder,'Tami Wroblewski','LTC Production',f'LTC Production - {c.month_name[month]} {year}.xlsx'))

mailItem.Body = '''
Hi Tami, 

Please find the attached '''+ c.month_name[month] +''' reports.

Please let me know if you need anything else.

Thanks,
Yu-Sheng Lee
Business Data Analyst
Yu-Sheng.Lee@mfin.com
Direct: 503.414.7590 
1125 NW Couch Street #900, Portland, OR 97209

'''
mailItem.Save()
# mailItem.Send()


####################################################################################################################################
## 





def copyRange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    #Loops through selected Rows
    for i in range(startRow,endRow + 1,1):
        #Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol,endCol+1,1):
            rowSelected.append(sheet.cell(row = i, column = j).value)
        #Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)
 
    return rangeSelected
        
 
#Paste range
#Paste data from copyRange into template sheet
def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving,copiedData):
    countRow = 0
    for i in range(startRow,endRow+1,1):
        countCol = 0
        for j in range(startCol,endCol+1,1):
            sheetReceiving.cell(row = i, column = j).value = copiedData[countRow][countCol]
            countCol += 1
        countRow += 1